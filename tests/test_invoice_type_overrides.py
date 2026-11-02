import sqlite3
import tempfile
import unittest
from contextlib import closing
from dataclasses import asdict, replace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import autoinvoice as app_module


class InvoiceTypeOverrideTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-types-")
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        data = root / "data"
        runtime = patch.multiple(
            app_module,
            APP_DIR=root,
            INVOICE_DIR=root / "invoices",
            USED_DIR=root / "used_invoices",
            VOID_DIR=root / "void_invoices",
            FAILED_DIR=root / "failed_invoices",
            PRINT_JOBS_DIR=root / "print_jobs",
            OFD_CACHE_DIR=root / "print_jobs" / "ofd_cache",
            DATA_DIR=data,
            DATABASE_FILE=data / "invoice_records.sqlite3",
            STATE_FILE=data / "used_records.json",
            FAILED_RECORD_FILE=data / "failed_records.json",
            SETTINGS_FILE=data / "app_settings.json",
            LEGACY_RUNTIME_FILES=(),
            DATABASE_INITIALIZED=False,
        )
        runtime.start()
        self.addCleanup(runtime.stop)
        app_module.ensure_dirs()
        self.traffic = app_module.CATEGORIES[1]
        self.dining = app_module.CATEGORIES[2]
        self.office = app_module.CATEGORIES[4]

    def invoice(self, filename="invoice.pdf", number="12345678", content=b"test invoice"):
        path = app_module.INVOICE_DIR / filename
        path.write_bytes(content)
        return app_module.Invoice(str(path), filename, 10000, self.office, "2026-09-01", number)

    def scan(self, invoices):
        by_path = {invoice.path: invoice for invoice in invoices}
        with patch.object(app_module, "parse_invoice", side_effect=lambda path: by_path[str(path)]):
            scanned, failures, reused, duplicates = app_module.scan_invoice_directory()
        self.assertEqual((failures, reused, duplicates), ([], [], []))
        return scanned

    def make_app(self, invoices):
        with patch.object(app_module.AutoInvoiceApp, "scan"):
            app = app_module.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        renderer = patch.object(app, "render_preview")
        renderer.start()
        self.addCleanup(renderer.stop)
        app.invoices = list(invoices)
        app.populate_tree(app.invoice_tree, app.invoices)
        app.update_available_summary()
        app.update()
        return app

    def test_override_survives_restart_rename_and_format_change(self):
        invoice = self.invoice()
        other = self.invoice("other.pdf", "87654321")
        app_module.save_invoice_type_override(invoice, self.traffic)
        app_module.DATABASE_INITIALIZED = False
        scanned = self.scan([invoice, other])
        self.assertEqual({item.invoice_number: item.invoice_type for item in scanned}, {
            invoice.invoice_number: self.traffic, other.invoice_number: self.office,
        })
        renamed = Path(invoice.path).with_name("renamed.ofd")
        Path(invoice.path).rename(renamed)
        invoice = replace(invoice, path=str(renamed), filename=renamed.name)
        scanned = self.scan([invoice, other])
        self.assertEqual(next(item for item in scanned if item.path == str(renamed)).invoice_type, self.traffic)

    def test_numberless_override_follows_content_not_filename(self):
        invoice = self.invoice(number="")
        app_module.save_invoice_type_override(invoice, self.traffic)
        renamed = Path(invoice.path).with_name("renamed.pdf")
        Path(invoice.path).rename(renamed)
        invoice = replace(invoice, path=str(renamed), filename=renamed.name)
        self.assertEqual(self.scan([invoice])[0].invoice_type, self.traffic)
        renamed.write_bytes(b"a different invoice")
        self.assertEqual(self.scan([invoice])[0].invoice_type, self.office)

    def test_replacing_and_removing_override(self):
        invoice = self.invoice()
        app_module.save_invoice_type_override(invoice, self.traffic)
        app_module.save_invoice_type_override(invoice, self.dining)
        self.assertEqual(self.scan([invoice])[0].invoice_type, self.dining)
        self.assertEqual(len(app_module.load_invoice_type_overrides()), 1)
        with self.assertRaises(ValueError):
            app_module.save_invoice_type_override(invoice, app_module.CATEGORY_ALL)
        self.assertEqual(self.scan([invoice])[0].invoice_type, self.dining)
        app_module.save_invoice_type_override(invoice, None)
        self.assertEqual(self.scan([invoice])[0].invoice_type, self.office)

    def test_existing_used_records_survive_schema_upgrade(self):
        invoice = self.invoice()
        app_module.save_used_invoice_record(asdict(invoice))
        with closing(sqlite3.connect(app_module.DATABASE_FILE)) as connection, connection:
            connection.execute("DROP TABLE invoice_type_overrides")
        app_module.DATABASE_INITIALIZED = False
        app_module.ensure_database()
        self.assertEqual(app_module.load_used_invoice_numbers(), {invoice.invoice_number})
        self.assertEqual(app_module.load_invoice_type_overrides(), {})

    def test_database_read_error_does_not_move_invoice(self):
        invoice = self.invoice()
        with patch.object(app_module, "load_invoice_type_overrides", side_effect=sqlite3.OperationalError("locked")):
            with self.assertRaises(sqlite3.OperationalError):
                app_module.scan_invoice_directory()
        self.assertTrue(Path(invoice.path).exists())
        self.assertEqual(list(app_module.FAILED_DIR.iterdir()), [])

    def test_menu_edit_updates_only_selected_invoice_and_preserves_combination(self):
        first = self.invoice()
        second = self.invoice("second.pdf", "87654321")
        app = self.make_app([first, second])
        app.result_invoices = [first, second]
        app.refresh_result_tree()
        app.invoice_tree.selection_set("0")
        app.invoice_tree.focus("0")
        app.result_tree.selection_set("0", "1")
        app.update()
        app.print_preview_pdf = app_module.PRINT_JOBS_DIR / "merged.pdf"
        app.print_preview_invoices = [first, second]
        app.preview_invoice = replace(first, path=str(app.print_preview_pdf), invoice_type="combination")
        app.prepare_invoice_type_menu()
        app.invoice_type_menu.invoke(0)
        self.assertEqual([item.invoice_type for item in app.invoices], [self.traffic, self.office])
        self.assertEqual([item.invoice_type for item in app.result_invoices], [self.traffic, self.office])
        self.assertEqual(app.invoice_tree.set("0", "type"), self.traffic)
        self.assertEqual(app.result_tree.set("0", "type"), self.traffic)
        self.assertTrue(app.invoice_tree.set("0", "status"))
        self.assertEqual(app.invoice_tree.selection(), ("0",))
        self.assertEqual(app.result_tree.selection(), ("0", "1"))
        self.assertIsNone(app.print_preview_pdf)
        self.assertEqual(app.print_preview_invoices, [])
        self.assertEqual(app.preview_invoice, app.invoices[0])
        app.category_vars[app_module.CATEGORY_ALL].set(False)
        app.category_vars[self.traffic].set(True)
        self.assertEqual(app.filtered_invoices(), [app.invoices[0]])
        self.assertEqual(self.scan([first, second])[0].invoice_type, self.traffic)
        with patch.object(app_module, "parse_invoice", return_value=first):
            app.invoice_type_menu.invoke(app.invoice_type_menu.index(app_module.tk.END))
        self.assertEqual(app.invoices[0].invoice_type, self.office)
        self.assertEqual(app_module.load_invoice_type_overrides(), {})

    def test_save_failure_leaves_ui_and_combination_unchanged(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        app.result_invoices = [invoice]
        app.refresh_result_tree()
        app.invoice_tree.selection_set("0")
        app.invoice_tree.focus("0")
        with patch.object(app_module, "save_invoice_type_override", side_effect=sqlite3.OperationalError("locked")), patch.object(app_module.messagebox, "showerror") as error:
            app.set_selected_invoice_type(self.traffic)
        error.assert_called_once()
        self.assertEqual(app.invoices, [invoice])
        self.assertEqual(app.result_invoices, [invoice])
        self.assertEqual(app.invoice_tree.set("0", "type"), self.office)
        self.assertEqual(app.invoice_type_var.get(), self.office)

    def test_context_menu_targets_clicked_row_and_blocks_edits_during_calculation(self):
        first = self.invoice()
        second = self.invoice("second.pdf", "87654321")
        app = self.make_app([first, second])
        app.invoice_tree.selection_set("0")
        event = SimpleNamespace(y=0, x_root=100, y_root=100)
        with patch.object(app.invoice_tree, "identify_row", return_value="1"), patch.object(app.invoice_context_menu, "tk_popup"), patch.object(app.invoice_context_menu, "grab_release"):
            self.assertEqual(app.show_invoice_type_menu(event), "break")
        self.assertEqual(app.invoice_tree.selection(), ("1",))
        app.set_selected_invoice_type(self.traffic)
        self.assertEqual(app.invoices[0], first)
        self.assertEqual(app.invoices[1].invoice_type, self.traffic)
        app.calculation_running = True
        app.prepare_invoice_type_menu()
        self.assertEqual(app.invoice_type_menu.entrycget(0, "state"), "disabled")
        app.set_selected_invoice_type(self.dining)
        self.assertEqual(app.invoices[1].invoice_type, self.traffic)
        app.calculation_running = False
        app.invoice_tree.selection_remove("1")
        app.prepare_invoice_type_menu()
        self.assertEqual(app.invoice_type_menu.entrycget(0, "state"), "disabled")

    def test_menu_and_summary_fit_default_and_minimum_window(self):
        app = self.make_app([self.invoice()])
        header = app.invoice_tree.master.grid_slaves(row=0, column=0)[0]
        for geometry in ("1180x760", "1020x640"):
            app.geometry(geometry)
            app.update()
            previous_right = 0
            for widget in sorted(header.winfo_children(), key=lambda item: item.winfo_x()):
                self.assertGreaterEqual(widget.winfo_x(), previous_right)
                previous_right = widget.winfo_x() + widget.winfo_width()
                self.assertLessEqual(previous_right, header.winfo_width())
            actions = app.result_tree.master.grid_slaves(row=2, column=0)[0]
            self.assertLessEqual(actions.winfo_y() + actions.winfo_height(), actions.master.winfo_height())


if __name__ == "__main__":
    unittest.main()
