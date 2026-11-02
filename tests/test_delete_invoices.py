import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import autoinvoice as m


class DeleteInvoiceTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-delete-")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        data = self.root / "data"
        runtime = patch.multiple(
            m, APP_DIR=self.root, INVOICE_DIR=self.root / "invoices",
            USED_DIR=self.root / "used_invoices", VOID_DIR=self.root / "void_invoices",
            FAILED_DIR=self.root / "failed_invoices", PRINT_JOBS_DIR=self.root / "print_jobs",
            OFD_CACHE_DIR=self.root / "print_jobs" / "ofd_cache", DATA_DIR=data,
            DATABASE_FILE=data / "invoice_records.sqlite3", STATE_FILE=data / "used_records.json",
            FAILED_RECORD_FILE=data / "failed_records.json", SETTINGS_FILE=data / "app_settings.json",
            LEGACY_RUNTIME_FILES=(), DATABASE_INITIALIZED=False,
        )
        runtime.start()
        self.addCleanup(runtime.stop)
        m.ensure_dirs()

    def invoice(self, filename="first.pdf", number="12345678", amount=10000):
        path = m.INVOICE_DIR / filename
        path.write_bytes(b"temporary test invoice")
        return m.Invoice(str(path), filename, amount, "办公", "2026-11-01", number)

    def make_app(self, invoices):
        with patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        renderer = patch.object(app, "render_preview")
        renderer.start()
        self.addCleanup(renderer.stop)
        app.invoices = list(invoices)
        app.result_invoices = list(invoices)
        app.populate_tree(app.invoice_tree, app.invoices)
        app.populate_tree(app.result_tree, app.result_invoices)
        app.update_available_summary()
        app.target_var.set("150")
        app.update_result_summary()
        app.update()
        self.assertTrue(hasattr(app, "invoice_context_menu"), "right-click menu with Delete is missing")
        return app

    def labels(self, menu):
        return [menu.entrycget(index, "label") for index in range(menu.index(m.tk.END) + 1) if menu.type(index) != "separator"]

    def select(self, app, index="0"):
        app.invoice_tree.selection_set(index)
        app.invoice_tree.focus(index)
        app.prepare_invoice_type_menu()

    def test_delete_is_only_in_right_click_menu_and_type_actions_remain_available(self):
        app = self.make_app([self.invoice()])
        self.assertIn("删除", self.labels(app.invoice_context_menu))
        self.assertNotIn("删除", self.labels(app.invoice_type_menu))
        for category in m.CATEGORIES[1:]:
            self.assertIn(category, self.labels(app.invoice_context_menu))
        self.assertIn("恢复自动识别", self.labels(app.invoice_context_menu))

    def test_confirmed_delete_targets_right_clicked_row_and_updates_both_lists(self):
        first = self.invoice()
        second = self.invoice("second.ofd", "87654321", 20000)
        third = self.invoice("third.png", "11223344", 30000)
        app = self.make_app([first, second, third])
        app.result_invoices = [first, second]
        app.populate_tree(app.result_tree, app.result_invoices)
        self.select(app, "0")
        event = SimpleNamespace(y=10, x_root=100, y_root=100)
        with patch.object(app.invoice_tree, "identify_row", return_value="1"), patch.object(app.invoice_context_menu, "tk_popup"), patch.object(app.invoice_context_menu, "grab_release"):
            app.show_invoice_type_menu(event)
        app.prepare_invoice_type_menu()
        with patch.object(m.messagebox, "askyesno", return_value=True) as confirmation:
            app.invoice_context_menu.invoke("删除")
        self.assertFalse(Path(second.path).exists())
        self.assertTrue(Path(first.path).is_file())
        self.assertTrue(Path(third.path).is_file())
        self.assertEqual(app.invoices, [first, third])
        self.assertEqual(app.result_invoices, [first])
        self.assertEqual(len(app.invoice_tree.get_children()), 2)
        self.assertEqual(len(app.result_tree.get_children()), 1)
        self.assertEqual(app.available_summary_var.get(), "共 2 张，合计 400.00 元")
        self.assertIn("100.00", app.summary_var.get())
        self.assertEqual(m.load_used_invoice_numbers(), set())
        self.assertEqual(list(m.USED_DIR.iterdir()), [])
        self.assertIn(second.filename, confirmation.call_args.args[1])
        self.assertIn("永久", confirmation.call_args.args[1])
        self.assertIn("回收站", confirmation.call_args.args[1])
        self.assertEqual(confirmation.call_args.kwargs["default"], m.messagebox.NO)

    def test_cancelled_confirmation_keeps_file_lists_and_preview(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        app.preview_invoice = invoice
        self.select(app)
        with patch.object(m.messagebox, "askyesno", return_value=False):
            app.invoice_context_menu.invoke("删除")
        self.assertTrue(Path(invoice.path).is_file())
        self.assertEqual(app.invoices, [invoice])
        self.assertEqual(app.result_invoices, [invoice])
        self.assertEqual(app.preview_invoice, invoice)
        self.assertEqual(app.invoice_tree.selection(), ("0",))

    def test_failed_file_delete_reports_error_without_removing_rows(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        self.select(app)
        with patch.object(m.messagebox, "askyesno", return_value=True), patch.object(Path, "unlink", side_effect=PermissionError("file is locked")), patch.object(m.messagebox, "showerror") as error:
            app.invoice_context_menu.invoke("删除")
        self.assertTrue(Path(invoice.path).is_file())
        self.assertEqual(app.invoices, [invoice])
        self.assertEqual(app.result_invoices, [invoice])
        self.assertEqual(len(app.invoice_tree.get_children()), 1)
        self.assertEqual(len(app.result_tree.get_children()), 1)
        self.assertIn("file is locked", error.call_args.args[1])

    def test_delete_is_disabled_without_selection_or_while_busy(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        for state in ("no_selection", "calculation_running", "scanning_running"):
            with self.subTest(state=state):
                self.select(app)
                if state == "no_selection":
                    app.invoice_tree.selection_remove("0")
                else:
                    setattr(app, state, True)
                app.prepare_invoice_type_menu()
                self.assertEqual(app.invoice_context_menu.entrycget("删除", "state"), "disabled")
                with patch.object(m.messagebox, "askyesno") as confirmation:
                    app.delete_selected_invoice()
                confirmation.assert_not_called()
                self.assertTrue(Path(invoice.path).is_file())
                if state != "no_selection":
                    setattr(app, state, False)

    def test_already_missing_file_is_removed_from_stale_lists_after_confirmation(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        self.select(app)
        Path(invoice.path).unlink()
        with patch.object(m.messagebox, "askyesno", return_value=True):
            app.invoice_context_menu.invoke("删除")
        self.assertEqual(app.invoices, [])
        self.assertEqual(app.result_invoices, [])
        self.assertEqual(app.available_summary_var.get(), "共 0 张，合计 0.00 元")

    def test_delete_invalidates_affected_combination_preview_but_keeps_other_invoice(self):
        first = self.invoice()
        second = self.invoice("second.pdf", "87654321", 20000)
        app = self.make_app([first, second])
        merged = m.PRINT_JOBS_DIR / "merged.pdf"
        merged.write_bytes(b"previous preview backup")
        app.print_preview_pdf = merged
        app.print_preview_invoices = [first, second]
        app.preview_invoice = m.replace(first, path=str(merged), invoice_type="组合")
        self.select(app, "1")
        with patch.object(m.messagebox, "askyesno", return_value=True):
            app.invoice_context_menu.invoke("删除")
        self.assertIsNone(app.print_preview_pdf)
        self.assertEqual(app.print_preview_invoices, [])
        self.assertEqual(app.preview_invoice, first)
        self.assertTrue(Path(first.path).is_file())
        self.assertEqual(merged.read_bytes(), b"previous preview backup")


if __name__ == "__main__":
    unittest.main()
