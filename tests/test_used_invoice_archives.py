import json
import sqlite3
import tempfile
import unittest
import zipfile
from contextlib import closing
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import autoinvoice as m
from test_combination_excel import read_sheet


class UsedInvoiceArchiveTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-archives-")
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

    def make_app(self, invoices=()):
        with patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        app.invoices = list(invoices)
        app.result_invoices = list(invoices)
        app.populate_tree(app.invoice_tree, app.invoices)
        app.populate_tree(app.result_tree, app.result_invoices)
        app.update()
        return app

    def invoice(self, filename="first.pdf", number="12345678", content=b"invoice"):
        path = m.INVOICE_DIR / filename
        if path.suffix == ".pdf":
            with m.fitz.open() as document:
                page = document.new_page()
                page.insert_text((30, 30), content.decode("ascii"))
                document.save(path)
        elif path.suffix == ".png":
            with m.Image.new("RGB", (300, 200), "white") as image:
                image.save(path)
        else:
            path.write_bytes(content)
        return m.Invoice(str(path), filename, 10000, "交通", "2026-10-01", number)

    def archive(self, action):
        with patch.object(m, "datetime", wraps=datetime) as clock, patch.object(m.messagebox, "showinfo"), patch.object(m.messagebox, "showwarning"):
            clock.now.return_value = datetime(2026, 11, 2, 10, 30)
            action()

    def pdf_invoice(self, filename="first.pdf", number="12345678", amount_cents=10000, pages=1):
        path = m.INVOICE_DIR / filename
        with m.fitz.open() as document:
            for index in range(pages):
                page = document.new_page()
                page.insert_text((30, 30), f"Invoice {number}, page {index + 1}")
            document.save(path)
        return m.Invoice(str(path), filename, amount_cents, "交通", "2026-10-01", number, pages=pages)

    def generate_preview(self, app):
        with patch.object(m.messagebox, "askyesno", return_value=False), patch.object(app, "render_preview"):
            app.generate_combination_preview()
        self.assertIsNotNone(app.print_preview_pdf)
        return app.print_preview_pdf

    def test_amount_and_content_inputs_fit_both_window_sizes(self):
        app = self.make_app()
        actions = app.calculate_button.master
        labels = [str(widget.cget("text")) for widget in actions.winfo_children() if isinstance(widget, m.ttk.Label)]
        self.assertIn("金额", labels)
        self.assertNotIn("报销金额", labels)
        self.assertIn("内容：", labels)
        entries = [widget for widget in actions.winfo_children() if isinstance(widget, m.ttk.Entry)]
        self.assertEqual(len(entries), 2)
        amount, content = entries
        self.assertLessEqual(int(amount.cget("width")), 10)
        self.assertEqual(str(amount.cget("textvariable")), str(app.target_var))
        self.assertEqual(content.get(), "报销")
        content.delete(0, m.tk.END)
        content.insert(0, "差旅")
        self.assertEqual(app.content_var.get(), "差旅")
        for size in ("1180x760", "1020x640"):
            app.geometry(size)
            app.update()
            right = 0
            for widget in sorted(actions.winfo_children(), key=lambda item: item.winfo_x()):
                self.assertGreaterEqual(widget.winfo_x(), right)
                self.assertGreaterEqual(widget.winfo_width(), widget.winfo_reqwidth())
                right = widget.winfo_x() + widget.winfo_width()
                self.assertLessEqual(right, actions.winfo_width())

    def test_selected_invoice_is_archived_and_unselected_combination_remains(self):
        first = self.invoice()
        second = self.invoice("second.ofd", "87654321")
        app = self.make_app([first, second])
        app.target_var.set("200")
        app.result_tree.selection_set("0")
        self.archive(app.mark_selected_used)
        destination = m.USED_DIR / "报销20261102" / "first.pdf"
        self.assertTrue(destination.is_file())
        self.assertFalse(Path(first.path).exists())
        self.assertTrue(Path(second.path).exists())
        self.assertEqual(app.result_invoices, [second])
        self.assertEqual(app.invoices, [second])
        self.assertEqual(len(app.result_tree.get_children()), 1)
        self.assertEqual(m.load_used_invoice_numbers(), {"12345678"})
        with closing(sqlite3.connect(m.DATABASE_FILE)) as connection:
            row = connection.execute("SELECT moved_to, target_amount FROM used_invoices").fetchone()
        self.assertEqual(row, (str(destination), "200"))
        records = json.loads(m.STATE_FILE.read_text(encoding="utf-8"))
        self.assertEqual(records[0]["moved_to"], str(destination))

    def test_custom_archive_reuses_same_day_directory_without_overwriting_files(self):
        first = self.invoice("same.pdf", content=b"new invoice")
        original_bytes = Path(first.path).read_bytes()
        second = self.invoice("second.png", "87654321")
        directory = m.USED_DIR / "差旅20261102"
        directory.mkdir()
        original = directory / "same.pdf"
        original.write_bytes(b"existing invoice")
        app = self.make_app([first, second])
        app.content_var.set("差旅")
        self.archive(app.mark_all_used)
        self.assertEqual(original.read_bytes(), b"existing invoice")
        self.assertTrue((directory / "same_1.pdf").is_file())
        self.assertEqual((directory / "same_1.pdf").read_bytes(), original_bytes)
        self.assertTrue((directory / "second.png").is_file())
        self.assertEqual(list(m.USED_DIR.iterdir()), [directory])
        self.assertEqual(app.result_invoices, [])
        self.assertEqual(m.load_used_invoice_numbers(), {"12345678", "87654321"})

    def test_large_archive_collision_sequence_still_returns_unused_filename(self):
        directory = m.USED_DIR / "报销20261102"
        directory.mkdir()
        (directory / "same.pdf").write_bytes(b"original")
        for index in range(1, 1000):
            (directory / f"same_{index}.pdf").touch()
        timestamp_file = directory / "same_20261102_103000.pdf"
        timestamp_file.write_bytes(b"keep this file")
        with patch.object(m, "datetime", wraps=datetime) as clock:
            clock.now.return_value = datetime(2026, 11, 2, 10, 30)
            destination = m.unique_destination(directory, "same.pdf")
        self.assertFalse(destination.exists())
        self.assertEqual(timestamp_file.read_bytes(), b"keep this file")

    def test_empty_content_uses_default_archive_name(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        app.content_var.set("   ")
        self.archive(app.mark_all_used)
        self.assertTrue((m.USED_DIR / "报销20261102" / "first.pdf").is_file())

    def test_content_is_sanitized_and_cannot_escape_used_directory(self):
        first = self.invoice()
        second = self.invoice("second.pdf", "87654321")
        app = self.make_app([first, second])
        app.content_var.set("差旅/甲:乙?")
        self.archive(lambda: app.mark_used([first]))
        self.assertTrue((m.USED_DIR / "差旅_甲_乙_20261102" / "first.pdf").is_file())
        app.content_var.set(r"..\..\outside")
        self.archive(lambda: app.mark_used([second]))
        self.assertTrue((m.USED_DIR / "_.._outside20261102" / "second.pdf").is_file())
        self.assertFalse((self.root / "outside20261102").exists())

    def test_windows_reserved_content_names_can_be_archived(self):
        app = self.make_app()
        for index, (content, folder) in enumerate((
            ("COM1.办公", "_COM1.办公20261102"),
            ("NUL.txt", "_NUL.txt20261102"),
            ("com².差旅", "_com².差旅20261102"),
        )):
            with self.subTest(content=content):
                invoice = self.invoice(f"reserved{index}.pdf", f"1234567{index}")
                app.content_var.set(content)
                self.archive(lambda: app.mark_used([invoice]))
                self.assertTrue((m.USED_DIR / folder / invoice.filename).is_file())

    def test_used_duplicate_in_dated_subdirectory_is_moved_to_void(self):
        invoice = self.invoice("ticket[1].pdf")
        original_bytes = Path(invoice.path).read_bytes()
        app = self.make_app([invoice])
        self.archive(app.mark_all_used)
        Path(invoice.path).write_bytes(b"duplicate invoice")
        with patch.object(m, "parse_invoice", return_value=invoice):
            available, failures, reused, duplicates = m.scan_invoice_directory()
        self.assertEqual((available, failures, duplicates), ([], [], []))
        self.assertEqual(len(reused), 1)
        self.assertEqual(reused[0]["moved_to_dir"], "void_invoices")
        self.assertTrue((m.VOID_DIR / "ticket[1].pdf").is_file())
        self.assertEqual((m.USED_DIR / "报销20261102" / "ticket[1].pdf").read_bytes(), original_bytes)
        self.assertFalse(Path(invoice.path).exists())

    def test_print_confirmation_uses_content_archive(self):
        invoice = self.invoice()
        app = self.make_app([invoice])
        app.content_var.set("办公")
        with patch.object(m.messagebox, "askyesno", return_value=True):
            self.archive(lambda: app.on_print_success([invoice]))
        self.assertTrue((m.USED_DIR / "办公20261102" / "first.pdf").is_file())
        self.assertEqual(app.result_invoices, [])

    def test_generated_pdf_and_excel_are_archived_with_original_invoices(self):
        first = self.pdf_invoice()
        second = self.pdf_invoice("second.pdf", "87654321", 20000)
        app = self.make_app([first, second])
        app.target_var.set("300")
        preview = self.generate_preview(app)
        original_pdf = preview.read_bytes()
        original_excel = preview.with_suffix(".xlsx").read_bytes()
        self.archive(app.mark_all_used)
        directory = m.USED_DIR / "报销20261102"
        self.assertTrue((directory / preview.name).is_file())
        self.assertTrue((directory / preview.with_suffix(".xlsx").name).is_file())
        self.assertEqual((directory / preview.name).read_bytes(), original_pdf)
        self.assertEqual((directory / preview.with_suffix(".xlsx").name).read_bytes(), original_excel)
        self.assertTrue((directory / "first.pdf").is_file())
        self.assertTrue((directory / "second.pdf").is_file())
        self.assertEqual(preview.read_bytes(), original_pdf)
        self.assertEqual(preview.with_suffix(".xlsx").read_bytes(), original_excel)

    def test_selected_subset_gets_its_own_pdf_and_excel(self):
        first = self.pdf_invoice()
        second = self.pdf_invoice("second.pdf", "87654321", 20000, pages=2)
        app = self.make_app([first, second])
        app.target_var.set("300")
        preview = self.generate_preview(app)
        app.result_tree.selection_set("0")
        self.archive(app.mark_selected_used)
        directory = m.USED_DIR / "报销20261102"
        workbooks = list(directory.glob("*.xlsx"))
        self.assertEqual(len(workbooks), 1)
        _, cells = read_sheet(workbooks[0])
        self.assertEqual(float(cells["D4"]["value"]), 100)
        self.assertEqual(cells["B5"]["value"], "1")
        self.assertEqual(cells["C8"]["value"], "12345678")
        self.assertEqual(m.count_pdf_pages(workbooks[0].with_suffix(".pdf")), 1)
        self.assertEqual(m.count_pdf_pages(preview), 3)
        self.assertTrue(Path(second.path).is_file())
        self.assertEqual(app.result_invoices, [second])

    def test_archive_without_preview_creates_pdf_and_excel(self):
        invoice = self.pdf_invoice()
        app = self.make_app([invoice])
        self.archive(app.mark_all_used)
        workbooks = list((m.USED_DIR / "报销20261102").glob("*.xlsx"))
        self.assertEqual(len(workbooks), 1)
        _, cells = read_sheet(workbooks[0])
        self.assertEqual(float(cells["D4"]["value"]), 100)
        self.assertEqual(m.count_pdf_pages(workbooks[0].with_suffix(".pdf")), 1)

    def test_new_preview_does_not_reuse_filename_of_existing_excel_only(self):
        invoice = self.pdf_invoice()
        existing = m.PRINT_JOBS_DIR / "发票组合_20261102_103000.xlsx"
        existing.write_bytes(b"preserve previous workbook")
        with patch.object(m, "datetime", wraps=datetime) as clock:
            clock.now.return_value = datetime(2026, 11, 2, 10, 30)
            pdf = m.merge_invoices_to_pdf([invoice])
        self.assertFalse(pdf.with_suffix(".xlsx").exists())
        self.assertEqual(existing.read_bytes(), b"preserve previous workbook")

    def test_bundle_filename_collision_keeps_both_pairs_and_updates_pdf_link(self):
        invoice = self.pdf_invoice()
        app = self.make_app([invoice])
        preview = self.generate_preview(app)
        directory = m.USED_DIR / "报销20261102"
        directory.mkdir()
        (directory / preview.name).write_bytes(b"existing PDF")
        (directory / preview.with_suffix(".xlsx").name).write_bytes(b"existing Excel")
        self.archive(app.mark_all_used)
        new_pdf = directory / f"{preview.stem}_1.pdf"
        new_excel = new_pdf.with_suffix(".xlsx")
        self.assertTrue(new_pdf.is_file())
        self.assertTrue(new_excel.is_file())
        self.assertEqual((directory / preview.name).read_bytes(), b"existing PDF")
        self.assertEqual((directory / preview.with_suffix(".xlsx").name).read_bytes(), b"existing Excel")
        with zipfile.ZipFile(new_excel) as archive:
            self.assertIn(new_pdf.name.encode("utf-8"), archive.read("xl/worksheets/_rels/sheet1.xml.rels"))
        self.assertEqual(m.count_pdf_pages(new_pdf), 1)

    def test_renamed_original_invoice_is_reflected_in_archived_excel(self):
        invoice = self.pdf_invoice("same.pdf")
        app = self.make_app([invoice])
        self.generate_preview(app)
        directory = m.USED_DIR / "报销20261102"
        directory.mkdir()
        (directory / "same.pdf").write_bytes(b"previous invoice")
        self.archive(app.mark_all_used)
        workbooks = list(directory.glob("*.xlsx"))
        self.assertEqual(len(workbooks), 1)
        _, cells = read_sheet(workbooks[0])
        self.assertEqual(cells["G8"]["value"], "same_1.pdf")
        self.assertEqual((directory / "same.pdf").read_bytes(), b"previous invoice")
        self.assertTrue((directory / "same_1.pdf").is_file())

    def test_partial_move_failure_exports_only_successfully_archived_invoices(self):
        first = self.pdf_invoice()
        second = self.pdf_invoice("second.pdf", "87654321", 20000)
        app = self.make_app([first, second])
        self.generate_preview(app)
        original_move = m.move_invoice

        def move(source, directory):
            if source == Path(second.path):
                raise PermissionError("second invoice is locked")
            return original_move(source, directory)

        with patch.object(m, "move_invoice", side_effect=move):
            self.archive(app.mark_all_used)
        workbooks = list((m.USED_DIR / "报销20261102").glob("*.xlsx"))
        self.assertEqual(len(workbooks), 1)
        _, cells = read_sheet(workbooks[0])
        self.assertEqual(float(cells["D4"]["value"]), 100)
        self.assertEqual(cells["B5"]["value"], "1")
        self.assertEqual(app.result_invoices, [second])
        self.assertTrue(Path(second.path).is_file())

    def test_archive_document_failure_keeps_recorded_invoices_and_removes_incomplete_outputs(self):
        invoice = self.pdf_invoice()
        app = self.make_app([invoice])
        with patch.object(m, "datetime", wraps=datetime) as clock, patch.object(m, "export_combination_excel", side_effect=PermissionError("Excel is locked")), patch.object(m.messagebox, "showwarning") as warning, patch.object(m.messagebox, "showinfo"):
            clock.now.return_value = datetime(2026, 11, 2, 10, 30)
            app.mark_all_used()
        directory = m.USED_DIR / "报销20261102"
        self.assertTrue((directory / "first.pdf").is_file())
        self.assertEqual(list(directory.glob("*.xlsx")), [])
        self.assertEqual(list(directory.glob("*.pdf")), [directory / "first.pdf"])
        self.assertEqual(m.load_used_invoice_numbers(), {"12345678"})
        self.assertEqual(app.result_invoices, [])
        self.assertIsNotNone(warning.call_args)
        self.assertIn("核对文件", warning.call_args.args[1])

    def test_partial_pdf_write_failure_does_not_leave_broken_preview(self):
        invoice = self.pdf_invoice()
        app = self.make_app([invoice])

        def fail_write(writer, output):
            output.write(b"%PDF-incomplete")
            raise OSError("disk write failed")

        with patch.object(m.PdfWriter, "write", autospec=True, side_effect=fail_write):
            self.archive(app.mark_all_used)
        self.assertEqual(list(m.PRINT_JOBS_DIR.glob("*.pdf")), [])
        self.assertEqual(list(m.PRINT_JOBS_DIR.glob("*.xlsx")), [])
        self.assertTrue((m.USED_DIR / "报销20261102" / "first.pdf").is_file())
        self.assertEqual(m.load_used_invoice_numbers(), {"12345678"})

    def test_workbook_copy_failure_removes_incomplete_archive_pair(self):
        invoice = self.pdf_invoice()
        app = self.make_app([invoice])
        preview = self.generate_preview(app)
        original_copy = m.shutil.copy2

        def copy(source, destination):
            if Path(source).suffix == ".xlsx":
                raise PermissionError("workbook cannot be read")
            return original_copy(source, destination)

        with patch.object(m.shutil, "copy2", side_effect=copy):
            self.archive(app.mark_all_used)
        directory = m.USED_DIR / "报销20261102"
        self.assertEqual(list(directory.iterdir()), [directory / "first.pdf"])
        self.assertTrue(preview.is_file())
        self.assertTrue(preview.with_suffix(".xlsx").is_file())
        self.assertEqual(m.load_used_invoice_numbers(), {"12345678"})


if __name__ == "__main__":
    unittest.main()
