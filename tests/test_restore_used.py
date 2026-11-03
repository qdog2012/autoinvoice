import json
import tempfile
import unittest
from contextlib import closing
from dataclasses import asdict, replace
from pathlib import Path
from unittest.mock import patch

import autoinvoice as m


class RestoreUsedArchiveTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-restore-")
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

    def archived_invoice(self, directory: Path, filename="26442000008512936336.pdf", number="26442000008512936336"):
        path = directory / filename
        with m.fitz.open() as document:
            page = document.new_page()
            page.insert_text((30, 30), f"Invoice {number}")
            page.insert_text((30, 100), "价税合计(大写)", fontname="china-s", fontsize=9)
            page.insert_text((200, 100), "(小写) 550.00", fontname="china-s", fontsize=9)
            document.save(path)
        return path

    def record(self, path: Path, number: str) -> dict:
        return {
            "path": str(path), "filename": path.name, "amount_cents": 55000,
            "invoice_type": "餐饮", "invoice_date": "2026年07月25日",
            "invoice_number": number, "seller": "简阳市海捞餐饮管理有限公司广州市福华路分公司",
            "pages": 1, "used_at": "2026-11-02 19:07:38",
            "target_amount": "12350.81", "moved_to": str(path),
        }

    def build_archive(self, name="报销20261102", count=1, base=26442000008512936330):
        directory = m.USED_DIR / name
        directory.mkdir(parents=True, exist_ok=True)
        numbers = []
        for index in range(count):
            number = str(base + index)
            path = self.archived_invoice(directory, f"{number}.pdf", number)
            record = self.record(path, number)
            m.save_used_invoice_record(record)
            m.append_json_record(m.STATE_FILE, record)
            numbers.append(number)
        (directory / "发票组合_20261102_190544.pdf").write_bytes(b"merged pdf")
        (directory / f"{name}.xlsx").write_bytes(b"review workbook")
        return directory, numbers

    def test_directories_and_invoice_files_exclude_archive_documents(self):
        self.build_archive("报销20261102")
        self.build_archive("差旅20261103", base=26442000008512939999)
        directories = m.used_archive_directories()
        # Newest-first by directory name, matching 内容YYYYMMDD ordering.
        self.assertEqual([path.name for path in directories], ["报销20261102", "差旅20261103"])
        files = m.archived_invoice_files(m.USED_DIR / "报销20261102")
        self.assertEqual([path.name for path in files], ["26442000008512936330.pdf"])

    def test_restore_moves_invoices_back_and_clears_both_records(self):
        directory, numbers = self.build_archive()
        moved, forgot, errors = m.restore_used_archive(directory)
        self.assertEqual(errors, [])
        self.assertEqual([path.name for path in moved], ["26442000008512936330.pdf"])
        self.assertEqual(forgot, set(numbers))

        restored = m.INVOICE_DIR / "26442000008512936330.pdf"
        self.assertTrue(restored.is_file())
        self.assertFalse((directory / "26442000008512936330.pdf").exists())
        self.assertNotIn(numbers[0], m.load_used_invoice_numbers())

        records = json.loads(m.STATE_FILE.read_text(encoding="utf-8"))
        self.assertEqual([record for record in records if record.get("invoice_number") in numbers], [])

        # The merged PDF and the review workbook stay where they are.
        self.assertTrue((directory / "发票组合_20261102_190544.pdf").is_file())
        self.assertTrue((directory / "报销20261102.xlsx").is_file())

    def test_restored_invoice_is_available_again_after_scan(self):
        directory, numbers = self.build_archive()
        m.restore_used_archive(directory)
        scanned, failures, reused, duplicates = m.scan_invoice_directory()
        self.assertEqual((failures, reused, duplicates), ([], [], []))
        self.assertEqual([invoice.invoice_number for invoice in scanned], numbers)
        self.assertEqual(scanned[0].amount_cents, 55000)

    def test_restore_keeps_other_batches_intact(self):
        directory, _numbers = self.build_archive()
        other_dir, other_numbers = self.build_archive("差旅20261103", base=26442000008512939999)
        m.restore_used_archive(directory)
        self.assertTrue((other_dir / f"{other_numbers[0]}.pdf").is_file())
        self.assertIn(other_numbers[0], m.load_used_invoice_numbers())

    def test_restore_reports_missing_or_empty_directory(self):
        empty = m.USED_DIR / "空目录"
        empty.mkdir(parents=True)
        with self.assertRaises(ValueError):
            m.restore_used_archive(m.USED_DIR / "不存在")
        with self.assertRaises(ValueError):
            m.restore_used_archive(empty)

    def test_existing_file_in_invoices_is_not_overwritten(self):
        directory, numbers = self.build_archive()
        existing = m.INVOICE_DIR / f"{numbers[0]}.pdf"
        existing.write_bytes(b"already here")
        moved, _forgot, errors = m.restore_used_archive(directory)
        self.assertEqual(errors, [])
        self.assertEqual(existing.read_bytes(), b"already here")
        self.assertEqual([path.name for path in moved], [f"{numbers[0]}_1.pdf"])

    def test_button_and_chooser_are_wired(self):
        self.build_archive()
        with patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        labels = [
            str(widget.cget("text"))
            for widget in app.calculate_button.master.winfo_children()
            if isinstance(widget, m.ttk.Button)
        ]
        self.assertIn("还原已使用", labels)

        with patch.object(app, "choose_used_archive", return_value=m.USED_DIR / "报销20261102"), \
                patch.object(m.messagebox, "askyesno", return_value=True), \
                patch.object(m.messagebox, "showinfo"), patch.object(m.messagebox, "showwarning"), \
                patch.object(app, "scan"):
            app.restore_used_invoices()
        self.assertTrue((m.INVOICE_DIR / "26442000008512936330.pdf").is_file())

    def test_cancelled_confirmation_changes_nothing(self):
        directory, numbers = self.build_archive()
        with patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        with patch.object(app, "choose_used_archive", return_value=directory), \
                patch.object(m.messagebox, "askyesno", return_value=False):
            app.restore_used_invoices()
        self.assertTrue((directory / f"{numbers[0]}.pdf").is_file())
        self.assertIn(numbers[0], m.load_used_invoice_numbers())


if __name__ == "__main__":
    unittest.main()
