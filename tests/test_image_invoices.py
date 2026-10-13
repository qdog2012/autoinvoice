import tempfile
import threading
import unittest
from dataclasses import asdict, replace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import autoinvoice as m


SAMPLE_TEXT = "发票号码:26442000008747028331\n开票日期:2026年07月30日\n餐费\n价税合计(大写)\n(小写)295.20"


def ocr_result(text=SAMPLE_TEXT):
    lines = text.splitlines()
    boxes = [[[0, index * 30], [250, index * 30], [250, index * 30 + 20], [0, index * 30 + 20]] for index in range(len(lines))]
    return SimpleNamespace(txts=lines, boxes=boxes)


class ImageInvoiceTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-images-")
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        data = root / "data"
        runtime = patch.multiple(
            m, APP_DIR=root, INVOICE_DIR=root / "invoices",
            USED_DIR=root / "used_invoices", VOID_DIR=root / "void_invoices",
            FAILED_DIR=root / "failed_invoices", PRINT_JOBS_DIR=root / "print_jobs",
            OFD_CACHE_DIR=root / "print_jobs" / "ofd_cache", DATA_DIR=data,
            DATABASE_FILE=data / "invoice_records.sqlite3", STATE_FILE=data / "used_records.json",
            FAILED_RECORD_FILE=data / "failed_records.json", SETTINGS_FILE=data / "app_settings.json",
            LEGACY_RUNTIME_FILES=(), DATABASE_INITIALIZED=False,
        )
        runtime.start()
        self.addCleanup(runtime.stop)
        self.engine = Mock(return_value=ocr_result())
        engine_patch = patch.object(m, "get_ocr_engine", return_value=self.engine)
        engine_patch.start()
        self.addCleanup(engine_patch.stop)
        m.ensure_dirs()

    def image(self, name="invoice.png", color="red", size=(300, 200)):
        path = m.INVOICE_DIR / name
        with m.Image.new("RGB", size, color) as image:
            image.save(path)
        return path

    def test_all_image_formats_parse_amount_number_date_and_category(self):
        for suffix in m.IMAGE_SUFFIXES | {".PNG", ".JPEG"}:
            with self.subTest(suffix=suffix):
                invoice = m.parse_invoice(self.image("invoice" + suffix))
                self.assertEqual(invoice.amount_cents, 29520)
                self.assertEqual(invoice.invoice_number, "26442000008747028331")
                self.assertEqual(invoice.invoice_date, "2026年07月30日")
                self.assertEqual(invoice.invoice_type, m.CATEGORIES[2])
                self.assertEqual(invoice.pages, 1)
                self.assertEqual(m.count_pdf_pages(m.invoice_pdf_path(invoice)), 1)

    def test_ocr_cache_survives_rename_and_invalidates_content_changes(self):
        path = self.image()
        first = m.parse_invoice(path)
        renamed = path.with_name("renamed.png")
        path.rename(renamed)
        self.assertEqual(m.parse_invoice(renamed).amount_cents, first.amount_cents)
        self.assertEqual(self.engine.call_count, 1)
        self.image("renamed.png", color="blue")
        m.parse_invoice(renamed)
        self.assertEqual(self.engine.call_count, 2)

    def test_total_labels_prevent_subtotal_or_invoice_number_as_amount(self):
        cases = {
            "金额合计100.00\n税额13.00\n(小写)¥113.00": 11300,
            "价税合计: 1,234. 56": 123456,
            "(小写) 295.20": 29520,
            "实收金额: ¥50.00": 5000,
            "¥292.28\n价税合计(大写)\n(小写)模糊": None,
            "(小写)\n发票号码:26442000008747028331": None,
            "(小写)29520": None,
            "税额¥2.92\n(小写)¥295.200": None,
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(m.extract_image_amount_cents(text), expected)

    def test_pdf_wins_duplicate_and_manual_type_survives_image_rescan(self):
        path = self.image("a.png")
        invoice = m.parse_invoice(path)
        m.save_invoice_type_override(invoice, m.CATEGORIES[1])
        scanned, failures, reused, duplicates = m.scan_invoice_directory()
        self.assertEqual(scanned[0].invoice_type, m.CATEGORIES[1])
        self.assertEqual((failures, reused, duplicates), ([], [], []))
        pdf = m.INVOICE_DIR / "z.pdf"
        with m.fitz.open() as document:
            document.new_page()
            document.save(pdf)
        original_parser = m.parse_invoice
        def parser(source):
            if source.suffix == ".pdf":
                return replace(invoice, path=str(source), filename=source.name)
            return original_parser(source)
        with patch.object(m, "parse_invoice", side_effect=parser):
            scanned, failures, reused, duplicates = m.scan_invoice_directory()
        self.assertEqual(len(scanned), 1)
        self.assertEqual(scanned[0].path, str(pdf))
        self.assertEqual(scanned[0].invoice_type, m.CATEGORIES[1])
        self.assertEqual(len(duplicates), 1)
        self.assertTrue((m.VOID_DIR / path.name).exists())

    def test_used_image_is_moved_and_failed_image_is_recorded(self):
        path = self.image()
        invoice = m.parse_invoice(path)
        m.save_used_invoice_record(asdict(invoice))
        scanned, failures, reused, duplicates = m.scan_invoice_directory()
        self.assertEqual(scanned, [])
        self.assertEqual(len(reused), 1)
        self.assertTrue((m.USED_DIR / path.name).exists())
        self.image("blank.png", color="white")
        self.engine.return_value = ocr_result("unreadable")
        scanned, failures, reused, duplicates = m.scan_invoice_directory()
        self.assertEqual(len(failures), 1)
        self.assertTrue((m.FAILED_DIR / "blank.png").exists())
        self.assertTrue(m.FAILED_RECORD_FILE.exists())

    def test_missing_ocr_engine_does_not_move_images(self):
        path = self.image()
        with patch.object(m, "get_ocr_engine", side_effect=RuntimeError("missing model")):
            with self.assertRaises(RuntimeError):
                m.scan_invoice_directory()
        self.assertTrue(path.exists())
        self.assertEqual(list(m.FAILED_DIR.iterdir()), [])

    def test_image_preview_orientation_alpha_and_cache(self):
        path = m.INVOICE_DIR / "transparent.png"
        with m.Image.new("RGBA", (300, 200), (0, 0, 0, 0)) as image:
            image.paste((255, 0, 0, 255), (50, 50, 100, 100))
            image.save(path)
        invoice = m.parse_invoice(path)
        pdf = m.invoice_pdf_path(invoice)
        with m.fitz.open(pdf) as document:
            page = document[0]
            self.assertAlmostEqual(page.rect.width / page.rect.height, 1.5, places=6)
            pixmap = page.get_pixmap()
            self.assertEqual(pixmap.pixel(0, 0), (255, 255, 255))
            self.assertEqual(pixmap.pixel(200, 200), (255, 0, 0))
        timestamp = pdf.stat().st_mtime_ns
        self.assertEqual(m.invoice_pdf_path(invoice), pdf)
        self.assertEqual(pdf.stat().st_mtime_ns, timestamp)
        self.image("transparent.png", color="blue")
        self.assertNotEqual(m.invoice_pdf_path(invoice), pdf)

    def test_multipage_tiff_and_mixed_pdf_image_combination(self):
        path = m.INVOICE_DIR / "invoice.tiff"
        with m.Image.new("RGB", (300, 200), "red") as first, m.Image.new("RGB", (200, 300), "blue") as second:
            first.save(path, save_all=True, append_images=[second])
        invoice = m.parse_invoice(path)
        self.assertEqual(invoice.pages, 2)
        self.assertEqual(m.count_pdf_pages(m.invoice_pdf_path(invoice)), 2)
        other = m.INVOICE_DIR / "other.pdf"
        with m.fitz.open() as document:
            document.new_page()
            document.save(other)
        other_invoice = replace(invoice, path=str(other), filename=other.name, amount_cents=5000, invoice_number="other")
        combination = m.find_best_combination([invoice, other_invoice], 34520)
        self.assertEqual(len(combination), 2)
        self.assertEqual(m.count_pdf_pages(m.merge_invoices_to_pdf(combination)), 3)

    def test_exif_orientation_is_applied_to_preview(self):
        path = m.INVOICE_DIR / "rotated.jpg"
        with m.Image.new("RGB", (200, 300), "red") as image:
            exif = m.Image.Exif()
            exif[274] = 6
            image.save(path, exif=exif)
        with m.fitz.open(m.invoice_pdf_path(m.parse_invoice(path))) as document:
            self.assertGreater(document[0].rect.width, document[0].rect.height)

    def test_background_scan_keeps_tk_event_loop_responsive(self):
        path = self.image()
        invoice = m.parse_invoice(path)
        with patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        release = threading.Event()
        finished = threading.Event()
        heartbeats = []
        def scanner(progress):
            if not release.wait(5):
                raise RuntimeError("scan test timeout")
            return [invoice], [], [], []
        original_done = app.on_scan_done
        def done(result, error):
            original_done(result, error)
            finished.set()
            app.quit()
        with patch.object(m, "scan_invoice_directory", side_effect=scanner), patch.object(app, "on_scan_done", side_effect=done):
            app.scan()
            self.assertTrue(app.scanning_running)
            self.assertIn("disabled", app.calculate_button.state())
            app.after(20, lambda: heartbeats.append(True))
            app.after(80, release.set)
            timeout = app.after(6000, app.quit)
            app.mainloop()
            app.after_cancel(timeout)
        self.assertTrue(finished.is_set())
        self.assertEqual(heartbeats, [True])
        self.assertEqual(app.invoices, [invoice])
        self.assertFalse(app.scanning_running)
        self.assertNotIn("disabled", app.calculate_button.state())


if __name__ == "__main__":
    unittest.main()
