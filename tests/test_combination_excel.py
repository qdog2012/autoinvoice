import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
import xml.etree.ElementTree as ET

import autoinvoice as m


NAMESPACE = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_sheet(path):
    with zipfile.ZipFile(path) as archive:
        strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        texts = ["".join(item.itertext()) for item in strings]
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        cells = {}
        for cell in sheet.findall(".//s:sheetData/s:row/s:c", NAMESPACE):
            value = cell.find("s:v", NAMESPACE)
            formula = cell.find("s:f", NAMESPACE)
            raw_value = value.text if value is not None else None
            cells[cell.attrib["r"]] = {
                "type": cell.attrib.get("t", "n"),
                "value": texts[int(raw_value)] if cell.attrib.get("t") == "s" else raw_value,
                "formula": formula.text if formula is not None else None,
            }
        return sheet, cells


class CombinationExcelTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-excel-")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.invoices = [
            m.Invoice("first.png", "first.png", 29520, "餐饮", "2026年07月30日", "26442000008747028331", "测试餐饮有限公司", 1),
            m.Invoice("second.pdf", "second.pdf", 5000, "交通", "2026-08-01", "00000001234567890123", "测试交通有限公司", 2),
        ]

    def test_numeric_amounts_text_identifiers_formulas_and_pdf_page_ranges(self):
        path = m.export_combination_excel(self.invoices, self.root / "combination.xlsx", 34000, [1, 2])
        sheet, cells = read_sheet(path)
        self.assertEqual(cells["C8"]["value"], self.invoices[0].invoice_number)
        self.assertEqual(cells["C9"]["value"], self.invoices[1].invoice_number)
        self.assertEqual(cells["C9"]["type"], "s")
        self.assertEqual(cells["D8"]["type"], "n")
        self.assertEqual(float(cells["D8"]["value"]), 295.2)
        self.assertEqual(cells["E8"]["type"], "n")
        self.assertEqual(cells["D4"]["formula"], "ROUND(SUM(D8:D9),2)")
        self.assertEqual(float(cells["D4"]["value"]), 345.2)
        self.assertEqual(float(cells["E4"]["value"]), 5.2)
        self.assertEqual(cells["B5"]["formula"], "COUNT(A8:A9)")
        self.assertEqual(cells["B5"]["value"], "2")
        self.assertEqual(cells["H8"]["value"], "1")
        self.assertEqual(cells["H9"]["value"], "2-3")
        pane = sheet.find(".//s:pane", NAMESPACE)
        self.assertEqual(pane.attrib["topLeftCell"], "D8")
        with zipfile.ZipFile(path) as archive:
            table = ET.fromstring(archive.read("xl/tables/table1.xml"))
            self.assertEqual(table.attrib["ref"], "A7:H9")

    def test_missing_target_and_date_do_not_become_zero(self):
        invoice = m.replace(self.invoices[0], invoice_date="", invoice_number="")
        path = m.export_combination_excel([invoice], self.root / "blank.xlsx", page_counts=[1])
        _, cells = read_sheet(path)
        self.assertNotIn("B4", cells)
        self.assertEqual(cells["E4"]["value"], None)
        self.assertEqual(cells["E8"]["value"], "")
        self.assertEqual(cells["C8"]["value"], "")

    def test_negative_difference_and_formula_like_source_text(self):
        invoice = m.replace(self.invoices[0], filename="=SUM(1,2).png", seller="=1+1")
        path = m.export_combination_excel([invoice], self.root / "literal.xlsx", 30000, [1])
        _, cells = read_sheet(path)
        self.assertEqual(float(cells["E4"]["value"]), -4.8)
        for address in ("F8", "G8"):
            self.assertEqual(cells[address]["type"], "s")
            self.assertIsNone(cells[address]["formula"])

    def test_difference_status_stays_numeric_and_formats_each_outcome_inline(self):
        expected_format = '"超出："#,##0.00" 元";"不足："#,##0.00" 元";"金额刚好";@'
        for target, difference in ((34000, 5.2), (35000, -4.8), (34520, 0.0), (None, None)):
            with self.subTest(target=target):
                path = m.export_combination_excel(self.invoices, self.root / f"status-{target}.xlsx", target, [1, 2])
                sheet, cells = read_sheet(path)
                self.assertEqual(cells["E4"]["formula"], 'IF(ISNUMBER(B4),ROUND(D4-B4,2),"")')
                value = cells["E4"]["value"]
                self.assertEqual(float(value) if value is not None else None, difference)
                self.assertNotIn("正数超出，负数不足", [cell["value"] for cell in cells.values()])
                self.assertIn("E4:F4", [item.attrib["ref"] for item in sheet.findall("s:mergeCells/s:mergeCell", NAMESPACE)])
                with zipfile.ZipFile(path) as archive:
                    styles = ET.fromstring(archive.read("xl/styles.xml"))
                cell = sheet.find('.//s:c[@r="E4"]', NAMESPACE)
                style = styles.find("s:cellXfs", NAMESPACE)[int(cell.attrib["s"])]
                number_formats = {item.attrib["numFmtId"]: item.attrib["formatCode"] for item in styles.findall("s:numFmts/s:numFmt", NAMESPACE)}
                self.assertEqual(number_formats[style.attrib["numFmtId"]], expected_format)
                self.assertEqual(style.find("s:alignment", NAMESPACE).attrib["horizontal"], "left")

    def test_empty_combination_or_invalid_page_counts_is_rejected(self):
        for invoices, counts in (([], []), (self.invoices, [1]), (self.invoices, [1, 0])):
            with self.assertRaises(ValueError):
                m.export_combination_excel(invoices, self.root / "invalid.xlsx", page_counts=counts)

    def test_preview_exports_in_pdf_order_and_opens_excel_when_requested(self):
        with patch.object(m, "ensure_dirs"), patch.object(m, "load_last_target_amount", return_value="340"), patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        app.result_invoices = list(self.invoices)
        pdf = self.root / "preview.pdf"
        with patch.object(m, "merge_invoices_to_pdf", return_value=pdf) as merger, patch.object(m, "count_pdf_pages", return_value=3), patch.object(m, "export_combination_excel", return_value=pdf.with_suffix(".xlsx")) as exporter, patch.object(app, "render_preview"), patch.object(m.messagebox, "askyesno", return_value=True), patch.object(m.os, "startfile") as opener:
            app.generate_combination_preview()
        expected = sorted(self.invoices, key=m.invoice_type_sort_key)
        merger.assert_called_once_with(expected)
        exporter.assert_called_once_with(expected, pdf.with_suffix(".xlsx"), 34000)
        opener.assert_called_once_with(pdf.with_suffix(".xlsx"))
        self.assertEqual(app.print_preview_invoices, expected)

    def test_export_failure_is_reported_without_replacing_existing_preview(self):
        with patch.object(m, "ensure_dirs"), patch.object(m, "load_last_target_amount", return_value="340"), patch.object(m.AutoInvoiceApp, "scan"):
            app = m.AutoInvoiceApp()
        self.addCleanup(app.destroy)
        app.result_invoices = list(self.invoices)
        previous_pdf = self.root / "previous.pdf"
        app.print_preview_pdf = previous_pdf
        with patch.object(m, "merge_invoices_to_pdf", return_value=self.root / "new.pdf"), patch.object(m, "count_pdf_pages", return_value=3), patch.object(m, "export_combination_excel", side_effect=PermissionError("file is open")), patch.object(m.messagebox, "showerror") as error, patch.object(m.messagebox, "askyesno") as success:
            app.generate_combination_preview()
        error.assert_called_once()
        success.assert_not_called()
        self.assertEqual(app.print_preview_pdf, previous_pdf)

    def test_repeated_pdf_generation_does_not_overwrite_prior_job(self):
        source = self.root / "invoice.pdf"
        with m.fitz.open() as document:
            document.new_page()
            document.save(source)
        invoice = m.replace(self.invoices[0], path=str(source), filename=source.name)
        with patch.object(m, "ensure_dirs"), patch.object(m, "PRINT_JOBS_DIR", self.root), patch.object(m, "datetime") as clock:
            clock.now.return_value = datetime(2026, 10, 13, 12, 0, 0)
            first = m.merge_invoices_to_pdf([invoice])
            second = m.merge_invoices_to_pdf([invoice])
        self.assertNotEqual(first, second)
        self.assertTrue(first.exists() and second.exists())


if __name__ == "__main__":
    unittest.main()
