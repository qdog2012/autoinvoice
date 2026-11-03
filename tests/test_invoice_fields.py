import tempfile
import unittest
from pathlib import Path

import autoinvoice as m


def build_pdf(path: Path, cells, page=(595.0, 400.0)):
    """Draw each cell as its own text object, like real e-invoices do."""
    with m.fitz.open() as document:
        pg = document.new_page(width=page[0], height=page[1])
        for x, y, text in cells:
            pg.insert_text((x, y), text, fontname="china-s", fontsize=9)
        document.save(path)


class SellerFromLayoutTests(unittest.TestCase):
    def test_right_hand_name_block_is_the_seller(self):
        rows = [
            (595.0, [
                (16.0, "购"), (32.0, "名称："), (57.0, "广州由榕管理咨询有限公司"),
                (301.0, "销"), (317.0, "名称："), (340.0, "广州真功夫快餐连锁管理有限公司"),
                (437.0, "914401016681074731"),
            ]),
        ]
        self.assertEqual(
            m.extract_seller_from_row_parts(rows),
            "广州真功夫快餐连锁管理有限公司",
        )

    def test_stacked_header_keeps_the_seller_next_to_the_tax_id_caption(self):
        rows = [
            (595.0, [
                (32.0, "名称："), (32.0, "统一社会信用代码/纳税人识别号："),
                (57.0, "广州由榕管理咨询有限公司"), (152.0, "91440101MA5AQ9F15Y"),
                (317.0, "名称："), (317.0, "统一社会信用代码/纳税人识别号："),
                (340.0, "中碧盈信（广州）石油有限公司"), (437.0, "91440113MA9YE0N669"),
            ]),
        ]
        self.assertEqual(m.extract_seller_from_row_parts(rows), "中碧盈信（广州）石油有限公司")

    def test_single_name_block_is_not_treated_as_seller(self):
        rows = [(595.0, [(32.0, "名称："), (57.0, "广州由榕管理咨询有限公司")])]
        self.assertEqual(m.extract_seller_from_row_parts(rows), "")

    def test_text_fallback_takes_the_trailing_name_on_a_shared_line(self):
        text = (
            "名称:广州由榕管理咨询有限公司 名称:广州市小老表餐饮有限公司\n"
            "购买方信息 销售方信息\n"
        )
        self.assertEqual(m.extract_seller(text), "广州市小老表餐饮有限公司")

    def test_text_fallback_keeps_the_labelled_block_behaviour(self):
        text = "销\n售\n方\n信\n息\n名称:简阳市海捞餐饮管理有限公司广州市福华路分公司\n"
        self.assertEqual(m.extract_seller(text), "简阳市海捞餐饮管理有限公司广州市福华路分公司")


class PdfFieldTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-fields-")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def test_pdf_text_comes_from_pymupdf(self):
        path = self.root / "text.pdf"
        build_pdf(path, [(30, 60, "发票号码：26447900220700149840"), (30, 90, "开票日期：2026年07月08日")])
        text, pages = m.read_pdf_text(path)
        self.assertEqual(pages, 1)
        self.assertIn("26447900220700149840", text)
        self.assertIn("2026年07月08日", text)

    def test_number_seller_and_totals_from_a_layout_split_across_objects(self):
        path = self.root / "fields.pdf"
        build_pdf(path, [
            (300, 60, "发票号码："), (380, 60, "26447900220700149840"),
            (300, 80, "开票日期："), (380, 80, "2026年07月08日"),
            (16, 120, "购"), (32, 120, "名称："), (57, 120, "广州由榕管理咨询有限公司"),
            (301, 120, "销"), (317, 120, "名称："), (340, 120, "广东联合电子服务股份有限公司"),
            (30, 200, "合 计"), (400, 200, "¥4.42"),
            (30, 240, "价税合计（大写）"), (150, 240, "肆圆肆角贰分"), (330, 240, "（小写） ¥4.42"),
        ])
        invoice = m.parse_invoice(path)
        self.assertEqual(invoice.invoice_number, "26447900220700149840")
        self.assertEqual(invoice.invoice_date, "2026年07月08日")
        self.assertEqual(invoice.seller, "广东联合电子服务股份有限公司")
        self.assertEqual(invoice.amount_cents, 442)


if __name__ == "__main__":
    unittest.main()
