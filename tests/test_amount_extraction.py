import tempfile
import unittest
from pathlib import Path

import autoinvoice as m


def build_invoice_pdf(path: Path, cells, page=(595.0, 400.0)):
    """Draw each cell as its own text object so reading order matches a real invoice."""
    with m.fitz.open() as document:
        pg = document.new_page(width=page[0], height=page[1])
        for x, y, text, size in cells:
            pg.insert_text((x, y), text, fontname="china-s", fontsize=size)
        document.save(path)


class RowAmountExtractionTests(unittest.TestCase):
    def test_uppercase_amount_parsing(self):
        cases = {
            "壹佰伍拾肆圆伍角整": 15450,
            "壹拾玖圆捌角整": 1980,
            "肆佰陆拾陆圆零伍分": 46605,
            "贰仟伍佰壹拾贰圆叁角": 251230,
            "柒佰肆拾陆圆整": 74600,
            "伍佰伍拾圆整": 55000,
            "壹仟贰佰零陆圆陆角伍分": 120665,
            "捌拾圆零壹分": 8001,
            "人民币壹佰圆整": 10000,
            "没有大写金额": None,
            "壹佰圆": 10000,
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(m.parse_uppercase_amount_cents(text), expected)

    def test_total_wins_over_subtotal_on_the_same_row(self):
        rows = [
            "项目名称 规格型号 数量 单价 金额 税率 税额",
            "餐费 1 619.81 619.81 6% 37.19",
            "餐费 -100.94 6% -6.06",
            "合 计 ¥518.87 ¥31.13",
            "价税合计(大写) 伍佰伍拾圆整 (小写) ¥550.00",
        ]
        self.assertEqual(m.amount_cents_from_visual_rows(rows), 55000)

    def test_total_before_its_label_is_still_found(self):
        rows = [
            "合 计 ¥518.87 ¥31.13",
            "(小写) ¥  550.00 价税合计(大写)",
            "伍佰伍拾圆整",
        ]
        self.assertEqual(m.amount_cents_from_visual_rows(rows), 55000)

    def test_currency_symbol_after_the_amount_is_accepted(self):
        rows = [
            "合         计 17.52¥ 2.28¥",
            "价税合计（大写） （小写）",
            "壹拾玖圆捌角整 19.80¥",
        ]
        self.assertEqual(m.amount_cents_from_visual_rows(rows), 1980)

    def test_uppercase_amount_is_the_fallback_when_digits_are_missing(self):
        rows = ["价税合计(大写)", "壹佰伍拾肆圆伍角整", "开票人:连泽森"]
        self.assertEqual(m.amount_cents_from_visual_rows(rows), 15450)

    def test_unrelated_currency_value_is_not_used_as_the_total(self):
        rows = ["¥292.28", "价税合计(大写)", "(小写)模糊"]
        self.assertIsNone(m.amount_cents_from_visual_rows(rows))

    def test_visual_rows_keep_row_order_and_rejoin_split_fragments(self):
        items = [
            (400.88, 258.85, 433.91, 270.19, "¥518.87"),
            (438.11, 281.47, 480.15, 290.47, "  550.00"),
            (175.84, 281.47, 229.84, 290.47, "伍佰伍拾圆整"),
            (411.11, 280.86, 448.62, 292.21, "(小写) ¥"),
            (52.88, 281.47, 115.88, 290.47, "价税合计(大写)"),
            (49.31, 260.80, 58.31, 269.80, "合"),
            (73.31, 260.80, 82.31, 269.80, "计"),
        ]
        rows = m.cluster_visual_rows(items)
        self.assertEqual(
            rows,
            [
                "合 计 ¥518.87",
                "价税合计(大写) 伍佰伍拾圆整 (小写) ¥   550.00",
            ],
        )


class PdfRowAmountTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="autoinvoice-rows-")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def test_total_placed_before_its_label_is_read_from_the_row(self):
        path = self.root / "total-first.pdf"
        build_invoice_pdf(
            path,
            [
                (30, 60, "合 计", 10),
                (470, 60, "¥518.87", 10),
                (30, 100, "伍佰伍拾圆整", 10),
                (200, 100, "(小写)", 10),
                (300, 100, "¥550.00", 10),
                (430, 100, "价税合计(大写)", 10),
            ],
        )
        self.assertEqual(m.parse_invoice(path).amount_cents, 55000)

    def test_symbol_after_the_total_is_read_from_the_row(self):
        path = self.root / "symbol-after.pdf"
        build_invoice_pdf(
            path,
            [
                (30, 60, "合 计", 10),
                (400, 60, "17.52¥", 10),
                (470, 60, "2.28¥", 10),
                (30, 100, "壹拾玖圆捌角整", 10),
                (250, 100, "19.80¥", 10),
                (150, 100, "（小写）", 10),
                (30, 130, "价税合计（大写）", 10),
            ],
        )
        self.assertEqual(m.parse_invoice(path).amount_cents, 1980)

    def test_line_item_amounts_do_not_override_the_total(self):
        path = self.root / "line-items.pdf"
        build_invoice_pdf(
            path,
            [
                (30, 40, "项目名称 金额 税额", 9),
                (30, 60, "餐费 619.81 37.19", 9),
                (30, 80, "餐费 -100.94 -6.06", 9),
                (30, 110, "合 计", 9),
                (400, 110, "¥518.87", 9),
                (500, 110, "¥31.13", 9),
                (30, 140, "价税合计(大写)", 9),
                (150, 140, "伍佰伍拾圆整", 9),
                (350, 140, "(小写) ¥550.00", 9),
            ],
        )
        self.assertEqual(m.parse_invoice(path).amount_cents, 55000)


if __name__ == "__main__":
    unittest.main()
