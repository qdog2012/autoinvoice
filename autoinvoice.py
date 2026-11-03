# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import posixpath
import re
import shutil
import sqlite3
import sys
import threading
import traceback
import unicodedata
import ctypes
import hashlib
import zipfile
from contextlib import closing, suppress
from ctypes import wintypes
from dataclasses import asdict, dataclass, replace
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Callable, Iterable
import xml.etree.ElementTree as ET

import tkinter as tk
from tkinter import messagebox, ttk

from pypdf import PdfReader, PdfWriter
import xlsxwriter

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover - dependency check is shown in the UI
    fitz = None

try:
    from PIL import Image, ImageOps, ImageTk, ImageWin
except Exception:  # pragma: no cover - dependency check is shown in the UI
    Image = ImageOps = ImageTk = ImageWin = None

try:
    import win32con
    import win32gui
    import win32print
    import win32ui
except Exception:  # pragma: no cover - non-Windows import guard
    win32con = win32gui = win32print = win32ui = None


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
INVOICE_DIR = APP_DIR / "invoices"
USED_DIR = APP_DIR / "used_invoices"
FAILED_DIR = APP_DIR / "failed_invoices"
VOID_DIR = APP_DIR / "void_invoices"
PRINT_JOBS_DIR = APP_DIR / "print_jobs"
OFD_CACHE_DIR = PRINT_JOBS_DIR / "ofd_cache"
DATA_DIR = APP_DIR / "data"
STATE_FILE = DATA_DIR / "used_records.json"
FAILED_RECORD_FILE = DATA_DIR / "failed_records.json"
SETTINGS_FILE = DATA_DIR / "app_settings.json"
DATABASE_FILE = DATA_DIR / "invoice_records.sqlite3"
LEGACY_RUNTIME_FILES = (
    (APP_DIR / "used_records.json", STATE_FILE),
    (APP_DIR / "failed_records.json", FAILED_RECORD_FILE),
    (APP_DIR / "app_settings.json", SETTINGS_FILE),
    (APP_DIR / "invoice_records.sqlite3", DATABASE_FILE),
)
APP_VERSION = "v1.1103.1442"

CATEGORY_ALL = "全部"
CATEGORIES = [CATEGORY_ALL, "交通", "餐饮", "住宿", "办公", "其他"]
CATEGORY_SORT_ORDER = {category: index for index, category in enumerate(CATEGORIES) if category != CATEGORY_ALL}
DATE_PRIORITY_WINDOW_CENTS = 5000
SMALL_INVOICE_THRESHOLD_CENTS = 5000
SMALL_INVOICE_MAX_COUNT = 2
DATABASE_INITIALIZED = False
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
SUPPORTED_INVOICE_SUFFIXES = {".pdf", ".ofd"} | IMAGE_SUFFIXES
PDF_SUFFIX = ".pdf"
OFD_SUFFIX = ".ofd"
OFD_RENDER_VERSION = "ofd-render-v3"
IMAGE_RENDER_VERSION = "image-render-v1"
OCR_CACHE_VERSION = "rapidocr-3.10-v1"
OCR_ENGINE = None
OCR_LOCK = threading.Lock()
PT_PER_MM = 72 / 25.4
CHINESE_FONT_CANDIDATES = (
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\simhei.ttf"),
    Path(r"C:\Windows\Fonts\simsun.ttc"),
    Path(r"C:\Windows\Fonts\msyh.ttf"),
    Path(r"C:\Windows\Fonts\arialuni.ttf"),
)

TYPE_KEYWORDS = {
    "交通": [
        "交通",
        "运输",
        "客运",
        "出租车",
        "网约车",
        "火车",
        "铁路",
        "机票",
        "航空",
        "行程单",
        "停车",
        "停车费",
        "停车场",
        "泊车",
        "车位",
        "车辆",
        "临保",
        "车牌号",
        "通行费",
        "路桥费",
        "高速",
        "地铁",
        "公交",
        "滴滴",
        "高德",
        "加油",
        "汽油",
    ],
    "住宿": ["住宿", "酒店", "宾馆", "旅馆", "民宿", "房费", "客房"],
    "餐饮": [
        "餐饮",
        "餐费",
        "美食",
        "饭店",
        "餐厅",
        "蔬菜",
        "肉及肉制品",
        "谷物",
        "调味品",
        "食品",
        "乳制品",
        "水果",
        "水产品",
    ],
    "办公": [
        "办公",
        "文具",
        "打印",
        "耗材",
        "软件",
        "信息系统",
        "网络服务",
        "云服务",
        "会员订阅",
        "通信",
        "宽带",
        "咨询",
        "服务费",
        "管理费",
        "月租",
        "租赁",
        "物业",
        "电脑",
        "电子",
        "腾讯云",
        "钉钉",
    ],
}


@dataclass(frozen=True)
class Invoice:
    path: str
    filename: str
    amount_cents: int
    invoice_type: str
    invoice_date: str = ""
    invoice_number: str = ""
    seller: str = ""
    pages: int = 0

    @property
    def amount(self) -> Decimal:
        return Decimal(self.amount_cents) / Decimal(100)

    @property
    def amount_text(self) -> str:
        return cents_to_text(self.amount_cents)


@dataclass(frozen=True)
class PrintDialogSelection:
    hdc: int
    hdevmode: int | None
    hdevnames: int | None
    from_page: int | None = None
    to_page: int | None = None


def ensure_dirs() -> None:
    for directory in (INVOICE_DIR, USED_DIR, FAILED_DIR, VOID_DIR, PRINT_JOBS_DIR, OFD_CACHE_DIR, DATA_DIR):
        directory.mkdir(parents=True, exist_ok=True)
    migrate_legacy_runtime_files()


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("\u00a5", "¥").replace("￥", "¥")
    return text


def cents_to_text(cents: int) -> str:
    return f"{Decimal(cents) / Decimal(100):.2f}"


def parse_money_to_cents(value: str) -> int | None:
    try:
        money = Decimal(value.replace(",", "").strip())
    except (InvalidOperation, AttributeError):
        return None
    if money < 0:
        return None
    money = money.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int(money * 100)


def read_pdf_text(path: Path) -> tuple[str, int]:
    """Read PDF text, preferring PyMuPDF.

    Some e-invoice generators (ETC toll receipts among them) embed digits in a
    way pypdf turns into space-separated single characters and mojibake, which
    silently drops the invoice number and date. PyMuPDF decodes those files
    correctly, so it is tried first and pypdf stays as the fallback.
    """
    if fitz is not None:
        try:
            text, pages = read_pdf_text_with_fitz(path)
        except Exception:
            text, pages = "", 0
        if text.strip():
            return text, pages

    text_parts: list[str] = []
    pages = 0
    try:
        reader = PdfReader(str(path))
        pages = len(reader.pages)
        for page in reader.pages:
            text_parts.append(page.extract_text() or "")
    except Exception:
        text_parts = []
    return normalize_text("\n".join(text_parts).strip()), pages


def read_pdf_text_with_fitz(path: Path) -> tuple[str, int]:
    if fitz is None:
        return "", 0
    with fitz.open(path) as doc:
        text = "\n".join(page.get_text("text") for page in doc)
        return normalize_text(text), len(doc)


def cluster_visual_row_parts(
    items: Iterable[tuple[float, float, float, float, str]],
) -> list[list[tuple[float, str]]]:
    """Rebuild visual text rows from positioned text runs, keeping x positions.

    Many e-invoices draw each cell as its own text object, so the reading order
    produced by the PDF text extractor is not the visual order. Grouping runs by
    vertical overlap and then ordering them by x restores the row a human sees.
    """
    rows: list[list] = []
    for x0, y0, x1, y1, text in sorted(items, key=lambda item: (item[1], item[0])):
        height = max(y1 - y0, 0.1)
        target = None
        best_ratio = 0.0
        for row in rows:
            overlap = min(row[1], y1) - max(row[0], y0)
            ratio = overlap / min(height, max(row[1] - row[0], 0.1))
            if ratio > best_ratio:
                best_ratio = ratio
                target = row
        if target is not None and best_ratio >= 0.45:
            target[0] = min(target[0], y0)
            target[1] = max(target[1], y1)
            target[2].append((x0, text))
        else:
            rows.append([y0, y1, [(x0, text)]])
    rows.sort(key=lambda row: row[0])
    return [sorted(row[2], key=lambda part: part[0]) for row in rows]


def cluster_visual_rows(items: Iterable[tuple[float, float, float, float, str]]) -> list[str]:
    return [" ".join(text for _, text in parts) for parts in cluster_visual_row_parts(items)]


def read_pdf_visual_row_parts(path: Path) -> list[tuple[float, list[tuple[float, str]]]]:
    """Return (page_width, row parts) for every visual row of the PDF."""
    if fitz is None:
        return []
    rows: list[tuple[float, list[tuple[float, str]]]] = []
    try:
        with fitz.open(path) as doc:
            for page in doc:
                items: list[tuple[float, float, float, float, str]] = []
                for block in page.get_text("dict").get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    for line in block.get("lines", []):
                        text = "".join(span.get("text", "") for span in line.get("spans", []))
                        if not text.strip():
                            continue
                        x0, y0, x1, y1 = line.get("bbox")
                        items.append((x0, y0, x1, y1, text))
                width = float(page.rect.width)
                for parts in cluster_visual_row_parts(items):
                    rows.append((width, parts))
    except Exception:
        return []
    return rows


def read_pdf_visual_rows(path: Path) -> list[str]:
    return [" ".join(text for _, text in parts) for _, parts in read_pdf_visual_row_parts(path)]


def get_ocr_engine():
    global OCR_ENGINE
    if OCR_ENGINE is None:
        try:
            from rapidocr import RapidOCR
            OCR_ENGINE = RapidOCR(params={
                "Global.log_level": "error",
                "EngineConfig.onnxruntime.intra_op_num_threads": 2,
                "EngineConfig.onnxruntime.inter_op_num_threads": 1,
            })
        except Exception as exc:
            raise RuntimeError(f"图片 OCR 初始化失败：{exc}") from exc
    return OCR_ENGINE


def image_rgb_frame(source):
    frame = ImageOps.exif_transpose(source.copy())
    if frame.mode in ("RGBA", "LA") or "transparency" in frame.info:
        rgba = frame.convert("RGBA")
        background = Image.new("RGB", rgba.size, "white")
        background.paste(rgba, mask=rgba.getchannel("A"))
        return background
    return frame.convert("RGB")


def ocr_output_text(result) -> str:
    if result.txts is None or result.boxes is None:
        return ""
    entries = []
    for box, text in zip(result.boxes, result.txts):
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        entries.append(((min(ys) + max(ys)) / 2, min(xs), max(max(ys) - min(ys), 1), text))
    rows = []
    # Join nearby OCR boxes into rows so labels stay next to their values.
    for center, left, height, text in sorted(entries):
        if rows and abs(center - rows[-1][0]) <= min(height, rows[-1][1]) * 0.5:
            rows[-1][2].append((left, text))
        else:
            rows.append((center, height, [(left, text)]))
    return normalize_text("\n".join(" ".join(text for _, text in sorted(items)) for _, _, items in rows))


def read_image_text(path: Path) -> tuple[str, int]:
    if Image is None:
        raise RuntimeError("缺少 Pillow，无法读取图片发票")
    with path.open("rb") as source:
        digest = hashlib.file_digest(source, "sha256").hexdigest()
    cache_dir = DATA_DIR / "ocr_cache"
    cache_path = cache_dir / f"{OCR_CACHE_VERSION}_{digest}.json"
    try:
        cached = json.loads(cache_path.read_text(encoding="utf-8"))
        if isinstance(cached.get("text"), str) and isinstance(cached.get("pages"), int) and cached["pages"] > 0:
            return cached["text"], cached["pages"]
    except (OSError, ValueError, AttributeError):
        pass
    parts = []
    with Image.open(path) as source:
        pages = getattr(source, "n_frames", 1)
        for page in range(pages):
            source.seek(page)
            frame = image_rgb_frame(source)
            with OCR_LOCK:
                result = get_ocr_engine()(frame)
            parts.append(ocr_output_text(result))
    text = "\n".join(parts)
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        temporary = cache_path.with_suffix(f".{threading.get_ident()}.tmp")
        temporary.write_text(json.dumps({"text": text, "pages": pages}, ensure_ascii=False), encoding="utf-8")
        temporary.replace(cache_path)
    except OSError:
        pass
    return text, pages


def extract_image_amount_cents(text: str) -> int | None:
    text = normalize_text(text)
    text = re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", text)
    row_cents = amount_cents_from_visual_rows(text.splitlines())
    if row_cents is not None:
        return row_cents
    # OCR may lose the currency symbol. Only allow that fallback beside a total label.
    for label in (r"小\s*写", r"价\s*税\s*合\s*计", r"(?:实收金额|金额合计|总金额)"):
        match = re.search(label + r"[)）:：\s]*[¥Y]?\s*([0-9][0-9, ]*\.[0-9]{2})(?![0-9])", text)
        if match:
            return parse_money_to_cents(match.group(1).replace(" ", ""))
        if re.search(label, text):
            return None
    values = [parse_money_to_cents(value.replace(" ", "")) for value in re.findall(r"¥\s*([0-9][0-9, ]*\.[0-9]{2})(?![0-9])", text)]
    return max((value for value in values if value is not None), default=None)


def parse_image_invoice(path: Path) -> Invoice:
    text, pages = read_image_text(path)
    amount_cents = extract_image_amount_cents(text)
    if amount_cents is None:
        raise ValueError("图片 OCR 未识别到发票金额，请检查图片清晰度及价税合计区域")
    number_match = re.search(r"发\s*票\s*号\s*码[:：\s]*([0-9][0-9 ]{7,40})", text)
    invoice_number = re.sub(r"\s+", "", number_match.group(1)) if number_match else extract_invoice_number(text)
    if not re.fullmatch(r"[0-9]{8,30}", invoice_number):
        invoice_number = ""
    return Invoice(
        path=str(path),
        filename=path.name,
        amount_cents=amount_cents,
        invoice_type=classify_invoice(text),
        invoice_date=extract_invoice_date(text),
        invoice_number=invoice_number,
        seller=extract_seller(text),
        pages=pages,
    )


def xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def iter_xml_elements(element: ET.Element, name: str):
    for item in element.iter():
        if xml_local_name(item.tag) == name:
            yield item


def first_xml_text(element: ET.Element, name: str) -> str:
    for item in iter_xml_elements(element, name):
        return "".join(item.itertext()).strip()
    return ""


def parse_ofd_box(value: str | None) -> tuple[float, float, float, float] | None:
    if not value:
        return None
    try:
        parts = [float(part) for part in re.split(r"\s+", value.strip()) if part]
    except Exception:
        return None
    if len(parts) != 4:
        return None
    return parts[0], parts[1], parts[2], parts[3]


def ofd_zip_join(base_dir: str, relative_path: str) -> str:
    relative_path = (relative_path or "").replace("\\", "/").strip()
    if not relative_path:
        return ""
    if relative_path.startswith("/"):
        return posixpath.normpath(relative_path.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, relative_path))


def read_ofd_xml(zip_file: zipfile.ZipFile, name: str) -> ET.Element:
    normalized = posixpath.normpath(name.replace("\\", "/")).lstrip("./")
    return ET.fromstring(zip_file.read(normalized))


def find_ofd_document(zip_file: zipfile.ZipFile) -> tuple[str, ET.Element]:
    names = set(zip_file.namelist())
    ofd_xml_name = "OFD.xml" if "OFD.xml" in names else next(
        (name for name in names if name.lower().endswith("/ofd.xml")),
        "",
    )
    if ofd_xml_name:
        ofd_root = read_ofd_xml(zip_file, ofd_xml_name)
        doc_root_text = first_xml_text(ofd_root, "DocRoot")
        if doc_root_text:
            doc_path = ofd_zip_join(posixpath.dirname(ofd_xml_name), doc_root_text)
            return doc_path, read_ofd_xml(zip_file, doc_path)

    document_name = next(
        (
            name
            for name in sorted(names)
            if name.lower().endswith("/document.xml") or name.lower() == "document.xml"
        ),
        "",
    )
    if not document_name:
        raise ValueError("OFD 文件缺少 Document.xml")
    return document_name, read_ofd_xml(zip_file, document_name)


def collect_ofd_custom_data(zip_file: zipfile.ZipFile) -> dict[str, str]:
    custom_data: dict[str, str] = {}
    for name in sorted(zip_file.namelist()):
        if not name.lower().endswith(".xml"):
            continue
        try:
            root = read_ofd_xml(zip_file, name)
        except Exception:
            continue
        for item in iter_xml_elements(root, "CustomData"):
            key = (item.attrib.get("Name") or item.attrib.get("name") or "").strip()
            value = "".join(item.itertext()).strip()
            if key and value:
                custom_data[key] = value
    return custom_data


def collect_ofd_text_runs(root: ET.Element) -> list[tuple[float, float, str]]:
    runs: list[tuple[float, float, str]] = []
    for text_object in iter_xml_elements(root, "TextObject"):
        box = parse_ofd_box(text_object.attrib.get("Boundary"))
        if box is None:
            continue
        box_x, box_y, _, _ = box
        for text_code in iter_xml_elements(text_object, "TextCode"):
            text = "".join(text_code.itertext()).strip()
            if not text:
                continue
            x = box_x + safe_float(text_code.attrib.get("X"), 0)
            y = box_y + safe_float(text_code.attrib.get("Y"), 0)
            runs.append((y, x, text))
    return runs


def ofd_visual_rows(runs: Iterable[tuple[float, float, str]]) -> list[str]:
    rows: list[list] = []
    for y, x, text in sorted(runs, key=lambda run: (run[0], run[1])):
        if rows and abs(y - rows[-1][0]) <= 3.0:
            rows[-1][2].append((x, text))
        else:
            rows.append([y, 0.0, [(x, text)]])
    return [" ".join(text for _, text in sorted(row[2], key=lambda part: part[0])) for row in rows]


def read_ofd_text(path: Path) -> tuple[str, int, dict[str, str], list[str], list[str]]:
    try:
        with zipfile.ZipFile(path) as zip_file:
            custom_data = collect_ofd_custom_data(zip_file)
            document_path, document_root = find_ofd_document(zip_file)
            doc_dir = posixpath.dirname(document_path)
            page_paths = [
                ofd_zip_join(doc_dir, page.attrib.get("BaseLoc", ""))
                for page in iter_xml_elements(document_root, "Page")
                if page.attrib.get("BaseLoc")
            ]
            if not page_paths:
                page_paths = sorted(
                    name
                    for name in zip_file.namelist()
                    if "/pages/" in name.lower() and name.lower().endswith("/content.xml")
                )

            body_values: list[str] = []
            visual_rows: list[str] = []
            all_values: list[str] = [f"{key}: {value}" for key, value in custom_data.items()]
            for page_path in page_paths:
                try:
                    page_root = read_ofd_xml(zip_file, page_path)
                except Exception:
                    continue
                runs = collect_ofd_text_runs(page_root)
                page_values = [text for _, _, text in sorted(runs)]
                body_values.extend(page_values)
                visual_rows.extend(ofd_visual_rows(runs))
                all_values.extend(page_values)

            if not all_values:
                for name in sorted(zip_file.namelist()):
                    if not name.lower().endswith(".xml"):
                        continue
                    try:
                        root = read_ofd_xml(zip_file, name)
                    except Exception:
                        continue
                    all_values.extend(text for _, _, text in sorted(collect_ofd_text_runs(root)))

            page_count = max(len(page_paths), 1)
            return (
                normalize_text("\n".join(all_values)),
                page_count,
                custom_data,
                body_values,
                [normalize_text(row) for row in visual_rows],
            )
    except zipfile.BadZipFile as exc:
        raise ValueError("OFD 文件格式无效或已损坏") from exc


def safe_float(value: object, default: float = 0.0) -> float:
    try:
        return float(str(value).strip())
    except Exception:
        return default


UPPERCASE_DIGITS = {
    "零": 0, "〇": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4,
    "伍": 5, "陆": 6, "柒": 7, "捌": 8, "玖": 9,
}
UPPERCASE_UNITS = {"拾": 10, "佰": 100, "仟": 1000}
UPPERCASE_AMOUNT_CHARS = "".join(UPPERCASE_DIGITS) + "拾佰仟万亿元圆角分整正"
UPPERCASE_AMOUNT_RE = re.compile(r"[" + UPPERCASE_AMOUNT_CHARS + r"]{2,}")
TOTAL_AMOUNT_LABEL_RE = re.compile(r"价\s*税\s*合\s*计|小\s*写")
AMOUNT_TOKEN_RE = re.compile(r"(?:[¥￥]\s*)?(\d{1,3}(?:,\d{3})+|\d+)\.(\d{2})(?!\d)")


def parse_uppercase_amount_cents(text: str) -> int | None:
    """Parse a Chinese uppercase amount such as 壹佰伍拾肆圆伍角整."""
    match = UPPERCASE_AMOUNT_RE.search(normalize_text(text))
    if not match:
        return None
    value = match.group(0)
    if "元" not in value and "圆" not in value:
        return None
    total = 0
    section = 0
    number = 0
    jiao = 0
    fen = 0
    for char in value:
        if char in UPPERCASE_DIGITS:
            number = UPPERCASE_DIGITS[char]
        elif char in UPPERCASE_UNITS:
            section += (number or 1) * UPPERCASE_UNITS[char]
            number = 0
        elif char == "万":
            section = (section + number) * 10000
            total += section
            section = 0
            number = 0
        elif char in ("元", "圆"):
            total += section + number
            section = 0
            number = 0
        elif char == "角":
            jiao = number
            number = 0
        elif char == "分":
            fen = number
            number = 0
    return total * 100 + jiao * 10 + fen


def amount_tokens_in_line(line: str) -> list[int]:
    line = re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", line)
    values: list[int] = []
    for match in AMOUNT_TOKEN_RE.finditer(line):
        start, end = match.span()
        if start > 0 and (line[start - 1].isdigit() or line[start - 1] == "."):
            continue
        if end < len(line) and (line[end].isdigit() or line[end] == "."):
            continue
        cents = parse_money_to_cents(match.group(1) + "." + match.group(2))
        if cents is not None:
            values.append(cents)
    return values


def choose_row_amount_cents(line: str, tokens: list[int]) -> int | None:
    if not tokens:
        return None
    if len(tokens) == 1:
        return tokens[0]
    uppercase_cents = parse_uppercase_amount_cents(line)
    if uppercase_cents is not None and uppercase_cents in tokens:
        return uppercase_cents
    # The 价税合计 value is the rightmost amount on its row.
    return tokens[-1]


def amount_cents_from_visual_rows(rows: Iterable[str]) -> int | None:
    """Read 价税合计 from visual rows instead of the raw PDF text stream."""
    lines = [row for row in rows if row and row.strip()]
    if not lines:
        return None
    for line in lines:
        if TOTAL_AMOUNT_LABEL_RE.search(line):
            tokens = amount_tokens_in_line(line)
            if tokens:
                return choose_row_amount_cents(line, tokens)
    for line in lines:
        if UPPERCASE_AMOUNT_RE.search(line):
            tokens = amount_tokens_in_line(line)
            if tokens:
                return choose_row_amount_cents(line, tokens)
    for line in lines:
        if UPPERCASE_AMOUNT_RE.search(line):
            uppercase_cents = parse_uppercase_amount_cents(line)
            if uppercase_cents is not None:
                return uppercase_cents
    return None


def money_values_in(text: str) -> list[int]:
    values: list[int] = []
    for match in re.finditer(r"¥\s*([0-9][0-9,]*(?:\.[0-9]{1,4})?)", text):
        cents = parse_money_to_cents(match.group(1))
        if cents is not None:
            values.append(cents)
    return values


def extract_amount_cents(text: str) -> int | None:
    text = normalize_text(text)
    row_cents = amount_cents_from_visual_rows(text.splitlines())
    if row_cents is not None:
        return row_cents

    candidates: list[int] = []

    for match in re.finditer(r"(?:小写|价税合计)", text):
        window = text[match.start() : match.start() + 500]
        candidates.extend(money_values_in(window))

    if candidates:
        return max(candidates)

    all_values = money_values_in(text)
    if all_values:
        return max(all_values)

    return None


def read_pdf_amount_cents(path: Path, text: str) -> int | None:
    row_cents = amount_cents_from_visual_rows(read_pdf_visual_rows(path))
    if row_cents is not None:
        return row_cents
    return extract_amount_cents(text)


def money_from_custom_value(value: object) -> int | None:
    text = normalize_text(str(value or ""))
    text = text.replace("元", "").strip()
    text = text.lstrip("¥").strip()
    return parse_money_to_cents(text)


def extract_ofd_amount_cents(
    custom_data: dict[str, str],
    text: str,
    visual_rows: Iterable[str] | None = None,
) -> int | None:
    compact_custom = {re.sub(r"\s+", "", key): value for key, value in custom_data.items()}
    for key in ("价税合计", "价税合计小写", "合计价税", "金额合计", "总金额"):
        cents = money_from_custom_value(compact_custom.get(key))
        if cents is not None:
            return cents

    amount_without_tax = money_from_custom_value(compact_custom.get("合计金额"))
    tax_amount = money_from_custom_value(compact_custom.get("合计税额"))
    if amount_without_tax is not None and tax_amount is not None:
        return amount_without_tax + tax_amount
    if amount_without_tax is not None:
        return amount_without_tax
    if visual_rows is not None:
        row_cents = amount_cents_from_visual_rows(visual_rows)
        if row_cents is not None:
            return row_cents
    return extract_amount_cents(text)


def classify_invoice(text: str) -> str:
    compact = re.sub(r"\s+", "", normalize_text(text))
    for category in ("交通", "住宿", "餐饮", "办公"):
        if any(keyword in compact for keyword in TYPE_KEYWORDS[category]):
            return category
    return "其他"


def normalize_invoice_date_text(date_text: str) -> str:
    date_text = normalize_text(date_text).strip()
    match = re.search(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", date_text)
    if match:
        year, month, day = match.groups()
        return f"{year}年{int(month):02d}月{int(day):02d}日"
    match = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", date_text)
    if match:
        year, month, day = match.groups()
        return f"{year}年{int(month):02d}月{int(day):02d}日"
    return date_text


def extract_invoice_date(text: str) -> str:
    text = normalize_text(text)
    match = re.search(r"20\d{2}年\d{1,2}月\d{1,2}日", text)
    if match:
        return normalize_invoice_date_text(match.group(0))
    match = re.search(r"20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}", text)
    return normalize_invoice_date_text(match.group(0)) if match else ""


def invoice_date_sort_key(date_text: str) -> int:
    match = re.search(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", date_text or "")
    if not match:
        match = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", date_text or "")
    if not match:
        return 99999999
    year, month, day = (int(part) for part in match.groups())
    return year * 10000 + month * 100 + day


def invoice_type_sort_key(invoice: Invoice) -> tuple:
    return (
        CATEGORY_SORT_ORDER.get(invoice.invoice_type, len(CATEGORY_SORT_ORDER)),
        invoice_date_sort_key(invoice.invoice_date),
        invoice.amount_cents,
        invoice.filename,
    )


def combination_date_key(invoices: list[Invoice], chosen: tuple[int, ...]) -> tuple[int, ...]:
    return tuple(sorted(invoice_date_sort_key(invoices[index].invoice_date) for index in chosen))


def combination_same_total_key(invoices: list[Invoice], chosen: tuple[int, ...]) -> tuple:
    return (
        len(chosen),
        combination_date_key(invoices, chosen),
        tuple(sorted(invoices[index].filename for index in chosen)),
    )


def lowest_set_bit_index(value: int) -> int:
    return (value & -value).bit_length() - 1


def reconstruct_earliest_combination(
    invoices: list[Invoice],
    amounts: list[int],
    small_invoices: list[bool],
    suffix_reachable: list[tuple[int, ...]],
    total: int,
) -> list[Invoice]:
    chosen: list[Invoice] = []
    remaining = total
    small_count = 0
    for index, invoice in enumerate(invoices):
        amount = amounts[index]
        next_small_count = small_count + int(small_invoices[index])
        if amount > remaining or next_small_count > SMALL_INVOICE_MAX_COUNT:
            continue
        remaining_after_invoice = remaining - amount
        remaining_small_slots = SMALL_INVOICE_MAX_COUNT - next_small_count
        next_reachable = suffix_reachable[index + 1]
        can_complete = any(
            (next_reachable[count] >> remaining_after_invoice) & 1
            for count in range(remaining_small_slots + 1)
        )
        if can_complete:
            chosen.append(invoice)
            remaining = remaining_after_invoice
            small_count = next_small_count
            if remaining == 0:
                break
    return chosen if remaining == 0 else []


def reconstruct_preferred_combination(
    invoices: list[Invoice],
    amounts: list[int],
    small_invoices: list[bool],
    suffix_reachable: list[tuple[int, ...]],
    total: int,
) -> list[Invoice]:
    single_matches = [invoice for invoice in invoices if invoice.amount_cents == total]
    if single_matches:
        return [
            min(
                single_matches,
                key=lambda invoice: (invoice_date_sort_key(invoice.invoice_date), invoice.filename),
            )
        ]
    return reconstruct_earliest_combination(invoices, amounts, small_invoices, suffix_reachable, total)


def combination_result_key(invoices: list[Invoice], total: int, target_cents: int) -> tuple:
    return (
        total - target_cents,
        len(invoices),
        tuple(invoice_date_sort_key(invoice.invoice_date) for invoice in invoices),
        tuple(sorted(invoice.filename for invoice in invoices)),
    )


def trim_redundant_invoices(invoices: list[Invoice], target_cents: int) -> list[Invoice]:
    invoice_list = list(invoices)
    overage = sum(invoice.amount_cents for invoice in invoice_list) - target_cents
    if overage <= 0:
        return invoice_list

    removable: dict[int, tuple[int, ...]] = {0: ()}
    for index, invoice in enumerate(invoice_list):
        amount = invoice.amount_cents
        if amount <= 0 or amount > overage:
            continue
        for subtotal, indices in list(removable.items()):
            next_total = subtotal + amount
            if next_total > overage:
                continue
            candidate = indices + (index,)
            existing = removable.get(next_total)
            if existing is None or len(candidate) > len(existing):
                removable[next_total] = candidate

    removed_total = max(removable)
    if removed_total <= 0:
        return invoice_list

    removed_indices = set(removable[removed_total])
    return [invoice for index, invoice in enumerate(invoice_list) if index not in removed_indices]


def extract_invoice_number(text: str) -> str:
    text = normalize_text(text)
    label_match = re.search(r"发票号码[:：\s]*([0-9]{8,30})", text)
    if label_match:
        return label_match.group(1)
    number_match = re.search(r"\b(2[0-9]{19,23})\b", text)
    return number_match.group(1) if number_match else ""


def extract_ofd_invoice_number(custom_data: dict[str, str], text: str) -> str:
    for key, value in custom_data.items():
        if "发票号码" in re.sub(r"\s+", "", key):
            match = re.search(r"[0-9]{8,30}", value)
            if match:
                return match.group(0)
    return extract_invoice_number(text)


def extract_seller(text: str) -> str:
    compact_lines = [line.strip() for line in normalize_text(text).splitlines() if line.strip()]
    for index, line in enumerate(compact_lines):
        if line.startswith("销售方") or line == "销":
            nearby = compact_lines[index : index + 12]
            for item in nearby:
                match = re.search(r"名称[:：]\s*(.+)", item)
                if match:
                    value = match.group(1).strip()
                    if value and "名称" not in value:
                        return value[:80]
    # Buyer and seller are often joined into one line (`名称: 买方 名称: 卖方`),
    # e.g. in OCR output; the trailing 名称: block is the seller.
    for line in compact_lines:
        labels = list(re.finditer(r"名\s*称\s*[:：]", line))
        if len(labels) < 2:
            continue
        value = line[labels[-1].end() :]
        value = re.split(r"名\s*称\s*[:：]|统一社会信用代码|纳税人识别号", value)[0]
        value = value.strip(" 　:：")
        if value and "名称" not in value:
            return value[:80]
    return ""


SELLER_NAME_LABEL_RE = re.compile(r"名\s*称\s*[:：]")
SELLER_STOP_RE = re.compile(r"统一社会信用代码|纳税人识别号|识别号")


def extract_seller_from_row_parts(
    rows: Iterable[tuple[float, list[tuple[float, str]]]],
) -> str:
    """Read the seller name from the right-hand 名称: block of the invoice header.

    The header draws buyer and seller side by side, each as `名称:` followed by
    the name, so the rightmost `名称:` on the page belongs to the seller.
    """
    for page_width, parts in rows:
        ordered = sorted(parts, key=lambda part: part[0])
        labels = [(x, text) for x, text in ordered if SELLER_NAME_LABEL_RE.search(text)]
        if len(labels) < 2:
            continue
        seller_labels = [(x, text) for x, text in labels if x >= page_width * 0.5]
        if not seller_labels:
            continue
        label_x, label_text = seller_labels[-1]
        pieces = SELLER_NAME_LABEL_RE.split(label_text, 1)
        name = pieces[1].strip() if len(pieces) > 1 else ""
        if not name:
            # The name is the first real entry to the right of the label; label
            # and tax-id captions can share that x, so skip them instead of
            # stopping at the first one.
            for x, text in ordered:
                if x <= label_x:
                    continue
                candidate = text.strip(" 　:：")
                if not candidate or "名称" in candidate or SELLER_STOP_RE.search(candidate):
                    continue
                name = candidate
                break
        name = name.strip(" 　:：")
        if name and "名称" not in name and not SELLER_STOP_RE.search(name):
            return name[:80]
    return ""


def extract_ofd_seller(custom_data: dict[str, str], body_values: list[str], text: str) -> str:
    seller_tax_id = ""
    for key, value in custom_data.items():
        if "销售方纳税人识别号" in re.sub(r"\s+", "", key):
            seller_tax_id = value.strip()
            break
    if seller_tax_id:
        for index, value in enumerate(body_values):
            if seller_tax_id in value:
                for candidate in reversed(body_values[max(0, index - 4) : index]):
                    candidate = candidate.strip(" ：:")
                    if candidate and not re.fullmatch(r"[0-9A-Z]{10,30}", candidate):
                        return candidate[:80]
    return extract_seller(text)


def parse_ofd_invoice(path: Path) -> Invoice:
    text, pages, custom_data, body_values, visual_rows = read_ofd_text(path)
    amount_cents = extract_ofd_amount_cents(custom_data, text, visual_rows)
    if amount_cents is None:
        raise ValueError("未识别到 OFD 发票金额")
    invoice_date = ""
    for key, value in custom_data.items():
        if "开票日期" in re.sub(r"\s+", "", key):
            invoice_date = normalize_invoice_date_text(value)
            break
    if not invoice_date:
        invoice_date = extract_invoice_date(text)
    return Invoice(
        path=str(path),
        filename=path.name,
        amount_cents=amount_cents,
        invoice_type=classify_invoice(text),
        invoice_date=invoice_date,
        invoice_number=extract_ofd_invoice_number(custom_data, text),
        seller=extract_ofd_seller(custom_data, body_values, text),
        pages=pages,
    )


def parse_invoice(path: Path) -> Invoice:
    suffix = path.suffix.lower()
    if suffix in IMAGE_SUFFIXES:
        return parse_image_invoice(path)
    if suffix == OFD_SUFFIX:
        return parse_ofd_invoice(path)
    if suffix != PDF_SUFFIX:
        raise ValueError(f"不支持的发票格式：{path.suffix}")

    text, pages = read_pdf_text(path)
    row_parts = read_pdf_visual_row_parts(path)
    amount_cents = amount_cents_from_visual_rows(
        " ".join(part for _, part in parts) for _, parts in row_parts
    )
    if amount_cents is None:
        amount_cents = extract_amount_cents(text)
    if amount_cents is None:
        raise ValueError("未识别到发票金额")
    seller = extract_seller_from_row_parts(row_parts) or extract_seller(text)
    return Invoice(
        path=str(path),
        filename=path.name,
        amount_cents=amount_cents,
        invoice_type=classify_invoice(text),
        invoice_date=extract_invoice_date(text),
        invoice_number=extract_invoice_number(text),
        seller=seller,
        pages=pages,
    )


def unique_destination(directory: Path, filename: str) -> Path:
    target = directory / filename
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    index = 1
    while True:
        candidate = directory / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def archive_legacy_runtime_file(source: Path) -> None:
    archive_dir = DATA_DIR / "legacy"
    archive_dir.mkdir(parents=True, exist_ok=True)
    target = unique_destination(archive_dir, source.name)
    shutil.move(str(source), str(target))


def merge_json_runtime_file(source: Path, destination: Path) -> bool:
    try:
        source_data = json.loads(source.read_text(encoding="utf-8"))
        destination_data = json.loads(destination.read_text(encoding="utf-8"))
    except Exception:
        return False

    if isinstance(source_data, list) and isinstance(destination_data, list):
        merged = destination_data + source_data
    elif isinstance(source_data, dict) and isinstance(destination_data, dict):
        merged = {**destination_data, **source_data}
    else:
        return False

    destination.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    source.unlink()
    return True


def migrate_legacy_runtime_files() -> None:
    for source, destination in LEGACY_RUNTIME_FILES:
        if not source.exists() or source.resolve() == destination.resolve():
            continue
        try:
            destination.parent.mkdir(parents=True, exist_ok=True)
            if not destination.exists():
                shutil.move(str(source), str(destination))
            elif source.suffix.lower() == ".json" and merge_json_runtime_file(source, destination):
                continue
            else:
                archive_legacy_runtime_file(source)
        except Exception:
            continue


def append_json_record(file_path: Path, record: dict) -> None:
    file_path.parent.mkdir(parents=True, exist_ok=True)
    records: list[dict] = []
    if file_path.exists():
        try:
            records = json.loads(file_path.read_text(encoding="utf-8"))
        except Exception:
            records = []
    records.append(record)
    file_path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def normalize_invoice_number_value(value: object) -> str:
    return str(value or "").strip()


def safe_filename_component(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value.strip())
    cleaned = cleaned.strip(" .")
    if not cleaned:
        return ""
    reserved_names = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        "CONIN$",
        "CONOUT$",
        *(f"COM{index}" for index in "123456789\u00b9\u00b2\u00b3"),
        *(f"LPT{index}" for index in "123456789\u00b9\u00b2\u00b3"),
    }
    if cleaned.partition(".")[0].rstrip(" ").upper() in reserved_names:
        cleaned = f"_{cleaned}"
    return cleaned


def unique_planned_filename(directory: Path, desired_name: str, allocated_names: set[str], reserved_names: set[str]) -> str:
    desired_path = directory / desired_name
    stem = desired_path.stem
    suffix = desired_path.suffix
    for index in range(0, 1000):
        filename = desired_name if index == 0 else f"{stem}_{index}{suffix}"
        key = filename.lower()
        if key not in allocated_names and key not in reserved_names:
            allocated_names.add(key)
            return filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{stem}_{timestamp}{suffix}"
    allocated_names.add(filename.lower())
    return filename


def unique_rename_temp_path(source: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    for index in range(0, 1000):
        candidate = source.with_name(f".autoinvoice_rename_{timestamp}_{index}_{source.name}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法创建临时重命名文件名：{source.name}")


def rename_invoice_files_by_number() -> tuple[int, int, list[str]]:
    ensure_dirs()
    invoice_paths = sorted(
        (
            path
            for path in INVOICE_DIR.iterdir()
            if path.is_file() and path.suffix.lower() in SUPPORTED_INVOICE_SUFFIXES
        ),
        key=invoice_file_sort_key,
    )
    entries: list[tuple[Path, str]] = []
    skipped: list[str] = []
    for path in invoice_paths:
        try:
            invoice = parse_invoice(path)
            invoice_number = normalize_invoice_number_value(invoice.invoice_number)
            if not invoice_number:
                skipped.append(f"{path.name}: 未识别到发票号码")
                continue
            filename_stem = safe_filename_component(invoice_number)
            if not filename_stem:
                skipped.append(f"{path.name}: 发票号码不能作为文件名")
                continue
            entries.append((path, f"{filename_stem}{path.suffix.lower()}"))
        except Exception as exc:
            skipped.append(f"{path.name}: {exc}")

    source_names = {path.name.lower() for path, _desired_name in entries}
    reserved_names = {
        path.name.lower()
        for path in INVOICE_DIR.iterdir()
        if path.is_file() and path.name.lower() not in source_names
    }
    allocated_names: set[str] = set()
    rename_plan: list[tuple[Path, Path]] = []
    unchanged = 0
    for path, desired_name in entries:
        if path.name.lower() == desired_name.lower():
            allocated_names.add(path.name.lower())
            unchanged += 1
            continue
        final_name = unique_planned_filename(INVOICE_DIR, desired_name, allocated_names, reserved_names)
        rename_plan.append((path, INVOICE_DIR / final_name))

    temp_plan: list[tuple[Path, Path, Path]] = []
    try:
        for source, target in rename_plan:
            temp_path = unique_rename_temp_path(source)
            source.rename(temp_path)
            temp_plan.append((temp_path, source, target))
    except Exception:
        for temp_path, original_path, _target in reversed(temp_plan):
            if temp_path.exists() and not original_path.exists():
                try:
                    temp_path.rename(original_path)
                except Exception:
                    pass
        raise

    try:
        for temp_path, _original_path, target in temp_plan:
            temp_path.rename(target)
    except Exception as exc:
        raise RuntimeError(f"重命名过程中断，请检查 invoices 目录中的临时文件：{exc}") from exc

    return len(rename_plan), unchanged, skipped


def int_or_zero(value: object) -> int:
    try:
        return int(value or 0)
    except Exception:
        return 0


def db_connect() -> sqlite3.Connection:
    DATABASE_FILE.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE_FILE)
    connection.row_factory = sqlite3.Row
    return connection


def ensure_database() -> None:
    global DATABASE_INITIALIZED
    if DATABASE_INITIALIZED:
        return
    ensure_dirs()
    with closing(db_connect()) as connection, connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS used_invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL,
                filename TEXT,
                amount_cents INTEGER,
                invoice_type TEXT,
                invoice_date TEXT,
                seller TEXT,
                pages INTEGER,
                source_path TEXT,
                moved_to TEXT,
                target_amount TEXT,
                used_at TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_used_invoices_number
            ON used_invoices(invoice_number)
            WHERE invoice_number <> ''
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS invoice_type_overrides (
                invoice_key TEXT PRIMARY KEY,
                invoice_type TEXT NOT NULL,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        migrate_used_records_json_to_db(connection)
    DATABASE_INITIALIZED = True


def migrate_used_records_json_to_db(connection: sqlite3.Connection) -> None:
    if not STATE_FILE.exists():
        return
    try:
        records = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return
    if not isinstance(records, list):
        return
    for record in records:
        if isinstance(record, dict) and normalize_invoice_number_value(record.get("invoice_number")):
            upsert_used_invoice_record(connection, record)


def upsert_used_invoice_record(connection: sqlite3.Connection, record: dict) -> None:
    invoice_number = normalize_invoice_number_value(record.get("invoice_number"))
    values = {
        "invoice_number": invoice_number,
        "filename": str(record.get("filename") or record.get("source_file") or ""),
        "amount_cents": int_or_zero(record.get("amount_cents")),
        "invoice_type": str(record.get("invoice_type") or ""),
        "invoice_date": str(record.get("invoice_date") or ""),
        "seller": str(record.get("seller") or ""),
        "pages": int_or_zero(record.get("pages")),
        "source_path": str(record.get("path") or record.get("source_path") or ""),
        "moved_to": str(record.get("moved_to") or ""),
        "target_amount": str(record.get("target_amount") or ""),
        "used_at": str(record.get("used_at") or now_text()),
    }
    if invoice_number:
        existing = connection.execute(
            "SELECT id FROM used_invoices WHERE invoice_number = ? LIMIT 1",
            (invoice_number,),
        ).fetchone()
        if existing:
            connection.execute(
                """
                UPDATE used_invoices
                SET filename = ?, amount_cents = ?, invoice_type = ?, invoice_date = ?,
                    seller = ?, pages = ?, source_path = ?, moved_to = ?,
                    target_amount = ?, used_at = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (
                    values["filename"],
                    values["amount_cents"],
                    values["invoice_type"],
                    values["invoice_date"],
                    values["seller"],
                    values["pages"],
                    values["source_path"],
                    values["moved_to"],
                    values["target_amount"],
                    values["used_at"],
                    existing["id"],
                ),
            )
            return
    connection.execute(
        """
        INSERT INTO used_invoices (
            invoice_number, filename, amount_cents, invoice_type, invoice_date,
            seller, pages, source_path, moved_to, target_amount, used_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            values["invoice_number"],
            values["filename"],
            values["amount_cents"],
            values["invoice_type"],
            values["invoice_date"],
            values["seller"],
            values["pages"],
            values["source_path"],
            values["moved_to"],
            values["target_amount"],
            values["used_at"],
        ),
    )


def save_used_invoice_record(record: dict) -> None:
    ensure_database()
    with closing(db_connect()) as connection, connection:
        upsert_used_invoice_record(connection, record)


def load_used_invoice_numbers() -> set[str]:
    ensure_database()
    with closing(db_connect()) as connection, connection:
        rows = connection.execute(
            "SELECT invoice_number FROM used_invoices WHERE invoice_number <> ''"
        ).fetchall()
    return {normalize_invoice_number_value(row["invoice_number"]) for row in rows}


COMBINATION_PDF_PREFIX = "发票组合"


def used_archive_directories() -> list[Path]:
    """Subdirectories of used_invoices that may hold archived invoices."""
    if not USED_DIR.exists():
        return []
    return sorted(
        (path for path in USED_DIR.iterdir() if path.is_dir() and not path.name.startswith(".")),
        key=lambda path: path.name,
        reverse=True,
    )


def archived_invoice_files(directory: Path) -> list[Path]:
    """Invoice files inside one archive directory, without the merged PDF."""
    if not directory.is_dir():
        return []
    return sorted(
        (
            path
            for path in directory.iterdir()
            if path.is_file()
            and path.suffix.lower() in SUPPORTED_INVOICE_SUFFIXES
            and not path.name.startswith("~$")
            and not path.name.startswith(COMBINATION_PDF_PREFIX)
        ),
        key=invoice_file_sort_key,
    )


def delete_used_invoice_records(invoice_numbers: Iterable[str]) -> int:
    numbers = {normalize_invoice_number_value(number) for number in invoice_numbers}
    numbers = {number for number in numbers if number}
    if not numbers:
        return 0
    ensure_database()
    placeholders = ", ".join("?" for _ in numbers)
    with closing(db_connect()) as connection, connection:
        cursor = connection.execute(
            f"DELETE FROM used_invoices WHERE invoice_number IN ({placeholders})",
            tuple(sorted(numbers)),
        )
        return cursor.rowcount


def remove_json_records(file_path: Path, invoice_numbers: Iterable[str]) -> int:
    """Drop records from the legacy JSON so a restart cannot replay them into SQLite."""
    numbers = {normalize_invoice_number_value(number) for number in invoice_numbers}
    numbers = {number for number in numbers if number}
    if not numbers or not file_path.exists():
        return 0
    try:
        records = json.loads(file_path.read_text(encoding="utf-8"))
    except Exception:
        return 0
    if not isinstance(records, list):
        return 0
    kept = [
        record
        for record in records
        if not (
            isinstance(record, dict)
            and normalize_invoice_number_value(record.get("invoice_number")) in numbers
        )
    ]
    removed = len(records) - len(kept)
    if removed:
        file_path.write_text(json.dumps(kept, ensure_ascii=False, indent=2), encoding="utf-8")
    return removed


def used_invoice_numbers_in_archive(directory: Path) -> dict[str, str]:
    """Map archived file name (lowercase) to its recorded invoice number."""
    ensure_database()
    target = directory.resolve()
    with closing(db_connect()) as connection, connection:
        rows = connection.execute(
            "SELECT filename, invoice_number, moved_to FROM used_invoices"
        ).fetchall()
    mapping: dict[str, str] = {}
    for row in rows:
        filename = str(row["filename"] or "").strip()
        moved_to = str(row["moved_to"] or "").strip()
        if not filename or not moved_to:
            continue
        try:
            parent = Path(moved_to).resolve().parent
        except OSError:
            continue
        if parent != target:
            continue
        mapping[filename.casefold()] = normalize_invoice_number_value(row["invoice_number"])
    return mapping


def invoice_number_for_restore(path: Path) -> str:
    stem = normalize_invoice_number_value(path.stem)
    if re.fullmatch(r"[0-9]{8,30}", stem):
        return stem
    try:
        return normalize_invoice_number_value(parse_invoice(path).invoice_number)
    except Exception:
        return ""


def restore_used_archive(directory: Path) -> tuple[list[Path], set[str], list[str]]:
    """Move one archived batch back to invoices/ and clear its used records."""
    if not directory.is_dir():
        raise ValueError(f"目录不存在：{directory}")
    files = archived_invoice_files(directory)
    if not files:
        raise ValueError("该子目录没有可还原的发票文件")
    recorded_numbers = used_invoice_numbers_in_archive(directory)
    ensure_dirs()
    moved: list[Path] = []
    numbers: set[str] = set()
    errors: list[str] = []
    for source in files:
        invoice_number = recorded_numbers.get(source.name.casefold(), "") or invoice_number_for_restore(source)
        try:
            destination = move_invoice(source, INVOICE_DIR)
        except Exception as exc:
            errors.append(f"{source.name}: {exc}")
            continue
        moved.append(destination)
        if invoice_number:
            numbers.add(invoice_number)
    if numbers:
        try:
            delete_used_invoice_records(numbers)
            remove_json_records(STATE_FILE, numbers)
        except Exception as exc:
            errors.append(f"已移动 {len(moved)} 张发票，但清除已使用记录失败：{exc}")
    return moved, numbers, errors


def invoice_type_override_key(invoice: Invoice) -> str:
    invoice_number = normalize_invoice_number_value(invoice.invoice_number)
    if invoice_number:
        return f"number:{invoice_number}"
    # Content identity keeps numberless invoices stable across file renames.
    with Path(invoice.path).open("rb") as source:
        digest = hashlib.file_digest(source, "sha256").hexdigest()
    return f"sha256:{digest}"


def load_invoice_type_overrides() -> dict[str, str]:
    ensure_database()
    with closing(db_connect()) as connection, connection:
        rows = connection.execute("SELECT invoice_key, invoice_type FROM invoice_type_overrides").fetchall()
    return {row["invoice_key"]: row["invoice_type"] for row in rows if row["invoice_type"] in CATEGORIES[1:]}


def save_invoice_type_override(invoice: Invoice, invoice_type: str | None) -> None:
    if invoice_type is not None and invoice_type not in CATEGORIES[1:]:
        raise ValueError("请选择有效的发票类型")
    invoice_key = invoice_type_override_key(invoice)
    ensure_database()
    with closing(db_connect()) as connection, connection:
        if invoice_type is None:
            connection.execute("DELETE FROM invoice_type_overrides WHERE invoice_key = ?", (invoice_key,))
        else:
            connection.execute(
                """
                INSERT INTO invoice_type_overrides (invoice_key, invoice_type) VALUES (?, ?)
                ON CONFLICT(invoice_key) DO UPDATE SET
                    invoice_type = excluded.invoice_type, updated_at = CURRENT_TIMESTAMP
                """,
                (invoice_key, invoice_type),
            )


def load_last_target_amount() -> str:
    if not SETTINGS_FILE.exists():
        return ""
    try:
        settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return ""
    value = settings.get("last_target_amount", "")
    return str(value).strip()


def save_last_target_amount(amount_text: str) -> None:
    try:
        settings = {}
        if SETTINGS_FILE.exists():
            try:
                settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            except Exception:
                settings = {}
        settings["last_target_amount"] = amount_text.strip()
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def move_invoice(path: Path, directory: Path) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    target = unique_destination(directory, path.name)
    return Path(shutil.move(str(path), str(target)))


def move_previously_used_invoice(path: Path) -> tuple[Path, str]:
    USED_DIR.mkdir(parents=True, exist_ok=True)
    VOID_DIR.mkdir(parents=True, exist_ok=True)
    if (USED_DIR / path.name).exists() or any(
        archived.is_file() and archived.name.casefold() == path.name.casefold()
        for archived in USED_DIR.rglob("*")
    ):
        return move_invoice(path, VOID_DIR), "void_invoices"
    return move_invoice(path, USED_DIR), "used_invoices"


def invoice_file_sort_key(path: Path) -> tuple[str, str]:
    return (path.stem.lower(), path.suffix.lower())


def invoice_duplicate_keep_key(invoice: Invoice) -> tuple[int, str]:
    suffix_rank = {PDF_SUFFIX: 0, OFD_SUFFIX: 1}.get(Path(invoice.path).suffix.lower(), 2)
    return suffix_rank, invoice.filename.lower()


def duplicate_record(moved_invoice: Invoice, kept_invoice: Invoice, target: Path) -> dict:
    return {
        "moved_at": now_text(),
        "source_file": moved_invoice.filename,
        "invoice_number": normalize_invoice_number_value(moved_invoice.invoice_number),
        "kept_file": kept_invoice.filename,
        "moved_to": str(target),
        "moved_to_dir": "void_invoices",
        "reason": "发票号码与 invoices 目录内其他文件重复",
    }


def scan_invoice_directory(progress: Callable[[int, int, str], None] | None = None) -> tuple[list[Invoice], list[dict], list[dict], list[dict]]:
    ensure_dirs()
    used_invoice_numbers = load_used_invoice_numbers()
    type_overrides = load_invoice_type_overrides()
    invoices: list[Invoice] = []
    failures: list[dict] = []
    reused: list[dict] = []
    duplicates: list[dict] = []
    seen_by_number: dict[str, Invoice] = {}
    invoice_paths = sorted(
        (
            path
            for path in INVOICE_DIR.iterdir()
            if path.is_file() and path.suffix.lower() in SUPPORTED_INVOICE_SUFFIXES
        ),
        key=invoice_file_sort_key,
    )
    if any(path.suffix.lower() in IMAGE_SUFFIXES for path in invoice_paths):
        with OCR_LOCK:
            get_ocr_engine()
    for index, path in enumerate(invoice_paths, 1):
        if progress is not None:
            progress(index, len(invoice_paths), path.name)
        try:
            invoice = parse_invoice(path)
            invoice_number = normalize_invoice_number_value(invoice.invoice_number)
            if invoice_number and invoice_number in used_invoice_numbers:
                target, target_dir = move_previously_used_invoice(path)
                reused.append(
                    {
                        "moved_at": now_text(),
                        "source_file": path.name,
                        "invoice_number": invoice_number,
                        "moved_to": str(target),
                        "moved_to_dir": target_dir,
                        "reason": "发票号码已在已使用数据库中",
                    }
                )
                continue
            if invoice_number and invoice_number in seen_by_number:
                existing = seen_by_number[invoice_number]
                if invoice_duplicate_keep_key(invoice) < invoice_duplicate_keep_key(existing):
                    existing_path = Path(existing.path)
                    target = move_invoice(existing_path, VOID_DIR)
                    invoices = [item for item in invoices if item.path != existing.path]
                    seen_by_number[invoice_number] = invoice
                    invoices.append(invoice)
                    duplicates.append(duplicate_record(existing, invoice, target))
                else:
                    target = move_invoice(path, VOID_DIR)
                    duplicates.append(duplicate_record(invoice, existing, target))
                continue
            invoices.append(invoice)
            if invoice_number:
                seen_by_number[invoice_number] = invoice
        except Exception as exc:
            target = move_invoice(path, FAILED_DIR)
            record = {
                "failed_at": now_text(),
                "source_file": path.name,
                "moved_to": str(target),
                "reason": str(exc),
            }
            append_json_record(FAILED_RECORD_FILE, record)
            failures.append(record)
    if type_overrides:
        for index, invoice in enumerate(invoices):
            invoice_type = type_overrides.get(invoice_type_override_key(invoice))
            if invoice_type is not None:
                invoices[index] = replace(invoice, invoice_type=invoice_type)
    invoices.sort(key=lambda item: (item.amount_cents, item.filename))
    return invoices, failures, reused, duplicates


def find_chinese_font_file() -> Path | None:
    for font_path in CHINESE_FONT_CANDIDATES:
        if font_path.exists():
            return font_path
    return None


def parse_ofd_color(element: ET.Element, preferred: tuple[str, ...] = ("FillColor", "StrokeColor")) -> tuple[float, float, float]:
    for color_name in preferred:
        for child in element:
            if xml_local_name(child.tag) != color_name:
                continue
            raw_value = child.attrib.get("Value", "")
            try:
                parts = [float(part) for part in raw_value.replace(",", " ").split()]
            except Exception:
                parts = []
            if len(parts) >= 3:
                return (
                    max(min(parts[0] / 255, 1), 0),
                    max(min(parts[1] / 255, 1), 0),
                    max(min(parts[2] / 255, 1), 0),
                )
    return (0, 0, 0)


def ofd_area_box(root: ET.Element, preferred: tuple[str, ...] = ("CropBox", "PhysicalBox", "ApplicationBox")) -> tuple[float, float, float, float] | None:
    for box_name in preferred:
        for item in iter_xml_elements(root, box_name):
            box = parse_ofd_box("".join(item.itertext()).strip())
            if box is not None and box[2] > 0 and box[3] > 0:
                return box
    return None


def ofd_object_extent(roots: Iterable[ET.Element]) -> tuple[float, float]:
    max_right = 0.0
    max_bottom = 0.0
    for root in roots:
        for item in root.iter():
            if not xml_local_name(item.tag).endswith("Object"):
                continue
            box = parse_ofd_box(item.attrib.get("Boundary"))
            if box is None:
                continue
            x, y, width, height = box
            max_right = max(max_right, x + width)
            max_bottom = max(max_bottom, y + height)
    return max_right, max_bottom


def ofd_page_metrics(page_root: ET.Element, template_roots: list[ET.Element]) -> tuple[float, float, float, float]:
    base_box = None
    for template_root in template_roots:
        base_box = ofd_area_box(template_root)
        if base_box is not None:
            break
    if base_box is None:
        base_box = ofd_area_box(page_root)

    if base_box is None:
        width_mm, height_mm = 210.0, 140.0
    else:
        _, _, width_mm, height_mm = base_box
    width_mm = max(width_mm, 60.0)
    height_mm = max(height_mm, 60.0)

    _, max_bottom = ofd_object_extent([*template_roots, page_root])
    y_scale = 1.0
    if max_bottom > height_mm + 1 and max_bottom < height_mm * 1.35:
        y_scale = max((height_mm - 0.5) / max_bottom, 0.75)
    return width_mm, height_mm, 1.0, y_scale


def ofd_rect_from_box(box: tuple[float, float, float, float], x_scale: float, y_scale: float) -> "fitz.Rect":
    x, y, width, height = box
    return fitz.Rect(
        x * x_scale * PT_PER_MM,
        y * y_scale * PT_PER_MM,
        (x + width) * x_scale * PT_PER_MM,
        (y + height) * y_scale * PT_PER_MM,
    )


def build_ofd_resource_map(zip_file: zipfile.ZipFile, document_path: str, document_root: ET.Element) -> dict[str, str]:
    doc_dir = posixpath.dirname(document_path)
    resource_paths: list[str] = []
    for tag_name in ("PublicRes", "DocumentRes"):
        for item in iter_xml_elements(document_root, tag_name):
            resource_path = ofd_zip_join(doc_dir, "".join(item.itertext()).strip())
            if resource_path:
                resource_paths.append(resource_path)

    resources: dict[str, str] = {}
    for resource_path in resource_paths:
        try:
            resource_root = read_ofd_xml(zip_file, resource_path)
        except Exception:
            continue
        resource_dir = posixpath.dirname(resource_path)
        for item in iter_xml_elements(resource_root, "MultiMedia"):
            resource_id = item.attrib.get("ID")
            media_file = first_xml_text(item, "MediaFile")
            if resource_id and media_file:
                resources[resource_id] = ofd_zip_join(resource_dir, media_file)
    return resources


def build_ofd_template_map(zip_file: zipfile.ZipFile, document_path: str, document_root: ET.Element) -> dict[str, ET.Element]:
    doc_dir = posixpath.dirname(document_path)
    templates: dict[str, ET.Element] = {}
    for item in iter_xml_elements(document_root, "TemplatePage"):
        template_id = item.attrib.get("ID")
        base_loc = item.attrib.get("BaseLoc", "")
        if not template_id or not base_loc:
            continue
        try:
            templates[template_id] = read_ofd_xml(zip_file, ofd_zip_join(doc_dir, base_loc))
        except Exception:
            continue
    return templates


def render_ofd_path_object(page, item: ET.Element, x_scale: float, y_scale: float) -> None:
    box = parse_ofd_box(item.attrib.get("Boundary"))
    if box is None:
        return
    base_x, base_y, _, _ = box
    path_data = first_xml_text(item, "AbbreviatedData")
    if not path_data:
        return
    tokens = re.findall(r"[A-Za-z]|[-+]?\d+(?:\.\d+)?", path_data)
    if any(token.isalpha() and token.upper() not in {"M", "L", "Z"} for token in tokens):
        return
    color = parse_ofd_color(item, ("StrokeColor", "FillColor"))
    line_width = max(safe_float(item.attrib.get("LineWidth"), 0.25) * PT_PER_MM, 0.2)
    current: tuple[float, float] | None = None
    start: tuple[float, float] | None = None
    index = 0

    def next_point() -> tuple[float, float] | None:
        nonlocal index
        if index + 1 >= len(tokens):
            return None
        try:
            x = float(tokens[index])
            y = float(tokens[index + 1])
        except Exception:
            return None
        index += 2
        return base_x + x, base_y + y

    def point_to_pt(point: tuple[float, float]) -> tuple[float, float]:
        return point[0] * x_scale * PT_PER_MM, point[1] * y_scale * PT_PER_MM

    while index < len(tokens):
        token = tokens[index]
        index += 1
        command = token.upper()
        if command == "M":
            current = next_point()
            start = current
        elif command == "L":
            next_item = next_point()
            if current is not None and next_item is not None:
                page.draw_line(point_to_pt(current), point_to_pt(next_item), color=color, width=line_width)
            current = next_item
        elif command == "Z":
            if current is not None and start is not None:
                page.draw_line(point_to_pt(current), point_to_pt(start), color=color, width=line_width)
                current = start
        elif command == "C":
            if index + 5 < len(tokens):
                try:
                    current = base_x + float(tokens[index + 4]), base_y + float(tokens[index + 5])
                except Exception:
                    pass
                index += 6
        elif command == "Q":
            if index + 3 < len(tokens):
                try:
                    current = base_x + float(tokens[index + 2]), base_y + float(tokens[index + 3])
                except Exception:
                    pass
                index += 4


def render_ofd_image_object(page, item: ET.Element, zip_file: zipfile.ZipFile, resource_map: dict[str, str], x_scale: float, y_scale: float) -> None:
    box = parse_ofd_box(item.attrib.get("Boundary"))
    resource_id = item.attrib.get("ResourceID")
    media_path = resource_map.get(str(resource_id or ""))
    if box is None or not media_path:
        return
    try:
        stream = zip_file.read(media_path)
        page.insert_image(ofd_rect_from_box(box, x_scale, y_scale), stream=stream, keep_proportion=False)
    except Exception:
        return


def render_ofd_text_object(page, item: ET.Element, font_file: Path | None, x_scale: float, y_scale: float) -> None:
    box = parse_ofd_box(item.attrib.get("Boundary"))
    if box is None:
        return
    box_x, box_y, _, _ = box
    font_size = max(safe_float(item.attrib.get("Size"), 3.0) * PT_PER_MM, 4.0)
    color = parse_ofd_color(item)
    for text_code in iter_xml_elements(item, "TextCode"):
        text = "".join(text_code.itertext())
        if not text:
            continue
        font_kwargs = {"fontname": "china-s"} if re.search(r"[\u4e00-\u9fff]", text) else {"fontname": "helv"}
        x = (box_x + safe_float(text_code.attrib.get("X"), 0)) * x_scale * PT_PER_MM
        y = (box_y + safe_float(text_code.attrib.get("Y"), 0)) * y_scale * PT_PER_MM
        try:
            page.insert_text((x, y), text, fontsize=font_size, color=color, overlay=True, **font_kwargs)
        except Exception:
            if font_file is None:
                continue
            try:
                page.insert_text(
                    (x, y),
                    text,
                    fontsize=font_size,
                    color=color,
                    overlay=True,
                    fontname="cnfont",
                    fontfile=str(font_file),
                )
            except Exception:
                continue


def render_ofd_xml_root(page, root: ET.Element, zip_file: zipfile.ZipFile, resource_map: dict[str, str], font_file: Path | None, x_scale: float, y_scale: float) -> None:
    for item in root.iter():
        name = xml_local_name(item.tag)
        if name == "PathObject":
            render_ofd_path_object(page, item, x_scale, y_scale)
        elif name == "ImageObject":
            render_ofd_image_object(page, item, zip_file, resource_map, x_scale, y_scale)
        elif name == "TextObject":
            render_ofd_text_object(page, item, font_file, x_scale, y_scale)


def extract_embedded_pdf_from_ofd(zip_file: zipfile.ZipFile, output_path: Path) -> bool:
    for name in sorted(zip_file.namelist()):
        if name.lower().endswith(".pdf"):
            output_path.write_bytes(zip_file.read(name))
            return True
    return False


def render_ofd_to_pdf(ofd_path: Path, output_path: Path) -> Path:
    if fitz is None:
        raise RuntimeError("缺少 PyMuPDF，无法转换 OFD")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ofd_path) as zip_file:
        if extract_embedded_pdf_from_ofd(zip_file, output_path):
            return output_path

        document_path, document_root = find_ofd_document(zip_file)
        doc_dir = posixpath.dirname(document_path)
        resource_map = build_ofd_resource_map(zip_file, document_path, document_root)
        template_map = build_ofd_template_map(zip_file, document_path, document_root)
        page_paths = [
            ofd_zip_join(doc_dir, page_item.attrib.get("BaseLoc", ""))
            for page_item in iter_xml_elements(document_root, "Page")
            if page_item.attrib.get("BaseLoc")
        ]
        if not page_paths:
            page_paths = sorted(
                name
                for name in zip_file.namelist()
                if "/pages/" in name.lower() and name.lower().endswith("/content.xml")
            )
        if not page_paths:
            raise ValueError("OFD 文件缺少页面内容")

        font_file = find_chinese_font_file()
        pdf = fitz.open()
        try:
            for page_path in page_paths:
                page_root = read_ofd_xml(zip_file, page_path)
                template_roots = [
                    template_map[template.attrib.get("TemplateID", "")]
                    for template in iter_xml_elements(page_root, "Template")
                    if template.attrib.get("TemplateID", "") in template_map
                ]
                width_mm, height_mm, x_scale, y_scale = ofd_page_metrics(page_root, template_roots)
                pdf_page = pdf.new_page(width=width_mm * PT_PER_MM, height=height_mm * PT_PER_MM)
                for template_root in template_roots:
                    render_ofd_xml_root(pdf_page, template_root, zip_file, resource_map, font_file, x_scale, y_scale)
                render_ofd_xml_root(pdf_page, page_root, zip_file, resource_map, font_file, x_scale, y_scale)
            pdf.save(output_path)
        finally:
            pdf.close()
    return output_path


def ofd_cache_pdf_path(ofd_path: Path) -> Path:
    stat = ofd_path.stat()
    identity = f"{OFD_RENDER_VERSION}|{ofd_path.resolve()}|{stat.st_size}|{stat.st_mtime_ns}"
    digest = hashlib.sha1(identity.encode("utf-8", errors="ignore")).hexdigest()[:12]
    safe_stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", ofd_path.stem).strip(" ._")[:80] or "invoice"
    return OFD_CACHE_DIR / f"{safe_stem}_{digest}.pdf"


def image_cache_pdf_path(path: Path) -> Path:
    with path.open("rb") as source:
        digest = hashlib.file_digest(source, "sha256").hexdigest()
    return PRINT_JOBS_DIR / "image_cache" / f"{IMAGE_RENDER_VERSION}_{digest}.pdf"


def render_image_to_pdf(path: Path, output_path: Path) -> Path:
    if Image is None or fitz is None:
        raise RuntimeError("缺少 Pillow/PyMuPDF，无法预览或打印图片发票")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(path) as source, fitz.open() as pdf:
        for index in range(getattr(source, "n_frames", 1)):
            source.seek(index)
            frame = image_rgb_frame(source)
            width, height = frame.size
            scale = 842 / max(width, height)
            page = pdf.new_page(width=width * scale, height=height * scale)
            pixmap = fitz.Pixmap(fitz.csRGB, width, height, frame.tobytes(), False)
            page.insert_image(page.rect, pixmap=pixmap)
        pdf.save(output_path, deflate=True)
    return output_path


def invoice_pdf_path(invoice: Invoice) -> Path:
    source = Path(invoice.path)
    if source.suffix.lower() == PDF_SUFFIX:
        return source
    if source.suffix.lower() == OFD_SUFFIX:
        target = ofd_cache_pdf_path(source)
        if not target.exists() or target.stat().st_mtime_ns < source.stat().st_mtime_ns:
            render_ofd_to_pdf(source, target)
        return target
    if source.suffix.lower() in IMAGE_SUFFIXES:
        target = image_cache_pdf_path(source)
        if not target.exists():
            render_image_to_pdf(source, target)
        return target
    raise RuntimeError(f"不支持预览或打印的发票格式：{source.suffix}")


def count_pdf_pages(path: Path) -> int:
    try:
        return len(PdfReader(str(path)).pages)
    except Exception:
        if fitz is None:
            return 0
        with fitz.open(path) as doc:
            return len(doc)


def unique_pdf_excel_destination(directory: Path, filename: str) -> Path:
    target = directory / filename
    index = 1
    while target.exists() or target.with_suffix(".xlsx").exists():
        target = directory / f"{Path(filename).stem}_{index}.pdf"
        index += 1
    return target


def merge_invoices_to_pdf(invoices: Iterable[Invoice]) -> Path:
    invoice_list = list(invoices)
    if not invoice_list:
        raise RuntimeError("没有可合并的发票")

    ensure_dirs()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = unique_pdf_excel_destination(PRINT_JOBS_DIR, f"发票组合_{timestamp}.pdf")
    writer = PdfWriter()

    for invoice in invoice_list:
        source = invoice_pdf_path(invoice)
        if not source.exists():
            raise FileNotFoundError(f"{invoice.filename} 不存在，无法合并")
        reader = PdfReader(str(source))
        for page in reader.pages:
            writer.add_page(page)

    try:
        with output_path.open("wb") as output:
            writer.write(output)
    except Exception:
        with suppress(OSError):
            output_path.unlink(missing_ok=True)
        raise
    return output_path


def export_combination_excel(
    invoices: Iterable[Invoice],
    output_path: Path,
    target_cents: int | None = None,
    page_counts: Iterable[int] | None = None,
) -> Path:
    invoice_list = list(invoices)
    if not invoice_list:
        raise ValueError("没有可导出的发票")
    counts = list(page_counts) if page_counts is not None else [
        count_pdf_pages(invoice_pdf_path(invoice)) for invoice in invoice_list
    ]
    if len(counts) != len(invoice_list) or any(count <= 0 for count in counts):
        raise ValueError("发票页数无效，无法生成核对表")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    total_cents = sum(invoice.amount_cents for invoice in invoice_list)
    total = total_cents / 100
    last_row = 7 + len(invoice_list)
    with xlsxwriter.Workbook(str(output_path), {"strings_to_formulas": False, "strings_to_urls": False}) as workbook:
        workbook.set_properties({"title": "发票组合核对表", "author": "AutoInvoice"})
        sheet = workbook.add_worksheet("发票核对")
        body = {"font_name": "Microsoft YaHei", "font_size": 10, "valign": "vcenter"}
        text_format = workbook.add_format(body)
        number_format = workbook.add_format({**body, "num_format": "#,##0.00", "align": "right"})
        id_format = workbook.add_format({**body, "num_format": "@", "align": "left", "quote_prefix": True})
        center_format = workbook.add_format({**body, "align": "center"})
        date_format = workbook.add_format({**body, "num_format": "yyyy-mm-dd", "align": "center"})
        timestamp_format = workbook.add_format({**body, "num_format": "yyyy-mm-dd hh:mm", "align": "left"})
        title_format = workbook.add_format({**body, "font_size": 14, "bold": True})
        label_format = workbook.add_format({**body, "bold": True})
        spaced_label_format = workbook.add_format({**body, "bold": True, "indent": 1})
        header_format = workbook.add_format({**body, "bold": True, "font_color": "#FFFFFF", "bg_color": "#374151", "align": "center"})
        file_format = workbook.add_format({**body, "text_wrap": True})
        summary_format = workbook.add_format({**body, "num_format": "#,##0.00", "bold": True, "align": "right"})
        difference_format = workbook.add_format({
            **body, "bold": True, "align": "left", "indent": 1,
            "num_format": '"超出："#,##0.00" 元";"不足："#,##0.00" 元";"金额刚好";@',
        })
        link_format = workbook.add_format({**body, "font_color": "#1565C0", "underline": True})
        sheet.hide_gridlines(2)
        sheet.set_default_row(23)
        sheet.set_column("A:A", 8, text_format)
        sheet.set_column("B:B", 19, text_format)
        sheet.set_column("C:C", 29, id_format)
        sheet.set_column("D:D", 16, number_format)
        sheet.set_column("E:E", 18, date_format)
        sheet.set_column("F:F", 32, text_format)
        sheet.set_column("G:G", 45, file_format)
        sheet.set_column("H:H", 14, text_format)
        sheet.set_row(1, 28)
        sheet.write_string("A2", "发票组合核对表", title_format)
        sheet.write_string("A3", "生成时间", label_format)
        sheet.write_datetime("B3", datetime.now(), timestamp_format)
        sheet.write_string("C3", "组合 PDF", spaced_label_format)
        sheet.write_url("D3", "external:" + output_path.with_suffix(".pdf").name, link_format, string="打开组合 PDF")
        sheet.write_string("A4", "报销金额", label_format)
        if target_cents is not None:
            sheet.write_number("B4", target_cents / 100, number_format)
        sheet.write_string("C4", "组合金额（元）", spaced_label_format)
        sheet.write_formula("D4", f"=ROUND(SUM(D8:D{last_row}),2)", summary_format, total)
        difference = (total_cents - target_cents) / 100 if target_cents is not None else ""
        sheet.merge_range("E4:F4", "", difference_format)
        sheet.write_formula("E4", '=IF(ISNUMBER(B4),ROUND(D4-B4,2),"")', difference_format, difference)
        sheet.write_string("A5", "发票张数", label_format)
        sheet.write_formula("B5", f"=COUNT(A8:A{last_row})", label_format, len(invoice_list))
        columns = ["序号", "类型", "发票号码", "金额（元）", "开票日期", "销售方", "文件名", "预览页码"]
        sheet.add_table(6, 0, last_row - 1, 7, {
            "name": "InvoiceDetails",
            "style": "Table Style Light 9",
            "columns": [{"header": name, "header_format": header_format} for name in columns],
        })
        next_page = 1
        for index, (invoice, pages) in enumerate(zip(invoice_list, counts), 7):
            sheet.write_number(index, 0, index - 6, center_format)
            sheet.write_string(index, 1, invoice.invoice_type, center_format)
            sheet.write_string(index, 2, invoice.invoice_number, id_format)
            sheet.write_number(index, 3, invoice.amount_cents / 100, number_format)
            date_key = invoice_date_sort_key(invoice.invoice_date)
            try:
                date = datetime.strptime(str(date_key), "%Y%m%d")
            except ValueError:
                sheet.write_string(index, 4, invoice.invoice_date, text_format)
            else:
                sheet.write_datetime(index, 4, date, date_format)
            sheet.write_string(index, 5, invoice.seller, file_format)
            sheet.write_string(index, 6, invoice.filename, file_format)
            end_page = next_page + pages - 1
            page_range = str(next_page) if pages == 1 else f"{next_page}-{end_page}"
            sheet.write_string(index, 7, page_range, center_format)
            lines = max(
                1,
                *(max(1, (sum(2 if unicodedata.east_asian_width(char) in ("W", "F") else 1 for char in value) + width - 1) // width)
                  for value, width in ((invoice.seller, 32), (invoice.filename, 45))),
            )
            sheet.set_row(index, max(23, lines * 15 + 5))
            next_page = end_page + 1
        sheet.freeze_panes(7, 3)
        sheet.set_landscape()
        sheet.set_paper(9)
        sheet.fit_to_pages(1, 0)
        sheet.repeat_rows(6)
        sheet.print_area(1, 0, last_row - 1, 7)
        sheet.set_margins(0.25, 0.25, 0.4, 0.4)
    return output_path


def archive_combination_documents(
    invoices: list[Invoice],
    directory: Path,
    target_cents: int | None,
    preview_pdf: Path | None = None,
) -> tuple[Path, Path]:
    if not invoices:
        raise ValueError("没有可归档的发票")
    if preview_pdf is None or not preview_pdf.is_file() or not preview_pdf.with_suffix(".xlsx").is_file():
        ordered = sorted(invoices, key=invoice_type_sort_key)
        preview_pdf = merge_invoices_to_pdf(ordered)
        try:
            export_combination_excel(ordered, preview_pdf.with_suffix(".xlsx"), target_cents)
        except Exception:
            for path in (preview_pdf, preview_pdf.with_suffix(".xlsx")):
                with suppress(OSError):
                    path.unlink(missing_ok=True)
            raise

    directory.mkdir(parents=True, exist_ok=True)
    destination_pdf = unique_destination(directory, preview_pdf.name)
    destination_excel = unique_destination(directory, f"{directory.name}.xlsx")
    source_excel = preview_pdf.with_suffix(".xlsx")
    try:
        shutil.copy2(preview_pdf, destination_pdf)
        if destination_pdf.name == preview_pdf.name:
            shutil.copy2(source_excel, destination_excel)
        else:
            # Preserve workbook edits while retargeting its relative PDF hyperlink.
            with zipfile.ZipFile(source_excel) as source, zipfile.ZipFile(destination_excel, "w") as output:
                for entry in source.infolist():
                    data = source.read(entry.filename)
                    if entry.filename.endswith(".rels"):
                        relationships = ET.fromstring(data)
                        changed = False
                        for relationship in relationships:
                            if relationship.get("Type", "").endswith("/hyperlink") and relationship.get("Target") == preview_pdf.name:
                                relationship.set("Target", destination_pdf.name)
                                changed = True
                        if changed:
                            data = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
                    output.writestr(entry, data)
    except Exception:
        for path in (destination_pdf, destination_excel):
            with suppress(OSError):
                path.unlink(missing_ok=True)
        raise
    return destination_pdf, destination_excel


def find_best_combination(invoices: list[Invoice], target_cents: int) -> list[Invoice]:
    if target_cents <= 0:
        return []
    available = sorted(
        (invoice for invoice in invoices if invoice.amount_cents > 0),
        key=lambda invoice: (invoice_date_sort_key(invoice.invoice_date), invoice.amount_cents, invoice.filename),
    )
    if not available:
        return []

    amounts = [invoice.amount_cents for invoice in available]
    small_invoices = [amount < SMALL_INVOICE_THRESHOLD_CENTS for amount in amounts]
    total_amount = sum(amounts)
    if total_amount < target_cents:
        return []

    limit = min(total_amount, target_cents + max(amounts))
    suffix_mask = (1 << (limit + 1)) - 1
    suffix_reachable: list[tuple[int, ...]] = [(0,) * (SMALL_INVOICE_MAX_COUNT + 1)] * (len(available) + 1)
    suffix_reachable[-1] = (1,) + (0,) * SMALL_INVOICE_MAX_COUNT

    for index in range(len(available) - 1, -1, -1):
        amount = amounts[index]
        small_cost = int(small_invoices[index])
        next_reachable = suffix_reachable[index + 1]
        current = list(next_reachable)
        if amount <= limit:
            for count in range(small_cost, SMALL_INVOICE_MAX_COUNT + 1):
                current[count] |= next_reachable[count - small_cost] << amount
        suffix_reachable[index] = tuple(bits & suffix_mask for bits in current)

    over_bits = 0
    for bits in suffix_reachable[0]:
        over_bits |= bits >> target_cents
    if over_bits == 0:
        return []

    min_over = lowest_set_bit_index(over_bits)
    min_total = target_cents + min_over
    eligible_max = min(limit, min_total + DATE_PRIORITY_WINDOW_CENTS)
    min_total_combo = reconstruct_preferred_combination(
        available, amounts, small_invoices, suffix_reachable, min_total
    )
    min_total_combo = trim_redundant_invoices(min_total_combo, target_cents)
    max_allowed_count = len(min_total_combo) if min_total_combo else None

    best_combo: list[Invoice] = []
    best_key: tuple | None = None
    for total in range(min_total, eligible_max + 1):
        if not any((bits >> total) & 1 for bits in suffix_reachable[0]):
            continue
        combo = reconstruct_preferred_combination(available, amounts, small_invoices, suffix_reachable, total)
        if not combo:
            continue
        combo = trim_redundant_invoices(combo, target_cents)
        combo_total = sum(invoice.amount_cents for invoice in combo)
        if combo_total < target_cents:
            continue
        if max_allowed_count is not None and len(combo) > max_allowed_count:
            continue
        key = combination_result_key(combo, combo_total, target_cents)
        if best_key is None or key < best_key:
            best_key = key
            best_combo = combo

    return trim_redundant_invoices(best_combo, target_cents)


def available_printers() -> list[str]:
    if win32print is None:
        return []
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        printers = win32print.EnumPrinters(flags)
        names = sorted({printer[2] for printer in printers if len(printer) > 2 and printer[2]})
        return names
    except Exception:
        return []


def default_printer() -> str:
    if win32print is None:
        return ""
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


PD_PAGENUMS = 0x00000002
PD_NOSELECTION = 0x00000004
PD_RETURNDC = 0x00000100
PD_USEDEVMODECOPIESANDCOLLATE = 0x00040000
PD_DISABLEPRINTTOFILE = 0x00080000
PD_HIDEPRINTTOFILE = 0x00100000


class PRINTDLGW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hDevMode", wintypes.HANDLE),
        ("hDevNames", wintypes.HANDLE),
        ("hDC", wintypes.HDC),
        ("Flags", wintypes.DWORD),
        ("nFromPage", wintypes.WORD),
        ("nToPage", wintypes.WORD),
        ("nMinPage", wintypes.WORD),
        ("nMaxPage", wintypes.WORD),
        ("nCopies", wintypes.WORD),
        ("hInstance", wintypes.HANDLE),
        ("lCustData", wintypes.LPARAM),
        ("lpfnPrintHook", wintypes.LPVOID),
        ("lpfnSetupHook", wintypes.LPVOID),
        ("lpPrintTemplateName", wintypes.LPCWSTR),
        ("lpSetupTemplateName", wintypes.LPCWSTR),
        ("hPrintTemplate", wintypes.HANDLE),
        ("hSetupTemplate", wintypes.HANDLE),
    ]


def show_windows_print_dialog(owner_hwnd: int, page_count: int) -> PrintDialogSelection | None:
    if win32ui is None:
        raise RuntimeError("缺少 pywin32，无法调用 Windows 打印接口")

    max_page = max(min(page_count, 65535), 1)
    dialog = PRINTDLGW()
    dialog.lStructSize = ctypes.sizeof(PRINTDLGW)
    dialog.hwndOwner = owner_hwnd
    dialog.Flags = (
        PD_RETURNDC
        | PD_NOSELECTION
        | PD_HIDEPRINTTOFILE
        | PD_DISABLEPRINTTOFILE
        | PD_USEDEVMODECOPIESANDCOLLATE
    )
    dialog.nFromPage = 1
    dialog.nToPage = max_page
    dialog.nMinPage = 1
    dialog.nMaxPage = max_page
    dialog.nCopies = 1

    result = ctypes.windll.comdlg32.PrintDlgW(ctypes.byref(dialog))
    if not result:
        error_code = ctypes.windll.comdlg32.CommDlgExtendedError()
        if error_code == 0:
            return None
        raise RuntimeError(f"Windows 打印设置窗口打开失败，错误码：{error_code}")

    from_page = to_page = None
    if dialog.Flags & PD_PAGENUMS:
        from_page = max(int(dialog.nFromPage), 1)
        to_page = max(int(dialog.nToPage), from_page)

    return PrintDialogSelection(
        hdc=int(dialog.hDC or 0),
        hdevmode=int(dialog.hDevMode or 0) or None,
        hdevnames=int(dialog.hDevNames or 0) or None,
        from_page=from_page,
        to_page=to_page,
    )


def release_print_dialog_selection(selection: PrintDialogSelection | None) -> None:
    if selection is None:
        return
    for handle in (selection.hdevmode, selection.hdevnames):
        if handle:
            try:
                ctypes.windll.kernel32.GlobalFree(wintypes.HGLOBAL(handle))
            except Exception:
                pass


def create_printer_dc(printer_name: str, duplex_label: str):
    if win32ui is None:
        raise RuntimeError("缺少 pywin32，无法调用 Windows 打印接口")

    duplex_value = {
        "单面": getattr(win32con, "DMDUP_SIMPLEX", 1),
        "双面长边": getattr(win32con, "DMDUP_VERTICAL", 2),
        "双面短边": getattr(win32con, "DMDUP_HORIZONTAL", 3),
    }.get(duplex_label, getattr(win32con, "DMDUP_SIMPLEX", 1))

    if win32print is not None and win32gui is not None:
        handle = None
        try:
            handle = win32print.OpenPrinter(printer_name)
            devmode = win32print.GetPrinter(handle, 2).get("pDevMode")
            if devmode is not None and duplex_value:
                devmode.Duplex = duplex_value
                devmode.Fields = devmode.Fields | getattr(win32con, "DM_DUPLEX", 0x1000)
            hdc = win32gui.CreateDC("WINSPOOL", printer_name, None, devmode)
            return win32ui.CreateDCFromHandle(hdc)
        except Exception:
            if handle is not None:
                win32print.ClosePrinter(handle)
            dc = win32ui.CreateDC()
            dc.CreatePrinterDC(printer_name)
            return dc
        finally:
            if handle is not None:
                try:
                    win32print.ClosePrinter(handle)
                except Exception:
                    pass

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer_name)
    return dc


def best_print_rotation(
    page_width: float,
    page_height: float,
    printable_width: int,
    printable_height: int,
) -> tuple[int, float]:
    normal_scale = min(printable_width / page_width, printable_height / page_height)
    rotated_scale = min(printable_width / page_height, printable_height / page_width)
    if rotated_scale > normal_scale:
        return 90, rotated_scale
    return 0, normal_scale


def print_pdf_to_dc(
    path: Path,
    dc,
    job_name: str | None = None,
    from_page: int | None = None,
    to_page: int | None = None,
) -> None:
    if fitz is None or Image is None or ImageWin is None:
        raise RuntimeError("缺少 PyMuPDF 或 Pillow，无法渲染 PDF 打印")

    printable_width = dc.GetDeviceCaps(getattr(win32con, "HORZRES", 8))
    printable_height = dc.GetDeviceCaps(getattr(win32con, "VERTRES", 10))
    if printable_width <= 0 or printable_height <= 0:
        raise RuntimeError("无法读取打印机可打印区域")

    dc.StartDoc(job_name or path.name)
    try:
        with fitz.open(path) as doc:
            first_index = max((from_page or 1) - 1, 0)
            last_index = min((to_page or len(doc)) - 1, len(doc) - 1)
            for page_index in range(first_index, last_index + 1):
                page = doc[page_index]
                page_width = max(float(page.rect.width), 1.0)
                page_height = max(float(page.rect.height), 1.0)
                rotation, scale = best_print_rotation(page_width, page_height, printable_width, printable_height)
                scale = max(scale, 1.0)
                matrix = fitz.Matrix(scale, scale)
                if rotation:
                    matrix = matrix.prerotate(rotation)
                pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)

                draw_width = min(printable_width, image.width)
                draw_height = min(printable_height, int(image.height * draw_width / image.width))
                if draw_height > printable_height:
                    draw_height = printable_height
                    draw_width = min(printable_width, int(image.width * draw_height / image.height))
                left = max((printable_width - draw_width) // 2, 0)
                top = max((printable_height - draw_height) // 2, 0)

                if image.width != draw_width or image.height != draw_height:
                    image = image.resize((draw_width, draw_height), Image.Resampling.LANCZOS)

                dc.StartPage()
                ImageWin.Dib(image).draw(dc.GetHandleOutput(), (left, top, left + draw_width, top + draw_height))
                dc.EndPage()
    finally:
        dc.EndDoc()


def print_pdf_file(path: Path, printer_name: str, duplex_label: str) -> None:
    dc = create_printer_dc(printer_name, duplex_label)
    try:
        print_pdf_to_dc(path, dc, job_name=path.name)
    finally:
        try:
            dc.DeleteDC()
        except Exception:
            pass


def print_pdf_with_dialog_selection(path: Path, selection: PrintDialogSelection) -> None:
    if not selection.hdc:
        raise RuntimeError("Windows 打印设置窗口没有返回有效打印机")
    dc = win32ui.CreateDCFromHandle(selection.hdc)
    try:
        print_pdf_to_dc(
            path,
            dc,
            job_name=path.name,
            from_page=selection.from_page,
            to_page=selection.to_page,
        )
    finally:
        try:
            dc.DeleteDC()
        except Exception:
            pass


def print_invoices(invoices: Iterable[Invoice], printer_name: str, copies: int, duplex_label: str) -> None:
    if not printer_name:
        raise RuntimeError("请选择打印机")
    copies = max(int(copies), 1)
    invoice_list = list(invoices)
    if not invoice_list:
        raise RuntimeError("没有可打印的发票")
    for _ in range(copies):
        for invoice in invoice_list:
            print_pdf_file(invoice_pdf_path(invoice), printer_name, duplex_label)


class AutoInvoiceApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        ensure_dirs()
        self.title(f"AutoInvoice 发票报销助手 {APP_VERSION}")
        self.geometry("1180x760")
        self.minsize(1020, 640)

        self.invoices: list[Invoice] = []
        self.result_invoices: list[Invoice] = []
        self.preview_invoice: Invoice | None = None
        self.print_preview_pdf: Path | None = None
        self.print_preview_invoices: list[Invoice] = []
        self.preview_page = 0
        self.preview_zoom = 1.0
        self.preview_photo = None
        self.calculation_running = False
        self.scanning_running = False
        self.calculate_button: ttk.Button | None = None

        self.target_var = tk.StringVar(value=load_last_target_amount())
        self.content_var = tk.StringVar(value="报销")
        self.category_vars = {
            category: tk.BooleanVar(value=category == CATEGORY_ALL)
            for category in CATEGORIES
        }
        self.invoice_type_var = tk.StringVar()
        self.status_var = tk.StringVar(value="准备就绪")
        self.summary_var = tk.StringVar(value="未计算")
        self.available_summary_var = tk.StringVar(value="共 0 张，合计 0.00 元")

        self.configure(bg="#f6f7f9")
        self.create_widgets()
        self.scan()

    def create_widgets(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f6f7f9")
        style.configure("Panel.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabel", background="#f6f7f9", foreground="#1f2937", font=("Microsoft YaHei UI", 9))
        style.configure("Panel.TLabel", background="#ffffff")
        style.configure("Summary.TLabel", font=("Microsoft YaHei UI", 10, "bold"), background="#ffffff")
        style.configure("TButton", font=("Microsoft YaHei UI", 9))
        style.configure("TMenubutton", font=("Microsoft YaHei UI", 9))
        style.configure("TCheckbutton", background="#f6f7f9", font=("Microsoft YaHei UI", 9))
        style.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=26)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))

        controls = ttk.Frame(self, padding=(14, 12, 14, 10))
        controls.pack(fill=tk.X)
        actions = ttk.Frame(controls)
        actions.pack(fill=tk.X)

        ttk.Label(actions, text="金额").pack(side=tk.LEFT)
        amount_entry = ttk.Entry(actions, textvariable=self.target_var, width=9)
        amount_entry.pack(side=tk.LEFT, padx=(6, 16))
        amount_entry.bind("<Return>", lambda _event: self.calculate())

        ttk.Label(actions, text="内容：").pack(side=tk.LEFT)
        ttk.Entry(actions, textvariable=self.content_var, width=16).pack(side=tk.LEFT, padx=(6, 16))

        self.calculate_button = ttk.Button(actions, text="计算组合", command=self.calculate)
        self.calculate_button.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(actions, text="扫描发票", command=self.scan).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(actions, text="发票目录", command=lambda: os.startfile(INVOICE_DIR)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(actions, text="自动文件名", command=self.rename_files_by_invoice_number).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(actions, text="还原已使用", command=self.restore_used_invoices).pack(side=tk.LEFT, padx=(0, 8))

        category_controls = ttk.Frame(controls)
        category_controls.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(category_controls, text="发票类型").pack(side=tk.LEFT, padx=(0, 6))
        for category, variable in self.category_vars.items():
            ttk.Checkbutton(
                category_controls,
                text=category,
                variable=variable,
                command=lambda selected=category: self.on_category_changed(selected),
            ).pack(side=tk.LEFT, padx=(0, 12))

        main = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 10))

        left = ttk.Frame(main)
        right = ttk.Frame(main)
        main.add(left, weight=3)
        main.add(right, weight=2)
        left.rowconfigure(0, weight=1, uniform="invoice_lists")
        left.rowconfigure(1, weight=1, uniform="invoice_lists")
        left.columnconfigure(0, weight=1)

        invoice_panel = ttk.Frame(left, style="Panel.TFrame", padding=10)
        invoice_panel.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        invoice_panel.rowconfigure(1, weight=1)
        invoice_panel.columnconfigure(0, weight=1)
        invoice_header = ttk.Frame(invoice_panel, style="Panel.TFrame")
        invoice_header.grid(row=0, column=0, sticky="ew")
        ttk.Label(invoice_header, text="可用发票", style="Summary.TLabel").pack(side=tk.LEFT)
        type_button = ttk.Menubutton(invoice_header, text="设置类型")
        type_button.pack(side=tk.LEFT, padx=(12, 0))
        self.invoice_type_menu = tk.Menu(type_button, tearoff=False, postcommand=self.prepare_invoice_type_menu)
        self.invoice_context_menu = tk.Menu(self, tearoff=False, postcommand=self.prepare_invoice_type_menu)
        for menu in (self.invoice_type_menu, self.invoice_context_menu):
            for category in CATEGORIES[1:]:
                menu.add_radiobutton(
                    label=category,
                    value=category,
                    variable=self.invoice_type_var,
                    command=lambda selected=category: self.set_selected_invoice_type(selected),
                )
            menu.add_separator()
            menu.add_command(label="恢复自动识别", command=lambda: self.set_selected_invoice_type(None))
        self.invoice_context_menu.add_separator()
        self.invoice_context_menu.add_command(label="删除", command=self.delete_selected_invoice)
        type_button.configure(menu=self.invoice_type_menu)
        ttk.Label(invoice_header, textvariable=self.available_summary_var, style="Summary.TLabel").pack(side=tk.RIGHT)
        self.invoice_tree = self.create_tree(invoice_panel, show_status=True)
        self.invoice_tree.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        self.invoice_tree.bind("<<TreeviewSelect>>", lambda _event: self.preview_from_tree(self.invoice_tree, self.invoices))
        self.invoice_tree.bind("<Double-1>", self.add_available_invoice_to_result)
        self.invoice_tree.bind("<Button-3>", self.show_invoice_type_menu)

        result_panel = ttk.Frame(left, style="Panel.TFrame", padding=10)
        result_panel.grid(row=1, column=0, sticky="nsew")
        result_panel.rowconfigure(1, weight=1)
        result_panel.columnconfigure(0, weight=1)
        result_header = ttk.Frame(result_panel, style="Panel.TFrame")
        result_header.grid(row=0, column=0, sticky="ew")
        ttk.Label(result_header, text="推荐组合", style="Summary.TLabel").pack(side=tk.LEFT)
        ttk.Label(result_header, textvariable=self.summary_var, style="Summary.TLabel").pack(side=tk.RIGHT)
        self.result_tree = self.create_tree(result_panel, selectmode="extended")
        self.result_tree.grid(row=1, column=0, sticky="nsew", pady=(8, 8))
        self.result_tree.bind("<<TreeviewSelect>>", lambda _event: self.preview_from_tree(self.result_tree, self.result_invoices))
        self.result_tree.bind("<Double-1>", self.remove_result_invoice_from_result)

        result_actions = ttk.Frame(result_panel, style="Panel.TFrame")
        result_actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(result_actions, text="标记已使用并移动", command=self.mark_selected_used).pack(side=tk.LEFT)
        ttk.Button(result_actions, text="全部标记使用并移动", command=self.mark_all_used).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(result_actions, text="清空组合", command=self.clear_result).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(result_actions, text="生成组合预览", command=self.generate_combination_preview).pack(side=tk.LEFT, padx=(8, 0))

        preview_panel = ttk.Frame(right, style="Panel.TFrame", padding=10)
        preview_panel.pack(fill=tk.BOTH, expand=True)
        preview_header = ttk.Frame(preview_panel, style="Panel.TFrame")
        preview_header.pack(fill=tk.X)
        ttk.Label(preview_header, text="PDF 预览", style="Summary.TLabel").pack(side=tk.LEFT)
        ttk.Button(preview_header, text="打印", command=self.print_preview_pdf_file).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(preview_header, text="下一页", command=lambda: self.change_preview_page(1)).pack(side=tk.RIGHT)
        ttk.Button(preview_header, text="上一页", command=lambda: self.change_preview_page(-1)).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(preview_header, text="放大", command=lambda: self.change_zoom(0.15)).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(preview_header, text="缩小", command=lambda: self.change_zoom(-0.15)).pack(side=tk.RIGHT)

        self.preview_label_var = tk.StringVar(value="选择一张发票预览")
        ttk.Label(preview_panel, textvariable=self.preview_label_var, style="Panel.TLabel").pack(anchor=tk.W, pady=(8, 6))
        self.preview_canvas = tk.Canvas(preview_panel, bg="#e5e7eb", highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", lambda _event: self.render_preview())

    def create_tree(self, parent, selectmode: str = "browse", show_status: bool = False) -> ttk.Treeview:
        columns = ("status", "amount", "type", "date", "number") if show_status else ("amount", "type", "date", "number")
        tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode=selectmode)
        if show_status:
            tree.heading("status", text="状态")
            tree.column("status", width=48, anchor=tk.CENTER, stretch=False)
        tree.heading("amount", text="金额")
        tree.heading("type", text="类型")
        tree.heading("date", text="日期")
        tree.heading("number", text="号码")
        tree.column("amount", width=80, anchor=tk.E)
        tree.column("type", width=70, anchor=tk.CENTER)
        tree.column("date", width=110, anchor=tk.CENTER)
        tree.column("number", width=190, anchor=tk.W)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=(8, 0))
        return tree

    def set_status(self, message: str) -> None:
        self.status_var.set(message)
        self.update_idletasks()

    def scan(self) -> None:
        if self.scanning_running or self.calculation_running:
            return
        self.scanning_running = True
        self.invoices = []
        self.clear_result()
        self.available_summary_var.set("正在扫描...")
        if self.calculate_button is not None:
            self.calculate_button.configure(state=tk.DISABLED)
        self.set_status("正在扫描 invoices...")

        def progress(index: int, total: int, filename: str) -> None:
            self.after(0, lambda: self.set_status(f"正在扫描 {index}/{total}：{filename}"))

        def worker() -> None:
            try:
                result = scan_invoice_directory(progress)
                self.after(0, lambda: self.on_scan_done(result, None))
            except Exception as exc:
                error = f"{exc}\n\n{traceback.format_exc()}"
                self.after(0, lambda: self.on_scan_done(None, error))

        threading.Thread(target=worker, daemon=True).start()

    def on_scan_done(self, result, error: str | None) -> None:
        self.scanning_running = False
        if self.calculate_button is not None:
            self.calculate_button.configure(state=tk.NORMAL)
        if error is not None:
            self.update_available_summary()
            messagebox.showerror("扫描失败", error)
            self.set_status("扫描失败")
            return
        try:
            self.invoices, failures, reused, duplicates = result
            self.populate_tree(self.invoice_tree, self.invoices)
            self.update_available_summary()
            self.clear_result()
            message = f"已识别 {len(self.invoices)} 张发票"
            if failures:
                message += f"，{len(failures)} 个文件已移到 failed_invoices"
            if reused:
                message += f"，{len(reused)} 张已使用发票已自动转移"
            if duplicates:
                message += f"，{len(duplicates)} 个重复号码文件已移到 void_invoices"
            self.set_status(message)
        except Exception as exc:
            messagebox.showerror("扫描失败", f"{exc}\n\n{traceback.format_exc()}")
            self.set_status("扫描失败")

    def rename_files_by_invoice_number(self) -> None:
        if self.scanning_running or self.calculation_running:
            return
        try:
            self.set_status("正在按发票号码重命名...")
            renamed, unchanged, skipped = rename_invoice_files_by_number()
            self.scan()
            message = f"已重命名 {renamed} 个文件，{unchanged} 个文件名无需修改"
            if skipped:
                message += f"，{len(skipped)} 个文件跳过"
            self.set_status(message)
            if skipped:
                details = "\n".join(skipped[:20])
                if len(skipped) > 20:
                    details += f"\n... 还有 {len(skipped) - 20} 个文件"
                messagebox.showwarning("自动文件名完成", f"{message}\n\n{details}")
            else:
                messagebox.showinfo("自动文件名完成", message)
        except Exception as exc:
            messagebox.showerror("自动文件名失败", f"{exc}\n\n{traceback.format_exc()}")
            self.set_status("自动文件名失败")

    def choose_used_archive(self, directories: list[Path]) -> Path | None:
        dialog = tk.Toplevel(self)
        dialog.title("选择要还原的已使用目录")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.geometry("480x380")
        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=f"选择 {USED_DIR} 下的一个子目录：").pack(anchor=tk.W)
        tree = ttk.Treeview(frame, columns=("name", "count"), show="headings", selectmode="browse")
        tree.heading("name", text="子目录")
        tree.heading("count", text="发票数")
        tree.column("name", width=320, anchor=tk.W)
        tree.column("count", width=80, anchor=tk.E)
        tree.pack(fill=tk.BOTH, expand=True, pady=(8, 8))
        for directory in directories:
            tree.insert("", tk.END, iid=directory.name, values=(directory.name, len(archived_invoice_files(directory))))
        if directories:
            first = directories[0].name
            tree.selection_set(first)
            tree.focus(first)
        chosen: dict[str, Path | None] = {"value": None}

        def confirm(_event=None) -> None:
            selection = tree.selection()
            if not selection:
                return
            chosen["value"] = USED_DIR / selection[0]
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(fill=tk.X)
        ttk.Button(buttons, text="取消", command=cancel).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="还原", command=confirm).pack(side=tk.RIGHT, padx=(0, 8))
        tree.bind("<Double-1>", confirm)
        dialog.bind("<Escape>", lambda _event: cancel())
        dialog.grab_set()
        dialog.wait_window()
        return chosen["value"]

    def restore_used_invoices(self) -> None:
        if self.scanning_running or self.calculation_running:
            return
        directories = used_archive_directories()
        if not directories:
            messagebox.showinfo("没有已使用记录", f"{USED_DIR} 下还没有已归档的子目录。")
            return
        chosen = self.choose_used_archive(directories)
        if chosen is None:
            return
        files = archived_invoice_files(chosen)
        if not files:
            messagebox.showinfo("没有可还原的发票", f"{chosen.name} 中没有发票文件。")
            return
        if not messagebox.askyesno(
            "确认还原",
            f"将把“{chosen.name}”中的 {len(files)} 张发票移回 invoices 目录，"
            f"并清除它们的“已使用”记录。\n\n"
            f"该目录中的合并 PDF 与 Excel 核对表保持不动。\n是否继续？",
            icon="warning",
            default=messagebox.NO,
            parent=self,
        ):
            return
        try:
            moved, numbers, errors = restore_used_archive(chosen)
        except Exception as exc:
            messagebox.showerror("还原失败", f"{exc}\n\n{traceback.format_exc()}")
            self.set_status("还原失败")
            return
        if moved:
            self.scan()
        detail = f"已把 {len(moved)} 张发票还原到 invoices，并清除 {len(numbers)} 条已使用记录。"
        if errors:
            messagebox.showwarning("部分还原未完成", f"{detail}\n\n" + "\n".join(errors))
        else:
            messagebox.showinfo("已还原", detail)
        self.set_status(f"已从 {chosen.name} 还原 {len(moved)} 张发票")

    def update_available_summary(self) -> None:
        total = sum(invoice.amount_cents for invoice in self.invoices)
        self.available_summary_var.set(f"共 {len(self.invoices)} 张，合计 {cents_to_text(total)} 元")

    def prepare_invoice_type_menu(self) -> None:
        invoice = self.selected_invoice_from_tree(self.invoice_tree, self.invoices)
        self.invoice_type_var.set(invoice.invoice_type if invoice is not None else "")
        state = tk.NORMAL if invoice is not None and not self.calculation_running and not self.scanning_running else tk.DISABLED
        for menu in (self.invoice_type_menu, self.invoice_context_menu):
            for index in range(menu.index(tk.END) + 1):
                if menu.type(index) != "separator":
                    menu.entryconfigure(index, state=state)

    def show_invoice_type_menu(self, event) -> str:
        item_id = self.invoice_tree.identify_row(event.y)
        if item_id:
            self.invoice_tree.selection_set(item_id)
            self.invoice_tree.focus(item_id)
            self.invoice_tree.focus_set()
            try:
                self.invoice_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.invoice_context_menu.grab_release()
        return "break"

    def delete_selected_invoice(self) -> None:
        if self.calculation_running or self.scanning_running:
            return
        invoice = self.selected_invoice_from_tree(self.invoice_tree, self.invoices)
        if invoice is None:
            return
        source = Path(invoice.path)
        if not messagebox.askyesno(
            "确认删除发票",
            f"确定永久删除这张发票吗？\n\n{invoice.filename}\n{source}\n\n不进入回收站，删除后无法恢复。",
            icon="warning", default=messagebox.NO, parent=self,
        ):
            return
        try:
            source.unlink(missing_ok=True)
        except OSError as exc:
            messagebox.showerror("删除失败", f"无法删除 {invoice.filename}：\n{exc}", parent=self)
            return
        self.refresh_after_moved({invoice.path})
        self.set_status(f"已删除：{invoice.filename}")

    def set_selected_invoice_type(self, invoice_type: str | None) -> None:
        invoice = self.selected_invoice_from_tree(self.invoice_tree, self.invoices)
        if invoice is None or self.calculation_running:
            return
        try:
            resolved_type = parse_invoice(Path(invoice.path)).invoice_type if invoice_type is None else invoice_type
            save_invoice_type_override(invoice, invoice_type)
        except Exception as exc:
            self.invoice_type_var.set(invoice.invoice_type)
            messagebox.showerror("设置类型失败", str(exc))
            return

        updated_invoice = replace(invoice, invoice_type=resolved_type)
        self.invoices = [updated_invoice if item.path == invoice.path else item for item in self.invoices]
        self.result_invoices = [updated_invoice if item.path == invoice.path else item for item in self.result_invoices]
        for tree, invoices in ((self.invoice_tree, self.invoices), (self.result_tree, self.result_invoices)):
            for index, item in enumerate(invoices):
                if item.path == invoice.path:
                    tree.set(str(index), "type", resolved_type)

        if any(item.path == invoice.path for item in self.print_preview_invoices):
            if self.preview_invoice is not None and self.preview_invoice.path == str(self.print_preview_pdf):
                self.preview_invoice = updated_invoice
                self.preview_page = 0
            self.reset_print_preview()
        if self.preview_invoice is not None and self.preview_invoice.path == invoice.path:
            self.preview_invoice = updated_invoice
        self.invoice_type_var.set(resolved_type)
        self.render_preview()
        self.set_status(f"已更新类型：{invoice.filename} → {resolved_type}")

    def populate_tree(self, tree: ttk.Treeview, invoices: list[Invoice]) -> None:
        tree.delete(*tree.get_children())
        show_status = "status" in tree["columns"]
        result_paths = {invoice.path for invoice in self.result_invoices} if show_status else set()
        for index, invoice in enumerate(invoices):
            values = [
                invoice.amount_text,
                invoice.invoice_type,
                invoice.invoice_date,
                invoice.invoice_number,
            ]
            if show_status:
                values.insert(0, "✅" if invoice.path in result_paths else "")
            tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=values,
            )

    def refresh_result_tree(self) -> None:
        self.reset_print_preview()
        self.populate_tree(self.result_tree, self.result_invoices)
        self.populate_tree(self.invoice_tree, self.invoices)
        self.update_result_summary()

    def add_available_invoice_to_result(self, _event=None) -> None:
        invoice = self.selected_invoice_from_tree(self.invoice_tree, self.invoices)
        if invoice is None:
            return
        if any(item.path == invoice.path for item in self.result_invoices):
            self.set_status(f"{invoice.filename} 已在推荐组合中")
            return

        self.result_invoices.append(invoice)
        self.result_invoices.sort(
            key=lambda item: (invoice_date_sort_key(item.invoice_date), item.amount_cents, item.filename)
        )
        self.refresh_result_tree()
        for item_id, item in enumerate(self.result_invoices):
            if item.path == invoice.path:
                iid = str(item_id)
                self.result_tree.selection_set(iid)
                self.result_tree.focus(iid)
                self.result_tree.see(iid)
                break
        self.preview_invoice = invoice
        self.preview_page = 0
        self.render_preview()
        self.set_status(f"已添加到推荐组合：{invoice.filename}")

    def remove_result_invoice_from_result(self, event) -> None:
        item_id = self.result_tree.identify_row(event.y)
        if not item_id:
            return
        try:
            index = int(item_id)
            removed = self.result_invoices.pop(index)
        except Exception:
            return

        next_index = min(index, len(self.result_invoices) - 1)
        self.refresh_result_tree()
        if next_index >= 0:
            iid = str(next_index)
            self.result_tree.selection_set(iid)
            self.result_tree.focus(iid)
            self.result_tree.see(iid)
            self.preview_invoice = self.result_invoices[next_index]
        else:
            self.preview_invoice = None
        self.preview_page = 0
        self.render_preview()
        self.set_status(f"已从推荐组合移除：{removed.filename}")

    def target_cents(self) -> int | None:
        value = self.target_var.get().strip()
        cents = parse_money_to_cents(value)
        if cents is None or cents <= 0:
            messagebox.showwarning("金额无效", "请输入大于 0 的金额，例如 1200.50")
            return None
        return cents

    def on_category_changed(self, category: str) -> None:
        if category == CATEGORY_ALL and self.category_vars[CATEGORY_ALL].get():
            for name, variable in self.category_vars.items():
                if name != CATEGORY_ALL:
                    variable.set(False)
        else:
            self.category_vars[CATEGORY_ALL].set(
                not any(variable.get() for name, variable in self.category_vars.items() if name != CATEGORY_ALL)
            )

    def filtered_invoices(self) -> list[Invoice]:
        categories = {name for name, variable in self.category_vars.items() if variable.get()}
        if not categories or CATEGORY_ALL in categories:
            return self.invoices
        return [invoice for invoice in self.invoices if invoice.invoice_type in categories]

    def calculate(self) -> None:
        if self.calculation_running or self.scanning_running:
            return
        self.clear_result()
        target = self.target_cents()
        if target is None:
            return
        save_last_target_amount(self.target_var.get().strip())
        candidates = list(self.filtered_invoices())
        self.calculation_running = True
        if self.calculate_button is not None:
            self.calculate_button.configure(state=tk.DISABLED)
        self.summary_var.set("正在计算...")
        self.set_status("正在计算组合...")

        def worker() -> None:
            try:
                result = find_best_combination(candidates, target)
                self.after(0, lambda: self.on_calculate_done(target, result, None))
            except Exception as exc:
                error = f"{exc}\n\n{traceback.format_exc()}"
                self.after(0, lambda: self.on_calculate_done(target, [], error))

        threading.Thread(target=worker, daemon=True).start()

    def on_calculate_done(self, target: int, result: list[Invoice], error: str | None) -> None:
        self.calculation_running = False
        if self.calculate_button is not None:
            self.calculate_button.configure(state=tk.NORMAL)
        if error is not None:
            messagebox.showerror("计算失败", error)
            self.summary_var.set("计算失败")
            self.set_status("计算失败")
            return
        self.result_invoices = result
        self.reset_print_preview()
        self.populate_tree(self.result_tree, self.result_invoices)
        self.populate_tree(self.invoice_tree, self.invoices)
        if not result:
            self.summary_var.set("没有找到可覆盖金额的组合")
            self.set_status("没有找到可覆盖金额的组合")
            return
        total = sum(invoice.amount_cents for invoice in result)
        over = total - target
        self.summary_var.set(f"合计 {cents_to_text(total)}，超出 {cents_to_text(over)}，{len(result)} 张")
        self.set_status("已计算最接近组合")
        self.preview_invoice = result[0]
        self.preview_page = 0
        self.render_preview()

    def clear_result(self) -> None:
        self.result_invoices = []
        self.preview_invoice = None
        self.preview_page = 0
        self.preview_photo = None
        self.reset_print_preview()
        self.populate_tree(self.result_tree, self.result_invoices)
        self.populate_tree(self.invoice_tree, self.invoices)
        self.render_preview()
        self.summary_var.set("未计算")

    def reset_print_preview(self) -> None:
        self.print_preview_pdf = None
        self.print_preview_invoices = []

    def update_result_summary(self) -> None:
        if not self.result_invoices:
            self.summary_var.set("未计算")
            return
        total = sum(invoice.amount_cents for invoice in self.result_invoices)
        target = parse_money_to_cents(self.target_var.get().strip())
        if target is not None and total >= target:
            self.summary_var.set(
                f"合计 {cents_to_text(total)}，超出 {cents_to_text(total - target)}，{len(self.result_invoices)} 张"
            )
        elif target is not None:
            self.summary_var.set(
                f"剩余合计 {cents_to_text(total)}，不足 {cents_to_text(target - total)}，{len(self.result_invoices)} 张"
            )
        else:
            self.summary_var.set(f"剩余合计 {cents_to_text(total)}，{len(self.result_invoices)} 张")

    def refresh_after_moved(self, moved_paths: set[str]) -> None:
        preview_path = self.preview_invoice.path if self.preview_invoice is not None else ""
        preview_was_combination = (
            self.print_preview_pdf is not None
            and preview_path == str(self.print_preview_pdf)
            and any(invoice.path in moved_paths for invoice in self.print_preview_invoices)
        )

        self.invoices = [invoice for invoice in self.invoices if invoice.path not in moved_paths]
        self.result_invoices = [invoice for invoice in self.result_invoices if invoice.path not in moved_paths]
        self.populate_tree(self.invoice_tree, self.invoices)
        self.populate_tree(self.result_tree, self.result_invoices)
        self.update_available_summary()
        self.update_result_summary()

        if preview_path in moved_paths or preview_was_combination:
            self.preview_invoice = self.result_invoices[0] if self.result_invoices else None
            self.preview_page = 0
            self.render_preview()

        if preview_was_combination or any(invoice.path in moved_paths for invoice in self.print_preview_invoices):
            self.reset_print_preview()

    def selected_invoice_from_tree(self, tree: ttk.Treeview, invoices: list[Invoice]) -> Invoice | None:
        selection = tree.selection()
        if not selection:
            return None
        focused = tree.focus()
        item_id = focused if focused in selection else selection[0]
        try:
            return invoices[int(item_id)]
        except Exception:
            return None

    def selected_invoices_from_tree(self, tree: ttk.Treeview, invoices: list[Invoice]) -> list[Invoice]:
        selected: list[Invoice] = []
        for item_id in tree.selection():
            try:
                selected.append(invoices[int(item_id)])
            except Exception:
                continue
        return selected

    def preview_from_tree(self, tree: ttk.Treeview, invoices: list[Invoice]) -> None:
        invoice = self.selected_invoice_from_tree(tree, invoices)
        if invoice is None:
            return
        self.preview_invoice = invoice
        self.preview_page = 0
        self.render_preview()

    def change_preview_page(self, delta: int) -> None:
        if self.preview_invoice is None:
            return
        pages = max(self.preview_invoice.pages, 1)
        self.preview_page = min(max(self.preview_page + delta, 0), pages - 1)
        self.render_preview()

    def change_zoom(self, delta: float) -> None:
        self.preview_zoom = min(max(self.preview_zoom + delta, 0.5), 2.5)
        self.render_preview()

    def render_preview(self) -> None:
        self.preview_canvas.delete("all")
        invoice = self.preview_invoice
        if invoice is None:
            self.preview_label_var.set("选择一张发票预览")
            return
        if fitz is None or Image is None or ImageTk is None:
            self.preview_label_var.set("缺少 PyMuPDF/Pillow，无法预览")
            return

        source_path = Path(invoice.path)
        if not source_path.exists():
            self.preview_label_var.set("文件不存在，可能已经移动")
            return

        try:
            path = invoice_pdf_path(invoice)
            with fitz.open(path) as doc:
                page_count = len(doc)
                self.preview_page = min(self.preview_page, page_count - 1)
                page = doc[self.preview_page]
                canvas_width = max(self.preview_canvas.winfo_width(), 200)
                canvas_height = max(self.preview_canvas.winfo_height(), 200)
                scale = min(canvas_width / page.rect.width, canvas_height / page.rect.height) * self.preview_zoom
                scale = max(scale, 0.3)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                self.preview_photo = ImageTk.PhotoImage(image)
                x = max((canvas_width - image.width) // 2, 0)
                y = max((canvas_height - image.height) // 2, 0)
                self.preview_canvas.create_image(x, y, image=self.preview_photo, anchor=tk.NW)
                self.preview_label_var.set(
                    f"{invoice.filename}  第 {self.preview_page + 1}/{page_count} 页  {invoice.amount_text} 元"
                )
        except Exception as exc:
            self.preview_label_var.set(f"预览失败：{exc}")

    def generate_combination_preview(self) -> None:
        if not self.result_invoices:
            messagebox.showwarning("没有组合", "请先输入金额并计算组合")
            return
        invoices = sorted(self.result_invoices, key=invoice_type_sort_key)

        try:
            merged_pdf = merge_invoices_to_pdf(invoices)
            page_count = count_pdf_pages(merged_pdf)
            excel_path = export_combination_excel(
                invoices, merged_pdf.with_suffix(".xlsx"), parse_money_to_cents(self.target_var.get().strip())
            )
            total = sum(invoice.amount_cents for invoice in invoices)
            self.preview_invoice = Invoice(
                path=str(merged_pdf),
                filename=merged_pdf.name,
                amount_cents=total,
                invoice_type="组合",
                pages=page_count,
            )
            self.print_preview_pdf = merged_pdf
            self.print_preview_invoices = invoices
            self.preview_page = 0
            self.render_preview()
            self.set_status(f"已生成组合预览及 Excel 核对表：{merged_pdf.name}")
        except Exception as exc:
            messagebox.showerror("生成组合预览失败", f"{exc}\n\n{traceback.format_exc()}")
            return

        if messagebox.askyesno(
            "组合预览已生成",
            f"已生成合并 PDF，共 {len(invoices)} 张发票、{page_count} 页。\n\nExcel 核对表：\n{excel_path}\n\n是否打开 Excel 核对表？",
        ):
            try:
                os.startfile(excel_path)
            except OSError as exc:
                messagebox.showwarning("无法打开核对表", f"核对表已保存，请使用 Excel 或 WPS 打开：\n{excel_path}\n\n{exc}")

    def print_preview_pdf_file(self) -> None:
        if self.preview_invoice is None:
            messagebox.showwarning("没有预览", "请先在列表中选择一张发票，或点击“生成组合预览”。")
            return

        source_path = Path(self.preview_invoice.path)
        if not source_path.exists():
            messagebox.showerror("文件不存在", "当前预览的发票文件不存在，可能已经移动。")
            if self.print_preview_pdf is not None and source_path == self.print_preview_pdf:
                self.reset_print_preview()
            return

        is_combination_preview = (
            self.print_preview_pdf is not None
            and source_path == self.print_preview_pdf
            and bool(self.print_preview_invoices)
        )
        try:
            print_path = source_path if is_combination_preview else invoice_pdf_path(self.preview_invoice)
        except Exception as exc:
            messagebox.showerror("准备打印失败", f"{exc}\n\n{traceback.format_exc()}")
            return
        invoices = list(self.print_preview_invoices) if is_combination_preview else [self.preview_invoice]
        page_count = max(count_pdf_pages(print_path), self.preview_invoice.pages, 1)
        selection: PrintDialogSelection | None = None
        try:
            selection = show_windows_print_dialog(self.winfo_id(), page_count)
        except Exception as exc:
            messagebox.showerror("打印设置失败", f"{exc}\n\n{traceback.format_exc()}")
            release_print_dialog_selection(selection)
            return

        if selection is None:
            self.set_status("已取消打印")
            return

        def worker() -> None:
            try:
                self.after(0, lambda: self.set_status("正在发送打印任务..."))
                print_pdf_with_dialog_selection(print_path, selection)
                self.after(0, lambda: self.on_print_success(invoices))
            except Exception as exc:
                error = f"{exc}\n\n{traceback.format_exc()}"
                self.after(0, lambda: messagebox.showerror("打印失败", error))
                self.after(0, lambda: self.set_status("打印失败"))
            finally:
                release_print_dialog_selection(selection)

        threading.Thread(target=worker, daemon=True).start()

    def on_print_success(self, invoices: list[Invoice]) -> None:
        self.set_status("打印任务已发送")
        label = "这张发票" if len(invoices) == 1 else "这些发票"
        if messagebox.askyesno("打印完成", f"打印任务已发送到 Windows 打印队列。是否将{label}标记为已使用并移动？"):
            self.mark_used(invoices)

    def mark_selected_used(self) -> None:
        invoices = self.selected_invoices_from_tree(self.result_tree, self.result_invoices)
        if not invoices:
            messagebox.showwarning("未选择发票", "请先在推荐组合列表中选择要标记的发票。")
            return
        self.mark_used(invoices)

    def mark_all_used(self) -> None:
        if not self.result_invoices:
            messagebox.showwarning("没有组合", "当前没有可标记的推荐组合")
            return
        self.mark_used(list(self.result_invoices))

    def mark_used(self, invoices: list[Invoice] | None = None) -> None:
        invoices = list(invoices or [])
        if not invoices:
            messagebox.showwarning("没有发票", "当前没有可标记的发票")
            return
        target = self.target_var.get().strip()
        content = safe_filename_component(self.content_var.get()) or "报销"
        archive_dir = USED_DIR / f"{content}{datetime.now():%Y%m%d}"
        preview_by_path = {invoice.path: invoice for invoice in self.print_preview_invoices}
        preview_pdf = self.print_preview_pdf if (
            len(invoices) == len(preview_by_path)
            and all(preview_by_path.get(invoice.path) == invoice for invoice in invoices)
        ) else None
        moved = 0
        moved_paths: set[str] = set()
        archived_invoices: list[Invoice] = []
        errors: list[str] = []
        for invoice in invoices:
            source = Path(invoice.path)
            if not source.exists():
                errors.append(f"{invoice.filename}: 文件不存在")
                continue
            try:
                destination = move_invoice(source, archive_dir)
                record = asdict(invoice)
                record.update(
                    {
                        "used_at": now_text(),
                        "target_amount": target,
                        "moved_to": str(destination),
                    }
                )
                save_used_invoice_record(record)
                append_json_record(STATE_FILE, record)
                moved += 1
                moved_paths.add(invoice.path)
                archived_invoices.append(replace(invoice, path=str(destination), filename=destination.name))
            except Exception as exc:
                errors.append(f"{invoice.filename}: {exc}")

        if archived_invoices:
            try:
                can_reuse_preview = moved == len(invoices) and all(
                    original.filename == archived.filename
                    for original, archived in zip(invoices, archived_invoices)
                )
                archive_combination_documents(
                    archived_invoices, archive_dir, parse_money_to_cents(target),
                    preview_pdf if can_reuse_preview else None,
                )
            except Exception as exc:
                errors.append(f"发票已归档到 {archive_dir}，但 PDF/Excel 核对文件归档失败：{exc}")
        if moved_paths:
            self.refresh_after_moved(moved_paths)
        if errors:
            messagebox.showwarning("部分归档未完成", "\n".join(errors))
        else:
            messagebox.showinfo("已标记", f"已归档 {moved} 张发票、合并 PDF 及 Excel 核对表到：\n{archive_dir}")


def print_scan_json() -> int:
    invoices, failures, reused, duplicates = scan_invoice_directory()
    payload = {
        "invoices": [asdict(invoice) for invoice in invoices],
        "failures": failures,
        "reused": reused,
        "duplicates": duplicates,
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def main() -> int:
    if "--scan-json" in sys.argv:
        return print_scan_json()
    app = AutoInvoiceApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
