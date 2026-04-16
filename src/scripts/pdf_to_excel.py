"""
pdf_to_excel.py  –  PDF→Excel 変換スクリプト
優先順位:
  1. pdfplumber テーブル検出（線ありPDF）
  2. EasyOCR + PyMuPDF（スキャンPDF・縦書き対応）
  3. PyMuPDF テキスト抽出（フォールバック）
  4. pdfplumber テキスト抽出（最終フォールバック）
CLI: python pdf_to_excel.py --input <path> --output <path>
stdout: JSON {"sheets": N, "tables": N, "pages": N}
"""

import argparse
import json
import io
import sys
from collections import defaultdict

import openpyxl
import pdfplumber

# ── オプション依存ライブラリの確認 ──────────────────────────────────────
try:
    import fitz  # PyMuPDF
    import numpy as np
    from PIL import Image
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import easyocr
    HAS_EASYOCR = True
except ImportError:
    HAS_EASYOCR = False


# ── ページ画像変換 ────────────────────────────────────────────────────

def _pdf_page_to_numpy(pdf_path: str, page_index: int, dpi: int = 150):
    """PyMuPDF でページを NumPy 配列（RGB）に変換する。"""
    doc = fitz.open(pdf_path)
    page = doc[page_index]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    doc.close()
    return np.array(img)


# ── EasyOCR 結果→行変換 ──────────────────────────────────────────────

def _easyocr_result_to_rows(ocr_result) -> list[list[str]]:
    """EasyOCR の出力をセルのリスト（行 × 列）に変換する。
    各要素は (bbox, text, confidence)。
    同じ Y 帯に属するテキストを同一行とみなし、X 座標でソートしてセルとする。
    """
    if not ocr_result:
        return []

    items: list[tuple[float, float, str]] = []
    for bbox, text, conf in ocr_result:
        if conf < 0.3:
            continue
        # bbox は [[x1,y1],[x2,y1],[x2,y2],[x1,y2]] 形式
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        x_left = min(p[0] for p in bbox)
        items.append((y_center, x_left, text))

    if not items:
        return []

    items.sort(key=lambda t: t[0])
    y_values = [t[0] for t in items]
    if len(y_values) > 1:
        gaps = [y_values[i + 1] - y_values[i] for i in range(len(y_values) - 1)]
        gaps_positive = [g for g in gaps if g > 0]
        threshold = max(8.0, sorted(gaps_positive)[len(gaps_positive) // 2] * 0.7) if gaps_positive else 8.0
    else:
        threshold = 8.0

    clusters: list[list[tuple[float, float, str]]] = []
    current: list[tuple[float, float, str]] = [items[0]]
    for item in items[1:]:
        if abs(item[0] - current[-1][0]) <= threshold:
            current.append(item)
        else:
            clusters.append(current)
            current = [item]
    clusters.append(current)

    rows: list[list[str]] = []
    for cluster in clusters:
        cluster.sort(key=lambda t: t[1])
        rows.append([t[2] for t in cluster])
    return rows


# ── PyMuPDF テキストフォールバック ──────────────────────────────────────

def _pymupdf_text_lines(pdf_path: str, page_index: int) -> list[list[str]]:
    """PyMuPDF の blocks API でテキスト行を取得する。"""
    doc = fitz.open(pdf_path)
    page = doc[page_index]
    blocks = page.get_text("blocks")
    doc.close()
    lines: list[list[str]] = []
    for block in sorted(blocks, key=lambda b: (b[1], b[0])):
        block_text = str(block[4]).strip()
        for line in block_text.splitlines():
            line = line.strip()
            if line:
                lines.append([line])
    return lines


# ── pdfplumber テキストフォールバック ────────────────────────────────────

def _pdfplumber_text_lines(page) -> list[list[str]]:
    """pdfplumber の extract_words() で Y グループ化してテキスト行を返す。"""
    words = page.extract_words(keep_blank_chars=True, x_tolerance=4, y_tolerance=4)
    if not words:
        raw = page.extract_text() or ""
        return [[ln] for ln in raw.splitlines() if ln.strip()]
    rows_by_y: dict[int, list] = defaultdict(list)
    for w in words:
        y_key = round(float(w.get("top", 0)) / 5) * 5
        rows_by_y[y_key].append(w)
    lines: list[list[str]] = []
    for y_key in sorted(rows_by_y.keys()):
        row = sorted(rows_by_y[y_key], key=lambda w: float(w.get("x0", 0)))
        line = " ".join(w["text"] for w in row).strip()
        if line:
            lines.append([line])
    return lines


# ── メイン ────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    total_tables = 0
    text_rows: list[list[str]] = []

    print(f"[pdf_to_excel] HAS_PYMUPDF={HAS_PYMUPDF} HAS_EASYOCR={HAS_EASYOCR}", file=sys.stderr)

    # EasyOCR エンジンを初期化
    ocr_reader = None
    if HAS_EASYOCR and HAS_PYMUPDF:
        try:
            ocr_reader = easyocr.Reader(['ja', 'en'], gpu=False, verbose=False)
            print("[pdf_to_excel] EasyOCR engine initialized.", file=sys.stderr)
        except Exception as e:
            print(f"[pdf_to_excel] EasyOCR init failed: {e}", file=sys.stderr)

    with pdfplumber.open(args.input) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, start=1):
            # pdfplumber でテーブル検出（線が明確なPDFには有効）
            tables = page.extract_tables() or []
            if tables:
                for t_idx, table in enumerate(tables):
                    total_tables += 1
                    sheet_name = (
                        f"P{page_num}" if len(tables) == 1 else f"P{page_num}-T{t_idx + 1}"
                    )
                    ws = wb.create_sheet(title=sheet_name[:31])
                    for row in table:
                        ws.append([str(cell) if cell is not None else "" for cell in row])
                continue

            lines: list[list[str]] = []

            # 1. EasyOCR
            if ocr_reader is not None and HAS_PYMUPDF:
                try:
                    img = _pdf_page_to_numpy(args.input, page_num - 1)
                    ocr_result = ocr_reader.readtext(img)
                    lines = _easyocr_result_to_rows(ocr_result)
                    print(f"[pdf_to_excel] p{page_num} EasyOCR: {len(lines)} rows", file=sys.stderr)
                except Exception as e:
                    print(f"[pdf_to_excel] EasyOCR page {page_num} failed: {e}", file=sys.stderr)

            # 2. PyMuPDF テキスト抽出
            if not lines and HAS_PYMUPDF:
                lines = _pymupdf_text_lines(args.input, page_num - 1)
                print(f"[pdf_to_excel] p{page_num} PyMuPDF fallback: {len(lines)} rows", file=sys.stderr)

            # 3. pdfplumber フォールバック
            if not lines:
                lines = _pdfplumber_text_lines(page)
                print(f"[pdf_to_excel] p{page_num} pdfplumber fallback: {len(lines)} rows", file=sys.stderr)

            text_rows.extend(lines)

    # テーブルがなければテキストを "Text" シートへ
    if total_tables == 0 and text_rows:
        ws = wb.create_sheet(title="Text")
        for row in text_rows:
            ws.append(row)

    if not wb.sheetnames:
        wb.create_sheet(title="Sheet1")

    wb.save(str(args.output))
    print(json.dumps({
        "sheets": len(wb.sheetnames),
        "tables": total_tables,
        "pages": total_pages,
        "textRows": len(text_rows),
        "hasPymupdf": HAS_PYMUPDF,
        "hasEasyocr": HAS_EASYOCR,
        "usedOcr": ocr_reader is not None,
    }))


if __name__ == "__main__":
    main()
