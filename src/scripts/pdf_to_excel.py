"""
pdf_to_excel.py  –  PDF→Excel 変換スクリプト
優先順位:
  1. PaddleOCR + PyMuPDF（最高品質。縦書き・スキャンPDF対応）
  2. PyMuPDF テキスト抽出（フォールバック）
  3. pdfplumber テキスト抽出（最終フォールバック）
CLI: python pdf_to_excel.py --input <path> --output <path>
stdout: JSON {"sheets": N, "tables": N, "pages": N}
"""

import argparse
import json
import io
from collections import defaultdict
from pathlib import Path

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
    from paddleocr import PaddleOCR
    HAS_PADDLE = True
except ImportError:
    HAS_PADDLE = False


# ── PaddleOCR パス ──────────────────────────────────────────────────────

def _pdf_page_to_numpy(pdf_path: str, page_index: int, dpi: int = 200):
    """PyMuPDF でページを NumPy 配列（RGB）に変換する。"""
    doc = fitz.open(pdf_path)
    page = doc[page_index]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    doc.close()
    return np.array(img)


def _ocr_result_to_rows(ocr_result) -> list[list[str]]:
    """PaddleOCR の出力をセルのリスト（行 × 列）に変換する。

    同じ Y 帯に属するテキストを同一行とみなし、X 座標でソートしてセルとする。
    """
    if not ocr_result or not ocr_result[0]:
        return []

    items: list[tuple[float, float, str]] = []
    for line in ocr_result[0]:
        box, (text, _conf) = line
        y_center = (box[0][1] + box[2][1]) / 2
        x_left = min(p[0] for p in box)
        items.append((y_center, x_left, text))

    if not items:
        return []

    # Y でソートしてクラスタリング
    items.sort(key=lambda t: t[0])
    y_values = [t[0] for t in items]
    # 行高さの中央値をクラスタリング閾値に使用（最低 8px）
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
        cluster.sort(key=lambda t: t[1])   # X でソート
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

    import sys
    print(f"[pdf_to_excel] HAS_PYMUPDF={HAS_PYMUPDF} HAS_PADDLE={HAS_PADDLE}", file=sys.stderr)

    # PaddleOCR エンジンを初期化（初回はモデルをダウンロード）
    ocr_engine = None
    if HAS_PADDLE and HAS_PYMUPDF:
        try:
            ocr_engine = PaddleOCR(use_textline_orientation=True, lang="japan")
            print("[pdf_to_excel] PaddleOCR engine initialized.", file=sys.stderr)
        except Exception as e:
            print(f"[pdf_to_excel] PaddleOCR init failed: {e}", file=sys.stderr)

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

            # テーブルが検出されなかった場合: テキスト抽出
            lines: list[list[str]] = []

            # 1. PaddleOCR（最高品質）
            if ocr_engine is not None and HAS_PYMUPDF:
                try:
                    img = _pdf_page_to_numpy(args.input, page_num - 1)
                    ocr_result = ocr_engine.ocr(img, cls=True)
                    lines = _ocr_result_to_rows(ocr_result)
                    print(f"[pdf_to_excel] p{page_num} PaddleOCR: {len(lines)} rows", file=sys.stderr)
                except Exception as e:
                    print(f"[pdf_to_excel] PaddleOCR page {page_num} failed: {e}", file=sys.stderr)

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
        "hasPaddle": HAS_PADDLE,
        "usedOcr": ocr_engine is not None,
    }))


if __name__ == "__main__":
    main()
