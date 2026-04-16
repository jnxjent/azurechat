"""
pdf_to_excel.py  –  PDF→Excel 変換スクリプト
優先順位:
  1. Azure Document Intelligence (prebuilt-layout) - 縦書き・複雑表・スキャン対応
  2. pdfplumber テーブル検出（線ありPDF・フォールバック）
  3. PyMuPDF テキスト抽出（フォールバック）
  4. pdfplumber テキスト抽出（最終フォールバック）
CLI: python pdf_to_excel.py --input <path> --output <path>
stdout: JSON {"sheets": N, "tables": N, "pages": N, "engine": "..."}
"""

import argparse
import json
import os
import re
import sys
from collections import defaultdict

import openpyxl
import pdfplumber

# ── オプション依存ライブラリ ──────────────────────────────────────────────
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    from azure.ai.documentintelligence import DocumentIntelligenceClient
    from azure.core.credentials import AzureKeyCredential
    HAS_DOC_INTEL = True
except ImportError:
    HAS_DOC_INTEL = False


# ── Doc Intelligence クライアント ─────────────────────────────────────────

def _get_doc_intel_client():
    if not HAS_DOC_INTEL:
        return None
    endpoint = os.environ.get("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT", "").strip()
    key = os.environ.get("AZURE_DOCUMENT_INTELLIGENCE_KEY", "").strip()
    if not endpoint or not key:
        return None
    return DocumentIntelligenceClient(endpoint=endpoint, credential=AzureKeyCredential(key))


# ── 金額正規化 ────────────────────────────────────────────────────────────

_AMOUNT_RE = re.compile(r"^[△▲(（]?[\d,，]+[)）]?$")

def _normalize_amount(text: str) -> str:
    """△1,234 / (1,234) → -1234、通常数字はそのまま返す。数字でなければ原文を返す。"""
    t = text.strip()
    if not t:
        return t
    negative = t.startswith(("△", "▲")) or (
        (t.startswith("(") and t.endswith(")")) or
        (t.startswith("（") and t.endswith("）"))
    )
    cleaned = re.sub(r"[△▲(（)）,，\s]", "", t)
    if not cleaned.lstrip("-").isdigit():
        return text
    val = int(cleaned)
    return str(-val if negative else val)


def _maybe_normalize(text: str) -> str:
    """金額っぽいセルだけ正規化する。"""
    t = text.strip()
    if _AMOUNT_RE.match(t):
        return _normalize_amount(t)
    return t


# ── Doc Intelligence メイン処理 ───────────────────────────────────────────

def _process_with_doc_intel(pdf_path: str, wb: openpyxl.Workbook) -> dict | None:
    client = _get_doc_intel_client()
    if client is None:
        print("[pdf_to_excel] Doc Intelligence: not configured, skipping.", file=sys.stderr)
        return None

    print("[pdf_to_excel] Using Azure Document Intelligence (prebuilt-layout).", file=sys.stderr)
    try:
        with open(pdf_path, "rb") as f:
            poller = client.begin_analyze_document(
                "prebuilt-layout",
                body=f,
                content_type="application/octet-stream",
            )
        result = poller.result()
    except Exception as e:
        print(f"[pdf_to_excel] Doc Intelligence failed: {e}", file=sys.stderr)
        return None

    total_pages = len(result.pages) if result.pages else 0
    total_tables = 0

    if result.tables:
        # ページごとにテーブルをグループ化してシート名を決める
        tables_by_page: dict[int, list] = defaultdict(list)
        for table in result.tables:
            page_num = (
                table.bounding_regions[0].page_number
                if table.bounding_regions else 1
            )
            tables_by_page[page_num].append(table)

        for page_num in sorted(tables_by_page):
            page_tables = tables_by_page[page_num]
            for t_idx, table in enumerate(page_tables):
                total_tables += 1
                sheet_name = (
                    f"P{page_num}"
                    if len(page_tables) == 1
                    else f"P{page_num}-T{t_idx + 1}"
                )
                ws = wb.create_sheet(title=sheet_name[:31])

                # グリッドを構築（結合セルは左上セルのみ書き込む）
                grid: dict[tuple[int, int], str] = {}
                for cell in table.cells:
                    content = _maybe_normalize(cell.content or "")
                    grid[(cell.row_index, cell.column_index)] = content

                for r in range(table.row_count):
                    row_data = [grid.get((r, c), "") for c in range(table.column_count)]
                    ws.append(row_data)

                print(
                    f"[pdf_to_excel] P{page_num}-T{t_idx+1}: "
                    f"{table.row_count}rows × {table.column_count}cols",
                    file=sys.stderr,
                )

    # テーブルが検出されなかった場合はパラグラフをテキストシートへ
    if total_tables == 0 and result.paragraphs:
        ws = wb.create_sheet(title="Text")
        for para in result.paragraphs:
            if para.content and para.content.strip():
                ws.append([para.content.strip()])
        print(
            f"[pdf_to_excel] No tables found; wrote {len(result.paragraphs)} paragraphs to 'Text'.",
            file=sys.stderr,
        )

    para_count = len(result.paragraphs) if result.paragraphs else 0
    return {
        "tables": total_tables,
        "pages": total_pages,
        "paragraphs": para_count,
        "engine": "doc_intelligence",
    }


# ── pdfplumber テキストフォールバック ────────────────────────────────────

def _pdfplumber_text_lines(page) -> list[list[str]]:
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


# ── PyMuPDF テキストフォールバック ──────────────────────────────────────

def _pymupdf_text_lines(pdf_path: str, page_index: int) -> list[list[str]]:
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


# ── メイン ────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    print(
        f"[pdf_to_excel] HAS_PYMUPDF={HAS_PYMUPDF} HAS_DOC_INTEL={HAS_DOC_INTEL}",
        file=sys.stderr,
    )

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── 1. Azure Document Intelligence ──────────────────────────────────
    doc_intel_result = _process_with_doc_intel(args.input, wb)

    if doc_intel_result is not None:
        # Doc Intelligence が成功した場合はそのまま保存
        if not wb.sheetnames:
            wb.create_sheet(title="Sheet1")
        wb.save(str(args.output))
        print(json.dumps({
            "sheets": len(wb.sheetnames),
            "tables": doc_intel_result["tables"],
            "pages": doc_intel_result["pages"],
            "engine": "doc_intelligence",
        }))
        return

    # ── 2. フォールバック: pdfplumber / PyMuPDF ─────────────────────────
    print("[pdf_to_excel] Falling back to pdfplumber/PyMuPDF.", file=sys.stderr)
    total_tables = 0
    text_rows: list[list[str]] = []

    with pdfplumber.open(args.input) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, start=1):
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

            if HAS_PYMUPDF:
                lines = _pymupdf_text_lines(args.input, page_num - 1)
                print(f"[pdf_to_excel] p{page_num} PyMuPDF: {len(lines)} rows", file=sys.stderr)

            if not lines:
                lines = _pdfplumber_text_lines(page)
                print(f"[pdf_to_excel] p{page_num} pdfplumber text: {len(lines)} rows", file=sys.stderr)

            text_rows.extend(lines)

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
        "engine": "pdfplumber",
        "textRows": len(text_rows),
    }))


if __name__ == "__main__":
    main()
