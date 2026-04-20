"""
pdf_to_excel.py  –  PDF/Word→Excel 変換スクリプト
優先順位（PDF）:
  1. Azure Document Intelligence (prebuilt-layout) - 縦書き・複雑表・スキャン対応
  2. pdfplumber テーブル検出（線ありPDF・フォールバック）
  3. PyMuPDF テキスト抽出（フォールバック）
  4. pdfplumber テキスト抽出（最終フォールバック）
優先順位（Word .docx）:
  1. Azure Document Intelligence (prebuilt-layout)
  2. python-docx（テーブル・段落抽出フォールバック）
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

try:
    import docx as python_docx
    HAS_PYTHON_DOCX = True
except ImportError:
    HAS_PYTHON_DOCX = False


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

    input_ext = os.path.splitext(pdf_path)[1].lower()
    content_type_map = {
        ".pdf":  "application/pdf",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    content_type = content_type_map.get(input_ext, "application/octet-stream")
    print(f"[pdf_to_excel] Using Azure Document Intelligence (prebuilt-layout) ext={input_ext}.", file=sys.stderr)
    try:
        with open(pdf_path, "rb") as f:
            poller = client.begin_analyze_document(
                "prebuilt-layout",
                body=f,
                content_type=content_type,
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


# ── LibreOffice: DOCX → PDF 変換（EMF 画像型 Word 対策） ────────────────

def _docx_to_pdf_with_libreoffice(docx_path: str) -> str | None:
    """LibreOffice --headless で .docx を .pdf に変換して返す。失敗時は None。"""
    import subprocess
    import tempfile
    out_dir = tempfile.mkdtemp()
    try:
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            capture_output=True,
            text=True,
            timeout=120,
        )
        print(f"[pdf_to_excel] LibreOffice rc={result.returncode} {result.stderr[:200]}", file=sys.stderr)
        if result.returncode != 0:
            return None
        base = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(out_dir, base + ".pdf")
        return pdf_path if os.path.exists(pdf_path) else None
    except FileNotFoundError:
        print("[pdf_to_excel] libreoffice not found.", file=sys.stderr)
        return None
    except subprocess.TimeoutExpired:
        print("[pdf_to_excel] LibreOffice conversion timed out.", file=sys.stderr)
        return None


# ── python-docx フォールバック（Word専用） ───────────────────────────────

def _process_with_python_docx(docx_path: str, wb: openpyxl.Workbook) -> dict:
    """python-docx で .docx のテーブルと段落をExcelに変換する。"""
    if not HAS_PYTHON_DOCX:
        print("[pdf_to_excel] python-docx not available.", file=sys.stderr)
        return {"tables": 0, "pages": 1, "engine": "none"}

    print("[pdf_to_excel] Falling back to python-docx.", file=sys.stderr)
    doc = python_docx.Document(docx_path)
    total_tables = 0
    text_rows: list[list[str]] = []

    # テーブル抽出
    for t_idx, table in enumerate(doc.tables):
        total_tables += 1
        sheet_name = f"Table{t_idx + 1}"
        ws = wb.create_sheet(title=sheet_name[:31])
        for row in table.rows:
            ws.append([cell.text.strip() for cell in row.cells])
        print(f"[pdf_to_excel] python-docx table {t_idx+1}: {len(table.rows)} rows", file=sys.stderr)

    # 段落テキスト（テーブルがない場合のみ）
    if total_tables == 0:
        for para in doc.paragraphs:
            if para.text.strip():
                text_rows.append([para.text.strip()])
        if text_rows:
            ws = wb.create_sheet(title="Text")
            for row in text_rows:
                ws.append(row)

    return {"tables": total_tables, "pages": 1, "engine": "python_docx"}


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

    input_ext = os.path.splitext(args.input)[1].lower()

    # ── .docx: python-docx を優先（DI が空成功を返す EMF 画像型 Word 対策） ──
    if input_ext == ".docx":
        docx_result = _process_with_python_docx(args.input, wb)
        got_content = docx_result["tables"] > 0 or any(
            ws.max_row and ws.max_row > 0 for ws in wb.worksheets
        )
        if got_content:
            if not wb.sheetnames:
                wb.create_sheet(title="Sheet1")
            wb.save(str(args.output))
            print(json.dumps({
                "sheets": len(wb.sheetnames),
                "tables": docx_result["tables"],
                "pages": docx_result["pages"],
                "engine": docx_result["engine"],
            }))
            return

        # python-docx で取れなかった → DI で再試行（テキスト型 DOCX への保険）
        print("[pdf_to_excel] python-docx found nothing, trying Doc Intelligence.", file=sys.stderr)
        wb2 = openpyxl.Workbook()
        wb2.remove(wb2.active)
        doc_intel_result = _process_with_doc_intel(args.input, wb2)
        if doc_intel_result is not None and (doc_intel_result["tables"] > 0 or doc_intel_result.get("paragraphs", 0) > 0):
            if not wb2.sheetnames:
                wb2.create_sheet(title="Sheet1")
            wb2.save(str(args.output))
            print(json.dumps({
                "sheets": len(wb2.sheetnames),
                "tables": doc_intel_result["tables"],
                "pages": doc_intel_result["pages"],
                "engine": "doc_intelligence",
            }))
            return

        # python-docx も DI も空 → LibreOffice で DOCX→PDF に変換して再度 DI を試みる
        print("[pdf_to_excel] Trying LibreOffice DOCX→PDF conversion.", file=sys.stderr)
        pdf_path = _docx_to_pdf_with_libreoffice(args.input)
        if pdf_path:
            wb3 = openpyxl.Workbook()
            wb3.remove(wb3.active)
            di_pdf_result = _process_with_doc_intel(pdf_path, wb3)
            if di_pdf_result is not None and (di_pdf_result["tables"] > 0 or di_pdf_result.get("paragraphs", 0) > 0):
                if not wb3.sheetnames:
                    wb3.create_sheet(title="Sheet1")
                wb3.save(str(args.output))
                print(json.dumps({
                    "sheets": len(wb3.sheetnames),
                    "tables": di_pdf_result["tables"],
                    "pages": di_pdf_result["pages"],
                    "engine": "doc_intelligence_via_pdf",
                }))
                return

        # 全手段で失敗 = LibreOffice 未インストールか変換不可
        print("[pdf_to_excel] All extraction methods failed for DOCX.", file=sys.stderr)
        ws_msg = wb.create_sheet(title="注意")
        ws_msg.append(["このWordファイルは画像埋め込み型のため、表データを抽出できませんでした。"])
        ws_msg.append(["PDFとして保存してからアップロードしてください。"])
        wb.save(str(args.output))
        print(json.dumps({"sheets": 1, "tables": 0, "pages": 1, "engine": "none"}))
        return

    # ── 1. Azure Document Intelligence（PDF 等） ─────────────────────────
    doc_intel_result = _process_with_doc_intel(args.input, wb)

    if doc_intel_result is not None:
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

    # ── 2. フォールバック（PDF） ──────────────────────────────────────────

    # PDF 以外（想定外拡張子）は空Excelで終了（.docx は上で処理済み）
    if input_ext not in (".pdf",):
        print(f"[pdf_to_excel] No fallback for {input_ext}.", file=sys.stderr)
        wb.create_sheet(title="Sheet1")
        wb.save(str(args.output))
        print(json.dumps({"sheets": 1, "tables": 0, "pages": 0, "engine": "none"}))
        return

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
