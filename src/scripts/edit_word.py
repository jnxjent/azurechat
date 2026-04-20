"""
edit_word.py  –  python-docx ベースの Word 編集スクリプト
CLI: python edit_word.py --input <path> --output <path> --plan <json_path>
stdout: JSON {"changedParagraphs": N, "totalParagraphs": N}
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor


def parse_hex_color(value: str | None):
    """6桁 RRGGBB → RGBColor。無効な値は None を返す。"""
    if not value:
        return None
    normalized = str(value).replace("#", "").strip().upper()
    if len(normalized) == 6 and all(ch in "0123456789ABCDEF" for ch in normalized):
        r = int(normalized[0:2], 16)
        g = int(normalized[2:4], 16)
        b = int(normalized[4:6], 16)
        return RGBColor(r, g, b)
    return None


def iter_all_paragraphs(doc):
    """段落とテーブルセル内の段落をすべてイテレートする。"""
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para


def apply_replace_text(doc, replacements: list) -> int:
    """replaceText を文書全体（段落・テーブル）に適用。変更した段落数を返す。"""
    changed = 0
    for para in iter_all_paragraphs(doc):
        para_changed = False
        for rep in replacements:
            find = rep.get("find", "")
            replace = rep.get("replace", "")
            if not find:
                continue
            for run in para.runs:
                if find in run.text:
                    run.text = run.text.replace(find, replace)
                    para_changed = True
        if para_changed:
            changed += 1
    return changed


def apply_format_runs(doc, formats: list) -> int:
    """formatRuns を段落に適用。matchText 省略時は全段落に適用。変更した段落数を返す。"""
    changed = 0
    for fmt in formats:
        match_text = str(fmt.get("matchText") or "").strip()
        apply_all = not match_text  # matchText が空 → 全段落対象
        bold = fmt.get("bold")
        italic = fmt.get("italic")
        font_size = fmt.get("fontSize")
        font_color = parse_hex_color(fmt.get("fontColor"))
        font_face = str(fmt.get("fontFace") or "").strip() or None

        if bold is None and italic is None and font_size is None and font_color is None and font_face is None:
            continue

        for para in iter_all_paragraphs(doc):
            if not apply_all and match_text not in para.text:
                continue
            runs = para.runs
            # runがない段落（空段落や特殊段落）には新しいrunを作って書式だけ設定
            if not runs and para.text:
                run = para.add_run(para.text)
                # 元のテキストをrunに移したので元の直接テキストは除去不要（add_runで追加）
                runs = [run]
            for run in runs:
                if bold is not None:
                    run.bold = bold
                if italic is not None:
                    run.italic = italic
                if font_size is not None:
                    run.font.size = Pt(float(font_size))
                if font_color is not None:
                    run.font.color.rgb = font_color
                if font_face is not None:
                    rPr = run._r.get_or_add_rPr()
                    rFonts = rPr.find(qn("w:rFonts"))
                    if rFonts is None:
                        rFonts = OxmlElement("w:rFonts")
                        rPr.insert(0, rFonts)
                    rFonts.set(qn("w:ascii"), font_face)
                    rFonts.set(qn("w:hAnsi"), font_face)
                    rFonts.set(qn("w:eastAsia"), font_face)
                    rFonts.set(qn("w:cs"), font_face)
            changed += 1

    return changed


def apply_add_paragraphs(doc, paragraphs: list) -> int:
    """addParagraphs をドキュメント末尾に追加。追加した段落数を返す。"""
    added = 0
    for para_def in paragraphs:
        text = str(para_def.get("text") or "").strip()
        if not text:
            continue
        style = str(para_def.get("style") or "Normal").strip()
        bold = para_def.get("bold")
        italic = para_def.get("italic")
        font_size = para_def.get("fontSize")
        font_color = parse_hex_color(para_def.get("fontColor"))

        try:
            para = doc.add_paragraph(text, style=style)
        except Exception:
            para = doc.add_paragraph(text)

        for run in para.runs:
            if bold is not None:
                run.bold = bold
            if italic is not None:
                run.italic = italic
            if font_size is not None:
                run.font.size = Pt(float(font_size))
            if font_color is not None:
                run.font.color.rgb = font_color
        added += 1

    return added


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--plan", required=True)
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    plan_path = Path(args.plan)

    plan: dict[str, Any] = json.loads(plan_path.read_text(encoding="utf-8"))
    doc = Document(str(input_path))
    total_paragraphs = sum(1 for _ in iter_all_paragraphs(doc))
    changed = 0

    changed += apply_replace_text(doc, plan.get("replaceText") or [])
    changed += apply_format_runs(doc, plan.get("formatRuns") or [])
    changed += apply_add_paragraphs(doc, plan.get("addParagraphs") or [])

    doc.save(str(output_path))
    print(json.dumps({"changedParagraphs": changed, "totalParagraphs": total_paragraphs}))


if __name__ == "__main__":
    main()
