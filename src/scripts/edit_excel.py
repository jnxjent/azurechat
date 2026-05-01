"""
edit_excel.py  –  openpyxl ベースの Excel 編集スクリプト
CLI: python edit_excel.py --input <path> --output <path> --plan <json_path>
stdout: JSON {"changedSheets": N, "totalSheets": N}
"""

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import column_index_from_string


def _replace_stdev_with_sqrt_sumproduct(formula: str) -> str:
    """STDEVP(range) / STDEV.P(range) \u3092\u5168Excel\u30d0\u30fc\u30b8\u30e7\u30f3\u4e92\u63db\u306e
    SQRT(SUMPRODUCT((range-AVERAGE(range))^2)/COUNT(range)) \u306b\u7f6e\u63db\u3059\u308b\u3002
    range \u90e8\u5206\u306f\u62ec\u5f27\u306e\u6df1\u3055\u3092\u8ffd\u3063\u3066\u6b63\u78ba\u306b\u53d6\u308a\u51fa\u3059\u3002"""
    pattern = re.compile(r"\bSTDEV(?:\.P|P)\s*\(", re.IGNORECASE)
    result = []
    i = 0
    while i < len(formula):
        m = pattern.search(formula, i)
        if not m:
            result.append(formula[i:])
            break
        result.append(formula[i:m.start()])
        # \u62ec\u5f27\u5185\u306e range \u3092\u53d6\u308a\u51fa\u3059
        depth = 1
        j = m.end()
        while j < len(formula) and depth > 0:
            if formula[j] == "(":
                depth += 1
            elif formula[j] == ")":
                depth -= 1
            j += 1
        rng = formula[m.end():j - 1]
        avg = f"AVERAGE({rng})"
        replacement = f"SQRT(SUMPRODUCT(({rng}-{avg})^2)/COUNT({rng}))"
        result.append(replacement)
        i = j
    return "".join(result)


def normalize_excel_formula(value: str) -> str:
    """Normalize LLM-generated Excel formulas before writing them with openpyxl."""
    if not value.startswith("="):
        return value

    # \u5168\u89d2@\u3068\u534a\u89d2@\u306e\u6697\u9ed9\u4ea4\u5dee\u6f14\u7b97\u5b50\u3092\u9664\u53bb
    normalized = value.replace("\uff20", "@")
    normalized = re.sub(r"@(?=[A-Za-z_])", "", normalized)
    # STDEV.P / STDEVP \u2192 SQRT(SUMPRODUCT(...)) \u306b\u5909\u63db\uff08\u5168Excel\u30d0\u30fc\u30b8\u30e7\u30f3\u5bfe\u5fdc\uff09
    normalized = _replace_stdev_with_sqrt_sumproduct(normalized)
    return normalized


def parse_hex_color(value: str | None) -> str | None:
    """6桁 RRGGBB → openpyxl 用 ARGB (FF + RRGGBB)。無効な値は None を返す。"""
    if not value:
        return None
    normalized = str(value).replace("#", "").strip().upper()
    if len(normalized) == 6 and all(ch in "0123456789ABCDEF" for ch in normalized):
        return "FF" + normalized
    return None


def apply_sheet_edits(ws, edits: dict) -> int:
    """setCells / replaceText をシートに適用。変更したセル数を返す。"""
    changed = 0

    for cell_edit in edits.get("setCells") or []:
        address = str(cell_edit.get("address") or "").strip()
        value = cell_edit.get("value")
        if address and value is not None:
            # Excel数式中の @ (暗黙交差演算子) を除去。LLMが誤って付けることがある
            if isinstance(value, str) and value.startswith("="):
                value = normalize_excel_formula(value)
            ws[address] = value
            changed += 1

    replacements = edits.get("replaceText") or []
    if replacements:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                original = cell.value
                new_val = original
                for rep in replacements:
                    find = rep.get("find", "")
                    replace = rep.get("replace", "")
                    if find and find in new_val:
                        new_val = new_val.replace(find, replace)
                if new_val != original:
                    cell.value = new_val
                    changed += 1

    return changed


def apply_format_edits(ws, fmt: dict) -> int:
    """1つの formatEdit エントリを適用。変更したセル数を返す。"""
    range_str = str(fmt.get("range") or "").strip()
    if not range_str:
        return 0

    bold = fmt.get("bold")
    font_color = parse_hex_color(fmt.get("fontColor"))
    fill_color = parse_hex_color(fmt.get("fillColor"))

    if bold is None and font_color is None and fill_color is None:
        return 0

    changed = 0
    try:
        cell_range = ws[range_str]
        # 単一セルの場合は list にする
        if not isinstance(cell_range, (list, tuple)):
            cell_range = [[cell_range]]
        for row in cell_range:
            row_cells = row if isinstance(row, (list, tuple)) else [row]
            for cell in row_cells:
                if bold is not None or font_color is not None:
                    f = cell.font if cell.font else Font()
                    kwargs: dict[str, Any] = {}
                    if bold is not None:
                        kwargs["bold"] = bold
                    if font_color is not None:
                        kwargs["color"] = font_color
                    cell.font = f.copy(**kwargs)
                if fill_color is not None:
                    cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
                changed += 1
    except Exception as e:
        print(f"[edit_excel] format range '{range_str}' failed: {e}", file=sys.stderr)

    return changed


def load_workbook_any(input_path: Path) -> openpyxl.Workbook:
    """xlsx / xlsm は openpyxl で直接ロード。xls は xlrd 経由で変換。"""
    suffix = input_path.suffix.lower()

    if suffix in (".xlsx", ".xlsm"):
        keep_vba = suffix == ".xlsm"
        return openpyxl.load_workbook(str(input_path), keep_vba=keep_vba)

    if suffix == ".xls":
        try:
            import xlrd  # type: ignore
        except ImportError:
            raise RuntimeError(
                "xlrd が未インストールのため .xls ファイルを読み込めません。"
                "startup.sh で pip install xlrd を実行してください。"
            )
        book = xlrd.open_workbook(str(input_path))
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # デフォルトシートを削除
        for sheet_idx in range(book.nsheets):
            sheet = book.sheet_by_index(sheet_idx)
            ws = wb.create_sheet(title=sheet.name)
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    ctype = sheet.cell_type(row_idx, col_idx)
                    if ctype == 0:  # empty
                        continue
                    ws.cell(
                        row=row_idx + 1,
                        column=col_idx + 1,
                        value=sheet.cell_value(row_idx, col_idx),
                    )
        return wb

    raise ValueError(f"Unsupported file format: {suffix}")


def resolve_sheet(wb: openpyxl.Workbook, sheet_name: str):
    """シート名で検索。完全一致 → 前方一致の順で返す。見つからなければ None。"""
    if not sheet_name or sheet_name == "*":
        return None  # 全シート対象を呼び元で処理
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    match = next(
        (s for s in wb.sheetnames if s.lower().startswith(sheet_name.lower())), None
    )
    return wb[match] if match else None


def resolve_column_index(ws, col_spec: str) -> int | None:
    """列指定を列番号（1始まり）に解決。
    - "A"〜"XFD" 形式はそのまま変換
    - ヘッダー名の場合は1行目を走査して列番号を返す
    """
    if not col_spec:
        return None
    col_spec = col_spec.strip()
    # アルファベットのみなら列記号として解釈
    if col_spec.isalpha():
        try:
            return column_index_from_string(col_spec)
        except Exception:
            pass
    # ヘッダー名として1行目を検索
    for cell in ws[1]:
        if str(cell.value or "").strip() == col_spec:
            return cell.column
    print(f"[edit_excel] resolve_column_index: '{col_spec}' not found in row 1", file=sys.stderr)
    return None


def apply_border_edits(ws, border_edit: dict) -> int:
    """指定範囲に罫線を適用する。
    edges: "all" | "outer" | "inner" | "top" | "bottom" | "left" | "right"
    style: "thin" | "medium" | "thick" | "hair" | "dashed" (デフォルト "thin")
    """
    range_str = str(border_edit.get("range") or "").strip()
    if not range_str:
        return 0
    style = str(border_edit.get("style") or "thin")
    edges = str(border_edit.get("edges") or "all").lower()
    side = Side(style=style)
    no_side = Side(style=None)

    try:
        cell_range = ws[range_str]
        if not isinstance(cell_range, (list, tuple)):
            cell_range = [[cell_range]]
        # 範囲の行数・列数を把握
        rows = cell_range
        max_r = len(rows) - 1
        changed = 0
        for ri, row in enumerate(rows):
            row_cells = row if isinstance(row, (list, tuple)) else [row]
            max_c = len(row_cells) - 1
            for ci, cell in enumerate(row_cells):
                b = cell.border.copy() if cell.border else Border()
                top = b.top
                bottom = b.bottom
                left = b.left
                right = b.right

                is_top_edge = ri == 0
                is_bottom_edge = ri == max_r
                is_left_edge = ci == 0
                is_right_edge = ci == max_c

                if edges == "all":
                    top = bottom = left = right = side
                elif edges == "outer":
                    if is_top_edge:    top = side
                    if is_bottom_edge: bottom = side
                    if is_left_edge:   left = side
                    if is_right_edge:  right = side
                elif edges == "inner":
                    if not is_top_edge:    top = side
                    if not is_bottom_edge: bottom = side
                    if not is_left_edge:   left = side
                    if not is_right_edge:  right = side
                elif edges == "top":    top = side
                elif edges == "bottom": bottom = side
                elif edges == "left":   left = side
                elif edges == "right":  right = side

                cell.border = Border(top=top, bottom=bottom, left=left, right=right)
                changed += 1
        return changed
    except Exception as e:
        print(f"[edit_excel] border range '{range_str}' failed: {e}", file=sys.stderr)
        return 0


def apply_copy_row_color(ws, target_col_spec: str, reference_col_spec: str, start_row: int = 2) -> int:
    """reference 列の各行の背景色を target 列の同じ行にコピーする。
    col_spec は列記号（"A"）またはヘッダー名（"対応"）のどちらでもよい。
    変更した行数を返す。
    """
    tgt_idx = resolve_column_index(ws, target_col_spec)
    ref_idx = resolve_column_index(ws, reference_col_spec)

    if tgt_idx is None or ref_idx is None:
        print(
            f"[edit_excel] copy_row_color: could not resolve columns "
            f"target='{target_col_spec}' ref='{reference_col_spec}'",
            file=sys.stderr,
        )
        return 0

    changed = 0
    max_row = ws.max_row
    for row_num in range(start_row, max_row + 1):
        ref_cell = ws.cell(row=row_num, column=ref_idx)
        tgt_cell = ws.cell(row=row_num, column=tgt_idx)

        ref_fill = ref_cell.fill
        if ref_fill and ref_fill.fill_type not in (None, "none", ""):
            # PatternFill をコピー（fgColor / bgColor ともにそのまま渡す）
            tgt_cell.fill = PatternFill(
                fill_type=ref_fill.fill_type,
                fgColor=ref_fill.fgColor.rgb if ref_fill.fgColor else "00000000",
                bgColor=ref_fill.bgColor.rgb if ref_fill.bgColor else "00000000",
            )
            changed += 1

    return changed


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
    wb = load_workbook_any(input_path)
    if wb.calculation is not None:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    total_sheets = len(wb.sheetnames)
    changed_sheets: set[str] = set()

    # sheetEdits
    for sheet_edit in plan.get("sheetEdits") or []:
        sheet_name = str(sheet_edit.get("sheetName") or "").strip()
        ws = resolve_sheet(wb, sheet_name)

        if ws is None:
            # "*" または未指定 → 全シート
            targets = wb.worksheets
        else:
            targets = [ws]

        for target_ws in targets:
            n = apply_sheet_edits(target_ws, sheet_edit)
            if n > 0:
                changed_sheets.add(target_ws.title)

    # formatEdits
    for fmt in plan.get("formatEdits") or []:
        sheet_name = str(fmt.get("sheetName") or "").strip()
        ws = resolve_sheet(wb, sheet_name)
        if ws is None:
            print(f"[edit_excel] formatEdits: sheet '{sheet_name}' not found", file=sys.stderr)
            continue
        n = apply_format_edits(ws, fmt)
        if n > 0:
            changed_sheets.add(ws.title)

    # borderEdits
    for be in plan.get("borderEdits") or []:
        sheet_name = str(be.get("sheetName") or "").strip()
        ws = resolve_sheet(wb, sheet_name)
        if ws is None:
            print(f"[edit_excel] borderEdits: sheet '{sheet_name}' not found", file=sys.stderr)
            continue
        n = apply_border_edits(ws, be)
        if n > 0:
            changed_sheets.add(ws.title)

    # copyRowColorEdits: 参照列の背景色を対象列の各行にコピー
    for crc in plan.get("copyRowColorEdits") or []:
        sheet_name = str(crc.get("sheetName") or "").strip()
        ws = resolve_sheet(wb, sheet_name)
        if ws is None:
            print(f"[edit_excel] copyRowColorEdits: sheet '{sheet_name}' not found", file=sys.stderr)
            continue
        target_col = str(crc.get("targetColumn") or "").strip()
        reference_col = str(crc.get("referenceColumn") or "").strip()
        start_row = int(crc.get("startRow") or 2)
        n = apply_copy_row_color(ws, target_col, reference_col, start_row)
        if n > 0:
            changed_sheets.add(ws.title)

    wb.save(str(output_path))

    print(json.dumps({"changedSheets": len(changed_sheets), "totalSheets": total_sheets}))


if __name__ == "__main__":
    main()
