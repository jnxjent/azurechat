import argparse
import json
from pathlib import Path
from typing import Any
import zipfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from lxml import etree


def normalize_hex(value: str | None) -> str | None:
    if not value:
        return None
    normalized = str(value).replace("#", "").strip().upper()
    if len(normalized) == 6 and all(ch in "0123456789ABCDEF" for ch in normalized):
        return normalized
    return None


def rgb_to_hsl(rgb: tuple[int, int, int]) -> tuple[float, float, float]:
    r, g, b = [c / 255.0 for c in rgb]
    max_c = max(r, g, b)
    min_c = min(r, g, b)
    l = (max_c + min_c) / 2
    if max_c == min_c:
        return 0.0, 0.0, l
    d = max_c - min_c
    s = d / (1 - abs(2 * l - 1))
    if max_c == r:
        h = 60 * (((g - b) / d) % 6)
    elif max_c == g:
        h = 60 * (((b - r) / d) + 2)
    else:
        h = 60 * (((r - g) / d) + 4)
    return h % 360, s, l


def hsl_to_rgb(h: float, s: float, l: float) -> tuple[int, int, int]:
    c = (1 - abs(2 * l - 1)) * s
    hh = h / 60
    x = c * (1 - abs((hh % 2) - 1))
    if 0 <= hh < 1:
        r1, g1, b1 = c, x, 0
    elif hh < 2:
        r1, g1, b1 = x, c, 0
    elif hh < 3:
        r1, g1, b1 = 0, c, x
    elif hh < 4:
        r1, g1, b1 = 0, x, c
    elif hh < 5:
        r1, g1, b1 = x, 0, c
    else:
        r1, g1, b1 = c, 0, x
    m = l - c / 2
    return (
        round((r1 + m) * 255),
        round((g1 + m) * 255),
        round((b1 + m) * 255),
    )


def recolor_preserving_tone(rgb: tuple[int, int, int], target_hex: str) -> str:
    hue, sat, lig = rgb_to_hsl(rgb)
    target_rgb = tuple(int(target_hex[i : i + 2], 16) for i in range(0, 6, 2))
    target_hue, _, _ = rgb_to_hsl(target_rgb)
    new_rgb = hsl_to_rgb(target_hue, max(sat, 0.35), lig)
    return "".join(f"{c:02X}" for c in new_rgb)


def shift_lightness(target_hex: str, lightness_delta: float) -> str:
    target_rgb = tuple(int(target_hex[i : i + 2], 16) for i in range(0, 6, 2))
    hue, sat, lig = rgb_to_hsl(target_rgb)
    next_lig = max(0.08, min(0.92, lig + lightness_delta))
    adjusted = hsl_to_rgb(hue, max(sat, 0.35), next_lig)
    return "".join(f"{c:02X}" for c in adjusted)


def is_neutral(rgb: tuple[int, int, int]) -> bool:
    hue, sat, lig = rgb_to_hsl(rgb)
    return sat < 0.12 or lig < 0.08 or lig > 0.95


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def apply_font_face(shape, font_face: str) -> bool:
    changed = False
    if not getattr(shape, "has_text_frame", False):
        return False
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.name != font_face:
                run.font.name = font_face
                changed = True
    return changed


def replace_text(shape, replacements: list[dict[str, str]]) -> bool:
    """テキストを置換する。shape.text = ... は書式・枠をリセットするため使わず、
    各 run のテキストを直接置換して書式を保持する。"""
    if not getattr(shape, "has_text_frame", False):
        return False

    changed = False
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            original_run = run.text
            updated_run = original_run
            for item in replacements:
                find = item.get("find") or ""
                rep = item.get("replace") or ""
                if find:
                    updated_run = updated_run.replace(find, rep)
            if updated_run != original_run:
                run.text = updated_run
                changed = True
    return changed


def try_get_rgb(color_format) -> tuple[int, int, int] | None:
    try:
        if color_format.type == MSO_COLOR_TYPE.RGB and color_format.rgb is not None:
            rgb = str(color_format.rgb)
            return tuple(int(rgb[i : i + 2], 16) for i in range(0, 6, 2))
    except Exception:
        return None
    return None


def recolor_fill(shape, target_hex: str) -> bool:
    try:
        fill = shape.fill
        if fill.type != MSO_FILL_TYPE.SOLID:
            return False
        rgb = try_get_rgb(fill.fore_color)
        if not rgb or is_neutral(rgb):
            return False
        fill.fore_color.rgb = RGBColor.from_string(recolor_preserving_tone(rgb, target_hex))
        return True
    except Exception:
        return False


def recolor_line(shape, target_hex: str) -> bool:
    """既存の枠線色を直接 XML で書き換える。
    python-pptx の shape.line アクセサは使わない
    （アクセスだけで <a:ln> が生成される／<a:noFill> が <a:solidFill> に変換されるため）。
    <a:ln><a:solidFill><a:srgbClr> が明示されている場合のみ処理。"""
    try:
        _ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        el = getattr(shape, "element", None)
        if el is None:
            return False
        ln_el = el.find(f".//{{{_ns}}}ln")
        if ln_el is None:
            return False
        solid = ln_el.find(f"{{{_ns}}}solidFill")
        if solid is None:
            return False  # noFill や空の <a:ln> はスキップ
        srgb = solid.find(f"{{{_ns}}}srgbClr")
        if srgb is None:
            return False  # schemeClr 等はスキップ
        val = srgb.get("val", "")
        if len(val) != 6:
            return False
        try:
            rgb = tuple(int(val[i : i + 2], 16) for i in range(0, 6, 2))
        except ValueError:
            return False
        if is_neutral(rgb):
            return False
        srgb.set("val", recolor_preserving_tone(rgb, target_hex))
        return True
    except Exception:
        return False


def recolor_slide_background(slide, target_hex: str) -> bool:
    """Recolor the slide's background fill (<p:bg>) if it is a non-neutral solid color."""
    try:
        background = slide.background
        fill = background.fill
        if fill.type != MSO_FILL_TYPE.SOLID:
            return False
        rgb = try_get_rgb(fill.fore_color)
        if not rgb or is_neutral(rgb):
            return False
        fill.fore_color.rgb = RGBColor.from_string(recolor_preserving_tone(rgb, target_hex))
        return True
    except Exception:
        return False


def recolor_table(shape, target_hex: str) -> bool:
    if not hasattr(shape, "has_table") or not shape.has_table:
        return False
    changed = False
    for row in shape.table.rows:
        for cell in row.cells:
            try:
                fill = cell.fill
                if fill.type == MSO_FILL_TYPE.SOLID:
                    rgb = try_get_rgb(fill.fore_color)
                    if rgb and not is_neutral(rgb):
                        fill.fore_color.rgb = RGBColor.from_string(
                            recolor_preserving_tone(rgb, target_hex)
                        )
                        changed = True
            except Exception:
                pass
    return changed


def update_theme_colors(pptx_path: Path, target_hex: str) -> None:
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    variant_map = {
        "accent1": target_hex,
        "accent2": shift_lightness(target_hex, 0.12),
        "accent3": shift_lightness(target_hex, -0.10),
        "accent4": shift_lightness(target_hex, 0.2),
        "accent5": shift_lightness(target_hex, -0.18),
        "accent6": shift_lightness(target_hex, 0.04),
        "hlink": target_hex,
        "folHlink": shift_lightness(target_hex, -0.12),
    }

    with zipfile.ZipFile(pptx_path, "r") as zin:
        entries: dict[str, bytes] = {}
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/theme/") and item.filename.endswith(".xml"):
                root = etree.fromstring(data)
                updated = False
                for name, color_hex in variant_map.items():
                    node = root.find(f".//a:clrScheme/a:{name}", ns)
                    if node is None:
                        continue
                    for child in list(node):
                        node.remove(child)
                    srgb = etree.SubElement(
                        node,
                        "{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr",
                    )
                    srgb.set("val", color_hex)
                    updated = True
                if updated:
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone="yes"
                    )
            entries[item.filename] = data

    with zipfile.ZipFile(pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for filename, data in entries.items():
            zout.writestr(filename, data)


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
    prs = Presentation(str(input_path))

    deck_edits = plan.get("deckEdits") or {}
    target_hex = normalize_hex(deck_edits.get("accentColor"))
    font_face = (deck_edits.get("fontFace") or "").strip() or None

    slide_edit_map = {
        int(item.get("slideIndex")): item
        for item in (plan.get("slideEdits") or [])
        if item.get("slideIndex") is not None
    }

    changed_slides: set[int] = set()

    for slide_index, slide in enumerate(prs.slides):
        slide_changed = False
        slide_edit = slide_edit_map.get(slide_index) or {}
        replacements = slide_edit.get("replaceText") or []

        if target_hex:
            if recolor_slide_background(slide, target_hex):
                slide_changed = True

        for shape in iter_shapes(slide.shapes):
            if target_hex:
                if recolor_fill(shape, target_hex):
                    slide_changed = True
                if recolor_line(shape, target_hex):
                    slide_changed = True
                if recolor_table(shape, target_hex):
                    slide_changed = True
            if font_face and apply_font_face(shape, font_face):
                slide_changed = True
            if replacements and replace_text(shape, replacements):
                slide_changed = True

        if slide_changed:
            changed_slides.add(slide_index)

    prs.save(str(output_path))
    if target_hex:
        update_theme_colors(output_path, target_hex)

    result = {
      "changedSlides": len(changed_slides),
      "totalSlides": len(prs.slides),
    }
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
