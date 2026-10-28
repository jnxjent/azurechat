"""Structural PPTX edits that the existing edit_pptx.py intentionally rejects.

Operations (plan["ops"], applied in order):
  add_slide   insert one content slide that reuses an existing slide's frame
  font_size   change font sizes of title / body / table text
  fill_slide  write a title and body into an existing empty slide

Existing slides are never re-generated. Only the inserted slide, or the
runs whose size is changed, differ from the input file.
"""
import argparse
import json
import re
import sys
import math
from copy import deepcopy

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt, Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN

MIN_PT = 10
MAX_PT = 44
DEFAULT_BODY_PT = 18
DEFAULT_TITLE_PT = 28
FOOTER_TYPES = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE}
TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def placeholder_type(shape):
    try:
        return shape.placeholder_format.type if shape.is_placeholder else None
    except Exception:
        return None


def shape_text(shape) -> str:
    # Read the XML directly: shape.text_frame adds an empty <p:txBody> to
    # autoshapes without text, which would modify slides we only inspect.
    body = shape._element.find(qn("p:txBody"))
    if body is None:
        return ""
    return "".join(t.text or "" for t in body.iter(qn("a:t"))).strip()


def same(a, b) -> bool:
    # python-pptx returns a new proxy object on every access; compare elements.
    return a is not None and b is not None and a._element is b._element


class RefFont:
    """Read-only run formatting. python-pptx's run.font adds an <a:rPr/> to
    runs that lack one, which would modify the slides we only read from."""

    def __init__(self, run):
        rpr = run._r.find(qn("a:rPr"))
        sz = rpr.get("sz") if rpr is not None else None
        self.size = Pt(int(sz) / 100) if sz and sz.isdigit() else None
        latin = rpr.find(qn("a:latin")) if rpr is not None else None
        self.name = latin.get("typeface") if latin is not None and not str(latin.get("typeface", "")).startswith("+") else None
        srgb = rpr.find(qn("a:solidFill") + "/" + qn("a:srgbClr")) if rpr is not None else None
        self.rgb = RGBColor.from_string(srgb.get("val")) if srgb is not None and len(srgb.get("val", "")) == 6 else None


def max_run_pt(shape) -> float:
    best = 0.0
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                size = RefFont(r).size
                if size:
                    best = max(best, size.pt)
    return best


def is_footer_like(shape, slide_h: int) -> bool:
    if placeholder_type(shape) in FOOTER_TYPES:
        return True
    text = shape_text(shape)
    top = shape.top or 0
    return bool(text) and top > slide_h * 0.88 and len(text) <= 40


def find_title(slide, slide_h: int):
    for shape in slide.shapes:
        if placeholder_type(shape) in TITLE_TYPES and shape_text(shape):
            return shape
    # Without a title placeholder, the slide heading is the top-most text in the
    # header area; a larger heading lower down is a body column heading.
    candidates = [s for s in slide.shapes if shape_text(s) and (s.top or 0) < slide_h * 0.2 and not is_footer_like(s, slide_h)]
    if not candidates:
        # Cover slides center their title: fall back to the largest text.
        texts = [s for s in slide.shapes if shape_text(s) and not is_footer_like(s, slide_h) and max_run_pt(s) >= 24]
        return max(texts, key=max_run_pt) if texts else None
    return min(candidates, key=lambda s: (round((s.top or 0) / (slide_h * 0.03)), -max_run_pt(s)))


def is_edge_decoration(shape, slide_w: int, slide_h: int) -> bool:
    """Thin bars/lines touching a slide edge belong to the frame, not the body."""
    if shape_text(shape):
        return False
    left, top = shape.left or 0, shape.top or 0
    width, height = shape.width or 0, shape.height or 0
    thin = width <= slide_w * 0.03 or height <= slide_h * 0.03
    touches = left <= slide_w * 0.02 or top <= slide_h * 0.02 or left + width >= slide_w * 0.98 or top + height >= slide_h * 0.98
    return thin and touches


def body_text_shapes(slide, slide_h: int):
    title = find_title(slide, slide_h)
    return [s for s in slide.shapes if not same(s, title) and shape_text(s) and not is_footer_like(s, slide_h)]


def first_run_font(shapes):
    for shape in shapes:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text.strip():
                    return RefFont(r)
    return None


def choose_reference(prs, slide_h: int):
    """Prefer a middle content slide that has a title and body text."""
    slides = list(prs.slides)
    # The cover (first slide) is the last resort: its frame differs from content slides.
    ordered = slides[1:-1] + slides[-1:] + slides[:1] if len(slides) > 2 else slides[1:] + slides[:1]
    for slide in ordered:
        if find_title(slide, slide_h) is not None and body_text_shapes(slide, slide_h):
            return slide
    for slide in ordered:
        if find_title(slide, slide_h) is not None:
            return slide
    return slides[0] if slides else None


def move_slide(prs, old_index: int, new_index: int):
    ids = prs.slides._sldIdLst
    items = list(ids)
    element = items[old_index]
    ids.remove(element)
    ids.insert(new_index, element)


def clone_frame(prs, ref, slide_h: int):
    """Copy the reference slide's frame (background, header band, logo, footer)
    but drop its body content, title text and charts."""
    new = prs.slides.add_slide(ref.slide_layout)
    for shape in list(new.shapes):
        shape._element.getparent().remove(shape._element)
    rid_map = {}
    for rid, rel in ref.part.rels.items():
        if rel.reltype.endswith("/slideLayout") or rel.reltype.endswith("/notesSlide") or rel.reltype.endswith("/chart"):
            continue
        if rel.is_external:
            rid_map[rid] = new.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rid_map[rid] = new.part.relate_to(rel.target_part, rel.reltype)
    bg = ref._element.cSld.find(qn("p:bg"))
    if bg is not None:
        new._element.cSld.insert(0, deepcopy(bg))
    title = find_title(ref, slide_h)
    title_bottom = (title.top + title.height) if title is not None else int(slide_h * 0.18)
    body_top = title_bottom
    for shape in ref.shapes:
        # Tables/charts (graphic frames, also inside groups) are body content, and
        # chart parts must not be shared between slides.
        if any(True for _ in shape._element.iter(qn("p:graphicFrame"))):
            continue
        top = shape.top or 0
        is_title = same(shape, title)
        keep = (is_title or is_footer_like(shape, slide_h) or top + (shape.height or 0) <= body_top
                or top >= slide_h * 0.9 or is_edge_decoration(shape, prs.slide_width, slide_h)
                or (not shape_text(shape) and top < slide_h * 0.2 and (shape.height or 0) <= slide_h * 0.03))
        if not keep:
            continue
        element = deepcopy(shape._element)
        for node in element.iter():
            for attr in list(node.attrib):
                if attr.startswith("{" + REL_NS + "}") and node.attrib[attr] in rid_map:
                    node.attrib[attr] = rid_map[node.attrib[attr]]
        new.shapes._spTree.append(element)
    return new, body_top


def new_title_shape(slide, slide_h: int):
    return find_title(slide, slide_h)


def set_shape_text(shape, text: str):
    tf = shape.text_frame
    first = tf.paragraphs[0]
    keep_rpr = None
    if first.runs and first.runs[0]._r.find(qn("a:rPr")) is not None:
        keep_rpr = deepcopy(first.runs[0]._r.find(qn("a:rPr")))
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for r in list(first.runs):
        r._r.getparent().remove(r._r)
    run = first.add_run()
    run.text = text
    if keep_rpr is not None:
        old = run._r.find(qn("a:rPr"))
        if old is not None:
            run._r.remove(old)
        run._r.insert(0, keep_rpr)


def apply_font(run, ref_font, size_pt: float, bold=None):
    run.font.size = Pt(size_pt)
    if ref_font is not None:
        if ref_font.name:
            run.font.name = ref_font.name
        if ref_font.rgb is not None:
            run.font.color.rgb = ref_font.rgb
    if bold is not None:
        run.font.bold = bold


def text_units(text: str) -> float:
    # Full-width characters take one em, ASCII roughly half.
    return sum(1.0 if ord(ch) > 0x2E7F else 0.55 for ch in text)


def lines_needed(texts, width_emu: int, size_pt: float, indent_em: float = 1.2) -> int:
    per_line = max(4.0, Emu(width_emu).pt / size_pt - indent_em)
    return sum(max(1, -(-int(text_units(t) * 100) // int(per_line * 100))) for t in texts)


def fit_bullets(bullets, width: int, height: int, start_pt: float):
    """Return (size, kept_bullets, omitted_count); never silently truncates text."""
    size = start_pt
    kept = list(bullets)
    while True:
        need = lines_needed(kept, width, size) * size * 1.45 + len(kept) * size * 0.4
        if need <= Emu(height).pt:
            return size, kept, len(bullets) - len(kept)
        if size > min(14, start_pt):
            size -= 1
            continue
        if len(kept) <= 1:
            return size, kept, len(bullets) - len(kept)
        kept = kept[:-1]


def add_bullets_box(slide, left, top, width, height, bullets, ref_font, start_pt):
    size, kept, omitted = fit_bullets(bullets, width, height, start_pt)
    box = slide.shapes.add_textbox(left, top, width, height)
    box.name = "AzureChat Added Body"
    tf = box.text_frame
    tf.word_wrap = True
    for i, text in enumerate(kept):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        ppr = p._p.get_or_add_pPr()
        ppr.set("marL", str(int(Pt(size * 1.2))))
        ppr.set("indent", str(-int(Pt(size * 1.2))))
        bu = ppr.makeelement(qn("a:buChar"), {"char": "•"})
        ppr.append(bu)
        p.space_after = Pt(size * 0.4)
        run = p.add_run()
        run.text = text
        apply_font(run, ref_font, size)
    return box, size, omitted


def font_rgb(font):
    return font.rgb if font is not None else None


def add_table(slide, left, top, width, height, columns, rows, ref_font, start_pt, header_rgb=None, warnings=None):
    columns = [str(c)[:30] for c in columns][:5]
    rows = [[str(c)[:80] for c in row][: len(columns)] + [""] * max(0, len(columns) - len(row)) for row in rows][:6]
    size = min(start_pt, 14)
    while size > body_min_pt(ref_font):
        row_lines = [max(lines_needed([cell], width // len(columns), size, 0.6) for cell in row) for row in rows]
        need = (sum(row_lines) + 1) * size * 1.5 + (len(rows) + 1) * size * 0.6
        if need <= Emu(height).pt:
            break
        size -= 1
    heights = lambda: [int(Pt(max(1,lines_needed(columns,width//len(columns),size,.6))*size*1.4+size*.6))] + [int(Pt(max(lines_needed([cell],width//len(columns),size,.6) for cell in row)*size*1.4+size*.6)) for row in rows]
    omitted = 0
    while rows and sum(heights()) > height:
        rows.pop()
        omitted += 1
    if not rows:
        raise ValueError("14ptの文字を維持したまま表を配置できません。内容を分けてください。")
    if omitted and warnings is not None:
        warnings.append(f"文字サイズを維持するため、表の{omitted}行を省略しました。")
    row_heights = heights()
    shape = slide.shapes.add_table(len(rows) + 1, len(columns), left, top, width, sum(row_heights))
    for i, value in enumerate(row_heights):
        shape.table.rows[i].height = value
    shape.name = "AzureChat Added Table"
    table = shape.table
    for c, name in enumerate(columns):
        cell = table.cell(0, c)
        cell.text = name
        if header_rgb is not None:
            # Match the deck's heading color instead of the theme's default blue.
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_rgb
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                apply_font(r, None, size, bold=True)
                if ref_font is not None and ref_font.name:
                    r.font.name = ref_font.name
    for r_i, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            cell = table.cell(r_i, c)
            cell.text = value
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    # Keep the table style's text colors (readable on its row fills).
                    apply_font(run, None, size)
                    if ref_font is not None and ref_font.name:
                        run.font.name = ref_font.name
    return shape, size


def add_sources(slide, left, top, width, sources, ref_font):
    box = slide.shapes.add_textbox(left, top, width, Pt(14 * max(1, len(sources))))
    box.name = "AzureChat Added Sources"
    tf = box.text_frame
    tf.word_wrap = True
    for i, src in enumerate(sources[:3]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        label = str(src.get("label") or src.get("url"))[:80]
        run.text = ("出典: " if i == 0 else "　　  ") + label
        apply_font(run, ref_font, 10)
        url = str(src.get("url") or "")
        if re.match(r"^https?://", url):
            run.hyperlink.address = url
    return box


# Structural additions only: existing font-size edit limits remain unchanged.
OUTER_MARGIN = Inches(0.5)
BLOCK_GAP = Inches(0.3)

def theme_colors(slide):
    result = {}
    try:
        master = slide.slide_layout.slide_master.part
        themes = [rel.target_part for rel in master.rels.values() if rel.reltype.endswith("/theme")]
        if themes:
            from lxml import etree
            root = etree.fromstring(themes[0].blob)
            for name in ("accent1", "dk2", "lt2"):
                element = root.find(".//" + qn("a:clrScheme") + "/" + qn("a:" + name))
                if element is not None and len(element):
                    value = element[0].get("val") or element[0].get("lastClr")
                    if value and re.fullmatch(r"[0-9a-fA-F]{6}", value):
                        result[name] = RGBColor.from_string(value)
    except (AttributeError, KeyError, ValueError):
        pass
    return result


def palette(slide, primary=None):
    colors = theme_colors(slide)
    main = primary if primary is not None and max(primary)-min(primary)>25 else colors.get("accent1") or primary or RGBColor(40, 92, 80)
    pale = RGBColor(*(round(c * .07 + 255 * .93) for c in main))
    return main, pale, colors.get("dk2") or RGBColor(40, 48, 48)


def body_min_pt(ref_font):
    size = ref_font.size.pt if ref_font is not None and ref_font.size else 14
    return min(14, max(10, size))


def design_text(slide, x, y, w, h, text, size, ref_font, color, bold=False, warnings=None, name="AzureChat Design Text"):
    text = str(text)
    # Never solve a crowded block by shrinking below the requested design size.
    original = text
    while text and lines_needed(text.split("\n"), w, size, .15) * size * 1.35 > Emu(h).pt:
        text = text[:-1]
    if text != original:
        text = text.rstrip() + "…"
        if warnings is not None:
            warnings.append("文字サイズを維持するため、追加スライドの長い文章を一部省略しました。")
    box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(text.split("\n")):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = line
        apply_font(run, ref_font, size, bold)
        run.font.color.rgb = color
    return box


def add_stats(slide, left, top, width, height, items, ref_font, primary, warnings):
    primary, pale, dark = palette(slide, primary)
    items = items[:4]
    if len(items) < 2:
        raise ValueError("数値パネルには根拠のある数値が2つ以上必要です。")
    gap = BLOCK_GAP
    card_w = (width - gap * (len(items) - 1)) // len(items)
    for i, item in enumerate(items):
        x = left + i * (card_w + gap)
        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, int(x), int(top), int(card_w), int(height))
        card.name = "AzureChat Stat Panel"
        card.fill.solid()
        card.fill.fore_color.rgb = pale
        card.line.fill.background()
        pad = Pt(14)
        design_text(slide, x+pad, top+pad, card_w-2*pad, height*.2, item["label"], 16, ref_font, dark, True, warnings)
        design_text(slide, x+pad, top+height*.28, card_w-2*pad, height*.3, item["value"], 32, ref_font, primary, True, warnings, "AzureChat Stat Value")
        basis = " / ".join(str(item.get(key) or "") for key in ("period", "unit") if item.get(key))
        design_text(slide, x+pad, top+height*.65, card_w-2*pad, height*.25, basis, 11, ref_font, dark, False, warnings, "AzureChat Stat Basis")


def company_card_column(columns):
    return next((i for i, name in enumerate(columns) if re.search(r"企業|会社|プレイヤー", str(name))), None)


def add_company_cards(slide, left, top, width, height, columns, rows, ref_font, header_rgb=None, warnings=None):
    """Editable company profiles; no decorative stripes or invented ranks."""
    warnings = warnings if warnings is not None else []
    primary, pale, dark = palette(slide, header_rgb)
    name_col = company_card_column(columns)
    gap = BLOCK_GAP
    count = len(rows)
    cols = min(3, count)
    nrows = (count + cols - 1) // cols
    card_w = (width - gap * (cols - 1)) // cols
    card_h = (height - gap * (nrows - 1)) // nrows
    size = body_min_pt(ref_font)
    for i, row in enumerate(rows):
        r, c = divmod(i, cols)
        row_count = min(cols, count-r*cols)
        offset = (width-(card_w*row_count+gap*(row_count-1)))//2
        x, y = left+offset+c*(card_w+gap), top+r*(card_h+gap)
        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, int(x), int(y), int(card_w), int(card_h))
        card.name = "AzureChat Company Card"
        card.fill.solid()
        card.fill.fore_color.rgb = pale
        card.line.fill.background()
        pad = Pt(12)
        inner_w = card_w-2*pad
        name = str(row[name_col])
        design_text(slide,x+pad,y+pad,inner_w,card_h*.26,name,18,ref_font,primary,True,warnings,"AzureChat Company Card Text")
        period_col = next((j for j,label in enumerate(columns) if re.search(r"年度|決算期|対象期間", str(label)) and not re.search(r"売上|数量|件数|金額|人数", str(label))), None)
        numeric = next((j for j, label in enumerate(columns) if j != name_col and re.search(r"売上|数量|件数|金額|人数", str(label))
                        and j < len(row) and re.fullmatch(r"[0-9][0-9,.]*(?:億円|百万円|万円|千円|円|件|人|台|トン)?", str(row[j]))
                        and re.search(r"(?:19|20)\d{2}.*(?:年|期)", str(label) + (str(row[period_col]) if period_col is not None and period_col<len(row) else ""))
                        and (re.search(r"[（(].+[）)]", str(label)) or re.search(r"億円|百万円|万円|千円|円|件|人|台|トン",str(row[j])))), None)
        body_y = y+card_h*.33
        details = "\n".join(f"{label}：{row[j]}" for j,label in enumerate(columns) if j not in (name_col,numeric,period_col if numeric is not None else None) and j<len(row) and str(row[j]).strip())
        if numeric is not None:
            match = re.fullmatch(r"([0-9][0-9,.]*)(.*)",str(row[numeric]))
            metric_value, metric_unit = match.group(1), match.group(2)
            design_text(slide,x+pad,body_y,inner_w,card_h*.28,metric_value,30,ref_font,primary,True,warnings,"AzureChat Company Metric")
            design_text(slide,x+pad,y+card_h*.64,inner_w,card_h*.22,str(columns[numeric]) + (" / " + str(row[period_col]) if period_col is not None and period_col<len(row) else "") + (" / " + metric_unit if metric_unit else ""),11,ref_font,dark,False,warnings,"AzureChat Metric Basis")
            body_y = y+card_h*.87
        if details:
            design_text(slide,x+pad,body_y,inner_w,max(Pt(16),y+card_h-pad-body_y),details,size,ref_font,dark,False,warnings,"AzureChat Company Card Text")


def add_visual_composition(slide, left, top, width, height, plan, ref_font, primary, warnings=None):
    """Render the same source-table composition selected by the TS planner.
    Shapes/text remain editable. No label, score or numerical value is invented.
    """
    columns, rows = plan["columns"], plan["rows"]
    primary = palette(slide, primary)[0]
    accent = primary
    for sh in slide.shapes:
        fill = sh._element.find(qn("p:spPr") + "/" + qn("a:solidFill") + "/" + qn("a:srgbClr"))
        if fill is not None and re.fullmatch(r"[0-9a-fA-F]{6}", fill.get("val", "")):
            rgb = RGBColor.from_string(fill.get("val"))
            if max(rgb) > 160 and min(rgb) < 140 and max(rgb) - min(rgb) > 40:
                accent = rgb
                break

    def rectangle(x, y, w, h, color, rounded=False):
        sh = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if rounded else MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                    int(x), int(y), max(1, int(w)), max(1, int(h)))
        sh.name = "AzureChat Visual " + plan["kind"]
        sh.fill.solid()
        sh.fill.fore_color.rgb = color
        sh.line.color.rgb = palette(slide, primary)[1]
        return sh

    def text(value, x, y, w, h, size=16, color=None, bold=False):
        return design_text(slide,x,y,w,h,value,max(body_min_pt(ref_font),size),ref_font,color or palette(slide,primary)[2],bold,warnings,
                           name="AzureChat Visual Text")

    if plan["kind"] == "bar":
        data = CategoryChartData()
        data.categories = [str(row[0]) for row in rows]
        data.add_series(columns[1], plan["values"])
        design_text(slide,left,top,width,Pt(26),columns[1],14,ref_font,palette(slide,primary)[2],False,warnings,"AzureChat Chart Basis")
        top += Pt(32)
        height -= Pt(32)
        chart_shape = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, int(left), int(top), int(width), int(height), data)
        chart_shape.name = "AzureChat Visual Native Bar"
        chart = chart_shape.chart
        chart.has_legend = False
        chart.has_title = False
        plot = chart.plots[0]
        plot.has_data_labels = True
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        plot.data_labels.font.size = Pt(14)
        series = chart.series[0]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = primary
        series.format.line.fill.background()
        chart.value_axis.minimum_scale = 0
        chart.value_axis.has_major_gridlines = True
        chart.value_axis.major_gridlines.format.line.color.rgb = palette(slide, primary)[1]
        chart.category_axis.tick_labels.font.size = Pt(14)
        chart.value_axis.tick_labels.font.size = Pt(10)
        return
    if plan["kind"] == "comparison":
        gap = BLOCK_GAP
        col_w = (width - gap) / 2
        header_h = height * .16
        row_h = (height - header_h) / len(rows)
        for side, col in enumerate((1, 2)):
            x = left + side * (col_w + gap)
            color = primary if side else accent
            rectangle(x, top, col_w, header_h, color)
            text(columns[col], x + Pt(12), top + Pt(7), col_w - Pt(24), header_h - Pt(14), 22, RGBColor(255, 255, 255), True)
            for i, row in enumerate(rows):
                y = top + header_h + i * row_h
                rectangle(x, y, col_w, row_h, palette(slide, primary)[1])
                text(row[0], x + Pt(12), y + Pt(6), col_w - Pt(24), row_h * .25, 11, color, True)
                text(row[col], x + Pt(12), y + row_h * .32, col_w - Pt(24), row_h * .58, 17)
        return
    gap = BLOCK_GAP
    card_w = (width - gap * (len(rows) - 1)) / len(rows)
    for i, row in enumerate(rows):
        x = left + i * (card_w + gap)
        rectangle(x, top, card_w, height, palette(slide, primary)[1], True)
        badge = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, int(x + Pt(12)), int(top + Pt(12)), Pt(30), Pt(30))
        badge.fill.solid()
        badge.fill.fore_color.rgb = primary
        badge.line.fill.background()
        text(str(i + 1) if plan["kind"] == "process" else "•", x + Pt(20), top + Pt(16), Pt(18), Pt(22), 16, RGBColor(255, 255, 255), True)
        text(row[0], x + Pt(12), top + height * .25, card_w - Pt(24), height * .2, 21, bold=True)
        text(row[1], x + Pt(12), top + height * .5, card_w - Pt(24), height * .42, 16)
        if plan["kind"] == "process" and i < len(rows) - 1:
            text("›", x + card_w, top + height * .43, gap, Pt(25), 23)


def write_content(slide, slide_w, slide_h, body_top, op, ref_body_font, warnings, slide_no, header_rgb=None):
    margin_x = int(max(OUTER_MARGIN, slide_w * 0.04))
    left, width = margin_x, slide_w - margin_x * 2
    top = body_top + BLOCK_GAP
    sources = [s for s in (op.get("sources") or []) if isinstance(s, dict) and s.get("url")]
    bottom = int(slide_h * 0.88) - (int(Pt(14 * min(3, len(sources)) + 6)) if sources else 0)
    caveats = [str(c).strip() for c in (op.get("caveats") or []) if str(c).strip()][:2]
    caveat_height = int(Pt(38)) if caveats else 0
    primary, pale, dark = palette(slide, header_rgb)
    key_message = str(op.get("keyMessage") or "").strip()
    if key_message:
        design_text(slide,left,top,width,Pt(38),key_message,22,ref_body_font,primary,True,warnings,"AzureChat Key Message")
        top += Pt(38) + BLOCK_GAP
    height = max(int(slide_h * 0.2), bottom - top - caveat_height)
    ref_size = ref_body_font.size.pt if ref_body_font is not None and ref_body_font.size else DEFAULT_BODY_PT
    ref_size = max(MIN_PT, min(24, ref_size))
    if op.get("kind") == "stats":
        add_stats(slide,left,top,width,height,op.get("stats") or [],ref_body_font,primary,warnings)
    elif op.get("kind") in ("table", "two_column") and op.get("table"):
        if op.get("kind") == "two_column":
            left_w = (width - BLOCK_GAP) * .38
            design_text(slide,left,top,left_w,height,"\n\n".join(op.get("bullets") or []),max(14,body_min_pt(ref_body_font)),ref_body_font,dark,False,warnings,"AzureChat Two Column Points")
            left += left_w + BLOCK_GAP
            width -= left_w + BLOCK_GAP
        table = op["table"]
        columns, rows = table.get("columns") or [], table.get("rows") or []
        composition = op.get("composition")
        if (isinstance(composition, dict) and composition.get("kind") in ("process", "classification", "comparison", "bar")
                and composition.get("columns") == columns and composition.get("rows") == rows):
            add_visual_composition(slide, left, top, width, height, composition, ref_body_font, primary, warnings)
        elif company_card_column(columns) is not None and 1 <= len(rows) <= 6 and len(columns) <= 3:
            add_company_cards(slide, left, top, width, height, columns, rows, ref_body_font, primary, warnings)
        else:
            add_table(slide, left, top, width, height, columns, rows, ref_body_font, ref_size, primary, warnings)
        if len(table.get("rows") or []) > 6 or len(table.get("columns") or []) > 5:
            warnings.append(f"P{slide_no}: 表は6行×5列までのため、超えた行・列を省略しました。")
    elif op.get("bullets"):
        bullets = [str(b).strip() for b in op["bullets"] if str(b).strip()][:6]
        panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, int(left), int(top), int(width), int(height))
        panel.name = "AzureChat Summary Panel"
        panel.fill.solid()
        panel.fill.fore_color.rgb = pale
        panel.line.fill.background()
        left, top, width, height = left+Pt(14), top+Pt(14), width-Pt(28), height-Pt(28)
        design_text(slide,left,top,width,height,"\n\n".join(bullets),max(body_min_pt(ref_body_font),min(18,ref_size)),ref_body_font,dark,False,warnings,"AzureChat Summary Text")
        omitted = 0
        if len(op["bullets"]) > 6:
            omitted += len(op["bullets"]) - 6
        if omitted:
            warnings.append(f"P{slide_no}: 収まらないため箇条書き{omitted}項目を省略しました。")
    left, width = margin_x, slide_w - margin_x * 2
    if caveats:
        box = slide.shapes.add_textbox(left, bottom - caveat_height, width, caveat_height)
        box.name = "AzureChat Qualifications"
        set_shape_text(box, "\n".join(caveats))
        for paragraph in box.text_frame.paragraphs:
            for run in paragraph.runs:
                apply_font(run, ref_body_font, 11)
    if sources:
        add_sources(slide, left, bottom + int(Pt(4)), width, sources, ref_body_font)


def op_add_slide(prs, op, warnings):
    slide_w, slide_h = prs.slide_width, prs.slide_height
    ref = choose_reference(prs, slide_h)
    if ref is None:
        raise ValueError("参照できるスライドがありません。")
    ref_body_font = first_run_font(body_text_shapes(ref, slide_h))
    new, body_top = clone_frame(prs, ref, slide_h)
    title = new_title_shape(new, slide_h)
    title_text = str(op.get("title") or "").strip()[:60].rstrip("。．.")
    if title is not None:
        set_shape_text(title, title_text)
    elif title_text:
        box = new.shapes.add_textbox(int(slide_w * 0.06), int(slide_h * 0.05), int(slide_w * 0.88), int(slide_h * 0.12))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = title_text
        apply_font(run, ref_body_font, DEFAULT_TITLE_PT, bold=True)
    count = len(prs.slides)
    position = op.get("position")
    if isinstance(position, int) and position < 0:
        position = count + position  # -2 = second from the end, counting the new slide
    position = count - 1 if not isinstance(position, int) else max(0, min(position, count - 1))
    for shape in new.shapes:
        text = shape_text(shape)
        if is_footer_like(shape, slide_h) and re.fullmatch(r"\d{1,3}", text):
            set_shape_text(shape, str(position + 1))
    if not op.get("blank"):
        header_rgb = font_rgb(first_run_font([title])) if title is not None else None
        write_content(new, slide_w, slide_h, body_top, op, ref_body_font, warnings, position + 1, header_rgb)
    move_slide(prs, count - 1, position)
    return position


def is_empty_slide(slide, slide_h: int) -> bool:
    return not any(shape_text(s) for s in slide.shapes if not is_footer_like(s, slide_h)) and not any(
        s.shape_type in (MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART) or getattr(s, "has_table", False) for s in slide.shapes
    )


def op_fill_slide(prs, op, warnings):
    slide_w, slide_h = prs.slide_width, prs.slide_height
    index = op.get("slideIndex")
    slides = list(prs.slides)
    if not isinstance(index, int) or not 0 <= index < len(slides):
        raise ValueError("指定したスライドがありません。")
    slide = slides[index]
    if not is_empty_slide(slide, slide_h):
        raise ValueError(f"P{index + 1}は空白スライドではないため、本文を書き込みませんでした。")
    ref = choose_reference(prs, slide_h)
    ref_body_font = first_run_font(body_text_shapes(ref, slide_h)) if ref is not None and not same(ref, slide) else None
    title_text = str(op.get("title") or "").strip()[:60]
    title_ph = next((s for s in slide.placeholders if placeholder_type(s) in TITLE_TYPES), None)
    body_top = int(slide_h * 0.2)
    if title_ph is not None:
        if title_text:
            title_ph.text_frame.text = title_text
        body_top = title_ph.top + title_ph.height
    elif title_text:
        ref_title = find_title(ref, slide_h) if ref is not None and not same(ref, slide) else None
        geo = (ref_title.left, ref_title.top, ref_title.width, ref_title.height) if ref_title is not None else (
            int(slide_w * 0.06), int(slide_h * 0.05), int(slide_w * 0.88), int(slide_h * 0.12))
        box = slide.shapes.add_textbox(*geo)
        box.name = "AzureChat Added Title"
        run = box.text_frame.paragraphs[0].add_run()
        run.text = title_text
        ref_title_font = first_run_font([ref_title]) if ref_title is not None else None
        apply_font(run, ref_title_font, ref_title_font.size.pt if ref_title_font is not None and ref_title_font.size else DEFAULT_TITLE_PT, bold=True)
        body_top = geo[1] + geo[3]
    body_ph = next((s for s in slide.placeholders if placeholder_type(s) not in TITLE_TYPES | FOOTER_TYPES), None)
    if body_ph is not None and op.get("bullets") and op.get("kind") != "table":
        body_ph._element.getparent().remove(body_ph._element)
        body_top = max(body_top, body_ph.top - int(slide_h * 0.03))
    write_content(slide, slide_w, slide_h, body_top, op, ref_body_font, warnings, index + 1)
    return index


def resize_runs(shape, mode, value, default_pt, sizes):
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            base = r.font.size.pt if r.font.size else default_pt
            size = value if mode == "absolute" else base + value
            size = max(MIN_PT, min(MAX_PT, size))
            r.font.size = Pt(size)
            sizes.append(size)


def overflows(shape) -> bool:
    if not shape.has_text_frame or not shape.width or not shape.height:
        return False
    texts, size = [], 0.0
    for p in shape.text_frame.paragraphs:
        texts.append(p.text or " ")
        for r in p.runs:
            if r.font.size:
                size = max(size, r.font.size.pt)
    size = size or DEFAULT_BODY_PT
    return lines_needed(texts, shape.width, size, 0.4) * size * 1.25 > Emu(shape.height).pt * 1.05


def op_font_size(prs, op, warnings):
    slide_h = prs.slide_height
    scope = op.get("scope") or "all"
    mode = "absolute" if op.get("mode") == "absolute" else "relative"
    value = float(op.get("value") or 0)
    if mode == "absolute" and not MIN_PT <= value <= MAX_PT:
        warnings.append(f"{value:g}ptは範囲外のため、{MIN_PT}〜{MAX_PT}ptに収めました。")
    targets = op.get("slides")
    changed, overflow = [], []
    for index, slide in enumerate(prs.slides):
        if isinstance(targets, list) and targets and index not in targets:
            continue
        title = find_title(slide, slide_h)
        sizes = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or is_footer_like(shape, slide_h):
                continue
            if getattr(shape, "has_table", False) and shape.has_table:
                if scope in ("table", "all"):
                    for row in shape.table.rows:
                        for cell in row.cells:
                            resize_runs(cell, mode, value, 14, sizes)
                continue
            if not shape_text(shape):
                continue
            is_title = same(shape, title)
            if (scope == "title" and is_title) or (scope == "body" and not is_title) or scope == "all":
                resize_runs(shape, mode, value, DEFAULT_TITLE_PT if is_title else DEFAULT_BODY_PT, sizes)
                if overflows(shape):
                    overflow.append(index + 1)
        if sizes:
            changed.append(index)
    if overflow:
        warnings.append("文字が枠からはみ出す可能性があります: " + "、".join(f"P{n}" for n in sorted(set(overflow))))
    return changed


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--plan")
    parser.add_argument("--review-slides")
    args = parser.parse_args()
    if args.review_slides is not None:
        selected = {int(value) for value in args.review_slides.split(",") if value.isdigit()}
        prs = Presentation(args.input)
        if not selected or min(selected) < 0 or max(selected) >= len(prs.slides):
            raise ValueError("Invalid review slide indices")
        for index in range(len(prs.slides)-1, -1, -1):
            if index not in selected:
                element = prs.slides._sldIdLst[index]
                prs.part.drop_rel(element.rId)
                prs.slides._sldIdLst.remove(element)
        prs.save(args.output)
        return
    if not args.plan:
        parser.error("--plan is required for editing")
    with open(args.plan, encoding="utf-8") as f:
        plan = json.load(f)
    prs = Presentation(args.input)
    before = len(prs.slides)
    try:
        warnings, changed, added = apply_ops(prs, plan)
    except ValueError as error:
        # Expected refusals (e.g. target slide is not empty): nothing is saved.
        print(json.dumps({"ok": False, "error": str(error)}))
        sys.exit(3)
    prs.save(args.output)
    # Keys follow edit_pptx.py so the route's runPythonEdit can read them.
    print(json.dumps({
        "ok": True,
        "slidesBefore": before,
        "totalSlides": len(prs.slides),
        "addedSlides": added,
        "changedSlides": len(set(changed)),
        "changedSlideIndices": sorted(set(changed)),
        "layoutWarnings": warnings,
    }))  # ASCII-escaped: Windows consoles would otherwise re-encode Japanese.


def apply_ops(prs, plan):
    warnings, changed, added = [], [], 0
    for op in plan.get("ops") or []:
        kind = op.get("type")
        if kind == "add_slide":
            if added >= 3:
                warnings.append("1回の依頼で追加できるのは3枚までのため、4枚目以降は追加していません。")
                continue
            index = op_add_slide(prs, op, warnings)
            changed = [i + 1 if i >= index else i for i in changed] + [index]
            added += 1
        elif kind == "fill_slide":
            changed.append(op_fill_slide(prs, op, warnings))
        elif kind == "font_size":
            changed.extend(op_font_size(prs, op, warnings))
    return warnings, changed, added


if __name__ == "__main__":
    main()
