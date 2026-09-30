"""Translate Japanese text in a DOCX while keeping its package and images."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import deque
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

from lxml import etree

from pdf_translate_to_pptx import (
    JAPANESE_RE,
    TARGET_LANGUAGES,
    TRANSLATION_BATCH_SIZE,
    _translate_batch,
)

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PARAGRAPH = f"{{{WORD_NS}}}p"
TEXT = f"{{{WORD_NS}}}t"
RUN = f"{{{WORD_NS}}}r"
RUN_PROPS = f"{{{WORD_NS}}}rPr"
RUN_FONTS = f"{{{WORD_NS}}}rFonts"
FONT_SIZE = f"{{{WORD_NS}}}sz"
COMPLEX_FONT_SIZE = f"{{{WORD_NS}}}szCs"
STORY_PART = re.compile(
    r"word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$"
)
MAX_BATCH_CHARS = 8000
TRANSLATED_SIZE_HALF_POINTS = "22"
JAPANESE_KANA_RE = re.compile(r"[\u3040-\u30ff\u31f0-\u31ff]")
WORD_TRANSLATION_FONTS = {
    "en": "Times New Roman",
    "pt": "Arial",
    "vi": "Arial",
    "id": "Arial",
    "zh-CN": "Microsoft YaHei",
    "ko": "Malgun Gothic",
    "es": "Arial",
    "fil": "Arial",
}


def apply_translation_font(properties: etree._Element, font_name: str) -> None:
    fonts = properties.find(RUN_FONTS)
    if fonts is None:
        fonts = etree.Element(RUN_FONTS)
        properties.insert(0, fonts)
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(f"{{{WORD_NS}}}{attribute}", font_name)
    for tag in (FONT_SIZE, COMPLEX_FONT_SIZE):
        size = properties.find(tag)
        if size is None:
            size = etree.SubElement(properties, tag)
        size.set(f"{{{WORD_NS}}}val", TRANSLATED_SIZE_HALF_POINTS)


def format_translated_run(text_node: etree._Element, font_name: str) -> None:
    run = next((parent for parent in text_node.iterancestors() if parent.tag == RUN), None)
    if run is None:
        return
    properties = run.find(RUN_PROPS)
    if properties is None:
        properties = etree.Element(RUN_PROPS)
        run.insert(0, properties)
    apply_translation_font(properties, font_name)


def update_default_font(styles: etree._Element, font_name: str) -> None:
    defaults = styles.find(f"{{{WORD_NS}}}docDefaults")
    if defaults is None:
        defaults = etree.Element(f"{{{WORD_NS}}}docDefaults")
        styles.insert(0, defaults)
    run_default = defaults.find(f"{{{WORD_NS}}}rPrDefault")
    if run_default is None:
        run_default = etree.SubElement(defaults, f"{{{WORD_NS}}}rPrDefault")
    properties = run_default.find(RUN_PROPS)
    if properties is None:
        properties = etree.SubElement(run_default, RUN_PROPS)
    apply_translation_font(properties, font_name)
    for style in styles.iter(f"{{{WORD_NS}}}style"):
        if style.get(f"{{{WORD_NS}}}styleId") != "Normal":
            continue
        properties = style.find(RUN_PROPS)
        if properties is None:
            properties = etree.SubElement(style, RUN_PROPS)
        apply_translation_font(properties, font_name)
        break


def paragraph_text_nodes(paragraph: etree._Element) -> list[etree._Element]:
    # A drawing can contain its own paragraphs. Keep their text separate.
    return [
        node
        for node in paragraph.iter(TEXT)
        if next((parent for parent in node.iterancestors() if parent.tag == PARAGRAPH), None)
        is paragraph
    ]


def translate_docx(input_path: Path, output_path: Path, target_language: str) -> dict:
    language = TARGET_LANGUAGES[target_language]
    font_name = WORD_TRANSLATION_FONTS[target_language]
    roots: dict[str, etree._Element] = {}
    parts: dict[str, tuple[str, list[etree._Element]]] = {}
    items: list[SimpleNamespace] = []

    with ZipFile(input_path) as source:
        for name in source.namelist():
            if not STORY_PART.fullmatch(name):
                continue
            root = etree.fromstring(
                source.read(name), etree.XMLParser(resolve_entities=False, no_network=True)
            )
            roots[name] = root
            for paragraph in root.iter(PARAGRAPH):
                nodes = paragraph_text_nodes(paragraph)
                value = "".join(node.text or "" for node in nodes)
                if not nodes or not JAPANESE_RE.search(value):
                    continue
                item_id = str(len(items))
                parts[item_id] = (name, nodes)
                items.append(SimpleNamespace(id=item_id, source=value))

        if not items:
            raise ValueError("Word文書に翻訳対象の日本語が見つかりませんでした。")

        changed_parts: set[str] = set()
        pending = deque(items)
        while pending:
            batch = []
            batch_chars = 0
            while pending and len(batch) < TRANSLATION_BATCH_SIZE:
                next_item = pending[0]
                if batch and batch_chars + len(next_item.source) > MAX_BATCH_CHARS:
                    break
                batch.append(pending.popleft())
                batch_chars += len(next_item.source)
            translated = _translate_batch(batch, language)
            for retry_number in range(1, 3):
                retry_items = []
                for item in batch:
                    text = translated.get(item.id, "").strip()
                    has_untranslated_japanese = (
                        JAPANESE_KANA_RE.search(text)
                        if target_language == "zh-CN"
                        else JAPANESE_RE.search(text)
                    )
                    if not text or has_untranslated_japanese:
                        retry_items.append(item)
                if not retry_items:
                    break
                print(
                    f"[word-translate] retry {retry_number} for paragraphs: "
                    + ", ".join(item.id for item in retry_items[:10]),
                    file=sys.stderr,
                )
                translated.update(_translate_batch(retry_items, language))
            for item in batch:
                text = translated.get(item.id, "").strip()
                has_untranslated_japanese = (
                    JAPANESE_KANA_RE.search(text)
                    if target_language == "zh-CN"
                    else JAPANESE_RE.search(text)
                )
                if not text or has_untranslated_japanese:
                    raise RuntimeError(
                        f"段落 {item.id} の{language.japanese_name}翻訳が完了しませんでした。"
                    )
                name, nodes = parts[item.id]
                nodes[0].text = text
                nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                format_translated_run(nodes[0], font_name)
                for node in nodes[1:]:
                    node.text = ""
                changed_parts.add(name)

        if "word/styles.xml" in source.namelist():
            styles = etree.fromstring(
                source.read("word/styles.xml"),
                etree.XMLParser(resolve_entities=False, no_network=True),
            )
            update_default_font(styles, font_name)
            roots["word/styles.xml"] = styles
            changed_parts.add("word/styles.xml")

        with ZipFile(output_path, "w") as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename in changed_parts:
                    data = etree.tostring(
                        roots[info.filename], encoding="UTF-8", xml_declaration=True
                    )
                output.writestr(info, data)

    return {
        "translatedParagraphs": len(items),
        "targetLanguage": target_language,
        "targetLanguageName": language.japanese_name,
        "fontName": font_name,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--target-language", choices=list(TARGET_LANGUAGES), default="en"
    )
    args = parser.parse_args()
    print(json.dumps(translate_docx(args.input, args.output, args.target_language)))


if __name__ == "__main__":
    main()
