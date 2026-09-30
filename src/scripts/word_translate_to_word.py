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
STORY_PART = re.compile(
    r"word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$"
)
MAX_BATCH_CHARS = 8000


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
            missing = [item for item in batch if not translated.get(item.id, "").strip()]
            if missing:
                translated.update(_translate_batch(missing, language))
            for item in batch:
                text = translated.get(item.id, "").strip()
                if not text or JAPANESE_RE.search(text):
                    raise RuntimeError(f"段落 {item.id} の英訳が完了しませんでした。")
                name, nodes = parts[item.id]
                nodes[0].text = text
                nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                for node in nodes[1:]:
                    node.text = ""
                changed_parts.add(name)

        with ZipFile(output_path, "w") as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename in changed_parts:
                    data = etree.tostring(
                        roots[info.filename], encoding="UTF-8", xml_declaration=True
                    )
                output.writestr(info, data)

    return {"translatedParagraphs": len(items), "targetLanguage": target_language}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--target-language", choices=["en"], default="en")
    args = parser.parse_args()
    print(json.dumps(translate_docx(args.input, args.output, args.target_language)))


if __name__ == "__main__":
    main()
