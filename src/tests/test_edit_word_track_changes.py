import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document


SCRIPT_PATH = Path(__file__).parents[1] / "scripts" / "edit_word.py"


class EditWordTrackChangesTests(unittest.TestCase):
    def test_two_replacements_are_recorded_as_revisions(self):
        with tempfile.TemporaryDirectory() as work_dir:
            work = Path(work_dir)
            source_path = work / "source.docx"
            output_path = work / "edited.docx"
            plan_path = work / "plan.json"

            document = Document()
            document.add_paragraph(
                "太平興産の業績への影響度合いと、稼働率とういう考え方を確認する。"
            )
            document.save(source_path)

            plan_path.write_text(
                json.dumps(
                    {
                        "replaceText": [
                            {
                                "find": "太平興産の業績への影響度合い",
                                "replace": "大平興産の業績への影響度合い",
                            },
                            {
                                "find": "稼働率とういう考え方",
                                "replace": "稼働率という考え方",
                            },
                        ],
                        "formatRuns": [],
                        "addParagraphs": [],
                        "trackChanges": True,
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            completed = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT_PATH),
                    "--input",
                    str(source_path),
                    "--output",
                    str(output_path),
                    "--plan",
                    str(plan_path),
                ],
                capture_output=True,
                text=True,
                check=True,
            )
            result = json.loads(completed.stdout)
            # changedParagraphs counts affected paragraphs, not replacements.
            self.assertEqual(1, result["changedParagraphs"])

            with zipfile.ZipFile(output_path) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")

            self.assertGreaterEqual(document_xml.count("<w:del "), 2)
            self.assertGreaterEqual(document_xml.count("<w:ins "), 2)
            self.assertIn("太平興産の業績への影響度合い", document_xml)
            self.assertIn("大平興産の業績への影響度合い", document_xml)
            self.assertIn("稼働率とういう考え方", document_xml)
            self.assertIn("稼働率という考え方", document_xml)


if __name__ == "__main__":
    unittest.main()
