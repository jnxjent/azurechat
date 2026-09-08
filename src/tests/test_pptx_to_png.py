import importlib.util
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


SCRIPT_PATH = Path(__file__).parents[1] / "scripts" / "pptx_to_png.py"
SPEC = importlib.util.spec_from_file_location("pptx_to_png", SCRIPT_PATH)
pptx_to_png = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(pptx_to_png)


class PptxToPngTests(unittest.TestCase):
    def test_scaled_conversion_retries_with_original_pptx(self):
        with tempfile.TemporaryDirectory() as work_dir:
            pptx_path = os.path.join(work_dir, "input.pptx")
            output_dir = os.path.join(work_dir, "pngs")
            Path(pptx_path).write_bytes(b"pptx")

            with (
                patch.object(pptx_to_png.platform, "system", return_value="Linux"),
                patch.object(
                    pptx_to_png,
                    "get_pptx_slide_dims_inches",
                    return_value=(pptx_to_png.WIDE_W_IN, pptx_to_png.WIDE_H_IN),
                ),
                patch.object(pptx_to_png, "create_scaled_pptx", return_value=True),
                patch.object(
                    pptx_to_png,
                    "convert_pptx_to_pdf_libreoffice",
                    side_effect=[None, os.path.join(work_dir, "output.pdf")],
                ) as convert,
                patch.object(
                    pptx_to_png,
                    "pdf_to_pngs",
                    return_value=[os.path.join(output_dir, "slide_0.png")],
                ) as render,
            ):
                result = pptx_to_png.convert_pptx_to_pngs(pptx_path, output_dir, 12)

            self.assertEqual(1, len(result))
            self.assertEqual(2, convert.call_count)
            self.assertEqual(pptx_path, convert.call_args_list[1].args[0])
            self.assertEqual(pptx_to_png.WIDE_W_IN, render.call_args.args[5])
            self.assertEqual(pptx_to_png.WIDE_H_IN, render.call_args.args[6])

    def test_standard_conversion_retries_same_pptx_once(self):
        with tempfile.TemporaryDirectory() as work_dir:
            pptx_path = os.path.join(work_dir, "input.pptx")
            output_dir = os.path.join(work_dir, "pngs")
            Path(pptx_path).write_bytes(b"pptx")

            with (
                patch.object(pptx_to_png.platform, "system", return_value="Linux"),
                patch.object(
                    pptx_to_png,
                    "get_pptx_slide_dims_inches",
                    return_value=(pptx_to_png.STD_W_IN, pptx_to_png.STD_H_IN),
                ),
                patch.object(
                    pptx_to_png,
                    "convert_pptx_to_pdf_libreoffice",
                    side_effect=[None, None],
                ) as convert,
            ):
                result = pptx_to_png.convert_pptx_to_pngs(pptx_path, output_dir, 12)

            self.assertEqual([], result)
            self.assertEqual(2, convert.call_count)
            self.assertEqual(pptx_path, convert.call_args_list[0].args[0])
            self.assertEqual(pptx_path, convert.call_args_list[1].args[0])

    def test_libreoffice_accepts_single_unexpected_pdf_name(self):
        with tempfile.TemporaryDirectory() as work_dir:
            pptx_path = os.path.join(work_dir, "input.pptx")
            Path(pptx_path).write_bytes(b"pptx")
            alternate_pdf = os.path.join(work_dir, "unexpected.pdf")
            Path(alternate_pdf).write_bytes(b"pdf")

            completed = SimpleNamespace(returncode=0, stdout="", stderr="")
            with (
                patch.object(pptx_to_png.shutil, "which", return_value="/usr/bin/soffice"),
                patch.object(pptx_to_png.subprocess, "run", return_value=completed),
            ):
                result = pptx_to_png.convert_pptx_to_pdf_libreoffice(
                    pptx_path, work_dir
                )

            self.assertEqual(alternate_pdf, result)


if __name__ == "__main__":
    unittest.main()
