import unittest
from pathlib import Path

from main import resolve_single_output_path


class MainPathTests(unittest.TestCase):
    def test_resolve_single_output_path_uses_output_directory(self):
        path = resolve_single_output_path("input/example.pdf", "result")

        self.assertEqual(path, str(Path("result") / "example.xlsx"))

    def test_resolve_single_output_path_accepts_xlsx_file(self):
        path = resolve_single_output_path("input/example.pdf", "result/custom.xlsx")

        self.assertEqual(path, str(Path("result/custom.xlsx")))


if __name__ == "__main__":
    unittest.main()
