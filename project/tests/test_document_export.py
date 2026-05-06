import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from converter import SmetaConverter
from document_export import DocumentExporter


class DocumentExportTests(unittest.TestCase):
    def test_raw_table_to_dataframe_preserves_first_row(self):
        exporter = DocumentExporter(ocr_extractor=None)
        table = [
            ["Header title", "", ""],
            ["1", "Code", "Name"],
        ]

        df = exporter.raw_table_to_dataframe(table)

        self.assertEqual(list(df.columns), ["Column_1", "Column_2", "Column_3"])
        self.assertEqual(df.iloc[0, 0], "Header title")
        self.assertEqual(df.iloc[1, 2], "Name")

    def test_raw_tables_to_dataframe_keeps_multiple_tables(self):
        exporter = DocumentExporter(ocr_extractor=None)
        tables = [
            [["A1", "B1"], ["A2", "B2"]],
            [["C1", "D1"]],
        ]

        df = exporter.raw_tables_to_dataframe(tables)

        self.assertEqual(list(df.columns), ["Column_1", "Column_2"])
        self.assertEqual(df.iloc[0].tolist(), ["A1", "B1"])
        self.assertEqual(df.iloc[1].tolist(), ["A2", "B2"])
        self.assertEqual(df.iloc[2].tolist(), ["", ""])
        self.assertEqual(df.iloc[3].tolist(), ["C1", "D1"])

    def test_append_page_sheets_creates_page_tabs(self):
        exporter = DocumentExporter(ocr_extractor=None)

        with tempfile.TemporaryDirectory() as tmp:
            output_path = Path(tmp) / "book.xlsx"
            SmetaConverter().save_to_excel(pd.DataFrame({"A": ["row"]}), str(output_path))

            pages = [
                {
                    "page_number": 1,
                    "sheet_name": "Page_01",
                    "method": "ocr_table",
                    "dataframe": pd.DataFrame({"Column_1": ["line1"], "Column_2": ["line2"]}),
                }
            ]

            exporter.append_page_sheets(str(output_path), pages)

            wb = load_workbook(output_path, data_only=True)
            self.assertIn("Page_01", wb.sheetnames)
            ws = wb["Page_01"]
            self.assertEqual(ws["A1"].value, "Method")
            self.assertEqual(ws["B1"].value, "ocr_table")
            self.assertEqual(ws["A3"].value, "line1")
            self.assertEqual(ws["B3"].value, "line2")
            self.assertEqual(ws.freeze_panes, "A3")
            wb.close()

    def test_normalize_dataframe_for_sheet_makes_headers_unique(self):
        exporter = DocumentExporter(ocr_extractor=None)
        df = pd.DataFrame([["a", "b", "c"]], columns=["", "Column_1", "Column_1"])

        normalized = exporter._normalize_dataframe_for_sheet(df)

        self.assertEqual(
            list(normalized.columns),
            ["Column_1", "Column_1_2", "Column_1_3"],
        )

    def test_bad_sparse_ocr_table_is_rejected(self):
        exporter = DocumentExporter(ocr_extractor=None)
        df = pd.DataFrame(
            [
                ["|", "", "", "", "", "", "", "", "", "", "", ""],
                ["", "05", "Г", "", "", "", "", "", "11", "", "", ""],
                ["Й", "", "100", "", "", "", "", "", "", "", "", ""],
                ["04", "", "", "", "", "", "", "", "", "", "", ""],
            ],
            columns=[f"Column_{i}" for i in range(1, 13)],
        )

        self.assertFalse(exporter._is_usable_ocr_table(df))


if __name__ == "__main__":
    unittest.main()
