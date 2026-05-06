import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

import pandas as pd

from extractor import SmetaExtractor
from main import SmetaProcessor


class ExtractorLogicTests(unittest.TestCase):
    def test_resource_row_splits_name_unit_and_numbers(self):
        extractor = SmetaExtractor()

        parsed = extractor._parse_row_by_type(
            ["1. 91.01.02-004", "Автогрейдеры среднего типа", "маш.час", "1,94", "9,94", "1919,78", "19082,61"],
            "RESOURCE",
            "3",
        )

        self.assertEqual(parsed["№ п/п"], "3")
        self.assertEqual(parsed["Код"], "1. 91.01.02-004")
        self.assertEqual(parsed["Наименование"], "Автогрейдеры среднего типа")
        self.assertEqual(parsed["Ед. изм."].lower(), "маш.час")
        self.assertEqual(parsed["Количество"], "1,94")
        self.assertEqual(parsed["Цена"], "9,94")
        self.assertEqual(parsed["Сумма"], "19082,61")

    def test_expense_row_splits_name_unit_and_numbers(self):
        extractor = SmetaExtractor()

        parsed = extractor._parse_row_by_type(
            ["Затраты труда рабочих (ср 2,3) чел.-ч 21,6 164,73 3558,17"],
            "EXPENSE",
            "3",
        )

        self.assertEqual(parsed["№ п/п"], "3")
        self.assertEqual(parsed["Наименование"], "Затраты труда рабочих (ср 2,3)")
        self.assertEqual(parsed["Ед. изм."].lower(), "чел.-ч")
        self.assertEqual(parsed["Количество"], "21,6")
        self.assertEqual(parsed["Цена"], "164,73")
        self.assertEqual(parsed["Сумма"], "3558,17")

    def test_resource_row_embedded_in_single_cell_is_detected_and_split(self):
        extractor = SmetaExtractor()

        row_type = extractor._determine_row_type(
            ["", "", "1. 91.01.02-004 Автогрейдеры среднего типа маш.час 1,94 9,94 1919,78 19082,61"]
        )
        parsed = extractor._parse_row_by_type(
            ["", "", "1. 91.01.02-004 Автогрейдеры среднего типа маш.час 1,94 9,94 1919,78 19082,61"],
            row_type,
            "3",
        )

        self.assertEqual(row_type, "RESOURCE")
        self.assertEqual(parsed["Код"], "1. 91.01.02-004")
        self.assertEqual(parsed["Наименование"], "Автогрейдеры среднего типа")
        self.assertEqual(parsed["Ед. изм."].lower(), "маш.час")
        self.assertEqual(parsed["Количество"], "1,94")
        self.assertEqual(parsed["Цена"], "9,94")
        self.assertEqual(parsed["Сумма"], "19082,61")

    def test_resource_row_prefers_real_unit_over_text_inside_name(self):
        extractor = SmetaExtractor()

        parsed = extractor._parse_row_by_type(
            ["1. 91.01.02-004 Автогрейдеры среднего типа, мощность 99 кВт (135 л.с.) маш.час 1,94 9,94 1919,78 19082,61"],
            "RESOURCE",
            "3",
        )

        self.assertEqual(parsed["Ед. изм."].lower(), "маш.час")
        self.assertEqual(parsed["Наименование"], "Автогрейдеры среднего типа, мощность 99 кВт (135 л.с.)")

    def test_position_row_uses_structured_cells_from_pdf_table(self):
        extractor = SmetaExtractor()

        parsed = extractor._parse_row_by_type(
            [
                "2",
                "ФССЦпг03-21-\n01-025",
                "Перевозка грузов автомобилями-\nсамосвалами",
                "1 т груза",
                "",
                "563,6\n(2,562*100)*2,2",
                "281,49",
                "158647,76",
                "",
                "158647,76",
                "",
                "",
                "",
                "",
            ],
            "POSITION",
            "",
        )

        self.assertEqual(parsed["№ п/п"], "2")
        self.assertEqual(parsed["Ед. изм."], "1 т груза")
        self.assertEqual(parsed["Количество"], "563,6")
        self.assertEqual(parsed["Цена"], "281,49")
        self.assertEqual(parsed["Сумма"], "158647,76")
        self.assertIn("(2,562*100)*2,2", parsed["Формула"])

    def test_resource_row_uses_structured_cells_from_pdf_table(self):
        extractor = SmetaExtractor()

        parsed = extractor._parse_row_by_type(
            [
                "",
                "3. 91.17.04-233",
                "Установки для сварки ручной дуговой\n(постоянного тока)",
                "маш.час",
                "0,66",
                "0,66",
                "44,14",
                "29,13",
                "",
                "29,13",
                "",
                "",
                "",
                "",
            ],
            "RESOURCE",
            "3",
        )

        self.assertEqual(parsed["№ п/п"], "3")
        self.assertEqual(parsed["Код"], "3. 91.17.04-233")
        self.assertEqual(parsed["Ед. изм."], "маш.час")
        self.assertEqual(parsed["Количество"], "0,66")
        self.assertEqual(parsed["Цена"], "44,14")
        self.assertEqual(parsed["Сумма"], "29,13")

    def test_merge_tables_keeps_ocr_style_columns_when_schema_missing(self):
        extractor = SmetaExtractor()
        table_a = pd.DataFrame([["a", "b"]], columns=["Column_1", "Column_1"])
        table_b = pd.DataFrame([["c", "d"]], columns=["Column_1", "Column_1"])

        merged = extractor.merge_tables([table_a, table_b])

        self.assertEqual(list(merged.columns), ["Column_1", "Column_1_1"])
        self.assertEqual(len(merged), 2)


class ProcessorMergeTests(unittest.TestCase):
    def test_process_single_file_merges_tables_when_enabled(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = Path(tmp) / "sample.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n%stub")
            output_path = Path(tmp) / "sample.xlsx"

            processor = SmetaProcessor()
            processor.ocr_extractor = Mock()
            processor.extractor = Mock()
            processor.converter = Mock()

            table_a = pd.DataFrame({"№ п/п": ["1"], "Наименование": ["A"]})
            table_b = pd.DataFrame({"№ п/п": ["2"], "Наименование": ["B"]})
            merged = pd.DataFrame({"№ п/п": ["1", "2"], "Наименование": ["A", "B"]})

            processor.ocr_extractor.is_scanned_pdf.return_value = False
            processor.extractor.extract_tables_from_pdf.return_value = [table_a, table_b]
            processor.extractor.merge_tables.return_value = merged
            processor.converter.save_to_excel.return_value = str(output_path)
            processor.document_exporter = Mock()

            processor.process_single_file(str(pdf_path), str(output_path), merge_tables=True)

            processor.extractor.merge_tables.assert_called_once()
            processor.document_exporter.build_pages.assert_called_once()
            processor.document_exporter.append_page_sheets.assert_called_once()
            saved_df = processor.converter.save_to_excel.call_args.args[0]
            self.assertEqual(len(saved_df), 2)


if __name__ == "__main__":
    unittest.main()
