"""
Project configuration for PDF-to-Excel estimate conversion.
"""

from pathlib import Path

INPUT_DIR_NAME = "input"
OUTPUT_DIR_NAME = "output"

PROJECT_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = PROJECT_DIR.parent
DEFAULT_INPUT_DIR = str(WORKSPACE_DIR / INPUT_DIR_NAME)
DEFAULT_OUTPUT_DIR = str(WORKSPACE_DIR / OUTPUT_DIR_NAME)

DEFAULT_SHEET_NAME = "Estimate"
SOURCE_TEXT_SHEET_NAME = "Source Text"
SOURCE_TEXT_SOURCE_LABEL = "Source"
SOURCE_TEXT_CONTENT_LABEL = "PDF Text"
PAGE_SHEET_PREFIX = "Page"

SMETA_COLUMNS = [
    "№ п/п",
    "Код",
    "Наименование",
    "Ед. изм.",
    "Количество",
    "Цена",
    "Сумма",
    "Формула",
    "Примечание",
]

MIN_TABLE_ROWS = 2
MIN_TABLE_COLS = 3

KEYWORDS_RESOURCES = ["затраты труда", "машины", "эксплуатация", "материалы"]
KEYWORDS_TOTALS = ["итого", "всего", "сумма", "стоимость"]

EXCEL_ENGINE = "openpyxl"
DROP_EMPTY_ROWS = True
DROP_EMPTY_COLS = True
