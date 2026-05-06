"""
Конфигурация проекта для структурированного извлечения смет.
"""

import os

# Настройки путей
DEFAULT_INPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input")
DEFAULT_OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# СТРОГАЯ СХЕМА КОЛОНОК (Standard Smeta Schema)
# Все извлеченные данные будут приводиться к этому виду
SMETA_COLUMNS = [
    "№ п/п", 
    "Код", 
    "Наименование", 
    "Ед. изм.", 
    "Количество", 
    "Цена", 
    "Сумма", 
    "Формула", 
    "Примечание"
]

# Пороги для распознавания таблиц
MIN_TABLE_ROWS = 2
MIN_TABLE_COLS = 3

# Маркеры для типов строк
KEYWORDS_RESOURCES = ["затраты труда", "машины", "эксплуатация", "материалы"]
KEYWORDS_TOTALS = ["итого", "всего", "сумма", "стоимость"]

# Настройки Excel
EXCEL_ENGINE = "openpyxl"
DEFAULT_SHEET_NAME = "Смета"

# Фильтрация пустых строк
DROP_EMPTY_ROWS = True
DROP_EMPTY_COLS = True
