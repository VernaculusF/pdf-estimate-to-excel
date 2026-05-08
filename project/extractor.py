"""
Модуль интеллектуального извлечения структурированных данных из смет.
"""

import os
import re
from typing import List, Optional, Dict, Any, Tuple

import pdfplumber
import pandas as pd

from config import (
    MIN_TABLE_ROWS, MIN_TABLE_COLS, 
    SMETA_COLUMNS, KEYWORDS_RESOURCES, KEYWORDS_TOTALS
)


class SmetaExtractor:
    """Извлекает данные из PDF и приводит их к строгой структуре сметы."""

    def __init__(self):
        # Регулярки для определения типов данных
        self.re_pos_number = re.compile(r'^\s*(\d+)\s*$', re.IGNORECASE)
        self.re_resource_number = re.compile(r'^\s*(\d+)\.\s*', re.IGNORECASE)
        self.re_formula = re.compile(r'[\*\/()+]')
        self.re_number = re.compile(r'^-?\d+([.,]\d+)?$')
        self.re_unit = re.compile(
            r'(?<!\w)(чел\.[-]?ч|маш\.[-]?час|м\.п\.?|м[²2³3]?|кг|шт\.?|компл\.?|руб\.?|%|см|мм|км|га|час|т)(?!\w)',
            re.IGNORECASE
        )

    def extract_tables_from_pdf(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Извлекает данные из PDF и возвращает список структурированных DataFrame.
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF-файл не найден: {pdf_path}")

        all_structured_data = []
        
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Извлекаем все таблицы со страницы
                tables = page.extract_tables()
                for table in tables:
                    if table and len(table) >= MIN_TABLE_ROWS:
                        df = self._process_table_to_structure(table)
                        if not df.empty:
                            all_structured_data.append(df)
        
        return all_structured_data

    def extract_raw_tables_from_pdf(
        self, pdf_path: str, ocr_extractor=None
    ) -> List[List[List[Any]]]:
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF-file not found: {pdf_path}")

        raw_tables: List[List[List[Any]]] = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_index, page in enumerate(pdf.pages):
                words = page.extract_words()
                is_scanned = len(words) == 0 and len(page.images) > 0

                if is_scanned and ocr_extractor is not None:
                    ocr_table = self._extract_raw_table_from_scanned_page(
                        pdf_path, page_index, ocr_extractor
                    )
                    if ocr_table and len(ocr_table) >= MIN_TABLE_ROWS:
                        raw_tables.append(ocr_table)
                else:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) >= MIN_TABLE_ROWS:
                            raw_tables.append(table)
        return raw_tables

    def _extract_raw_table_from_scanned_page(
        self,
        pdf_path: str,
        page_index: int,
        ocr_extractor,
    ) -> Optional[List[List[str]]]:
        import pypdfium2 as pdfium

        try:
            pdf_doc = pdfium.PdfDocument(pdf_path)
            bitmap = pdf_doc[page_index].render(scale=ocr_extractor.dpi / 72)
            pil_image = bitmap.to_pil()
            pdf_doc.close()
        except Exception:
            return None

        processed = ocr_extractor._preprocess_image(pil_image)

        table_top = self._detect_table_top(processed)

        if table_top and table_top > 20:
            cropped = processed.crop((0, table_top, processed.width, processed.height))
        else:
            cropped = processed

        try:
            import pytesseract
            ocr_data = pytesseract.image_to_data(
                cropped,
                lang=ocr_extractor.lang,
                output_type=pytesseract.Output.DATAFRAME,
                config=r'--oem 1 --psm 6',
            )
        except Exception:
            return None

        df_words = ocr_data[
            (ocr_data['text'].notna())
            & (ocr_data['text'].astype(str).str.strip() != '')
            & (ocr_data['conf'] > 20)
        ].copy()

        if len(df_words) < 5:
            return None

        df_words['text'] = df_words['text'].astype(str)

        col_lines = ocr_extractor._detect_table_lines(cropped)
        if col_lines and len(col_lines) >= 3:
            table_data = ocr_extractor._extract_with_column_lines(
                df_words, col_lines, cropped.height
            )
        else:
            df_words = ocr_extractor._group_words_into_lines(df_words)
            table_data = ocr_extractor._group_lines_into_table(df_words)

        if not table_data:
            return None

        table_data = self._filter_table_rows(table_data)

        if not table_data:
            return None

        max_cols = max(len(row) for row in table_data)
        normalized = [
            row + [''] * (max_cols - len(row)) for row in table_data
        ]
        return normalized

    def _detect_table_top(self, image) -> Optional[int]:
        import cv2
        import numpy as np

        gray = np.array(image.convert("L") if hasattr(image, "convert") else image)
        _, threshold = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(80, gray.shape[1] // 5), 1))
        horizontal = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, h_kernel)

        v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, gray.shape[0] // 15)))
        vertical = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, v_kernel)

        combined = cv2.bitwise_or(horizontal, vertical)
        row_sums = combined.sum(axis=1) / 255
        width = gray.shape[1]
        min_line_width = width * 0.25

        for y in range(gray.shape[0]):
            if row_sums[y] >= min_line_width:
                return max(y - 5, 0)

        return None

    def _filter_table_rows(self, table_data: List[List[str]]) -> List[List[str]]:
        filtered = []
        found_header = False
        col_number_re = re.compile(r'^\s*[\|\[\]]*\s*\d+\s*[\|\[\]]*\s*$')

        for row in table_data:
            non_empty = [c.strip() for c in row if c.strip()]
            if not non_empty:
                continue

            joined = ' '.join(non_empty).lower()

            if not found_header:
                if 'наименование' in joined and ('ед' in joined or 'изм' in joined):
                    found_header = True
                    continue
                continue

            col_num_count = sum(
                1 for c in row if c.strip() and col_number_re.match(c.strip())
            )
            if col_num_count >= 5:
                continue

            if re.match(r'^\s*страница\s+\d+\s*$', joined, re.IGNORECASE):
                continue

            all_empty_or_noise = all(
                len(c.strip()) <= 1 or c.strip() in ('|', '[', ']', '‚', "'")
                for c in row
            )
            if all_empty_or_noise:
                continue

            filtered.append(row)

        return filtered

    def build_raw_estimate_dataframe(self, raw_tables: List[List[List[Any]]]) -> pd.DataFrame:
        max_cols = 0
        cleaned_tables: List[List[List[str]]] = []

        for table in raw_tables:
            cleaned_table = [
                [str(cell).replace("\n", " ").strip() if cell is not None else "" for cell in row]
                for row in table
                if any(cell is not None and str(cell).strip() != "" for cell in row)
            ]
            if not cleaned_table:
                continue
            table_max = max(len(row) for row in cleaned_table)
            max_cols = max(max_cols, table_max)
            cleaned_tables.append(cleaned_table)

        if not cleaned_tables:
            return pd.DataFrame()

        normalized_rows: List[List[str]] = []
        for cleaned_table in cleaned_tables:
            table_max = max(len(row) for row in cleaned_table)
            for row in cleaned_table:
                row_len = len(row)
                if row_len < max_cols:
                    gap = max_cols - row_len
                    aligned_row = [row[0]] + [""] * gap + row[1:]
                else:
                    aligned_row = row + [""] * (max_cols - row_len)
                normalized_rows.append(aligned_row)

        columns = [f"Column_{index + 1}" for index in range(max_cols)]
        return pd.DataFrame(normalized_rows, columns=columns)

    def _process_table_to_structure(self, raw_table: List[List[Any]]) -> pd.DataFrame:
        """
        Превращает «сырую» таблицу из PDF в строго структурированный DataFrame.
        """
        structured_rows = []
        
        # Очистка от пустых строк в самом начале и конце
        raw_table = [row for row in raw_table if any(cell is not None and str(cell).strip() != '' for cell in row)]
        
        # Состояние для отслеживания иерархии
        current_pos = ""
        
        for row in raw_table:
            # 1. Склеиваем все ячейки в одну строку для анализа, если нужно
            # Но сначала очищаем каждую ячейку от \n
            clean_row = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
            
            # Пропускаем абсолютно пустые строки
            if not any(clean_row):
                continue
                
            # 2. Определяем тип строки
            row_type = self._determine_row_type(clean_row)
            
            # 3. Парсим строку в соответствии с типом
            parsed_row = self._parse_row_by_type(clean_row, row_type, current_pos)
            
            # Если строка определена как позиция, обновляем текущий контекст
            if row_type == "POSITION":
                current_pos = parsed_row.get("№ п/п", "")
                
            # Добавляем в итоговый список, если строка несет смысл
            if parsed_row:
                structured_rows.append(parsed_row)
        
        return pd.DataFrame(structured_rows, columns=SMETA_COLUMNS)

    def _determine_row_type(self, row: List[str]) -> str:
        """
        Классифицирует строку сметы.
        """
        non_empty_cells = [cell for cell in row if cell]

        # Проверка на заголовок (игнорируем)
        if any("наименование" in s.lower() or "ед.изм" in s.lower() for s in row):
            return "HEADER"
            
        # Проверка на итоговые строки
        if any(any(kw in s.lower() for kw in KEYWORDS_TOTALS) for s in row):
            return "TOTAL"
            
        # Проверка на позицию (начинается с числа)
        if non_empty_cells and self.re_pos_number.match(non_empty_cells[0]):
            return "POSITION"
            
        # Проверка на ресурс (начинается с 1. или содержит код)
        if any(self.re_resource_number.match(cell) for cell in non_empty_cells):
            return "RESOURCE"
            
        # Проверка на затраты
        if any(any(kw in s.lower() for kw in KEYWORDS_RESOURCES) for s in row):
            return "EXPENSE"
            
        return "UNKNOWN"

    def _parse_row_by_type(self, row: List[str], row_type: str, current_pos: str) -> Optional[Dict[str, Any]]:
        """
        Распределяет данные по строгим колонкам.
        """
        if row_type == "HEADER":
            return None
            
        parsed = {col: "" for col in SMETA_COLUMNS}
        
        # Склеиваем разорванные строки (убираем лишние пробелы внутри ячеек)
        clean_row = [re.sub(r'\s+', ' ', cell).strip() for cell in row]
        
        if row_type == "POSITION":
            parsed["№ п/п"] = clean_row[0]
            parsed["Код"] = clean_row[1] if len(clean_row) > 1 else ""
            parsed["Наименование"] = clean_row[2] if len(clean_row) > 2 else ""
            if self._looks_like_structured_cost_row(clean_row):
                self._fill_structured_cost_cols(clean_row, parsed)
            else:
                self._fill_numeric_cols(clean_row, parsed)
            
        elif row_type == "RESOURCE":
            parsed["№ п/п"] = current_pos
            if self._looks_like_structured_cost_row(clean_row):
                parsed["Код"] = clean_row[1] if len(clean_row) > 1 else ""
                parsed["Наименование"] = clean_row[2] if len(clean_row) > 2 else ""
                self._fill_structured_cost_cols(clean_row, parsed)
            else:
                resource_cell_index = next(
                    (idx for idx, cell in enumerate(clean_row) if self.re_resource_number.match(cell)),
                    0,
                )
                resource_cell = clean_row[resource_cell_index] if clean_row else ""
                code_match = re.match(r'^\s*(\d+\.\s*[\d.:-]+)', resource_cell)
                if code_match:
                    parsed["Код"] = code_match.group(1).strip()
                    resource_tail = resource_cell[code_match.end():].strip()
                    extra_cells = clean_row[resource_cell_index + 1:]
                    full_text = " ".join([resource_tail] + extra_cells).strip()
                else:
                    parsed["Код"] = clean_row[resource_cell_index] if len(clean_row) > resource_cell_index else ""
                    full_text = " ".join(clean_row[resource_cell_index + 1:]).strip()
                self._split_text_row(full_text, parsed)
            
        elif row_type == "EXPENSE":
            full_text = " ".join(clean_row)
            parsed["№ п/п"] = current_pos
            if self._looks_like_structured_cost_row(clean_row):
                parsed["Наименование"] = clean_row[2] if len(clean_row) > 2 else full_text
                self._fill_structured_cost_cols(clean_row, parsed, prefer_price_index=5)
            else:
                self._split_text_row(full_text, parsed)
            
        elif row_type == "TOTAL":
            parsed["Наименование"] = " ".join(clean_row[:len(clean_row)//2 + 1])
            parsed["Сумма"] = clean_row[-1] if clean_row else ""
            
        elif row_type == "UNKNOWN":
            parsed["Наименование"] = " ".join(clean_row)
            parsed["№ п/п"] = current_pos
            
        return parsed

    def _split_text_row(self, full_text: str, parsed: Dict[str, Any]) -> None:
        unit_match = self._find_best_unit_match(full_text)
        if not unit_match:
            parsed["Наименование"] = full_text.strip()
            numbers = re.findall(r'\d+[.,]?\d*', full_text)
            if len(numbers) >= 3:
                parsed["Количество"] = numbers[-3]
                parsed["Цена"] = numbers[-2]
                parsed["Сумма"] = numbers[-1]
            elif len(numbers) == 2:
                parsed["Количество"] = numbers[-2]
                parsed["Сумма"] = numbers[-1]
            return

        unit_pos = unit_match.start()
        parsed["Наименование"] = full_text[:unit_pos].strip()
        parsed["Ед. изм."] = unit_match.group(1)

        rest = full_text[unit_match.end():].strip()
        numbers = re.findall(r'\d+[.,]?\d*', rest)
        if len(numbers) >= 3:
            parsed["Количество"] = numbers[0]
            parsed["Цена"] = numbers[1]
            parsed["Сумма"] = numbers[-1]
        elif len(numbers) == 2:
            parsed["Количество"] = numbers[0]
            parsed["Сумма"] = numbers[1]

    def _looks_like_structured_cost_row(self, row: List[str]) -> bool:
        return len(row) >= 8 and any(cell.strip() for cell in row[2:8])

    def _fill_structured_cost_cols(
        self,
        row: List[str],
        parsed: Dict[str, Any],
        prefer_price_index: int = 6,
    ) -> None:
        parsed["Ед. изм."] = row[3].strip() if len(row) > 3 else ""

        quantity_candidates = []
        if len(row) > 4:
            quantity_candidates.append(row[4])
        if len(row) > 5:
            quantity_candidates.append(row[5])
        parsed["Количество"] = self._extract_primary_number(quantity_candidates)

        price_candidates = []
        if len(row) > prefer_price_index:
            price_candidates.append(row[prefer_price_index])
        if len(row) > 6 and prefer_price_index != 6:
            price_candidates.append(row[6])
        if len(row) > 5:
            price_candidates.append(row[5])
        parsed["Цена"] = self._extract_primary_number(price_candidates)

        sum_candidates = []
        for index in (7, 9, 11, 6, 10, 12, 13):
            if len(row) > index:
                sum_candidates.append(row[index])
        parsed["Сумма"] = self._extract_primary_number(sum_candidates)

        formulas = []
        for cell in row[4:]:
            if cell and self.re_formula.search(cell):
                formulas.append(re.sub(r"\s+", " ", cell).strip())
        if formulas:
            parsed["Формула"] = " | ".join(formulas)

    def _extract_primary_number(self, cells: List[str]) -> str:
        for cell in cells:
            if not cell:
                continue
            match = re.search(r'-?\d+(?:[.,]\d+)?', str(cell))
            if match:
                return match.group(0)
        return ""

    def _find_best_unit_match(self, full_text: str):
        matches = list(self.re_unit.finditer(full_text))
        if not matches:
            return None

        best_match = None
        for match in matches:
            rest = full_text[match.end():].strip()
            numbers = re.findall(r'\d+[.,]?\d*', rest)
            if len(numbers) >= 2:
                best_match = match

        return best_match or matches[0]

    def _fill_numeric_cols(self, row: List[str], parsed: Dict[str, Any]):
        """
        Интеллектуальное заполнение колонок Количество, Цена, Сумма и Формула.
        """
        # Берем последние несколько колонок, там обычно цифры
        # Идем с конца строки
        numeric_values = []
        for cell in reversed(row):
            cell = cell.strip()
            if not cell: continue
            
            # Разделяем число и формулу
            if self.re_formula.search(cell):
                numeric_values.append(("FORMULA", cell))
            elif self.re_number.match(cell.replace(' ', '').replace(',', '.')):
                numeric_values.append(("VALUE", cell))
            else:
                numeric_values.append(("TEXT", cell))
                
        # Распределяем по колонкам (Сумма -> Цена -> Кол-во)
        # Это упрощенная логика, так как структура PDF может меняться
        if len(numeric_values) >= 1:
            val, text = numeric_values[0]
            if val == "FORMULA": parsed["Формула"] = text
            else: parsed["Сумма"] = text
            
        if len(numeric_values) >= 2:
            val, text = numeric_values[1]
            if val == "FORMULA": parsed["Формула"] += f" | {text}"
            else: parsed["Цена"] = text
            
        if len(numeric_values) >= 3:
            val, text = numeric_values[2]
            if val == "FORMULA": parsed["Формула"] += f" | {text}"
            else: parsed["Количество"] = text

    def merge_tables(self, tables: List[pd.DataFrame]) -> pd.DataFrame:
        """Объединяет таблицы, приводя их к единой схеме SMETA_COLUMNS."""
        if not tables: 
            return pd.DataFrame(columns=SMETA_COLUMNS)
        
        aligned_tables = []
        generic_tables = []
        for df in tables:
            if df.empty:
                continue
            # Убираем дубли колонок
            cols = pd.Series(df.columns)
            for dup in cols[cols.duplicated()].unique():
                dup_indices = cols[cols == dup].index.tolist()
                cols.iloc[dup_indices] = [f"{dup}_{i}" if i != 0 else dup for i in range(len(dup_indices))]
            df.columns = cols
            generic_tables.append(df.copy())
            # Приводим к схеме
            aligned = pd.DataFrame(columns=SMETA_COLUMNS)
            for col in SMETA_COLUMNS:
                if col in df.columns:
                    aligned[col] = df[col]
            aligned_tables.append(aligned)
        
        if not aligned_tables:
            return pd.DataFrame(columns=SMETA_COLUMNS)

        has_structured_data = any(not aligned.empty and len(aligned.columns) > 0 for aligned in aligned_tables)
        if not has_structured_data:
            return pd.concat(generic_tables, ignore_index=True, sort=False)
            
        return pd.concat(aligned_tables, ignore_index=True)
