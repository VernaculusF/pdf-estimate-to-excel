"""
Модуль интеллектуального извлечения структурированных данных из смет.
"""

import os
import re
import tempfile
from typing import List, Optional, Dict, Any, Tuple

import pdfplumber
import pandas as pd
import pypdfium2 as pdfium
import pytesseract
from PIL import Image

from config import (
    MIN_TABLE_ROWS, MIN_TABLE_COLS, 
    SMETA_COLUMNS, KEYWORDS_RESOURCES, KEYWORDS_TOTALS
)
from ocr_extractor import OCRExtractor


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
                # Prefer find_tables which handles spanning cells better
                try:
                    found_tables = page.find_tables()
                    tables = [t.extract() for t in found_tables]
                except Exception:
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

        all_tables: List[tuple] = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_index, page in enumerate(pdf.pages):
                words = page.extract_words()
                is_scanned = len(words) == 0 and len(page.images) > 0

                if is_scanned and ocr_extractor is not None:
                    ocr_table = self._extract_raw_table_from_scanned_page(
                        pdf_path, page_index, ocr_extractor
                    )
                    if ocr_table and len(ocr_table) >= MIN_TABLE_ROWS:
                        all_tables.append((page_index, 'ocr', ocr_table))
                else:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) >= MIN_TABLE_ROWS:
                            all_tables.append((page_index, 'plumber', table))

        plumber_tables = [t for _, src, t in all_tables if src == 'plumber']
        if plumber_tables:
            canonical_cols = self._canonical_col_count(plumber_tables)
            result: List[List[List[Any]]] = []
            for _, src, table in all_tables:
                if src == 'plumber':
                    result.append(table)
                else:
                    ot_cols = len(table[0]) if table else 0
                    if abs(ot_cols - canonical_cols) <= 2:
                        result.append(table)
            return result

        return [t for _, _, t in all_tables]

    @staticmethod
    def _canonical_col_count(tables: List[List[List[Any]]]) -> int:
        from collections import Counter
        counts = Counter(
            len(t[0]) for t in tables if t
        )
        return counts.most_common(1)[0][0] if counts else 0

    def _extract_raw_table_from_scanned_page(
        self,
        pdf_path: str,
        page_index: int,
        ocr_extractor,
    ) -> Optional[List[List[str]]]:
        try:
            pil_image = ocr_extractor.render_pdf_page(
                pdf_path, page_index, ocr_extractor.dpi
            )
        except Exception:
            return None

        processed = ocr_extractor._preprocess_image(pil_image)

        table_top = self._detect_table_top(processed)

        # Only crop if there is a lot of whitespace/noise above the table
        if table_top and table_top > 80:
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

        # Fallback keyword-based table top detection when line detection failed
        if table_top is None:
            keyword_top = ocr_extractor._find_table_top_by_keywords(df_words)
            if keyword_top is not None and keyword_top > 80:
                table_top = keyword_top
                cropped = processed.crop((0, table_top, processed.width, processed.height))
                try:
                    ocr_data = pytesseract.image_to_data(
                        cropped,
                        lang=ocr_extractor.lang,
                        output_type=pytesseract.Output.DATAFRAME,
                        config=r'--oem 1 --psm 6',
                    )
                    df_words = ocr_data[
                        (ocr_data['text'].notna())
                        & (ocr_data['text'].astype(str).str.strip() != '')
                        & (ocr_data['conf'] > 20)
                    ].copy()
                    df_words['text'] = df_words['text'].astype(str)
                except Exception:
                    pass

        line_detect_src = pil_image
        if table_top and table_top > 80:
            line_detect_src = pil_image.crop(
                (0, table_top, pil_image.width, pil_image.height)
            )
        col_lines = ocr_extractor._detect_table_lines(line_detect_src)
        h_lines = ocr_extractor._detect_horizontal_lines(line_detect_src)
        if col_lines and len(col_lines) >= 3:
            table_data = ocr_extractor._extract_with_column_lines(
                df_words, col_lines, cropped.height, h_lines=h_lines,
            )
            # Validate: if too many empty/garbage rows fall back to gap-based grouping
            filled_rows = [
                r for r in table_data
                if sum(1 for c in r if str(c).strip()) >= 2
            ]
            if len(filled_rows) < max(3, len(table_data) * 0.3):
                df_words = ocr_extractor._group_words_into_lines(df_words)
                table_data = ocr_extractor._group_lines_into_table(df_words)
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

        # Skip the top 15% of the page to avoid document-header lines
        skip_top = int(gray.shape[0] * 0.08)
        first_line_y = None
        for y in range(skip_top, gray.shape[0]):
            if row_sums[y] >= min_line_width:
                first_line_y = y
                break

        if first_line_y is None:
            return None

        # Find the topmost horizontal line (table top border)
        # We only crop if the line is far enough from top (there is significant header noise)
        # and if there is enough whitespace between top and the line
        if first_line_y > 80:
            # Check there are no strong horizontal lines in upper half
            upper_region = row_sums[:first_line_y]
            strong_lines_in_upper = sum(1 for val in upper_region if val >= min_line_width * 0.5)
            if strong_lines_in_upper <= 1:
                return max(first_line_y - 5, 0)

        # If table starts near top, don't crop
        return None

    def _filter_table_rows(self, table_data: List[List[str]]) -> List[List[str]]:
        filtered = []
        found_header = False
        col_number_re = re.compile(r'^\s*[\|\[\]{}]*\s*\d+\s*[\|\[\]{}]*\s*$')
        header_keywords = {
            "наименование", "стоимость", "работ", "затрат", "обоснование", "код",
            "ед", "изм", "кол", "п/п", "сумма", "цена", "индекс", "контракт",
            "инфляц", "период", "выполнен", "примечани", "макс", "начал",
            "номер", "позиция", "шифр", "ресурс", "труд", "материал", "цены",
        }
        noise_chars = {'|', '[', ']', '{', '}', '\u201a', "'", '-', '_', '°', ' ', ''}

        for row in table_data:
            non_empty = [c.strip() for c in row if c.strip()]
            if not non_empty:
                continue

            joined = ' '.join(non_empty).lower()
            joined_no_space = joined.replace(' ', '')

            # Detect header row
            is_header = False
            if len(non_empty) >= 2:
                is_header = any(kw in joined_no_space for kw in header_keywords)

            if not found_header:
                if is_header:
                    found_header = True
                    # Keep the header row so _detect_header_rows in
                    # build_raw_estimate_dataframe can use it
                    filtered.append(row)
                continue

            # Do NOT drop data rows that merely contain a keyword inside a value.
            # Only skip a row after the header if it looks like a *repeated* header
            # (most non-empty cells are short keyword fragments).
            if is_header:
                keyword_cell_count = sum(
                    1 for c in non_empty
                    if any(kw in c.lower().replace(' ', '') for kw in header_keywords)
                )
                if keyword_cell_count >= max(2, len(non_empty) * 0.6):
                    continue

            col_num_count = sum(
                1 for c in row if c.strip() and col_number_re.match(c.strip())
            )
            if col_num_count >= max(3, len(non_empty) // 2):
                continue

            if re.match(r'^\s*страница\s+\d+\s*$', joined, re.IGNORECASE):
                continue

            # Drop rows with generic Column_N placeholders
            if any(re.match(r'^Column_\d+$', c.strip(), re.IGNORECASE) for c in row):
                continue

            all_empty_or_noise = all(
                len(c.strip()) <= 1 or set(c.strip()).issubset(noise_chars)
                for c in row
            )
            if all_empty_or_noise:
                continue

            filtered.append(row)

        return filtered

    def build_raw_estimate_dataframe(self, raw_tables: List[List[List[Any]]]) -> pd.DataFrame:
        target_cols = self._canonical_col_count(raw_tables)
        if target_cols == 0:
            target_cols = max(
                (len(row) for t in raw_tables for row in t),
                default=0,
            )
        if target_cols == 0:
            return pd.DataFrame()

        normalized_rows: List[List[str]] = []
        for table in raw_tables:
            for row in table:
                if not any(
                    cell is not None and str(cell).strip() != ""
                    for cell in row
                ):
                    continue
                clean = [
                    OCRExtractor._clean_ocr_cell(
                        str(cell).replace("\n", " ").strip()
                    )
                    if cell is not None
                    else ""
                    for cell in row
                ]
                if len(clean) < target_cols:
                    gap = target_cols - len(clean)
                    if gap >= 3 and len(clean) >= 2:
                        clean = [clean[0]] + [""] * gap + clean[1:]
                    else:
                        clean += [""] * gap
                elif len(clean) > target_cols:
                    clean = clean[:target_cols]
                normalized_rows.append(clean)

        if not normalized_rows:
            return pd.DataFrame()

        # Try to detect real headers from the first rows
        header_rows, data_rows = self._detect_header_rows(normalized_rows, target_cols)
        if header_rows:
            columns = self._merge_header_rows(header_rows, target_cols)
        else:
            columns = [f"Column_{i + 1}" for i in range(target_cols)]

        # Filter garbage / phantom rows
        data_rows = [row for row in data_rows if not self._is_garbage_row(row)]

        if not data_rows:
            return pd.DataFrame()

        df = pd.DataFrame(data_rows, columns=columns)
        df = df.replace(r'^\s*$', None, regex=True)
        df = df.dropna(how='all').reset_index(drop=True)
        return df

    @staticmethod
    def _detect_header_rows(rows: List[List[Any]], num_cols: int) -> Tuple[List[List[str]], List[List[str]]]:
        if not rows:
            return [], []

        header_keywords = {
            "наименование", "стоимость", "работ", "затрат", "обоснование", "код",
            "ед", "изм", "кол", "п/п", "сумма", "цена", "индекс", "контракт",
            "инфляц", "период", "выполнен", "примечани", "макс", "начал",
            "номер", "позиция", "шифр", "ресурс", "труд", "материал", "цены",
        }

        max_header_scan = min(5, len(rows))
        header_row_indices = []

        for i, row in enumerate(rows[:max_header_scan]):
            str_row = [str(c).strip() if c is not None else "" for c in row]
            non_empty = [c for c in str_row if c]
            if len(non_empty) < 2:
                continue

            joined = " ".join(non_empty).lower()
            joined_no_space = joined.replace(" ", "")

            has_keyword = any(kw in joined_no_space for kw in header_keywords)
            if not has_keyword:
                continue

            # Reject if any of the first two cells looks like a position number
            # (digits like "1.", "12 ", or a single letter like "В")
            position_re = re.compile(r"^\s*[\dА-ЯA-Z][\d\s.)\-/\\]*\s*$")
            if any(position_re.match(c) for c in str_row[:2] if c):
                continue

            # Reject totals
            if any(kw in joined_no_space for kw in {"итого", "всего", "всего:", "итого:"}):
                continue

            # Reject rows that contain large monetary / numeric values
            # (headers rarely have 5+ digit numbers in data columns)
            has_large_number = any(
                len(re.sub(r"[^\d]", "", c)) >= 5
                for c in str_row[2:] if c
            )
            if has_large_number:
                continue

            header_row_indices.append(i)

        # Only accept consecutive header rows starting from index 0
        if header_row_indices and header_row_indices[0] == 0:
            consecutive = [0]
            for idx in header_row_indices[1:]:
                if idx == consecutive[-1] + 1:
                    consecutive.append(idx)
                else:
                    break
            header = [rows[i] for i in consecutive]
            data = rows[len(consecutive):]
            return header, data

        return [], rows

    @staticmethod
    def _merge_header_rows(header_rows: List[List[Any]], num_cols: int) -> List[str]:
        columns = [""] * num_cols
        for row in header_rows:
            for i, cell in enumerate(row[:num_cols]):
                cell = str(cell).strip() if cell is not None else ""
                if not cell:
                    continue
                if cell in columns[i]:
                    continue
                if columns[i]:
                    columns[i] += " " + cell
                else:
                    columns[i] = cell
        for i in range(num_cols):
            if not columns[i]:
                columns[i] = f"Column_{i + 1}"
        return columns

    @staticmethod
    def _is_garbage_row(row: List[Any]) -> bool:
        str_row = [str(c).strip() if c is not None else "" for c in row]
        non_empty = [c for c in str_row if c]
        if not non_empty:
            return True

        joined = " ".join(non_empty).lower()
        joined_no_space = joined.replace(" ", "")

        # Page numbers
        if re.match(r"^\s*страница\s+\d+\s*$", joined, re.IGNORECASE):
            return True

        # Generic column placeholders only
        if all(re.match(r"^Column_\d+$", c, re.IGNORECASE) for c in non_empty):
            return True

        # Only noise characters (allow single long cell if it contains real text)
        noise_chars = {"|", "[", "]", "{", "}", "\u201a", "'", "-", "_", "°", " ", ""}
        all_noise = all(
            len(c) <= 1 or set(c).issubset(noise_chars)
            for c in str_row
        )
        if all_noise:
            return True

        # A single cell that is just a short number (like "1", "2", "3") with no other text
        if len(non_empty) == 1:
            only = non_empty[0]
            if re.match(r"^\d+[.)]?$", only):
                return True

        # Rows that look like scattered random short tokens without any Russian word
        if len(non_empty) <= 2:
            has_russian = any(
                bool(re.search(r"[а-яёА-ЯЁ]", c)) for c in non_empty
            )
            has_meaningful_number = any(
                re.search(r"\d{3,}", c) for c in non_empty
            )
            if not has_russian and not has_meaningful_number:
                return True

        return False

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
