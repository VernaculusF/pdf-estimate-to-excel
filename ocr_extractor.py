"""
OCR-модуль для извлечения табличных данных из сканированных PDF.
"""

import os
from pathlib import Path
from typing import List, Optional

import pandas as pd
import pytesseract
import pypdfium2 as pdfium
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np

from config import MIN_TABLE_ROWS, MIN_TABLE_COLS, DROP_EMPTY_ROWS, DROP_EMPTY_COLS

# Указываем путь к Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


class OCRExtractor:
    """Извлекает таблицы из сканированных PDF с помощью OCR."""

    def __init__(self, dpi: int = 400, lang: str = "rus"):
        self.dpi = dpi
        self.lang = lang

    def is_scanned_pdf(self, pdf_path: str, sample_pages: int = 2) -> bool:
        """
        Проверяет, является ли PDF сканированным (без текстового слоя).
        """
        import pdfplumber
        
        with pdfplumber.open(pdf_path) as pdf:
            pages_to_check = min(sample_pages, len(pdf.pages))
            scanned_pages = 0
            
            for i in range(pages_to_check):
                page = pdf.pages[i]
                text = page.extract_text()
                words = page.extract_words()
                images = page.images
                
                if (not text or len(words) == 0) and len(images) > 0:
                    scanned_pages += 1
            
            return scanned_pages >= pages_to_check // 2 + 1

    def _preprocess_image(self, image: Image.Image) -> Image.Image:
        """
        Предобработка изображения для улучшения OCR.
        """
        # Конвертируем в grayscale
        img = image.convert('L')
        
        # Увеличиваем контраст
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)
        
        # Увеличиваем резкость
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0)
        
        # Применяем фильтр для удаления шума
        img = img.filter(ImageFilter.MedianFilter(size=3))
        
        # Бинаризация с помощью Otsu thresholding
        import cv2
        img_array = np.array(img)
        _, binary = cv2.threshold(img_array, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        img = Image.fromarray(binary)
        
        return img

    def extract_tables_from_pdf(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Извлекает таблицы из сканированного PDF.
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF-файл не найден: {pdf_path}")

        print(f"[OCR] Обработка скана: {Path(pdf_path).name}")
        
        # Конвертируем PDF в изображения с помощью pypdfium2
        try:
            pdf_doc = pdfium.PdfDocument(pdf_path)
            images = []
            for page in pdf_doc:
                bitmap = page.render(scale=self.dpi / 72)
                pil_image = bitmap.to_pil()
                images.append(pil_image)
        except Exception as e:
            print(f"[OCR] Ошибка конвертации PDF в изображения: {e}")
            return []
        
        tables = []
        
        for page_num, image in enumerate(images, 1):
            print(f"[OCR] Обработка страницы {page_num}/{len(images)}")
            
            # Предобработка изображения
            processed_image = self._preprocess_image(image)
            
            # Извлекаем таблицу со страницы
            df = self._extract_table_from_image(processed_image, page_num)
            if df is not None and len(df) >= MIN_TABLE_ROWS:
                tables.append(df)
        
        return tables

    def _extract_table_from_image(self, image: Image.Image, page_num: int) -> Optional[pd.DataFrame]:
        """
        Извлекает таблицу из изображения с помощью OCR.
        """
        try:
            custom_config = r'--oem 1 --psm 3'
            ocr_data = pytesseract.image_to_data(
                image, 
                lang=self.lang, 
                output_type=pytesseract.Output.DATAFRAME,
                config=custom_config
            )
        except Exception as e:
            print(f"[OCR] Ошибка OCR на странице {page_num}: {e}")
            return None
        
        # Фильтруем только распознанный текст с достаточной уверенностью
        df_words = ocr_data[
            (ocr_data['text'].notna()) & 
            (ocr_data['text'].str.strip() != '') & 
            (ocr_data['conf'] > 25)  # Уверенность > 25%
        ].copy()
        
        if len(df_words) < 10:
            return None
        
        # Группируем слова в строки по координате top (y)
        df_words = self._group_words_into_lines(df_words)
        
        # Группируем строки в таблицу по координате left (x)
        table_data = self._group_lines_into_table(df_words)
        
        if not table_data or len(table_data) < MIN_TABLE_ROWS:
            return None
        
        # Создаем DataFrame
        header = table_data[0]
        data = table_data[1:]
        
        # Если первая строка не похожа на заголовок, используем ее как данные
        if not self._looks_like_header(header):
            data = table_data
            header = [f"Column_{i+1}" for i in range(len(table_data[0]))]
        
        # Уравниваем длину строк
        max_cols = max(len(row) for row in table_data)
        normalized_data = []
        for row in data:
            normalized_row = row + [''] * (max_cols - len(row))
            normalized_data.append(normalized_row)
        
        header = header + [''] * (max_cols - len(header))
        
        df = pd.DataFrame(normalized_data, columns=header[:max_cols])
        
        # Очистка
        df = self._clean_dataframe(df)
        
        return df

    def _group_words_into_lines(self, df_words: pd.DataFrame, y_threshold: int = 15) -> pd.DataFrame:
        """
        Группирует слова в строки на основе близости по координате Y.
        """
        # Сортируем по Y, затем по X
        df_words = df_words.sort_values(['top', 'left']).reset_index(drop=True)
        
        line_ids = []
        current_line = 0
        prev_top = None
        
        for _, row in df_words.iterrows():
            if prev_top is None or abs(row['top'] - prev_top) > y_threshold:
                current_line += 1
            line_ids.append(current_line)
            prev_top = row['top']
        
        df_words['line_id'] = line_ids
        return df_words

    def _group_lines_into_table(self, df_words: pd.DataFrame, x_threshold: int = 30) -> List[List[str]]:
        """
        Группирует слова в строки в таблицу на основе координат X.
        """
        table_data = []
        
        for line_id in sorted(df_words['line_id'].unique()):
            line_words = df_words[df_words['line_id'] == line_id].sort_values('left')
            
            # Объединяем слова в ячейки на основе близости по X
            cells = []
            current_cell = []
            prev_right = None
            
            for _, word in line_words.iterrows():
                if prev_right is not None and (word['left'] - prev_right) > x_threshold:
                    # Начинаем новую ячейку
                    if current_cell:
                        cells.append(' '.join(current_cell))
                    current_cell = [word['text']]
                else:
                    current_cell.append(word['text'])
                
                prev_right = word['left'] + word['width']
            
            if current_cell:
                cells.append(' '.join(current_cell))
            
            if cells:
                table_data.append(cells)
        
        return table_data

    def _looks_like_header(self, row: List[str]) -> bool:
        """Проверяет, похожа ли строка на заголовок таблицы."""
        if not row or len(row) < MIN_TABLE_COLS:
            return False
        
        header_keywords = [
            "наименование", "количество", "цена", "сумма", "ед.", 
            "изм", "№", "п/п", "номер", "стоимость", "итого"
        ]
        
        row_text = " ".join(str(cell).lower() for cell in row if cell)
        
        if any(keyword in row_text for keyword in header_keywords):
            return True
        
        text_count = sum(1 for cell in row if cell and not self._is_number(str(cell)))
        return text_count >= len(row) * 0.7

    def _is_number(self, value: str) -> bool:
        """Проверяет, является ли строка числом."""
        cleaned = value.replace(" ", "").replace(",", ".").replace("\xa0", "")
        try:
            float(cleaned)
            return True
        except ValueError:
            return False

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Очищает DataFrame."""
        if df.empty:
            return df
        
        df = df.replace(r'^\s*$', None, regex=True)
        
        if DROP_EMPTY_ROWS:
            df = df.dropna(how='all')
        
        if DROP_EMPTY_COLS:
            df = df.dropna(axis=1, how='all')
        
        for col in df.columns:
            try:
                if pd.api.types.is_object_dtype(df[col]):
                    df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
            except Exception:
                pass
        
        return df
