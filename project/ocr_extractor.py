"""
OCR-модуль для извлечения табличных данных из сканированных PDF.
"""

import os
import shutil
from pathlib import Path
from typing import List, Optional

import cv2
import numpy as np
import pandas as pd
import pytesseract
import pypdfium2 as pdfium
from PIL import Image, ImageEnhance

from config import MIN_TABLE_ROWS, DROP_EMPTY_ROWS, DROP_EMPTY_COLS

# Auto-detect tesseract path
_tesseract_path = shutil.which("tesseract")
if _tesseract_path:
    pytesseract.pytesseract.tesseract_cmd = _tesseract_path
elif os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# tessdata bundled with the project
tessdata_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tessdata")
if os.path.isdir(tessdata_dir):
    os.environ['TESSDATA_PREFIX'] = tessdata_dir + os.sep


class OCRExtractor:
    """Извлекает таблицы из сканированных PDF с помощью OCR."""

    def __init__(self, dpi: int = 300, lang: str = "rus"):
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
        image = self._normalize_orientation(image)
        img = image.convert('L')

        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.5)

        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(1.5)

        img_array = np.array(img)
        _, binary = cv2.threshold(img_array, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        img = Image.fromarray(binary)

        return img

    def _normalize_orientation(self, image: Image.Image) -> Image.Image:
        try:
            osd = pytesseract.image_to_osd(image, lang="osd")
        except Exception:
            try:
                sys_tessdata = "/usr/share/tesseract-ocr/4.00/tessdata"
                osd = pytesseract.image_to_osd(
                    image,
                    config=f'--tessdata-dir "{sys_tessdata}"',
                )
            except Exception:
                return image

        try:
            match = None
            for line in osd.splitlines():
                if "Rotate:" in line:
                    match = line.split(":", 1)[1].strip()
                    break
            if match is None:
                return image

            angle = int(match)
            if angle == 90:
                return image.rotate(270, expand=True)
            if angle == 180:
                return image.rotate(180, expand=True)
            if angle == 270:
                return image.rotate(90, expand=True)
        except Exception:
            return image

        return image

    def extract_tables_from_pdf(self, pdf_path: str) -> List[pd.DataFrame]:
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF-файл не найден: {pdf_path}")

        print(f"[OCR] Обработка скана: {Path(pdf_path).name}")

        try:
            pdf_doc = pdfium.PdfDocument(pdf_path)
            images = []
            for page in pdf_doc:
                bitmap = page.render(scale=self.dpi / 72)
                pil_image = bitmap.to_pil()
                images.append(pil_image)
            pdf_doc.close()
        except Exception as e:
            print(f"[OCR] Ошибка конвертации PDF в изображения: {e}")
            return []

        tables = []

        for page_num, image in enumerate(images, 1):
            print(f"[OCR] Обработка страницы {page_num}/{len(images)}")

            processed_image = self._preprocess_image(image)

            df = self._extract_table_from_image(processed_image, page_num)
            if df is not None and len(df) >= MIN_TABLE_ROWS:
                tables.append(df)

        return tables

    def extract_text_from_image(self, image: Image.Image) -> str:
        try:
            return pytesseract.image_to_string(
                image,
                lang=self.lang,
                config=r'--oem 1 --psm 6',
            )
        except Exception:
            return ""

    def extract_full_text_from_pdf(self, pdf_path: str) -> str:
        """Extract full text from a scanned PDF using OCR."""
        if not os.path.exists(pdf_path):
            return ""

        try:
            pdf_doc = pdfium.PdfDocument(pdf_path)
            chunks = []
            for page in pdf_doc:
                bitmap = page.render(scale=self.dpi / 72)
                pil_image = bitmap.to_pil()
                processed = self._preprocess_image(pil_image)
                text = pytesseract.image_to_string(
                    processed,
                    lang=self.lang,
                    config=r'--oem 1 --psm 6',
                )
                if text and text.strip():
                    chunks.append(text)
            pdf_doc.close()
            return "\n".join(chunks)
        except Exception:
            return ""

    def _extract_table_from_image(self, image: Image.Image, page_num: int) -> Optional[pd.DataFrame]:
        try:
            custom_config = r'--oem 1 --psm 6'
            ocr_data = pytesseract.image_to_data(
                image,
                lang=self.lang,
                output_type=pytesseract.Output.DATAFRAME,
                config=custom_config
            )
        except Exception as e:
            print(f"[OCR] Ошибка OCR на странице {page_num}: {e}")
            return None

        df_words = ocr_data[
            (ocr_data['text'].notna()) &
            (ocr_data['text'].astype(str).str.strip() != '') &
            (ocr_data['conf'] > 20)
        ].copy()

        if len(df_words) < 5:
            return None

        df_words['text'] = df_words['text'].astype(str)

        table_lines = self._detect_table_lines(image)

        if table_lines is not None and len(table_lines) > 1:
            table_data = self._extract_with_column_lines(df_words, table_lines, image.height)
        else:
            df_words = self._group_words_into_lines(df_words)
            table_data = self._group_lines_into_table(df_words)

        if not table_data or len(table_data) < MIN_TABLE_ROWS:
            return None

        max_cols = max(len(row) for row in table_data)
        normalized_data = []
        for row in table_data:
            normalized_row = row + [''] * (max_cols - len(row))
            normalized_data.append(normalized_row)

        header = [f"Column_{i+1}" for i in range(max_cols)]
        df = pd.DataFrame(normalized_data, columns=header)
        df = self._clean_dataframe(df)

        return df

    def _detect_table_lines(self, image: Image.Image) -> Optional[List[int]]:
        """Detect vertical column separator lines in the image."""
        try:
            gray = np.array(image.convert("L"))
            _, threshold = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, gray.shape[0] // 10)))
            vertical = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, kernel)

            col_sums = vertical.sum(axis=0) / 255
            height = gray.shape[0]
            min_line_height = height * 0.15

            lines: List[int] = []
            in_line = False
            line_start = 0

            for x, val in enumerate(col_sums):
                if val >= min_line_height:
                    if not in_line:
                        line_start = x
                        in_line = True
                else:
                    if in_line:
                        lines.append((line_start + x) // 2)
                        in_line = False

            if in_line:
                lines.append((line_start + len(col_sums)) // 2)

            if len(lines) >= 3:
                return sorted(lines)
            return None
        except Exception:
            return None

    def _extract_with_column_lines(
        self,
        df_words: pd.DataFrame,
        col_lines: List[int],
        img_height: int,
    ) -> List[List[str]]:
        """Use detected column lines to assign words to cells."""
        right_bound = max(
            int(df_words['left'].max() + df_words['width'].max() + 50),
            col_lines[-1] + 50,
        )
        col_boundaries = [0] + col_lines + [right_bound]

        df_words = df_words.sort_values(['top', 'left']).reset_index(drop=True)

        line_ids: List[int] = []
        current_line = 0
        prev_top = None
        y_threshold = max(12, int(img_height * 0.01))

        for _, row in df_words.iterrows():
            if prev_top is None or abs(row['top'] - prev_top) > y_threshold:
                current_line += 1
            line_ids.append(current_line)
            prev_top = row['top']

        df_words = df_words.copy()
        df_words['line_id'] = line_ids

        table_data: List[List[str]] = []
        for line_id in sorted(df_words['line_id'].unique()):
            line_words = df_words[df_words['line_id'] == line_id].sort_values('left')

            cells = [''] * (len(col_boundaries) - 1)
            for _, word in line_words.iterrows():
                word_center = word['left'] + word['width'] / 2
                for ci in range(len(col_boundaries) - 1):
                    if col_boundaries[ci] <= word_center < col_boundaries[ci + 1]:
                        if cells[ci]:
                            cells[ci] += ' ' + word['text']
                        else:
                            cells[ci] = word['text']
                        break

            if any(c.strip() for c in cells):
                table_data.append(cells)

        return table_data

    def _group_words_into_lines(self, df_words: pd.DataFrame, y_threshold: int = 15) -> pd.DataFrame:
        df_words = df_words.sort_values(['top', 'left']).reset_index(drop=True)

        line_ids: List[int] = []
        current_line = 0
        prev_top = None

        for _, row in df_words.iterrows():
            if prev_top is None or abs(row['top'] - prev_top) > y_threshold:
                current_line += 1
            line_ids.append(current_line)
            prev_top = row['top']

        df_words = df_words.copy()
        df_words['line_id'] = line_ids
        return df_words

    def _group_lines_into_table(self, df_words: pd.DataFrame, x_threshold: int = 30) -> List[List[str]]:
        table_data: List[List[str]] = []

        for line_id in sorted(df_words['line_id'].unique()):
            line_words = df_words[df_words['line_id'] == line_id].sort_values('left')

            cells: List[str] = []
            current_cell: List[str] = []
            prev_right = None

            for _, word in line_words.iterrows():
                if prev_right is not None and (word['left'] - prev_right) > x_threshold:
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

    def _is_number(self, value: str) -> bool:
        cleaned = value.replace(" ", "").replace(",", ".").replace("\xa0", "")
        try:
            float(cleaned)
            return True
        except ValueError:
            return False

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
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
