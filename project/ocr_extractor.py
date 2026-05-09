"""
OCR-модуль для извлечения табличных данных из сканированных PDF.
"""

import os
import shutil
import tempfile
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
tessdata_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tessdata_new")
if os.path.isdir(tessdata_dir):
    os.environ['TESSDATA_PREFIX'] = tessdata_dir + os.sep


class OCRExtractor:
    """Извлекает таблицы из сканированных PDF с помощью OCR."""

    def __init__(self, dpi: int = 300, lang: str = "rus"):
        self.dpi = dpi
        self.lang = lang

    def normalize_pdf_orientation(self, pdf_path: str) -> Optional[str]:
        """Detect per-page orientation via Tesseract OSD and physically rotate scans.
        For text-based PDFs (majority of pages have text layer) returns None
        so pdfplumber keeps working with original coordinates.
        For scanned PDFs renders every page, rotates the image physically,
        and writes a new temporary PDF."""
        try:
            import pdfplumber

            with pdfplumber.open(pdf_path) as pdf:
                pages_to_check = min(3, len(pdf.pages))
                text_pages = 0
                for i in range(pages_to_check):
                    page = pdf.pages[i]
                    if page.extract_text():
                        text_pages += 1
                # Text-based PDF – do not touch it (pdfplumber handles rotation itself)
                if text_pages >= (pages_to_check // 2 + 1):
                    return None

            pdf_doc = pdfium.PdfDocument(pdf_path)
            images: List[Image.Image] = []
            needs_save = False
            for i in range(len(pdf_doc)):
                page = pdf_doc[i]
                bitmap = page.render(scale=300 / 72)
                pil_image = bitmap.to_pil()
                gray = pil_image.convert('L')
                try:
                    osd = pytesseract.image_to_osd(gray, output_type=pytesseract.Output.DICT)
                    orientation = osd.get("orientation", 0)
                    confidence = osd.get("orientation_conf", 0)
                    if confidence >= 5 and orientation in (90, 180, 270):
                        needs_save = True
                        if orientation == 90:
                            pil_image = pil_image.rotate(270, expand=True)
                        elif orientation == 270:
                            pil_image = pil_image.rotate(90, expand=True)
                        elif orientation == 180:
                            pil_image = pil_image.rotate(180, expand=True)
                except Exception:
                    pass
                images.append(pil_image)
            pdf_doc.close()

            if needs_save and images:
                fd, tmp_path = tempfile.mkstemp(suffix='.pdf')
                os.close(fd)
                first = images[0]
                rest = images[1:] if len(images) > 1 else []
                first.save(tmp_path, save_all=True, append_images=rest, resolution=300.0)
                return tmp_path
        except Exception as exc:
            print(f"[OCR] Failed to normalize orientation: {exc}")
        return None

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
        """Light preprocessing: grayscale, contrast, sharpness.
        We intentionally skip OTSU binarisation because it destroys
        faint text on scanned estimates and hurts Tesseract accuracy."""
        img = image.convert('L')

        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.5)

        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(1.5)

        return img

    @staticmethod
    def render_pdf_page(pdf_path: str, page_index: int, dpi: int = 300) -> Image.Image:
        """Render a PDF page to PIL Image, respecting PDF rotation metadata."""
        pdf_doc = pdfium.PdfDocument(pdf_path)
        page = pdf_doc[page_index]
        bitmap = page.render(scale=dpi / 72)
        pil_image = bitmap.to_pil()
        pdf_doc.close()
        return pil_image

    @staticmethod
    def get_pdf_page_count(pdf_path: str) -> int:
        """Return the number of pages in a PDF."""
        pdf_doc = pdfium.PdfDocument(pdf_path)
        count = len(pdf_doc)
        pdf_doc.close()
        return count

    def extract_tables_from_pdf(self, pdf_path: str) -> List[pd.DataFrame]:
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF-файл не найден: {pdf_path}")

        print(f"[OCR] Обработка скана: {Path(pdf_path).name}")

        try:
            page_count = self.get_pdf_page_count(pdf_path)
        except Exception as e:
            print(f"[OCR] Ошибка открытия PDF: {e}")
            return []

        tables = []

        for page_index in range(page_count):
            page_num = page_index + 1
            print(f"[OCR] Обработка страницы {page_num}/{page_count}")
            try:
                original_image = self.render_pdf_page(pdf_path, page_index, self.dpi)
            except Exception as e:
                print(f"[OCR] Ошибка рендеринга страницы {page_num}: {e}")
                continue

            processed_image = self._preprocess_image(original_image)

            df = self._extract_table_from_image(
                processed_image, page_num, original_image=original_image,
            )
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
            page_count = self.get_pdf_page_count(pdf_path)
            chunks = []
            for page_index in range(page_count):
                pil_image = self.render_pdf_page(pdf_path, page_index, self.dpi)
                processed = self._preprocess_image(pil_image)
                text = pytesseract.image_to_string(
                    processed,
                    lang=self.lang,
                    config=r'--oem 1 --psm 6',
                )
                if text and text.strip():
                    chunks.append(text)
            return "\n".join(chunks)
        except Exception:
            return ""

    def _detect_table_top_from_original(
        self, original_image: Image.Image,
    ) -> Optional[int]:
        """Detect the top of the table area using vertical lines in the
        original (non-binarized) rendered image."""
        try:
            gray = np.array(original_image.convert("L"))
            _, threshold = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

            h_kernel = cv2.getStructuringElement(
                cv2.MORPH_RECT, (max(80, gray.shape[1] // 5), 1),
            )
            horizontal = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, h_kernel)

            v_kernel = cv2.getStructuringElement(
                cv2.MORPH_RECT, (1, max(40, gray.shape[0] // 15)),
            )
            vertical = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, v_kernel)

            combined = cv2.bitwise_or(horizontal, vertical)
            row_sums = combined.sum(axis=1) / 255
            width = gray.shape[1]
            height = gray.shape[0]
            min_line_width = width * 0.25

            # Skip the very top of the page (headers / title block)
            min_y = max(40, int(height * 0.08))
            first_line_y = None
            for y in range(min_y, len(row_sums)):
                if row_sums[y] >= min_line_width:
                    first_line_y = y
                    break

            if first_line_y is None:
                return None

            # Only crop if table starts well below top margin
            if first_line_y > min_y:
                upper_region = row_sums[:first_line_y]
                strong_lines_in_upper = sum(
                    1 for val in upper_region if val >= min_line_width * 0.5
                )
                if strong_lines_in_upper <= 1:
                    return max(first_line_y - 5, 0)
        except Exception:
            pass
        return None

    @staticmethod
    def _find_table_top_by_keywords(df_words: pd.DataFrame) -> Optional[int]:
        """Use OCR word positions to find the first row that looks like a table header."""
        header_keywords = {
            "№", "п/п", "наименование", "обоснование", "код", "ед", "изм",
            "кол", "цена", "сумма", "стоимость", "работ", "ресур", "затрат",
            "примечание", "формула", "позиция", "шифр", "смета", "объект",
        }
        if df_words.empty:
            return None
        words = df_words.sort_values("top").reset_index(drop=True)
        y_threshold = max(8, int(words["top"].max() * 0.005))
        line_id = 0
        prev_top = None
        line_tops: List[int] = []
        line_texts: List[str] = []
        for _, row in words.iterrows():
            if prev_top is None or abs(row["top"] - prev_top) > y_threshold:
                line_id += 1
                line_tops.append(row["top"])
                line_texts.append(row["text"].lower())
            else:
                line_texts[-1] += " " + row["text"].lower()
            prev_top = row["top"]

        for top, joined in zip(line_tops, line_texts):
            match_score = sum(1 for kw in header_keywords if kw in joined)
            if match_score >= 2:
                return max(0, int(top - 10))
        return None

    def _extract_table_from_image(
        self,
        image: Image.Image,
        page_num: int,
        original_image: Optional[Image.Image] = None,
    ) -> Optional[pd.DataFrame]:
        table_top = None
        line_detect_src = original_image if original_image is not None else image
        if original_image is not None:
            table_top = self._detect_table_top_from_original(original_image)

        table_lines = self._detect_table_lines(line_detect_src)

        ocr_image = image
        # Only crop if there is a lot of whitespace/noise above the table
        if table_top is not None and table_top > 80:
            ocr_image = image.crop((0, table_top, image.width, image.height))
            if table_lines:
                table_lines = self._detect_table_lines(
                    line_detect_src.crop(
                        (0, table_top, line_detect_src.width, line_detect_src.height)
                    )
                )

        try:
            custom_config = r'--oem 1 --psm 6'
            ocr_data = pytesseract.image_to_data(
                ocr_image,
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

        # Fallback: keyword-based table-top detection when line detection failed
        if table_top is None:
            keyword_top = self._find_table_top_by_keywords(df_words)
            if keyword_top is not None and keyword_top > 80:
                table_top = keyword_top
                ocr_image = image.crop((0, table_top, image.width, image.height))
                # Re-run OCR on cropped region for cleaner results
                try:
                    ocr_data = pytesseract.image_to_data(
                        ocr_image,
                        lang=self.lang,
                        output_type=pytesseract.Output.DATAFRAME,
                        config=custom_config
                    )
                    df_words = ocr_data[
                        (ocr_data['text'].notna()) &
                        (ocr_data['text'].astype(str).str.strip() != '') &
                        (ocr_data['conf'] > 20)
                    ].copy()
                    df_words['text'] = df_words['text'].astype(str)
                except Exception:
                    pass

        h_lines = self._detect_horizontal_lines(
            line_detect_src.crop(
                (0, table_top, line_detect_src.width, line_detect_src.height)
            ) if table_top and table_top > 20 else line_detect_src
        )

        if table_lines is not None and len(table_lines) > 1:
            table_data = self._extract_with_column_lines(
                df_words, table_lines, ocr_image.height, h_lines=h_lines,
            )
            # Validate column-line extraction: reject if too many garbage/empty rows
            if table_data and len(table_data) >= MIN_TABLE_ROWS:
                filled_rows = [
                    r for r in table_data
                    if sum(1 for c in r if str(c).strip()) >= 2
                ]
                if len(filled_rows) < max(3, MIN_TABLE_ROWS):
                    table_data = None
            if not table_data:
                df_words = self._group_words_into_lines(df_words)
                table_data = self._group_lines_into_table(df_words)
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

        # Try to detect real headers from the first rows
        from extractor import SmetaExtractor
        header_rows, data_rows = SmetaExtractor._detect_header_rows(normalized_data, max_cols)
        if header_rows:
            columns = SmetaExtractor._merge_header_rows(header_rows, max_cols)
        else:
            columns = [f"Column_{i+1}" for i in range(max_cols)]
        df = pd.DataFrame(data_rows, columns=columns)
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

    def _detect_horizontal_lines(self, image: Image.Image) -> Optional[List[int]]:
        """Detect horizontal row separator lines in the image."""
        try:
            gray = np.array(image.convert("L"))
            _, threshold = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

            kernel = cv2.getStructuringElement(
                cv2.MORPH_RECT, (max(80, gray.shape[1] // 5), 1),
            )
            horizontal = cv2.morphologyEx(threshold, cv2.MORPH_OPEN, kernel)

            row_sums = horizontal.sum(axis=1) / 255
            width = gray.shape[1]
            min_line_width = width * 0.20

            lines: List[int] = []
            in_line = False
            line_start = 0

            for y, val in enumerate(row_sums):
                if val >= min_line_width:
                    if not in_line:
                        line_start = y
                        in_line = True
                else:
                    if in_line:
                        lines.append((line_start + y) // 2)
                        in_line = False

            if in_line:
                lines.append((line_start + len(row_sums)) // 2)

            if len(lines) >= 2:
                return sorted(lines)
            return None
        except Exception:
            return None

    def _extract_with_column_lines(
        self,
        df_words: pd.DataFrame,
        col_lines: List[int],
        img_height: int,
        h_lines: Optional[List[int]] = None,
    ) -> List[List[str]]:
        """Use detected column and horizontal lines to assign words to cells.

        When horizontal lines are available, words between two consecutive
        horizontal lines are grouped into the same table row regardless of
        their y-position.  This correctly handles multi-line cell content
        (e.g. long descriptions that wrap inside a cell).
        """
        right_bound = max(
            int(df_words['left'].max() + df_words['width'].max() + 50),
            col_lines[-1] + 50,
        )
        col_boundaries = [0] + col_lines + [right_bound]
        num_cols = len(col_boundaries) - 1

        df_words = df_words.sort_values(['top', 'left']).reset_index(drop=True)

        if h_lines and len(h_lines) >= 2:
            row_boundaries = h_lines + [img_height + 10]

            table_data: List[List[str]] = []
            for ri in range(len(row_boundaries) - 1):
                y_top = row_boundaries[ri]
                y_bot = row_boundaries[ri + 1]
                band = df_words[
                    (df_words['top'] + df_words['height'] / 2 > y_top)
                    & (df_words['top'] + df_words['height'] / 2 <= y_bot)
                ].sort_values(['top', 'left'])

                if band.empty:
                    continue

                cells = [''] * num_cols
                for _, word in band.iterrows():
                    word_center = word['left'] + word['width'] / 2
                    for ci in range(num_cols):
                        if col_boundaries[ci] <= word_center < col_boundaries[ci + 1]:
                            if cells[ci]:
                                cells[ci] += ' ' + word['text']
                            else:
                                cells[ci] = word['text']
                            break

                if any(c.strip() for c in cells):
                    table_data.append(cells)

            return table_data

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

        table_data = []
        for line_id in sorted(df_words['line_id'].unique()):
            line_words = df_words[df_words['line_id'] == line_id].sort_values('left')

            cells = [''] * num_cols
            for _, word in line_words.iterrows():
                word_center = word['left'] + word['width'] / 2
                for ci in range(num_cols):
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

    def _group_lines_into_table(self, df_words: pd.DataFrame, x_threshold: Optional[int] = None) -> List[List[str]]:
        table_data: List[List[str]] = []

        # Compute adaptive x-threshold from word gaps across all lines
        if x_threshold is None:
            all_gaps: List[int] = []
            for line_id in sorted(df_words['line_id'].unique()):
                line_words = df_words[df_words['line_id'] == line_id].sort_values('left')
                prev_r = None
                for _, word in line_words.iterrows():
                    if prev_r is not None:
                        gap = int(word['left'] - prev_r)
                        if gap > 0:
                            all_gaps.append(gap)
                    prev_r = word['left'] + word['width']
            if all_gaps:
                # Use a high percentile to separate real columns from intra-cell word gaps
                img_width = int(df_words['left'].max() + df_words['width'].max())
                p80 = int(np.percentile(all_gaps, 80))
                p90 = int(np.percentile(all_gaps, 90))
                # Choose a threshold between P80 and P90, but at least 80 px
                x_threshold = max(80, p80 + (p90 - p80) // 3, int(img_width * 0.012))
            else:
                x_threshold = 80

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

    @staticmethod
    def _clean_ocr_cell(text: str) -> str:
        """Remove common OCR artifacts from a single cell value."""
        if not text or not isinstance(text, str):
            return text
        import re as _re
        cleaned = text.strip()
        cleaned = cleaned.strip('|[]{}‚')
        cleaned = _re.sub(r'\|+', '', cleaned)
        cleaned = _re.sub(r'[\[\]{}]', '', cleaned)
        cleaned = _re.sub(r'\s{2,}', ' ', cleaned)
        cleaned = cleaned.strip()
        if not cleaned or all(c in '|[]{}‚\' °' for c in cleaned):
            return ''
        return cleaned

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df

        for col in df.columns:
            try:
                df[col] = df[col].apply(
                    lambda x: self._clean_ocr_cell(str(x)) if pd.notna(x) else x
                )
            except Exception:
                pass

        df = df.replace(r'^\s*$', None, regex=True)

        if DROP_EMPTY_ROWS:
            df = df.dropna(how='all')

        if DROP_EMPTY_COLS:
            df = df.dropna(axis=1, how='all')

        # Remove garbage rows (single-char noise, page numbers, generic placeholders)
        if not df.empty:
            from extractor import SmetaExtractor
            mask = df.apply(
                lambda row: not SmetaExtractor._is_garbage_row(
                    [str(c) if pd.notna(c) else "" for c in row]
                ),
                axis=1,
            )
            df = df[mask]

        return df
