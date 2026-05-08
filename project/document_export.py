from io import BytesIO
from typing import Optional

import pdfplumber
import pypdfium2 as pdfium
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from config import HEADER_SHEET_NAME


class DocumentExporter:
    def __init__(self, ocr_extractor, header_render_dpi: int = 180):
        self.ocr_extractor = ocr_extractor
        self.header_render_dpi = header_render_dpi

    def append_header_sheet(self, excel_path: str, pdf_path: str) -> None:
        header_image = self.build_header_image(pdf_path)
        if header_image is None:
            return

        workbook = load_workbook(excel_path)
        if HEADER_SHEET_NAME in workbook.sheetnames:
            del workbook[HEADER_SHEET_NAME]

        worksheet = workbook.create_sheet(HEADER_SHEET_NAME, 0)
        worksheet.sheet_view.showGridLines = False

        image_buffer = BytesIO()
        image_buffer.name = "header.png"
        header_image.save(image_buffer, format="PNG")
        image_buffer.seek(0)

        worksheet.add_image(XLImage(image_buffer), "A1")
        workbook.save(excel_path)
        workbook.close()

    def build_header_image(self, pdf_path: str):
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return None

            first_page = pdf.pages[0]
            render_scale = self.header_render_dpi / 72
            pdf_doc = pdfium.PdfDocument(pdf_path)
            try:
                image = pdf_doc[0].render(scale=render_scale).to_pil()
            finally:
                pdf_doc.close()

            bottom = self._detect_header_bottom(first_page, image.height)
            bottom = max(200, min(bottom, image.height))
            return image.crop((0, 0, image.width, bottom))

    def _detect_header_bottom(self, first_page, image_height: int) -> int:
        default_bottom = int(image_height * 0.46)

        try:
            tables = first_page.find_tables()
        except Exception:
            tables = []

        if not tables:
            return default_bottom

        first_table_top = min(table.bbox[1] for table in tables)
        page_height = max(float(first_page.height), 1.0)
        top_ratio = first_table_top / page_height
        margin_ratio = 0.015
        return int(image_height * max(top_ratio - margin_ratio, 0.15))
