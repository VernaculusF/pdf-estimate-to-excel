from typing import List, Optional

import pandas as pd
import pdfplumber
import pypdfium2 as pdfium
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from config import PAGE_SHEET_PREFIX
from quality_report import normalize_visible_text


META_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
META_FONT = Font(bold=True)
WRAP_TOP = Alignment(vertical="top", wrap_text=True)
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)


class DocumentExporter:
    def __init__(self, ocr_extractor):
        self.ocr_extractor = ocr_extractor

    def build_pages(self, pdf_path: str) -> List[dict]:
        pages = []
        pdf_file = str(pdf_path)

        with pdfplumber.open(pdf_file) as pdf:
            pdf_doc = pdfium.PdfDocument(pdf_file)
            try:
                for page_number, page in enumerate(pdf.pages, start=1):
                    image = pdf_doc[page_number - 1].render(
                        scale=self.ocr_extractor.dpi / 72
                    ).to_pil()
                    pages.append(self._build_single_page(page, image, page_number))
            finally:
                pdf_doc.close()

        return pages

    def append_page_sheets(self, excel_path: str, pages: List[dict]) -> None:
        workbook = load_workbook(excel_path)

        for page in pages:
            sheet_name = page["sheet_name"]
            if sheet_name in workbook.sheetnames:
                del workbook[sheet_name]

            dataframe = self._normalize_dataframe_for_sheet(page["dataframe"])
            worksheet = workbook.create_sheet(sheet_name)
            worksheet["A1"] = "Method"
            worksheet["B1"] = page["method"]
            worksheet["A2"] = "Page"
            worksheet["B2"] = page["page_number"]

            for row_index, row in enumerate(
                dataframe_to_rows(dataframe, index=False, header=False),
                start=3,
            ):
                for col_index, value in enumerate(row, start=1):
                    worksheet.cell(row=row_index, column=col_index, value=value)

            self._style_page_sheet(worksheet, dataframe)

        workbook.save(excel_path)
        workbook.close()

    def raw_table_to_dataframe(self, table: List[List[Optional[str]]]) -> pd.DataFrame:
        max_cols = max(len(row) for row in table) if table else 0
        normalized_rows = []
        for row in table:
            normalized = [(cell or "").replace("\n", " ").strip() for cell in row]
            normalized += [""] * (max_cols - len(normalized))
            normalized_rows.append(normalized)

        columns = [f"Column_{index + 1}" for index in range(max_cols)]
        return pd.DataFrame(normalized_rows, columns=columns)

    def raw_tables_to_dataframe(self, tables: List[List[List[Optional[str]]]]) -> pd.DataFrame:
        if not tables:
            return pd.DataFrame()

        max_cols = max((len(row) for table in tables for row in table), default=0)
        columns = [f"Column_{index + 1}" for index in range(max_cols)]
        normalized_rows = []

        for table_index, table in enumerate(tables):
            if table_index > 0 and normalized_rows:
                normalized_rows.append([""] * max_cols)

            for row in table:
                normalized = [(cell or "").replace("\n", " ").strip() for cell in row]
                normalized += [""] * (max_cols - len(normalized))
                normalized_rows.append(normalized)

        return pd.DataFrame(normalized_rows, columns=columns)

    def text_to_dataframe(self, text: str) -> pd.DataFrame:
        lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
        return pd.DataFrame({"Column_1": lines or [""]})

    def _build_single_page(self, page, image, page_number: int) -> dict:
        text = page.extract_text(layout=True) or page.extract_text() or ""
        raw_tables = page.extract_tables()

        if raw_tables:
            page_df = self.raw_tables_to_dataframe(raw_tables)
            return self._page_result(page_number, "pdf_table", page_df)

        if normalize_visible_text(text):
            return self._page_result(page_number, "pdf_text", self.text_to_dataframe(text))

        processed = self.ocr_extractor._preprocess_image(image)
        ocr_table = self.ocr_extractor._extract_table_from_image(processed, page_number)
        if self._is_usable_ocr_table(ocr_table):
            return self._page_result(page_number, "ocr_table", ocr_table)

        ocr_text = self.ocr_extractor.extract_text_from_image(processed)
        return self._page_result(page_number, "ocr_text", self.text_to_dataframe(ocr_text))

    def _page_result(self, page_number: int, method: str, dataframe: pd.DataFrame) -> dict:
        return {
            "page_number": page_number,
            "sheet_name": f"{PAGE_SHEET_PREFIX}_{page_number:02d}",
            "method": method,
            "dataframe": dataframe.fillna(""),
        }

    def _normalize_dataframe_for_sheet(self, dataframe: pd.DataFrame) -> pd.DataFrame:
        normalized = dataframe.copy()
        headers = []
        seen = {}

        for index, column in enumerate(normalized.columns, start=1):
            name = str(column).strip() if column is not None else ""
            if not name:
                name = f"Column_{index}"

            count = seen.get(name, 0)
            seen[name] = count + 1
            headers.append(name if count == 0 else f"{name}_{count + 1}")

        normalized.columns = headers
        return normalized.fillna("")

    def _is_usable_ocr_table(self, dataframe: Optional[pd.DataFrame]) -> bool:
        if dataframe is None or dataframe.empty:
            return False

        normalized = self._normalize_dataframe_for_sheet(dataframe)
        filled_per_row = normalized.astype(str).apply(
            lambda row: sum(1 for value in row if value.strip()),
            axis=1,
        )
        if filled_per_row.empty:
            return False

        avg_filled = float(filled_per_row.mean())
        max_filled = int(filled_per_row.max())
        column_count = len(normalized.columns)
        non_empty_values = [
            value.strip()
            for value in normalized.astype(str).to_numpy().flatten().tolist()
            if value and value.strip()
        ]
        avg_value_length = (
            sum(len(value) for value in non_empty_values) / len(non_empty_values)
            if non_empty_values
            else 0.0
        )

        if column_count >= 8 and avg_filled < 2.4:
            return False
        if column_count >= 12 and max_filled < 5:
            return False
        if column_count >= 8 and avg_filled < 3.2 and avg_value_length < 5.0:
            return False

        return True

    def _style_page_sheet(self, worksheet, dataframe: pd.DataFrame) -> None:
        worksheet.freeze_panes = "A3"

        for cell_ref in ("A1", "A2"):
            worksheet[cell_ref].font = META_FONT
            worksheet[cell_ref].fill = META_FILL
            worksheet[cell_ref].border = THIN_BORDER
            worksheet[cell_ref].alignment = CENTER_WRAP

        for cell_ref in ("B1", "B2"):
            worksheet[cell_ref].border = THIN_BORDER
            worksheet[cell_ref].alignment = WRAP_TOP

        for row in worksheet.iter_rows(min_row=3):
            for cell in row:
                cell.border = THIN_BORDER
                cell.alignment = WRAP_TOP

        self._auto_adjust_columns(worksheet, dataframe)

    def _auto_adjust_columns(self, worksheet, dataframe: pd.DataFrame) -> None:
        for column_index, column_name in enumerate(dataframe.columns, start=1):
            values = [str(column_name)]
            for value in dataframe.iloc[:, column_index - 1].tolist():
                if value is not None:
                    values.append(str(value))

            max_length = max((len(value) for value in values), default=10)
            letter = get_column_letter(column_index)

            if len(dataframe.columns) == 1:
                width = min(max(max_length + 4, 40), 120)
            elif column_index <= 2:
                width = min(max(max_length + 3, 12), 28)
            else:
                width = min(max(max_length + 2, 10), 22)

            worksheet.column_dimensions[letter].width = width
