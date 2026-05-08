"""
Модуль конвертации и сохранения данных смет в формат Excel.
"""

import os
from typing import List, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from config import DEFAULT_SHEET_NAME


class SmetaConverter:
    """Конвертирует DataFrame смет в профессионально оформленный Excel-файл."""

    def __init__(self):
        self.sheet_name = DEFAULT_SHEET_NAME
        # Стили
        self.header_font = Font(bold=True, size=11, color="FFFFFF")
        self.header_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        self.center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def save_to_excel(
        self,
        df: pd.DataFrame,
        output_path: str,
        sheet_name: Optional[str] = None
    ) -> str:
        """
        Сохраняет DataFrame в Excel-файл с жестким форматированием.
        """
        if df is None:
            raise ValueError("DataFrame is None")
        
        sheet = sheet_name or self.sheet_name
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        
        # Записываем данные
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                
                if r_idx == 1:
                    cell.font = Font(bold=True, size=11, color="FFFFFF")
                    cell.fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
        
        self._auto_adjust_columns(ws)
        ws.freeze_panes = "A2"
        wb.save(output_path)
        return output_path

    def _is_numeric_or_short(self, value) -> bool:
        """Определяет, нужно ли центрировать ячейку."""
        if value is None: return True
        val_str = str(value)
        if len(val_str) < 10: return True
        try:
            float(val_str.replace(',', '.').replace(' ', ''))
            return True
        except ValueError:
            return False

    def _auto_adjust_columns(self, ws):
        """Автоматический расчет ширины колонок на основе содержимого."""
        for col in ws.columns:
            max_length = 0
            column_letter = get_column_letter(col[0].column)
            
            for cell in col:
                try:
                    if cell.value:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except:
                    pass
            
            adjusted_width = min(max_length + 4, 50)
            ws.column_dimensions[column_letter].width = max(adjusted_width, 10)
