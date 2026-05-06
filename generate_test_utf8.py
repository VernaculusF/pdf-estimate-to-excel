"""
Скрипт для генерации тестового PDF-файла сметы с корректной кириллицей.
Использует fpdf2 с поддержкой UTF-8.
"""

from fpdf import FPDF
import os


class SmetaPDF(FPDF):
    def header(self):
        self.set_font('DejaVu', 'B', 14)
        self.cell(0, 10, 'СМЕТА № 15/2024', ln=True, align='C')
        self.set_font('DejaVu', '', 10)
        self.cell(0, 8, 'Объект: Ремонт офисного помещения', ln=True, align='C')
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('DejaVu', 'I', 8)
        self.cell(0, 10, f'Страница {self.page_no()}', align='C')


FONT_DIR = "fonts/dejavu-fonts-ttf-2.37/ttf/"


def create_test_pdf(output_path: str = "test_smeta_utf8.pdf"):
    """Создает тестовый PDF-файл с типичной таблицей сметы."""
    
    pdf = SmetaPDF()
    
    # Добавляем шрифт DejaVu с поддержкой кириллицы
    pdf.add_font('DejaVu', '', FONT_DIR + 'DejaVuSans.ttf')
    pdf.add_font('DejaVu', 'B', FONT_DIR + 'DejaVuSans-Bold.ttf')
    pdf.add_font('DejaVu', 'I', FONT_DIR + 'DejaVuSans-Oblique.ttf')
    
    pdf.add_page()
    
    # Данные таблицы
    headers = ["№ п/п", "Наименование работ", "Ед. изм.", "Кол-во", "Цена за ед., руб.", "Сумма, руб."]
    data = [
        ["1", "Демонтаж старых перегородок", "м²", "45,00", "850,00", "38 250,00"],
        ["2", "Устройство новых перегородок из ГКЛ", "м²", "32,50", "1 250,00", "40 625,00"],
        ["3", "Шпаклевка стен под обои", "м²", "120,00", "450,00", "54 000,00"],
        ["4", "Покраска потолка водоэмульсионной краской", "м²", "48,00", "380,00", "18 240,00"],
        ["5", "Укладка ламината", "м²", "48,00", "1 100,00", "52 800,00"],
        ["6", "Установка дверей межкомнатных", "шт.", "4", "8 500,00", "34 000,00"],
        ["7", "Монтаж электроточек (розетки/выключатели)", "шт.", "18", "650,00", "11 700,00"],
        ["8", "Прокладка кабеля ВВГнг 3х2.5", "м.п.", "85", "95,00", "8 075,00"],
        ["9", "Установка светильников светодиодных", "шт.", "12", "2 400,00", "28 800,00"],
        ["10", "Монтаж кондиционера", "шт.", "2", "15 000,00", "30 000,00"],
    ]
    
    # Заголовки таблицы
    pdf.set_fill_color(54, 96, 146)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('DejaVu', 'B', 9)
    pdf.set_line_width(0.3)
    
    col_widths = [15, 75, 20, 20, 30, 30]
    row_height = 8
    
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], row_height, header, border=1, align='C', fill=True)
    pdf.ln()
    
    # Данные
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('DejaVu', '', 9)
    pdf.set_fill_color(255, 255, 255)
    
    for row in data:
        for i, cell in enumerate(row):
            align = 'C' if i != 1 else 'L'
            pdf.cell(col_widths[i], row_height, cell, border=1, align=align)
        pdf.ln()
    
    # Итоговая строка
    pdf.set_font('DejaVu', 'B', 9)
    pdf.set_fill_color(217, 226, 243)
    pdf.cell(sum(col_widths[:4]), row_height, '', border=1, fill=True)
    pdf.cell(col_widths[4], row_height, 'ИТОГО:', border=1, align='C', fill=True)
    pdf.cell(col_widths[5], row_height, '316 490,00', border=1, align='C', fill=True)
    pdf.ln()
    
    # Примечания
    pdf.ln(5)
    pdf.set_font('DejaVu', 'I', 8)
    pdf.cell(0, 5, '* Цены указаны с учетом НДС 20%', ln=True)
    pdf.cell(0, 5, '** Срок выполнения работ: 15 рабочих дней', ln=True)
    
    pdf.output(output_path)
    print(f"[OK] Тестовый PDF с UTF-8 создан: {output_path}")
    
    return output_path


def create_second_test_pdf(output_path: str = "test_smeta2_utf8.pdf"):
    """Создает второй тестовый PDF с другой структурой."""
    
    pdf = FPDF()
    pdf.add_font('DejaVu', '', FONT_DIR + 'DejaVuSans.ttf')
    pdf.add_font('DejaVu', 'B', FONT_DIR + 'DejaVuSans-Bold.ttf')
    
    pdf.add_page()
    pdf.set_font('DejaVu', 'B', 14)
    pdf.cell(0, 10, 'Локальная смета № 42', ln=True, align='C')
    pdf.set_font('DejaVu', '', 10)
    pdf.cell(0, 8, 'На отделочные работы', ln=True, align='C')
    pdf.ln(5)
    
    data = [
        ["№", "Наименование", "Ед.", "Кол-во", "Цена", "Сумма"],
        ["1", "Подготовка основания", "м²", "200", "120", "24000"],
        ["2", "Грунтовка", "м²", "200", "85", "17000"],
        ["3", "Штукатурка", "м²", "150", "650", "97500"],
        ["4", "Шпаклевка", "м²", "150", "320", "48000"],
        ["5", "Поклейка обоев", "м²", "180", "280", "50400"],
        ["", "", "", "", "ИТОГО", "236900"],
    ]
    
    col_widths = [15, 80, 20, 25, 30, 30]
    row_height = 8
    
    for row_idx, row in enumerate(data):
        if row_idx == 0:
            pdf.set_fill_color(68, 114, 196)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font('DejaVu', 'B', 9)
        elif row_idx == len(data) - 1:
            pdf.set_fill_color(211, 211, 211)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font('DejaVu', 'B', 9)
        else:
            pdf.set_fill_color(255, 255, 255)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font('DejaVu', '', 9)
        
        for i, cell in enumerate(row):
            align = 'C'
            fill = row_idx == 0 or row_idx == len(data) - 1
            pdf.cell(col_widths[i], row_height, cell, border=1, align=align, fill=fill)
        pdf.ln()
    
    pdf.output(output_path)
    print(f"[OK] Второй тестовый PDF с UTF-8 создан: {output_path}")
    
    return output_path


if __name__ == "__main__":
    create_test_pdf("test_smeta_utf8.pdf")
    create_second_test_pdf("test_smeta2_utf8.pdf")
    print("\n[OK] Все тестовые файлы с UTF-8 созданы успешно!")
