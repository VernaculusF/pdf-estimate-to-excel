"""
Скрипт для генерации тестового PDF-файла сметы.
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
import os


def create_test_pdf(output_path: str = "test_smeta.pdf"):
    """Создает тестовый PDF-файл с типичной таблицей сметы."""
    
    # Пробуем зарегистрировать шрифт с поддержкой кириллицы
    # Если не получится, используем стандартный
    try:
        # Проверим наличие системных шрифтов Windows
        font_paths = [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/Arial.ttf",
        ]
        
        font_registered = False
        for font_path in font_paths:
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('Arial', font_path))
                font_registered = True
                break
        
        if not font_registered:
            # Попробуем использовать встроенный Helvetica
            pass
            
    except Exception:
        pass
    
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Заголовок документа
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=1,  # По центру
        spaceAfter=20,
        textColor=colors.HexColor('#1a5490')
    )
    
    elements.append(Paragraph("СМЕТА № 15/2024", title_style))
    elements.append(Paragraph("Объект: Ремонт офисного помещения", styles["Normal"]))
    elements.append(Spacer(1, 20))
    
    # Данные таблицы
    data = [
        ["№ п/п", "Наименование работ", "Ед. изм.", "Количество", "Цена за ед., руб.", "Сумма, руб."],
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
        ["", "", "", "", "ИТОГО:", "316 490,00"],
    ]
    
    # Создаем таблицу
    table = Table(data, colWidths=[1.2*cm, 6.5*cm, 1.8*cm, 2.2*cm, 2.8*cm, 2.5*cm])
    
    # Стили таблицы
    style = TableStyle([
        # Заголовок
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#366092')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        
        # Данные
        ('BACKGROUND', (0, 1), (-1, -2), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -2), colors.black),
        ('ALIGN', (0, 1), (0, -2), 'CENTER'),  # № п/п по центру
        ('ALIGN', (2, 1), (-1, -2), 'CENTER'),  # Цифровые колонки по центру
        ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # Наименование по левому краю
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -2), 9),
        ('BOTTOMPADDING', (0, 1), (-1, -2), 8),
        
        # Итоговая строка
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#d9e2f3')),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
        ('ALIGN', (0, -1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 10),
        
        # Границы
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#366092')),
        ('LINEABOVE', (0, -1), (-1, -1), 2, colors.grey),
        
        # Отступы
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ])
    
    table.setStyle(style)
    elements.append(table)
    
    # Добавляем примечание
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("* Цены указаны с учетом НДС 20%", styles["Italic"]))
    elements.append(Paragraph("** Срок выполнения работ: 15 рабочих дней", styles["Italic"]))
    
    # Собираем PDF
    doc.build(elements)
    print(f"[OK] Тестовый PDF создан: {output_path}")
    
    return output_path


def create_second_test_pdf(output_path: str = "test_smeta2.pdf"):
    """Создает второй тестовый PDF с другой структурой."""
    
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph("Локальная смета № 42", styles["Heading1"]))
    elements.append(Paragraph("На отделочные работы", styles["Heading2"]))
    elements.append(Spacer(1, 15))
    
    data = [
        ["№", "Наименование", "Ед.", "Кол-во", "Цена", "Сумма"],
        ["1", "Подготовка основания", "м²", "200", "120", "24000"],
        ["2", "Грунтовка", "м²", "200", "85", "17000"],
        ["3", "Штукатурка", "м²", "150", "650", "97500"],
        ["4", "Шпаклевка", "м²", "150", "320", "48000"],
        ["5", "Поклейка обоев", "м²", "180", "280", "50400"],
        ["", "", "", "", "ИТОГО", "236900"],
    ]
    
    table = Table(data, colWidths=[1.5*cm, 7*cm, 2*cm, 2.5*cm, 3*cm, 3*cm])
    
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472c4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
    ])
    
    table.setStyle(style)
    elements.append(table)
    
    doc.build(elements)
    print(f"[OK] Второй тестовый PDF создан: {output_path}")
    
    return output_path


if __name__ == "__main__":
    create_test_pdf("test_smeta.pdf")
    create_second_test_pdf("test_smeta2.pdf")
    print("\n[OK] Все тестовые файлы созданы успешно!")
