"""
Быстрая проверка работы OCR pipeline с улучшенными настройками.
"""

import pypdfium2 as pdfium
import pytesseract
from pathlib import Path

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Проверяем pypdfium2 + OCR
test_pdf = Path("input/file-2.pdf")
if test_pdf.exists():
    try:
        pdf_doc = pdfium.PdfDocument(str(test_pdf))
        page = pdf_doc[0]
        bitmap = page.render(scale=300/72)
        image = bitmap.to_pil()
        print(f"Rendered page: {image.size}")
        
        # Пробуем OCR с конфигом для таблиц
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(image, lang='rus+eng', config=custom_config)
        
        # Сохраняем текст в файл
        with open("ocr_test_text_v2.txt", "w", encoding="utf-8") as f:
            f.write(text)
        print("OCR text saved to ocr_test_text_v2.txt")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
else:
    print("Test PDF not found")
