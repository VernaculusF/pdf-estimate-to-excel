"""
Диагностический скрипт для анализа проблемных PDF.
"""

import pdfplumber
from pathlib import Path
import json

input_dir = Path("input")
problem_files = ["file-1.PDF", "file-2.pdf", "file.pdf", "NMCK.pdf", "OSR.pdf", "otdelochny_039_e_raboty_039.pdf"]

results = {}

for fname in problem_files:
    fpath = input_dir / fname
    if not fpath.exists():
        continue
    
    print(f"\n{'='*60}")
    print(f"Analyzing: {fname}")
    print(f"{'='*60}")
    
    with pdfplumber.open(fpath) as pdf:
        print(f"Pages: {len(pdf.pages)}")
        
        for i, page in enumerate(pdf.pages[:2]):  # Анализируем первые 2 страницы
            print(f"\n--- Page {i+1} ---")
            
            # Проверяем, есть ли изображения
            images = page.images
            print(f"Images on page: {len(images)}")
            
            # Извлекаем текст
            text = page.extract_text()
            if text:
                lines = text.strip().split('\n')
                print(f"Text lines: {len(lines)}")
                print(f"First 10 lines:")
                for line in lines[:10]:
                    print(f"  {line[:120]}")
            else:
                print("No text extracted!")
                # Проверяем, есть ли изображение, которое может быть сканом
                if len(images) > 0:
                    print("WARNING: Page contains images but no text - likely a SCAN!")
            
            # Пробуем извлечь слова
            words = page.extract_words()
            print(f"Words extracted: {len(words)}")
            if words:
                print(f"First 10 words: {[w['text'] for w in words[:10]]}")

print("\n[OK] Diagnostic completed")
