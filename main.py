#!/usr/bin/env python3
"""
PDF Смета → Excel Конвертер

Автоматически извлекает табличные данные из PDF-файлов смет
и сохраняет их в формат Excel (.xlsx).

Использование:
    python main.py --input папка_с_pdf --output папка_для_xlsx
    python main.py --file смета.pdf --output результат.xlsx
"""

import os
import sys
import argparse
import logging
from pathlib import Path
from typing import List, Optional

import pandas as pd
from tqdm import tqdm

from config import DEFAULT_INPUT_DIR, DEFAULT_OUTPUT_DIR
from extractor import SmetaExtractor
from ocr_extractor import OCRExtractor
from converter import SmetaConverter
from quality_report import build_reports_for_pairs, summarize_quality


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class SmetaProcessor:
    """Главный класс для обработки смет."""

    def __init__(self):
        self.extractor = SmetaExtractor()
        self.ocr_extractor = OCRExtractor()
        self.converter = SmetaConverter()

    def process_single_file(
        self,
        pdf_path: str,
        output_path: Optional[str] = None,
        merge_tables: bool = True,
        force_ocr: bool = False
    ) -> str:
        """
        Обрабатывает один PDF-файл.
        
        Args:
            pdf_path: Путь к PDF-файлу
            output_path: Путь для сохранения (опционально)
            merge_tables: Объединять ли таблицы в одну
            force_ocr: Принудительно использовать OCR
            
        Returns:
            Путь к созданному Excel-файлу
        """
        pdf_path = Path(pdf_path)
        
        if not pdf_path.exists():
            raise FileNotFoundError(f"Файл не найден: {pdf_path}")
        
        if output_path is None:
            output_dir = Path(DEFAULT_OUTPUT_DIR)
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / f"{pdf_path.stem}.xlsx"
        else:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Обработка файла: {pdf_path.name}")
        
        # Определяем тип PDF (текстовый или скан)
        if force_ocr:
            is_scanned = True
        else:
            is_scanned = self.ocr_extractor.is_scanned_pdf(str(pdf_path))
        
        if is_scanned:
            logger.info(f"Обнаружен сканированный PDF, используется OCR")
            tables = self.ocr_extractor.extract_tables_from_pdf(str(pdf_path))
        else:
            # Извлекаем таблицы стандартным способом
            tables = self.extractor.extract_tables_from_pdf(str(pdf_path))
        
        if not tables:
            logger.warning(f"В файле {pdf_path.name} не найдено таблиц")
            # Создаем пустой файл с сообщением
            df = pd.DataFrame({"Сообщение": ["В PDF-файле не обнаружены таблицы"]})
            self.converter.save_to_excel(df, str(output_path))
            return str(output_path)
        
        logger.info(f"Найдено таблиц: {len(tables)}")
        
        # Берем самую большую таблицу вместо слияния (избегаем конфликтов схем)
        if len(tables) > 1:
            df = max(tables, key=lambda t: len(t))
            logger.info(f"Выбрана самая большая таблица, строк: {len(df)}")
        else:
            df = tables[0]
            logger.info(f"Выбрана основная таблица, строк: {len(df)}")
        
        # Приведение к строгой схеме колонок БЕЗОПАСНЫМ способом
        from config import SMETA_COLUMNS
        df = df.reset_index(drop=True)
        
        # Убираем дубли в именах колонок
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique():
            dup_indices = cols[cols == dup].index.tolist()
            cols.iloc[dup_indices] = [f"{dup}_{i}" if i != 0 else dup for i in range(len(dup_indices))]
        df.columns = cols
        
        # Создаем новый DataFrame по схеме и копируем данные
        aligned_df = pd.DataFrame(columns=SMETA_COLUMNS)
        for col in SMETA_COLUMNS:
            if col in df.columns:
                aligned_df[col] = df[col]
        
        # Если после приведения к схеме ничего не осталось, используем исходный df
        if aligned_df.dropna(how='all').empty and not df.dropna(how='all').empty:
            df.columns = [f"Column_{i+1}" for i in range(len(df.columns))]
        else:
            df = aligned_df
        
        # Сохраняем в Excel
        result_path = self.converter.save_to_excel(df, str(output_path))
        logger.info(f"Сохранено: {result_path}")
        
        return result_path

    def process_directory(
        self,
        input_dir: str,
        output_dir: Optional[str] = None,
        merge_tables: bool = True,
        force_ocr: bool = False
    ) -> List[tuple[str, str]]:
        """
        Обрабатывает все PDF-файлы в директории.
        
        Args:
            input_dir: Папка с PDF-файлами
            output_dir: Папка для сохранения результатов
            merge_tables: Объединять ли таблицы в одну
            force_ocr: Принудительно использовать OCR
            
        Returns:
            Список пар (PDF-файл, созданный Excel-файл)
        """
        input_dir = Path(input_dir)
        
        if not input_dir.exists():
            raise FileNotFoundError(f"Директория не найдена: {input_dir}")
        
        if output_dir is None:
            output_dir = Path(DEFAULT_OUTPUT_DIR)
        else:
            output_dir = Path(output_dir)
        
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Находим все PDF-файлы
        pdf_files = list(input_dir.glob("*.pdf"))
        
        if not pdf_files:
            logger.warning(f"В директории {input_dir} не найдено PDF-файлов")
            return []
        
        logger.info(f"Найдено PDF-файлов: {len(pdf_files)}")
        
        results = []
        
        for pdf_file in tqdm(pdf_files, desc="Обработка PDF", unit="файл"):
            try:
                output_path = output_dir / f"{pdf_file.stem}.xlsx"
                result = self.process_single_file(
                    str(pdf_file),
                    str(output_path),
                    merge_tables=merge_tables,
                    force_ocr=force_ocr
                )
                results.append((str(pdf_file), result))
            except Exception as e:
                logger.error(f"Ошибка при обработке {pdf_file.name}: {str(e)}")
                continue
        
        logger.info(f"Обработано успешно: {len(results)} из {len(pdf_files)}")
        logger.info(f"Результаты сохранены в: {output_dir}")
        
        return results


def resolve_single_output_path(pdf_path: str, output_arg: Optional[str]) -> str:
    """Resolve --output for single-file mode as either a folder or .xlsx path."""
    output = Path(output_arg or DEFAULT_OUTPUT_DIR)
    if output.suffix.lower() == ".xlsx":
        return str(output)
    return str(output / f"{Path(pdf_path).stem}.xlsx")


def create_parser() -> argparse.ArgumentParser:
    """Создает парсер аргументов командной строки."""
    parser = argparse.ArgumentParser(
        description="Конвертер смет из PDF в Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python main.py --input ./сметы --output ./результаты
  python main.py --file смета.pdf
  python main.py --input ./сметы --no-merge
        """
    )
    
    parser.add_argument(
        "--input", "-i",
        type=str,
        default=DEFAULT_INPUT_DIR,
        help="Папка с PDF-файлами (по умолчанию: ./input)"
    )
    
    parser.add_argument(
        "--output", "-o",
        type=str,
        default=DEFAULT_OUTPUT_DIR,
        help="Папка для сохранения Excel (по умолчанию: ./output)"
    )
    
    parser.add_argument(
        "--file", "-f",
        type=str,
        help="Обработать один конкретный PDF-файл"
    )
    
    parser.add_argument(
        "--no-merge",
        action="store_true",
        help="Не объединять таблицы из одного PDF в одну"
    )
    
    parser.add_argument(
        "--use-ocr",
        action="store_true",
        help="Принудительно использовать OCR для всех PDF"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Подробный вывод"
    )

    parser.add_argument(
        "--no-quality-report",
        action="store_true",
        help="Не создавать отчет сохранности символов PDF -> Excel"
    )
    
    return parser


def main():
    """Главная функция."""
    parser = create_parser()
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    processor = SmetaProcessor()
    
    try:
        if args.file:
            # Обработка одного файла
            output_path = resolve_single_output_path(args.file, args.output)
            result = processor.process_single_file(
                args.file,
                output_path=output_path,
                merge_tables=not args.no_merge,
                force_ocr=args.use_ocr
            )
            print(f"\n[OK] Готово! Результат: {result}")
            if not args.no_quality_report:
                report = build_reports_for_pairs([(args.file, result)], args.output)
                print(f"[OK] Отчет качества: {report['paths']['xlsx']}")
                print(summarize_quality(report["records"]))
        else:
            # Обработка директории
            results = processor.process_directory(
                args.input,
                args.output,
                merge_tables=not args.no_merge,
                force_ocr=args.use_ocr
            )
            
            if results:
                print(f"\n[OK] Обработано файлов: {len(results)}")
                print(f"[OK] Результаты сохранены в: {args.output}")
                if not args.no_quality_report:
                    report = build_reports_for_pairs(results, args.output)
                    print(f"[OK] Отчет качества XLSX: {report['paths']['xlsx']}")
                    print(f"[OK] Отчет качества JSON: {report['paths']['json']}")
                    print(summarize_quality(report["records"]))
            else:
                print("\n[!] Не найдено PDF-файлов для обработки")
                print(f"  Убедитесь, что в папке '{args.input}' есть файлы .pdf")
                
    except FileNotFoundError as e:
        logger.error(f"Ошибка: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Неожиданная ошибка: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
