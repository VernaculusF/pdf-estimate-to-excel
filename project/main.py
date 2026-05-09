#!/usr/bin/env python3
"""
Convert estimate PDFs into Excel workbooks.
"""

import argparse
import logging
import os
import sys
import tempfile
from pathlib import Path
from typing import List, Optional

import pandas as pd
import pypdfium2 as pdfium
from tqdm import tqdm

from config import DEFAULT_INPUT_DIR, DEFAULT_OUTPUT_DIR, SMETA_COLUMNS
from converter import SmetaConverter
from document_export import DocumentExporter
from extractor import SmetaExtractor
from ocr_extractor import OCRExtractor
from quality_report import (
    append_source_text_sheet,
    build_reports_for_pairs,
    extract_pdf_text,
    summarize_quality,
)


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


class SmetaProcessor:
    """Process estimate PDFs and save them as Excel workbooks."""

    def __init__(self) -> None:
        self.extractor = SmetaExtractor()
        self.ocr_extractor = OCRExtractor()
        self.converter = SmetaConverter()
        self.document_exporter = DocumentExporter(self.ocr_extractor)

    def process_single_file(
        self,
        pdf_path: str,
        output_path: Optional[str] = None,
        merge_tables: bool = True,
        force_ocr: bool = False,
    ) -> str:
        pdf_path_obj = Path(pdf_path)
        if not pdf_path_obj.exists():
            raise FileNotFoundError(f"File not found: {pdf_path_obj}")

        if output_path is None:
            output_dir = Path(DEFAULT_OUTPUT_DIR)
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path_obj = output_dir / f"{pdf_path_obj.stem}.xlsx"
        else:
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)

        logger.info("Processing file: %s", pdf_path_obj.name)

        oriented_tmp = self.ocr_extractor.normalize_pdf_orientation(str(pdf_path_obj))
        working_pdf = oriented_tmp if oriented_tmp else str(pdf_path_obj)

        try:
            return self._do_process(working_pdf, str(pdf_path_obj), str(output_path_obj), merge_tables, force_ocr)
        finally:
            if oriented_tmp and os.path.exists(oriented_tmp):
                os.remove(oriented_tmp)

    def _do_process(
        self,
        working_pdf: str,
        original_pdf: str,
        output_path: str,
        merge_tables: bool,
        force_ocr: bool,
    ) -> str:
        output_path_obj = Path(output_path)

        is_scanned = force_ocr or self.ocr_extractor.is_scanned_pdf(working_pdf)
        raw_tables = []
        if is_scanned:
            logger.info("Detected scanned PDF, using OCR extraction")
            tables = self.ocr_extractor.extract_tables_from_pdf(working_pdf)
        else:
            tables = self.extractor.extract_tables_from_pdf(working_pdf)
            raw_tables = self.extractor.extract_raw_tables_from_pdf(
                working_pdf, ocr_extractor=self.ocr_extractor
            )

        if not tables:
            logger.warning("No tables detected in %s", Path(original_pdf).name)
            empty_df = pd.DataFrame({"Message": ["No tables were detected in the PDF file."]})
            self.converter.save_to_excel(empty_df, str(output_path_obj))
            return str(output_path_obj)

        logger.info("Detected %s table(s)", len(tables))

        if merge_tables and len(tables) > 1:
            df = self.extractor.merge_tables(tables)
            logger.info("Merged tables into %s row(s)", len(df))
        else:
            df = tables[0]
            logger.info("Using a single table with %s row(s)", len(df))

        df = self._align_columns(df.reset_index(drop=True))
        df = df.dropna(how="all").reset_index(drop=True)

        if raw_tables:
            raw_df = self.extractor.build_raw_estimate_dataframe(raw_tables)
            raw_df = raw_df.dropna(how="all").reset_index(drop=True)
            result_path = self.converter.save_raw_estimate_to_excel(raw_df, str(output_path_obj))
        else:
            result_path = self.converter.save_to_excel(df, str(output_path_obj))

        self.document_exporter.append_header_sheet(result_path, working_pdf)

        pdf_text, source_method, _ = extract_pdf_text(working_pdf)
        if pdf_text:
            append_source_text_sheet(result_path, pdf_text, source_method)
            logger.info("Added Source Text sheet (method=%s)", source_method)

        logger.info("Saved workbook: %s", result_path)
        return result_path

    def process_directory(
        self,
        input_dir: str,
        output_dir: Optional[str] = None,
        merge_tables: bool = True,
        force_ocr: bool = False,
    ) -> List[tuple[str, str]]:
        input_path = Path(input_dir)
        if not input_path.exists():
            raise FileNotFoundError(f"Directory not found: {input_path}")

        output_path = Path(output_dir or DEFAULT_OUTPUT_DIR)
        output_path.mkdir(parents=True, exist_ok=True)

        pdf_files = list(input_path.glob("*.pdf")) + list(input_path.glob("*.PDF"))
        pdf_files = list({f: None for f in pdf_files}.keys())
        if not pdf_files:
            logger.warning("No PDF files found in %s", input_path)
            return []

        logger.info("Found %s PDF file(s)", len(pdf_files))
        results: List[tuple[str, str]] = []

        for pdf_file in tqdm(pdf_files, desc="Processing PDFs", unit="file"):
            try:
                result = self.process_single_file(
                    str(pdf_file),
                    str(output_path / f"{pdf_file.stem}.xlsx"),
                    merge_tables=merge_tables,
                    force_ocr=force_ocr,
                )
                results.append((str(pdf_file), result))
            except Exception as exc:
                logger.error("Failed to process %s: %s", pdf_file.name, exc)

        logger.info("Completed %s of %s file(s)", len(results), len(pdf_files))
        logger.info("Output directory: %s", output_path)
        return results

    def _align_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        columns = pd.Series(df.columns)
        for duplicate in columns[columns.duplicated()].unique():
            duplicate_indices = columns[columns == duplicate].index.tolist()
            columns.iloc[duplicate_indices] = [
                f"{duplicate}_{index}" if index != 0 else duplicate
                for index in range(len(duplicate_indices))
            ]
        df.columns = columns

        aligned_df = pd.DataFrame(columns=SMETA_COLUMNS)
        for column in SMETA_COLUMNS:
            if column in df.columns:
                aligned_df[column] = df[column]

        if aligned_df.dropna(how="all").empty and not df.dropna(how="all").empty:
            current_cols = list(df.columns)
            has_generic = any(str(c).startswith("Column_") for c in current_cols)
            # If columns already look like real headers, keep them
            if not has_generic:
                return df
            # Try to detect headers from first data rows when columns are generic
            from extractor import SmetaExtractor
            rows = df.head(10).astype(str).values.tolist()
            num_cols = len(rows[0]) if rows else 0
            header_rows, data_rows = SmetaExtractor._detect_header_rows(rows, num_cols)
            if header_rows:
                new_cols = SmetaExtractor._merge_header_rows(header_rows, num_cols)
                df = pd.DataFrame(data_rows + rows[len(header_rows):], columns=new_cols)
                return df
            df.columns = [f"Column_{index + 1}" for index in range(len(df.columns))]
            return df

        return aligned_df


def resolve_single_output_path(pdf_path: str, output_arg: Optional[str]) -> str:
    output = Path(output_arg or DEFAULT_OUTPUT_DIR)
    if output.suffix.lower() == ".xlsx":
        return str(output)
    return str(output / f"{Path(pdf_path).stem}.xlsx")


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert estimate PDFs into Excel workbooks.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python main.py --input ./input --output ./output\n"
            "  python main.py --file ./input/example.pdf --output ./output\n"
            "  python main.py --input ./input --use-ocr\n"
        ),
    )
    parser.add_argument(
        "--input",
        "-i",
        type=str,
        default=DEFAULT_INPUT_DIR,
        help="Input folder with PDF files",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        default=DEFAULT_OUTPUT_DIR,
        help="Output folder for generated Excel files",
    )
    parser.add_argument(
        "--file",
        "-f",
        type=str,
        help="Process a single PDF file",
    )
    parser.add_argument(
        "--no-merge",
        action="store_true",
        help="Do not merge multiple tables from the same PDF into one sheet",
    )
    parser.add_argument(
        "--use-ocr",
        action="store_true",
        help="Force OCR for every PDF file",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose logging",
    )
    parser.add_argument(
        "--no-quality-report",
        action="store_true",
        help="Skip the quality report generation step",
    )
    return parser


def main() -> None:
    parser = create_parser()
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    processor = SmetaProcessor()

    try:
        if args.file:
            result = processor.process_single_file(
                args.file,
                output_path=resolve_single_output_path(args.file, args.output),
                merge_tables=not args.no_merge,
                force_ocr=args.use_ocr,
            )
            print(f"\n[OK] Done: {result}")
            if not args.no_quality_report:
                report = build_reports_for_pairs([(args.file, result)], str(Path(result).parent))
                print(f"[OK] Quality report: {report['paths']['xlsx']}")
                print(summarize_quality(report["records"]))
            return

        results = processor.process_directory(
            args.input,
            args.output,
            merge_tables=not args.no_merge,
            force_ocr=args.use_ocr,
        )
        if not results:
            print(f"\n[!] No PDF files found in '{args.input}'")
            return

        print(f"\n[OK] Processed files: {len(results)}")
        print(f"[OK] Output directory: {args.output}")
        if not args.no_quality_report:
            report = build_reports_for_pairs(results, args.output)
            print(f"[OK] XLSX quality report: {report['paths']['xlsx']}")
            print(f"[OK] JSON quality report: {report['paths']['json']}")
            print(summarize_quality(report["records"]))
    except FileNotFoundError as exc:
        logger.error("%s", exc)
        sys.exit(1)
    except Exception as exc:
        logger.error("Unexpected error: %s", exc)
        if args.verbose:
            import traceback

            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
