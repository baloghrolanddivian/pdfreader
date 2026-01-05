from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader


def extract_page_text(pages: Iterable) -> str:
    """Return concatenated text from PDF pages."""
    return "\n".join((page.extract_text() or "").strip() for page in pages)


def read_pdf(path: Path) -> str:
    reader = PdfReader(path)
    return extract_page_text(reader.pages)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read and print PDF text content.")
    parser.add_argument("pdf_path", type=Path, help="Path to the PDF file to read")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    pdf_path: Path = args.pdf_path
    if not pdf_path.is_file():
        parser.error(f"PDF not found: {pdf_path}")

    text = read_pdf(pdf_path)
    print(text)


if __name__ == "__main__":
    main()
