from __future__ import annotations

import argparse
import csv
import io
import json
import sys
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

FIELDNAMES = [
    "termek",
    "szin",
    "meret",
    "m2",
    "db",
    "ossz_m2",
    "egyseg_ar",
    "netto_ar",
    "kod",
]
DEFAULT_TRANSLATIONS = Path(__file__).resolve().parent.parent / "translations.json"


def extract_page_text(pages: Iterable) -> str:
    """Return concatenated text from PDF pages."""
    return "\n".join((page.extract_text() or "").strip() for page in pages)


def read_pdf(path: Path) -> str:
    reader = PdfReader(path)
    return extract_page_text(reader.pages)


def normalize_numeric(value: str) -> str:
    """Strip whitespace and normalise decimal separators."""
    cleaned = value.replace("\u00a0", " ").replace(" ", "").replace(",", ".")
    return cleaned


def is_numeric(value: str) -> bool:
    try:
        float(normalize_numeric(value))
        return True
    except ValueError:
        return False


def format_currency(value: str) -> str:
    """Return currency-style numeric text using comma decimals for Excel."""
    return normalize_numeric(value).replace(".", ",")


TEXT_FIXES = {
    "Magasfényû antracit": "Magasfényű antracit",
    "Magasfényû krém": "Magasfényű krém",
    "Magasfényû latte": "Magasfényű latte",
    "Magasfényû fehér": "Magasfényű fehér",
    "Matt ANTRAZITE": "Matt antracit",
}
MOJIBAKE_FIXES = {
    "├í": "á",
    "├ę": "é",
    "├ş": "í",
    "├│": "ó",
    "├Â": "ö",
    "┼Ĺ": "ő",
    "├║": "ú",
    "├╝": "ü",
    "┼▓": "ű",
    "├ü": "Á",
    "├ë": "É",
    "├Ź": "Í",
    "├ô": "Ó",
    "├ľ": "Ö",
    "┼É": "Ő",
    "├Ü": "Ú",
    "├ť": "Ü",
    "┼░": "Ű",
}


def fix_mojibake(value: str) -> str:
    fixed = value
    for bad, good in MOJIBAKE_FIXES.items():
        fixed = fixed.replace(bad, good)
    return fixed


def normalize_text(value: str) -> str:
    """Fix common accent/whitespace issues from PDF extraction."""
    fixed = fix_mojibake(value)
    fixed = fixed.replace("\u00a0", " ").replace("  ", " ").strip()
    fixed = fixed.replace("û", "ű").replace("Â", "")
    fixed = " ".join(fixed.split())
    return TEXT_FIXES.get(fixed, fixed)


def parse_rows(text: str) -> list[dict[str, str]]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    rows: list[dict[str, str]] = []
    in_table = False
    i = 0

    while i < len(lines):
        line = lines[i]

        if line == "Nettó ár":
            in_table = True
            i += 1
            continue

        if line.startswith("Összesítve"):
            in_table = False

        if in_table and line.isdigit() and i + 8 < len(lines):
            termek, szin, meret = lines[i + 1 : i + 4]
            m2 = lines[i + 4]
            db = lines[i + 5]
            ossz_m2 = lines[i + 6]
            egyseg_ar = lines[i + 7]
            netto_ar = lines[i + 8]

            numeric_values = (m2, db, ossz_m2, egyseg_ar, netto_ar)
            if "x" in meret and all(is_numeric(value) for value in numeric_values):
                rows.append(
                    {
                        "termek": termek,
                        "szin": szin,
                        "meret": meret,
                        "m2": normalize_numeric(m2),
                        "db": normalize_numeric(db),
                        "ossz_m2": normalize_numeric(ossz_m2),
                        "egyseg_ar": normalize_numeric(egyseg_ar),
                        "netto_ar": normalize_numeric(netto_ar),
                    }
                )
                i += 9

                if i < len(lines) and lines[i].startswith("Négyzetméterár"):
                    i += 1
                continue

        i += 1

    return rows


def load_translations(path: Path) -> dict:
    if not path.is_file():
        raise FileNotFoundError(f"Translation table not found: {path}")
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def apply_translations(
    rows: list[dict[str, str]], table: dict
) -> list[dict[str, str]]:
    products = table.get("products", {})
    colors = table.get("colors", {})
    standard_sizes = set(table.get("standard_sizes", []))

    translated: list[dict[str, str]] = []
    for row in rows:
        termek = normalize_text(row["termek"])
        szin = normalize_text(row["szin"])
        meret = row["meret"].replace(" ", "")

        product_meta = products.get(termek, {})
        color_meta = colors.get(szin, {})

        product_code = product_meta.get("code", termek)
        color_code = color_meta.get("code", szin)
        model_code = product_code.split("_")[-1]
        is_standard_size = meret in standard_sizes

        if meret == "718x250":
            # Speciális kód: NFAH_<színkód>_<méret>
            kod = "_".join(filter(None, ["NFAH", model_code, color_code, meret]))
        elif is_standard_size:
            size_part = meret
            code_parts = [product_code, color_code, size_part]
            kod = "_".join(filter(None, code_parts))
        else:
            # Egyedi méret: NFAY_<színkód>_<méret>
            kod = "_".join(filter(None, ["NFAY", model_code, color_code, meret]))

        translated.append(
            {
                **row,
                "termek": product_meta.get("name", termek),
                "szin": color_meta.get("name", szin),
                "meret": meret,
                "kod": kod,
            }
        )

    return translated


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read a PDF invoice and extract line items."
    )
    parser.add_argument("pdf_path", type=Path, help="Path to the PDF file to read")
    parser.add_argument(
        "--raw",
        action="store_true",
        help="Print the full extracted text instead of parsed rows.",
    )
    parser.add_argument(
        "--delimiter",
        default=";",
        help="CSV delimiter to use when printing parsed rows.",
    )
    parser.add_argument(
        "--translations",
        type=Path,
        default=DEFAULT_TRANSLATIONS,
        help=f"Path to JSON translation table (default: {DEFAULT_TRANSLATIONS.name}).",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    pdf_path: Path = args.pdf_path
    if not pdf_path.is_file():
        parser.error(f"PDF not found: {pdf_path}")

    text = read_pdf(pdf_path)
    if args.raw:
        print(text)
        return

    rows = parse_rows(text)
    translations = load_translations(args.translations)
    rows = apply_translations(rows, translations)
    for row in rows:
        row["egyseg_ar"] = format_currency(row["egyseg_ar"])
        row["netto_ar"] = format_currency(row["netto_ar"])

    try:
        sys.stdout.reconfigure(encoding="utf-8", newline="")
    except (AttributeError, ValueError):
        if hasattr(sys.stdout, "buffer"):
            sys.stdout = io.TextIOWrapper(
                sys.stdout.buffer, encoding="utf-8", newline=""
            )

    writer = csv.DictWriter(
        sys.stdout,
        fieldnames=FIELDNAMES,
        delimiter=args.delimiter,
        lineterminator="\n",
    )
    try:
        writer.writeheader()
        writer.writerows(rows)
    except BrokenPipeError:
        # Allow piping to commands that close the stream early.
        sys.exit(0)


if __name__ == "__main__":
    main()
