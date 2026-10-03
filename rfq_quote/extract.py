"""Extract the two supported pipe-separated text PDF table layouts."""

import re
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

import pdfplumber


@dataclass(frozen=True)
class SourceLine:
    """A source location and its original text."""

    page: int
    line: int
    text: str


@dataclass(frozen=True)
class Item:
    """One requested item; an empty part number is intentionally allowed."""

    line_no: str
    part_no: str
    description: str
    qty: Decimal
    unit: str
    source: SourceLine


@dataclass(frozen=True)
class Unparsed:
    """A source line that could not be interpreted as an item."""

    source: SourceLine
    reason: str


@dataclass(frozen=True)
class Extraction:
    """Parsed items and all unrecognised nonempty source lines."""

    items: list[Item]
    unparsed: list[Unparsed]


_LAYOUTS = (
    ("Line", "Part no", "Description", "Qty", "Unit"),
    ("Line", "Qty", "Unit", "Description", "Part no"),
)
_METADATA = ("RFQ:", "Requester:", "Delivery:", "Notes:", "Page:")


def _quantity(text: str) -> Decimal:
    """Parse explicit positive quantities, rejecting ambiguous decimal groups."""
    if not text:
        raise ValueError("quantity is empty")
    if re.fullmatch(r"\d+\.\d{3}", text):
        raise ValueError(f"quantity {text!r} is ambiguous (thousands or decimal?)")
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        raise ValueError(f"quantity {text!r} is not a number in the supported format "
                         "(use unsigned digits with an optional decimal point)")
    qty = Decimal(text)
    if qty <= 0:
        raise ValueError(f"quantity {text!r} must be positive")
    return qty


def extract_pdf(path: str | Path) -> Extraction:
    """Read supported text tables, retaining every unrecognised nonempty line.

    Known metadata prefixes and exact table headers are excluded. Quantities
    must be positive unsigned integers or dot decimals, excluding decimals with
    exactly three fractional digits (ambiguous thousands/decimal notation).
    Commas and wrapped rows are not supported.
    PDFs without any supported header raise ValueError (including image PDFs).
    """
    items, unparsed = [], []
    layout = None
    found_header = False
    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            for line_no, text in enumerate((page.extract_text() or "").splitlines(), 1):
                text = text.strip()
                if not text:
                    continue
                source = SourceLine(page_no, line_no, text)
                cells = tuple(cell.strip() for cell in text.split("|"))
                if cells in _LAYOUTS:
                    layout = cells
                    found_header = True
                    continue
                if text.startswith(_METADATA):
                    continue
                if layout is None or len(cells) != len(layout):
                    unparsed.append(Unparsed(source, "Unsupported row structure"))
                    continue
                row = dict(zip(layout, cells))
                try:
                    qty = _quantity(row["Qty"])
                    if not row["Line"].isdigit() or not row["Description"]:
                        raise ValueError("Missing item number or description")
                except ValueError as error:
                    unparsed.append(Unparsed(source, f"Invalid item fields: {error}"))
                    continue
                items.append(Item(row["Line"], row["Part no"], row["Description"],
                                  qty, row["Unit"], source))
    if not found_header:
        raise ValueError("No supported text table header found; OCR is not supported")
    return Extraction(items, unparsed)
