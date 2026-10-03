"""Generate deterministic, entirely synthetic catalogue, prices and text PDFs."""

import csv
from pathlib import Path

from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas


_FAMILIES = (
    ("TRM", "Terminal block", "ea"),
    ("CBL", "Control cable", "m"),
    ("FUS", "Fuse holder", "ea"),
    ("REL", "Relay socket", "ea"),
    ("GLD", "Cable gland", "ea"),
    ("DIN", "Mounting rail", "m"),
    ("SWI", "Selector switch", "ea"),
    ("LMP", "Indicator lamp", "ea"),
    ("JNC", "Junction enclosure", "ea"),
    ("LUG", "Crimp lug", "ea"),
)


def _write_csv(path, headers, rows):
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.writer(stream)
        writer.writerow(headers)
        writer.writerows(rows)


def _write_pdf(path, requester, layout, rows):
    pdf = canvas.Canvas(str(path), pagesize=landscape(A4), invariant=1)
    pdf.setTitle("Synthetic RFQ demo")
    pdf.setAuthor("Fictional demo data")
    pdf.setFont("Courier", 9)
    y = 555
    for line in (f"RFQ: {path.stem.upper()} - synthetic demonstration",
                 f"Requester: {requester} (fictional)",
                 "Delivery: Synthetic workshop, Example district",
                 "Notes: Draft request only; all names and parts are invented"):
        pdf.drawString(30, y, line)
        y -= 17
    widths = {"Line": 4, "Part no": 12, "Description": 34, "Qty": 5, "Unit": 6}
    pdf.drawString(30, y - 8, " | ".join(column.ljust(widths[column]) for column in layout))
    pdf.line(30, y - 13, 810, y - 13)
    y -= 32
    for row in rows:
        if isinstance(row, str):
            text = row
        else:
            text = " | ".join(str(row[column]).ljust(widths[column]) for column in layout)
        pdf.drawString(30, y, text)
        y -= 23
    pdf.save()


def make_samples(destination: str | Path | None = None) -> dict[str, Path]:
    """Create 50 parts, approved prices and two 12-item text RFQs locally.

    Each RFQ has 11 parsed rows and one deliberately malformed quantity. Price
    validity is fixed: ordinary prices last through 2099, one expired in 2020.
    Regenerating these sample assets intentionally replaces those sample files.
    """
    destination = Path(destination) if destination is not None else Path(__file__).resolve().parent
    destination.mkdir(parents=True, exist_ok=True)
    catalogue, prices = [], []
    for family_index, (prefix, description, unit) in enumerate(_FAMILIES):
        for variant in range(1, 6):
            part_no = f"{prefix}-{100 + variant}"
            catalogue.append((part_no, f"{description} type {variant}", unit))
            cents = 125 + family_index * 211 + variant * 73
            prices.append((part_no, f"{cents // 100}.{cents % 100:02d}", "USD",
                           "2020-01-01" if part_no == "REL-101" else "2099-12-31"))
    paths = {name: destination / name for name in
             ("catalogue.csv", "price_list.csv", "rfq_a.pdf", "rfq_b.pdf")}
    _write_csv(paths["catalogue.csv"], ("part_no", "description", "unit"), catalogue)
    _write_csv(paths["price_list.csv"], ("part_no", "unit_price", "currency", "valid_until"), prices)
    rows_a = [
        ("TRM-101", "Terminal block type 1", "24", "ea"),
        ("CBL-102", "Control cable type 2", "35.5", "m"),
        ("FUS-103", "Fuse holder type 3", "10", "ea"),
        ("GLD-104", "Cable gland type 4", "12", "ea"),
        ("LMP-105", "Indicator lamp type 5", "6", "ea"),
        (" trm - 102 ", "Terminal blok type 2", "18", "pcs"),
        ("", "Selector swich type 2", "4", "ea"),
        ("ZZZ-900", "Flux harmonizer module", "2", "ea"),
        ("LUG-101", "Crimp lug type 1", "50", "box?"),
        ("REL-101", "Relay socket type 1", "8", "ea"),
        ("DIN-103", "Mounting rail type 3", "7", "m"),
        ("JNC-102", "Junction enclosure type 2", "several", "ea"),
    ]
    rows_b = [
        ("SWI-104", "Selector switch type 4", "8", "ea"),
        ("JNC-105", "Junction enclosure type 5", "3", "ea"),
        ("FUS-101", "Fuse holder type 1", "20", "ea"),
        ("CBL-104", "Control cable type 4", "60", "m"),
        ("GLD-101", "Cable gland type 1", "30", "ea"),
        ("lmp 102", "Indictor lamp type 2", "10", "each"),
        ("", "Junction enclousure type 3", "2", "ea"),
        ("QPX-777", "Polar resonance coupler", "1", "ea"),
        ("DIN-102", "Mounting rail type 2", "4", "length?"),
        ("REL-101", "Relay socket type 1", "5", "ea"),
        ("LUG-105", "Crimp lug type 5", "40", "ea"),
        ("TRM-103", "Terminal block type 3", "TBD", "ea"),
    ]
    for name, requester, layout, values in (
        ("rfq_a.pdf", "Velmora Demo Assembly", ("Line", "Part no", "Description", "Qty", "Unit"), rows_a),
        ("rfq_b.pdf", "Orvessa Example Works", ("Line", "Qty", "Unit", "Description", "Part no"), rows_b),
    ):
        rows = [dict(zip(("Line", "Part no", "Description", "Qty", "Unit"), (i, *row)))
                for i, row in enumerate(values, 1)]
        _write_pdf(paths[name], requester, layout, rows)
    return paths


if __name__ == "__main__":
    for name, path in make_samples().items():
        print(f"Created: {path}")
