"""Offline command-line entry point for the quote draft demo."""

import argparse
import sys
from datetime import date, datetime, timezone
from pathlib import Path

from .export import export_quote, summarize
from .extract import extract_pdf
from .match import load_catalogue, load_prices, match_items


def main(argv: list[str] | None = None) -> int:
    """Run extraction, matching and export; return zero only after a saved draft."""
    parser = argparse.ArgumentParser(description="Text RFQ PDF to a human-reviewed Excel draft")
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--catalogue", type=Path, required=True)
    parser.add_argument("--prices", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, default=Path("out"))
    parser.add_argument("--today", type=date.fromisoformat, default=None,
                        help="Price evaluation date YYYY-MM-DD (default: local today)")
    args = parser.parse_args(argv)
    started_at = datetime.now(timezone.utc)
    today = args.today if args.today is not None else date.today()
    try:
        catalogue = load_catalogue(args.catalogue)
        prices = load_prices(args.prices)
        extraction = extract_pdf(args.pdf)
        matches = match_items(extraction.items, catalogue, prices, today)
        output = export_quote(matches, extraction.unparsed, args.pdf, args.out_dir, today,
                              started_at, (args.catalogue, args.prices))
    except Exception as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    counts = summarize(matches, extraction.unparsed)
    print(" / ".join(f"{key}: {counts[key]}" for key in
                     ("priced", "needs review", "not in catalogue", "unparsed")))
    print(f"Output: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
