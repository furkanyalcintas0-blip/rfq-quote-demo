# RFQ PDF → quote draft (Excel)

A small offline, rule-based demo for a distributor receiving an RFQ. It reads
items from a text PDF, matches a catalogue and an approved price list, and
exports an Excel **draft for human review**. All companies, part codes,
descriptions and prices are synthetic. There is no real customer data or
production-use claim.

Built by Furkan Yalcintas together with an AI coding agent; every result here was run and checked by a human.

## Run

Use Python 3.10+. Run from this repository's directory. Create and activate a
virtual environment, then install the dependencies once. With dependencies
available, sample generation, quoting and tests run offline.

```sh
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python samples/make_samples.py
python -m rfq_quote.cli samples/rfq_a.pdf --catalogue samples/catalogue.csv --prices samples/price_list.csv --today 2026-10-03
python -m rfq_quote.cli samples/rfq_b.pdf --catalogue samples/catalogue.csv --prices samples/price_list.csv --today 2026-10-03
python -m unittest discover -s tests
```

On Windows PowerShell, activate with `.venv\Scripts\Activate.ps1` instead.
`requirements.txt` lists the dependencies. Regenerating samples replaces only
the four sample assets; the catalogue and price list have 50 fictional parts.
Each PDF contains 12 requested items, including one unparseable quantity.

Both examples, evaluated on `2026-10-03`, produce:

```text
priced: 7 / needs review: 3 / not in catalogue: 1 / unparsed: 1
Output: out/run-YYYYmmdd-HHMMSS/quote-draft.xlsx
```

The counts are the same because both PDFs come from the same generator and
contain the same anomaly categories: one description-only request, one unknown
code, one unit mismatch, one expired price and one unparseable quantity. They
use two different column layouts and different requested parts. These are
controlled synthetic examples, not independent evidence of general PDF accuracy.

Use `--today 2026-10-03` to reproduce the price-validity check; otherwise the
local current date is used. One example price expired in 2020; the others are
valid through 2099. Run-folder timestamps are UTC. Same-second runs get numeric
suffixes, and earlier outputs are never overwritten. `--out-dir PATH` changes
the output root. Errors have a nonzero exit code and leave no partial workbook.

## Decisions and audit trail

* Exact part code: use only that part's approved price.
* Unique code match ignoring spaces, hyphens and case: price with a note.
* Missing code with a similar description (similarity at least 0.75):
  **no price and no selected part**; show at most three candidates for review.
  Scores are string similarities, not probabilities or validated confidence
  estimates.
* Unknown explicit code, or no similar description: `not in catalogue`.
* Ambiguous normalised codes, missing prices, unclear/mismatched units: review.
* `valid_until` earlier than the evaluation date: `price expired`, no price.
  The final validity day is included. Expired rows count under `needs review`.
* Zero or negative unit price: `invalid price`, no price; counts under
  `needs review`. Malformed price CSV values still fail explicitly.

No substitute is selected automatically. Only explicit unit aliases (`each`,
`pcs`, `piece`, `pieces` to `ea`; `metre`, `meter` to `m`) are accepted; there
is no pack or unit conversion. Exact codes take precedence over descriptions;
description agreement is not independently verified for code matches.

The workbook has `Quote draft`, `Needs review`, and `Run info` sheets. Unparsed
rows appear in both the draft and review sheets with source text and location.
Run info includes the input filename and SHA256, reference-file hashes, UTC
start time, elapsed time before workbook saving, evaluation date and counts.
Line amounts round half up to two decimals; currency is explicit per line.
There is no cross-currency total, tax, discount or shipping calculation.

## Limits

Only text-based PDFs and these two example column layouts have been tested.
Tables must have the exact sample headers and `|` delimiters. Multi-page
tables with repeated headers are supported; wrapped descriptions, other
layouts and comma-formatted quantities are not. Unknown nonempty lines are
retained as `unparsed`; exact recognised headers and lines beginning `RFQ:`,
`Requester:`, `Delivery:`, `Notes:` or `Page:` are treated as metadata.
A PDF without a supported header fails explicitly.

Quantities accept only positive plain integers (`^\d+$`) or explicit dot
decimals (`^\d+\.\d+$`). A decimal with exactly three fractional digits, such
as `1.500`, is ambiguous (thousands or decimal?) and becomes `unparsed`.
Decimals with one, two, or four or more fractional digits are accepted, for
example `1.5`, `1.50`, and `1.5000`. Commas, underscores, exponent notation,
internal spaces, signs, zero and empty quantities are rejected: `1,5`, `1_000`,
`1e3`, `1 000`, `+5`, `-2`, `0`, and an empty field. Each rejected row remains
in the workbook with its source location, original text and a readable reason.

No scanned/image PDF support, OCR, LLM, ERP integration or automatic email.
There is no inventory check, engineering compatibility check or substitute
selection. Similarity candidates need a human decision. Excel stores numbers
as floating point, so very large or unusually precise values are outside this
demo's intended scope. This is a draft generator, not a final quote system.
