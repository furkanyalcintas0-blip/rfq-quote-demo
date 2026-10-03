"""Conservative code matching and approved-price lookup; never substitute."""

import csv
import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path

from .extract import Item


@dataclass(frozen=True)
class Part:
    """A catalogue part and its sales unit."""

    part_no: str
    description: str
    unit: str


@dataclass(frozen=True)
class Price:
    """An approved price, with an inclusive validity end date."""

    part_no: str
    unit_price: Decimal
    currency: str
    valid_until: date


@dataclass(frozen=True)
class Candidate:
    """A description suggestion for human review, without a price."""

    part: Part
    score: float


@dataclass(frozen=True)
class Match:
    """A requested item, decision, explanation, and optional approved price."""

    item: Item
    status: str
    reason: str
    part: Part | None = None
    price: Price | None = None
    candidates: tuple[Candidate, ...] = ()


def normalize_code(code: str) -> str:
    """Ignore only whitespace, hyphens, and letter case in part numbers."""
    return re.sub(r"[\s-]+", "", code).upper()


def _read_csv(path, columns):
    with Path(path).open(newline="", encoding="utf-8-sig") as stream:
        reader = csv.DictReader(stream)
        if not set(columns).issubset(reader.fieldnames or []):
            raise ValueError(f"{Path(path).name}: missing required CSV columns")
        rows = []
        seen = set()
        for number, row in enumerate(reader, 2):
            values = {column: (row.get(column) or "").strip() for column in columns}
            if not all(values.values()) or None in row:
                raise ValueError(f"{Path(path).name}: invalid CSV row {number}")
            if values["part_no"] in seen:
                raise ValueError(f"{Path(path).name}: duplicate part_no at row {number}")
            seen.add(values["part_no"])
            rows.append(values)
        if not rows:
            raise ValueError(f"{Path(path).name}: empty data")
        return rows


def load_catalogue(path: str | Path) -> dict[str, Part]:
    """Read catalogue CSV; reject empty fields and duplicate exact codes."""
    return {row["part_no"]: Part(**row)
            for row in _read_csv(path, ("part_no", "description", "unit"))}


def load_prices(path: str | Path) -> dict[str, Price]:
    """Read prices; reject malformed data, retain nonpositive amounts for review."""
    prices = {}
    for row in _read_csv(path, ("part_no", "unit_price", "currency", "valid_until")):
        try:
            amount = Decimal(row["unit_price"])
            if not amount.is_finite():
                raise ValueError("Price must be finite")
            if not re.fullmatch(r"[A-Z]{3}", row["currency"]):
                raise ValueError("Currency must be a three-letter uppercase code")
            prices[row["part_no"]] = Price(row["part_no"], amount, row["currency"],
                                           date.fromisoformat(row["valid_until"]))
        except (InvalidOperation, ValueError) as error:
            raise ValueError(f"Invalid price for {row['part_no']}: {error}") from error
    return prices


def _description(text):
    return " ".join(re.findall(r"[a-z0-9]+", text.lower()))


def _unit(unit):
    return {"each": "ea", "pcs": "ea", "piece": "ea", "pieces": "ea",
            "metre": "m", "meter": "m"}.get(unit.lower().strip(), unit.lower().strip())


def match_items(items: list[Item], catalogue: dict[str, Part],
                prices: dict[str, Price], today: date | None = None) -> list[Match]:
    """Match codes only for pricing; suggest up to three description candidates.

    Unknown explicit codes are never replaced. Description-only requests use a
    0.75 SequenceMatcher threshold. Unit mismatches, normalisation collisions,
    missing, nonpositive, nonfinite and expired prices require human review.
    Supply today to make expiry checks deterministic; a price valid_until ==
    today is valid.
    """
    today = today if today is not None else date.today()
    normalized = {}
    for part in catalogue.values():
        normalized.setdefault(normalize_code(part.part_no), []).append(part)
    matches = []
    for item in items:
        part = catalogue.get(item.part_no)
        note = "Exact part number match"
        if part is None and item.part_no:
            choices = normalized.get(normalize_code(item.part_no), [])
            if len(choices) > 1:
                matches.append(Match(item, "needs review", "Normalized code is ambiguous",
                                     candidates=tuple(Candidate(p, 1.0) for p in choices[:3])))
                continue
            if choices:
                part = choices[0]
                note = f"Normalized code match: {item.part_no} -> {part.part_no}"
        if part is None:
            candidates = ()
            if not item.part_no:
                scored = [Candidate(p, SequenceMatcher(None, _description(item.description),
                                                       _description(p.description)).ratio())
                          for p in catalogue.values()]
                candidates = tuple(sorted((c for c in scored if c.score >= 0.75),
                                          key=lambda c: (-c.score, c.part.part_no))[:3])
            matches.append(Match(item, "needs review" if candidates else "not in catalogue",
                                 "Description only; human must select a part" if candidates
                                 else "No catalogue code match or sufficiently similar description",
                                 candidates=candidates))
            continue
        if not item.unit or _unit(item.unit) != _unit(part.unit):
            matches.append(Match(item, "needs review", f"{note}; unit unclear or mismatched: "
                                 f"{item.unit or '(missing)'} vs {part.unit}", part))
            continue
        price = prices.get(part.part_no)
        if price is None:
            matches.append(Match(item, "needs review", f"{note}; no approved price", part))
        elif not price.unit_price.is_finite() or price.unit_price <= 0:
            matches.append(Match(item, "invalid price", f"{note}; unit price must be positive "
                                 "and finite", part))
        elif price.valid_until < today:
            matches.append(Match(item, "price expired", f"{note}; price expired on "
                                 f"{price.valid_until.isoformat()}", part))
        else:
            matches.append(Match(item, "priced", note, part, price))
    return matches
