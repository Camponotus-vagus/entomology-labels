"""
Collection dates in the international entomological convention.

Labels write the month as a lower-case Roman numeral -- ``20.viii.2026`` --
because an all-numeric date is ambiguous the moment a specimen crosses a
border: 08/09/2026 is September in Bergen and August in Boston. The Roman
numeral removes the ambiguity, and this module is what converts to it.

Parsing is deliberately forgiving and normalisation is non-destructive.
Real labels carry things like ``viii.2026`` or ``ex larva 2026``, and the
field has always been free text, so anything unrecognised is passed through
untouched rather than rejected.
"""

import logging
import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import List, Optional, Union

logger = logging.getLogger(__name__)

#: Roman numerals for months 1-12, as written on specimen labels.
ROMAN_MONTHS = (
    "i",
    "ii",
    "iii",
    "iv",
    "v",
    "vi",
    "vii",
    "viii",
    "ix",
    "x",
    "xi",
    "xii",
)

#: Lookup from a Roman numeral to its month number.
_ROMAN_TO_MONTH = {roman: index + 1 for index, roman in enumerate(ROMAN_MONTHS)}

#: Years outside this range are almost certainly a typo or a bad EXIF tag.
MIN_YEAR = 1700
MAX_YEAR = 2200

_ROMAN_PATTERN = "|".join(sorted(ROMAN_MONTHS, key=len, reverse=True))

_ISO_RE = re.compile(r"^(\d{4})[-:](\d{1,2})[-:](\d{1,2})(?:[ T].*)?$")
_ROMAN_RE = re.compile(rf"^(\d{{1,2}})\.({_ROMAN_PATTERN})\.(\d{{4}})$", re.IGNORECASE)
_ROMAN_MONTH_YEAR_RE = re.compile(rf"^({_ROMAN_PATTERN})\.(\d{{4}})$", re.IGNORECASE)
_NUMERIC_RE = re.compile(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$")
_YEAR_RE = re.compile(r"^(\d{4})$")
_RANGE_RE = re.compile(
    rf"^(\d{{1,2}})\s*[-–]\s*(\d{{1,2}})\.({_ROMAN_PATTERN})\.(\d{{4}})$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class DateParts:
    """A parsed collection date. Day and month may be absent."""

    year: int
    month: Optional[int] = None
    day: Optional[int] = None
    day_end: Optional[int] = None
    ambiguous: bool = False

    def to_roman(self) -> str:
        """Render these parts in the entomological convention."""
        if self.month is None:
            return str(self.year)

        roman = ROMAN_MONTHS[self.month - 1]
        if self.day is None:
            return f"{roman}.{self.year}"
        if self.day_end is not None:
            return f"{self.day:02d}-{self.day_end:02d}.{roman}.{self.year}"
        return f"{self.day:02d}.{roman}.{self.year}"


def _valid(year: int, month: Optional[int] = None, day: Optional[int] = None) -> bool:
    """Return whether the given parts form a real calendar date."""
    if not (MIN_YEAR <= year <= MAX_YEAR):
        return False
    if month is not None and not (1 <= month <= 12):
        return False
    if day is not None:
        if month is None:
            return False
        try:
            date(year, month, day)
        except ValueError:
            return False
    return True


def to_roman_date(value: Union[date, datetime]) -> str:
    """Format a date or datetime in the entomological convention.

    Args:
        value: The date to format

    Returns:
        The date as ``DD.roman.YYYY``, e.g. ``20.viii.2026``
    """
    if isinstance(value, datetime):
        value = value.date()
    return f"{value.day:02d}.{ROMAN_MONTHS[value.month - 1]}.{value.year}"


def parse_date(text: str) -> Optional[DateParts]:
    """Parse a collection date written in any of the accepted forms.

    Accepts ISO (``2026-08-20``), EXIF (``2026:08:20 09:14:32``), the Roman
    convention (``20.viii.2026``), a day range (``20-24.viii.2026``), a
    day-first numeric date (``20/08/2026``), a month and year
    (``viii.2026``), or a bare year.

    Args:
        text: The date as written

    Returns:
        The parsed parts, or None if the text is not a date in any known form
    """
    if not text:
        return None

    cleaned = text.strip()

    match = _RANGE_RE.match(cleaned)
    if match:
        day, day_end, roman, year = match.groups()
        month = _ROMAN_TO_MONTH[roman.lower()]
        if _valid(int(year), month, int(day)) and _valid(int(year), month, int(day_end)):
            return DateParts(int(year), month, int(day), int(day_end))
        return None

    match = _ROMAN_RE.match(cleaned)
    if match:
        day, roman, year = match.groups()
        month = _ROMAN_TO_MONTH[roman.lower()]
        return DateParts(int(year), month, int(day)) if _valid(int(year), month, int(day)) else None

    match = _ROMAN_MONTH_YEAR_RE.match(cleaned)
    if match:
        roman, year = match.groups()
        month = _ROMAN_TO_MONTH[roman.lower()]
        return DateParts(int(year), month) if _valid(int(year), month) else None

    match = _ISO_RE.match(cleaned)
    if match:
        year, month, day = (int(g) for g in match.groups())
        return DateParts(year, month, day) if _valid(year, month, day) else None

    match = _NUMERIC_RE.match(cleaned)
    if match:
        # Day-first: the European convention, and the one this project's
        # existing data uses. Flagged as ambiguous when it could be read
        # either way, so callers can warn instead of guessing silently.
        day, month, year = (int(g) for g in match.groups())
        if not _valid(year, month, day):
            return None
        return DateParts(year, month, day, ambiguous=day <= 12 and day != month)

    match = _YEAR_RE.match(cleaned)
    if match:
        year = int(match.group(1))
        return DateParts(year) if _valid(year) else None

    return None


def normalize_date(text: str) -> str:
    """Convert a date to the Roman convention, leaving it alone if unparseable.

    Args:
        text: The date as written

    Returns:
        The normalised date, or the original text unchanged
    """
    parts = parse_date(text)
    if parts is None:
        return text
    return parts.to_roman()


def validate_date(text: str) -> List[str]:
    """Describe anything questionable about a hand-written date.

    Args:
        text: The date as written

    Returns:
        Human-readable warnings; empty when the date looks fine
    """
    if not text or not text.strip():
        return []

    parts = parse_date(text)
    if parts is None:
        return [f'"{text}" is not a recognised date format and will print as written']

    warnings = []
    if parts.ambiguous:
        warnings.append(
            f'"{text}" was read day-first as {parts.to_roman()}; '
            f"write the month as a Roman numeral to remove the ambiguity"
        )
    if parts.to_roman() != text.strip():
        warnings.append(f'"{text}" is not in the usual label form ({parts.to_roman()})')
    return warnings
