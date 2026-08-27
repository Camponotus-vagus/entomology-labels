"""Tests for collection date parsing and normalization."""

from datetime import date, datetime

import pytest

from entomology_labels.dates import (
    ROMAN_MONTHS,
    normalize_date,
    parse_date,
    to_roman_date,
    validate_date,
)


class TestParsing:
    @pytest.mark.parametrize(
        "text,expected",
        [
            ("2026-08-20", "20.viii.2026"),  # ISO
            ("2026:08:20 09:14:32", "20.viii.2026"),  # EXIF
            ("20.viii.2026", "20.viii.2026"),  # already correct
            ("20.VIII.2026", "20.viii.2026"),  # upper case
            ("  20.viii.2026  ", "20.viii.2026"),  # padded
            ("20/08/2026", "20.viii.2026"),  # day-first numeric
            ("20-08-2026", "20.viii.2026"),
            ("20.08.2026", "20.viii.2026"),
            ("20-24.viii.2026", "20-24.viii.2026"),  # a range of days
            ("viii.2026", "viii.2026"),  # month and year only
            ("2026", "2026"),  # year only
            ("1.i.2026", "01.i.2026"),  # day is zero-padded
        ],
    )
    def test_accepted_forms(self, text, expected):
        assert normalize_date(text) == expected

    @pytest.mark.parametrize(
        "text",
        [
            "32.viii.2026",  # no such day
            "20.xiii.2026",  # no such month
            "2026-13-01",
            "2026-02-30",
            "31.ii.2026",
            "ex larva 2026",  # legitimate free text
            "summer 2026",
            "",
            "not a date",
            "20.viii.1200",  # implausible year
        ],
    )
    def test_rejected_forms(self, text):
        assert parse_date(text) is None

    def test_normalizing_is_idempotent(self):
        once = normalize_date("2026-08-20")

        assert normalize_date(once) == once

    def test_unparseable_text_passes_through_untouched(self):
        """The date field is free text and always has been."""
        assert normalize_date("ex larva 2026") == "ex larva 2026"

    def test_vi_and_vii_are_distinguished(self):
        """The typo this format is most prone to."""
        assert parse_date("15.vi.2024").month == 6
        assert parse_date("15.vii.2024").month == 7
        assert normalize_date("15.vi.2024") != normalize_date("15.vii.2024")

    def test_every_month_round_trips(self):
        for number, roman in enumerate(ROMAN_MONTHS, start=1):
            assert normalize_date(f"2026-{number:02d}-05") == f"05.{roman}.2026"


class TestToRomanDate:
    def test_formats_a_date(self):
        assert to_roman_date(date(2026, 8, 20)) == "20.viii.2026"

    def test_accepts_a_datetime_as_exif_supplies(self):
        assert to_roman_date(datetime(2026, 8, 20, 14, 19, 27)) == "20.viii.2026"

    def test_january_and_december_are_not_off_by_one(self):
        assert to_roman_date(date(2026, 1, 1)) == "01.i.2026"
        assert to_roman_date(date(2026, 12, 31)) == "31.xii.2026"


class TestValidation:
    def test_a_correctly_written_date_produces_no_warnings(self):
        assert validate_date("20.viii.2026") == []

    def test_an_empty_date_produces_no_warnings(self):
        assert validate_date("") == []
        assert validate_date("   ") == []

    def test_an_ambiguous_numeric_date_is_flagged(self):
        warnings = validate_date("05/06/2026")

        assert any("day-first" in w for w in warnings)

    def test_an_unambiguous_numeric_date_is_not_called_ambiguous(self):
        """Day 20 cannot be a month, so there is nothing to warn about."""
        warnings = validate_date("20/08/2026")

        assert not any("day-first" in w for w in warnings)

    def test_unrecognised_text_is_flagged_as_printing_verbatim(self):
        warnings = validate_date("ex larva 2026")

        assert any("print as written" in w for w in warnings)
