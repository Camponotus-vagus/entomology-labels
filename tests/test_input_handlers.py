"""Tests for the input handlers module."""

import json
from pathlib import Path

import pytest

from entomology_labels.config import MAX_COPIES_PER_ENTRY
from entomology_labels.input_handlers import _parse_count, _validate_file_path, load_data


@pytest.fixture
def json_file(tmp_path):
    """Write a JSON label file and return its path."""

    def _write(labels):
        path = tmp_path / "labels.json"
        path.write_text(json.dumps({"labels": labels}), encoding="utf-8")
        return path

    return _write


class TestValidateFilePath:
    """Tests for _validate_file_path."""

    def test_accepts_readable_file(self, tmp_path):
        path = tmp_path / "data.json"
        path.write_text("[]", encoding="utf-8")
        assert _validate_file_path(path) == path.resolve()

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            _validate_file_path(tmp_path / "nope.json")

    def test_directory_raises(self, tmp_path):
        with pytest.raises(ValueError, match="Not a file"):
            _validate_file_path(tmp_path)


class TestLoadData:
    """Tests for the load_data dispatcher."""

    def test_loads_json_file(self, json_file):
        path = json_file([{"location_line1": "Italia", "code": "N1"}])
        labels = load_data(path)
        assert len(labels) == 1
        assert labels[0].code == "N1"

    def test_loads_txt_key_value_file(self, tmp_path):
        path = tmp_path / "labels.txt"
        path.write_text("location1: Italia\ncode: N1\n", encoding="utf-8")
        labels = load_data(path)
        assert len(labels) == 1
        assert labels[0].code == "N1"

    def test_unsupported_extension_raises(self, tmp_path):
        path = tmp_path / "labels.pdf"
        path.write_text("x", encoding="utf-8")
        with pytest.raises(ValueError, match="Unsupported file format"):
            load_data(path)


class TestParseCount:
    """Tests for the count validation shared by every loader."""

    def test_missing_count_uses_default(self):
        assert _parse_count(None) == 1
        assert _parse_count("") == 1

    def test_numeric_values(self):
        assert _parse_count(3) == 3
        assert _parse_count("4") == 4
        assert _parse_count("5.0") == 5

    def test_non_numeric_raises(self):
        with pytest.raises(ValueError, match="Invalid count value"):
            _parse_count("abc")

    def test_negative_raises(self):
        with pytest.raises(ValueError, match="cannot be negative"):
            _parse_count(-1)

    def test_above_limit_raises(self):
        with pytest.raises(ValueError, match="exceeds the maximum"):
            _parse_count(MAX_COPIES_PER_ENTRY + 1)


class TestLoaderEdgeCases:
    """Edge cases that used to abort a whole file."""

    def test_empty_docx_table_is_skipped(self, tmp_path):
        docx = pytest.importorskip("docx")

        path = tmp_path / "labels.docx"
        doc = docx.Document()
        doc.add_table(rows=0, cols=3)  # a table shell with no header row
        doc.add_paragraph("Italia")
        doc.add_paragraph("Milano")
        doc.add_paragraph("N1")
        doc.save(path)

        labels = load_data(path)
        assert any(label.code == "N1" for label in labels)

    def test_zero_is_not_treated_as_missing(self, tmp_path):
        pytest.importorskip("pandas")

        path = tmp_path / "labels.csv"
        path.write_text("location_line1,code,date\nItalia,0,2024\n", encoding="utf-8")

        labels = load_data(path)
        assert len(labels) == 1
        assert labels[0].code == "0"

    def test_malformed_csv_reports_its_own_error(self, tmp_path):
        pytest.importorskip("pandas")
        import pandas as pd

        path = tmp_path / "broken.csv"
        path.write_text('location_line1,code\n"Italia,N1\n', encoding="utf-8")

        # The comma-parse failure is the real one, and it must not be masked
        # by a second attempt with a different delimiter.
        with pytest.raises(pd.errors.ParserError):
            load_data(path)


class TestCountExpansion:
    """The copy count is bounded before any large list is materialised."""

    def test_count_expands_labels(self, json_file):
        path = json_file([{"code": "N1", "count": 3}])
        assert len(load_data(path)) == 3

    def test_huge_count_is_rejected(self, json_file):
        path = json_file([{"code": "N1", "count": 5_000_000}])
        with pytest.raises(ValueError, match="exceeds the maximum"):
            load_data(path)

    def test_non_numeric_count_is_rejected(self, json_file):
        path = json_file([{"code": "N1", "count": "abc"}])
        with pytest.raises(ValueError, match="Invalid count value"):
            load_data(path)

    def test_negative_count_is_rejected(self, json_file):
        path = json_file([{"code": "N1", "count": -5}])
        with pytest.raises(ValueError, match="cannot be negative"):
            load_data(path)


class TestShippedExamples:
    """The example files are documentation; they must load and agree."""

    EXAMPLES = Path(__file__).resolve().parent.parent / "examples"

    def test_every_example_format_loads(self):
        for name in ("example_labels.json", "example_labels.csv", "example_labels.txt"):
            labels = load_data(self.EXAMPLES / name)
            assert labels, f"{name} produced no labels"

    def test_the_three_formats_describe_the_same_data(self):
        """They are hand-maintained siblings and drift silently otherwise."""
        loaded = {
            name: [label.to_dict() for label in load_data(self.EXAMPLES / name)]
            for name in ("example_labels.json", "example_labels.csv", "example_labels.txt")
        }

        reference = loaded["example_labels.json"]
        for name, labels in loaded.items():
            assert labels == reference, f"{name} disagrees with example_labels.json"

    def test_non_ascii_localities_survive_the_round_trip(self):
        """Diacritics in place names are the common case outside English."""
        localities = {
            label.location_line2 for label in load_data(self.EXAMPLES / "example_labels.json")
        }

        assert "Bergen, Fløyen" in localities


class TestTabularCellsReadAsText:
    """Spreadsheet cells come through exactly as written."""

    def _load(self, tmp_path, text, name="t.csv"):
        pytest.importorskip("pandas")
        path = tmp_path / name
        path.write_text(text, encoding="utf-8")
        return load_data(path)

    def test_numeric_code_with_a_blank_cell_is_not_a_float(self, tmp_path):
        labels = self._load(tmp_path, "location1,code,elevation\nIt,,\nIt,5,310\n")
        assert labels[1].code == "5"
        assert labels[1].elevation == "310"

    def test_leading_zeros_survive(self, tmp_path):
        assert self._load(tmp_path, "location1,code\nIt,007\n")[0].code == "007"

    def test_na_is_namibia_not_a_missing_value(self, tmp_path):
        labels = self._load(tmp_path, "location1,location2,code\nNA,Windhoek,N1\n")
        assert labels[0].location_line1 == "NA"

    def test_duplicate_headings_are_rejected(self, tmp_path):
        with pytest.raises(ValueError, match="Duplicate column"):
            self._load(tmp_path, "Code,code,location1\nA,B,C\n")

    def test_utf8_bom_is_accepted(self, tmp_path):
        pytest.importorskip("pandas")
        path = tmp_path / "bom.csv"
        path.write_bytes("﻿location1,code\nIt,N1\n".encode("utf-8"))
        assert load_data(path)[0].location_line1 == "It"


class TestCountParsing:
    def test_whole_float_is_accepted(self):
        assert _parse_count("5.0") == 5

    @pytest.mark.parametrize("value", ["2.7", "-0.5", "inf", "1e400", "nan", True])
    def test_non_counts_are_rejected(self, value):
        with pytest.raises(ValueError):
            _parse_count(value)


class TestTxtFormatDetection:
    def test_a_time_does_not_switch_to_key_value(self, tmp_path):
        path = tmp_path / "simple.txt"
        path.write_text("Italia\nRoma\nN1\n15.vi.2024\ncaught at 12:30\n", encoding="utf-8")
        labels = load_data(path)
        assert labels[0].code == "N1"
        assert labels[0].additional_info == "caught at 12:30"

    def test_key_value_blocks_split_on_whitespace_lines(self, tmp_path):
        path = tmp_path / "kv.txt"
        path.write_text("location1: a\ncode: 1\n  \nlocation1: b\ncode: 2\n", encoding="utf-8")
        assert [label.location_line1 for label in load_data(path)] == ["a", "b"]

    def test_bom_does_not_hide_the_first_key(self, tmp_path):
        path = tmp_path / "kv.txt"
        path.write_bytes("﻿location1: a\ncode: 1\n".encode("utf-8"))
        assert load_data(path)[0].location_line1 == "a"


class TestStructuredShape:
    @pytest.mark.parametrize("body", ["null", '{"labels": {"a": 1}}', '["text"]'])
    def test_malformed_documents_raise_value_error(self, tmp_path, body):
        path = tmp_path / "bad.json"
        path.write_text(body, encoding="utf-8")
        if body == "null":
            assert load_data(path) == []
        else:
            with pytest.raises(ValueError):
                load_data(path)

    def test_json_with_bom_loads(self, tmp_path):
        path = tmp_path / "bom.json"
        path.write_bytes('﻿[{"code": "N1"}]'.encode("utf-8"))
        assert load_data(path)[0].code == "N1"
