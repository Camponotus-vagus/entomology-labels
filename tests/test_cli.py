"""Tests for the command line interface."""

import json
import os

import pytest
from click.testing import CliRunner

from entomology_labels.cli import cli


@pytest.fixture
def runner():
    return CliRunner()


@pytest.fixture
def labels_file(tmp_path):
    path = tmp_path / "labels.json"
    path.write_text(
        json.dumps({"labels": [{"location_line1": "Italia", "code": "N1"}]}),
        encoding="utf-8",
    )
    return path


class TestGenerate:
    """The generate command, end to end."""

    def test_generates_html(self, runner, labels_file, tmp_path):
        out = tmp_path / "labels.html"
        result = runner.invoke(cli, ["generate", str(labels_file), "-o", str(out)])
        assert result.exit_code == 0, result.output
        assert out.exists()
        assert "N1" in out.read_text(encoding="utf-8")

    def test_creates_missing_output_directory(self, runner, labels_file, tmp_path):
        out = tmp_path / "reports" / "2026" / "labels.html"
        result = runner.invoke(cli, ["generate", str(labels_file), "-o", str(out)])
        assert result.exit_code == 0, result.output
        assert out.exists()

    def test_rejects_unknown_output_format(self, runner, labels_file, tmp_path):
        result = runner.invoke(cli, ["generate", str(labels_file), "-o", str(tmp_path / "x.txt")])
        assert result.exit_code != 0
        assert "Unsupported output format" in result.output

    @pytest.mark.skipif(
        hasattr(os, "geteuid") and os.geteuid() == 0,
        reason="root ignores permission bits, so the directory stays writable",
    )
    def test_rejects_unwritable_output_directory(self, runner, labels_file, tmp_path):
        readonly = tmp_path / "readonly"
        readonly.mkdir()
        readonly.chmod(0o500)
        try:
            result = runner.invoke(
                cli, ["generate", str(labels_file), "-o", str(readonly / "labels.html")]
            )
            assert result.exit_code != 0
            assert "Cannot write to output directory" in result.output
        finally:
            readonly.chmod(0o700)

    def test_rejects_zero_rows(self, runner, labels_file, tmp_path):
        result = runner.invoke(
            cli,
            ["generate", str(labels_file), "-o", str(tmp_path / "o.html"), "--rows", "0"],
        )
        assert result.exit_code != 0
        assert "labels_per_row must be between" in result.output

    def test_rejects_injected_font_family(self, runner, labels_file, tmp_path):
        result = runner.invoke(
            cli,
            [
                "generate",
                str(labels_file),
                "-o",
                str(tmp_path / "o.html"),
                "--font-family",
                "Arial; } body{display:none",
            ],
        )
        assert result.exit_code != 0
        assert "font_family must match" in result.output

    def test_reports_oversized_count_from_input_file(self, runner, tmp_path):
        bad = tmp_path / "bad.json"
        bad.write_text(json.dumps({"labels": [{"code": "N1", "count": 5_000_000}]}), "utf-8")
        result = runner.invoke(cli, ["generate", str(bad), "-o", str(tmp_path / "o.html")])
        assert result.exit_code != 0
        assert "exceeds the maximum" in result.output


class TestSequence:
    """The sequence command."""

    def test_generates_a_range(self, runner, tmp_path):
        out = tmp_path / "seq.html"
        result = runner.invoke(
            cli,
            [
                "sequence",
                "--location1",
                "Italia",
                "--location2",
                "Milano",
                "--prefix",
                "N",
                "--start",
                "1",
                "--end",
                "5",
                "-o",
                str(out),
            ],
        )
        assert result.exit_code == 0, result.output
        assert "5 sequential labels" in result.output
        assert out.exists()

    def test_rejects_inverted_range(self, runner, tmp_path):
        result = runner.invoke(
            cli,
            [
                "sequence",
                "--location1",
                "Italia",
                "--location2",
                "Milano",
                "--prefix",
                "N",
                "--start",
                "20",
                "--end",
                "1",
                "-o",
                str(tmp_path / "seq.html"),
            ],
        )
        assert result.exit_code != 0
        assert "must be greater than or equal to" in result.output

    def test_rejects_oversized_grid(self, runner, tmp_path):
        result = runner.invoke(
            cli,
            [
                "sequence",
                "--location1",
                "Italia",
                "--location2",
                "Milano",
                "--prefix",
                "N",
                "--start",
                "1",
                "--end",
                "1",
                "--cols",
                "200000",
                "-o",
                str(tmp_path / "seq.html"),
            ],
        )
        assert result.exit_code != 0
        assert "labels_per_column must be between" in result.output


class TestTemplate:
    """The template command writes a starter file that loads back in."""

    def test_json_template_round_trips(self, runner, tmp_path):
        target = tmp_path / "template.json"
        assert runner.invoke(cli, ["template", str(target)]).exit_code == 0

        out = tmp_path / "labels.html"
        result = runner.invoke(cli, ["generate", str(target), "-o", str(out)])
        assert result.exit_code == 0, result.output
        assert out.exists()
