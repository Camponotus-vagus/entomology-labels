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


class TestProbeElevation:
    """A reading of exactly zero is a reading, not a missing value."""

    def _probe_output(self, tmp_path, altitude_m, exif_builder):
        from click.testing import CliRunner

        from entomology_labels.cli import cli

        path = tmp_path / "shore.ORF"
        path.write_bytes(exif_builder(altitude_m=altitude_m))
        return CliRunner().invoke(cli, ["photos", "probe", str(tmp_path)]).output

    def test_sea_level_is_printed_as_a_number(self, tmp_path, exif_builder):
        """A fjord shore reads 0 m; `or "-"` made that look like no GPS fix."""
        output = self._probe_output(tmp_path, 0.0, exif_builder)

        elevation_line = next(ln for ln in output.splitlines() if "elevation" in ln)
        assert "0" in elevation_line
        assert elevation_line.strip().endswith("-") is False

    def test_a_normal_altitude_still_prints(self, tmp_path, exif_builder):
        output = self._probe_output(tmp_path, 312.2, exif_builder)

        assert "312.2" in output


class TestReviewRegressions:
    def test_photos_group_exists_when_run_as_a_module(self):
        import subprocess
        import sys

        result = subprocess.run(
            [sys.executable, "-m", "entomology_labels.cli", "photos", "--help"],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0, result.stderr
        assert "scan" in result.stdout

    def test_a_json_draft_is_refused(self, runner, tmp_path):
        result = runner.invoke(
            cli, ["photos", "scan", str(tmp_path), "-o", str(tmp_path / "d.json")]
        )
        assert result.exit_code != 0
        assert "Use .yaml or .yml" in result.output

    def test_output_cannot_overwrite_the_input(self, runner, tmp_path):
        docx = pytest.importorskip("docx")
        source = tmp_path / "in.docx"
        document = docx.Document()
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text, table.cell(0, 1).text = "location1", "code"
        table.cell(1, 0).text, table.cell(1, 1).text = "Italia", "N1"
        document.save(source)
        before = source.read_bytes()

        result = runner.invoke(cli, ["generate", str(source), "-o", str(source)])

        assert result.exit_code != 0
        assert "overwrite" in result.output
        assert source.read_bytes() == before

    def test_sequence_reads_label_config_json(self, runner, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        (tmp_path / "label_config.json").write_text(
            json.dumps({"labels_per_row": 3, "labels_per_column": 4}), encoding="utf-8"
        )
        result = runner.invoke(
            cli,
            [
                "sequence",
                "--location1",
                "a",
                "--location2",
                "b",
                "--prefix",
                "N",
                "--start",
                "1",
                "--end",
                "30",
                "-o",
                "s.html",
            ],
        )
        assert result.exit_code == 0, result.output
        assert "Page 3 of 3" in (tmp_path / "s.html").read_text(encoding="utf-8")

    def test_site_column_without_sites_file_warns(self, runner, tmp_path):
        source = tmp_path / "s.csv"
        source.write_text("site,code\nFLOYEN,N1\n", encoding="utf-8")
        result = runner.invoke(cli, ["generate", str(source), "-o", str(tmp_path / "o.html")])
        assert result.exit_code == 0
        assert "no --sites file" in result.output

    def test_subcommand_help_needs_no_exifread(self, tmp_path):
        import subprocess
        import sys

        # Block the import the way a missing package would.
        (tmp_path / "exifread.py").write_text("raise ImportError('blocked')\n", encoding="utf-8")
        env = dict(os.environ, PYTHONPATH=str(tmp_path))
        result = subprocess.run(
            [sys.executable, "-m", "entomology_labels.cli", "photos", "scan", "--help"],
            capture_output=True,
            text=True,
            env=env,
        )
        assert result.returncode == 0, result.stderr
        assert "--geocode" in result.stdout
