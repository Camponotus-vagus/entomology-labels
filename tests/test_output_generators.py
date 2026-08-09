"""Tests for the output generators."""

import pytest

from entomology_labels.label_generator import Label, LabelConfig, LabelGenerator
from entomology_labels.output_generators import (
    _escape_html,
    generate_docx,
    generate_html,
)


@pytest.fixture
def generator():
    """A generator holding one ordinary label."""
    gen = LabelGenerator(LabelConfig(labels_per_row=2, labels_per_column=2))
    gen.add_label(Label(location_line1="Italia", location_line2="Milano", code="N1", date="2024"))
    return gen


class TestEscapeHtml:
    """Label text must not be able to inject markup."""

    def test_escapes_markup(self):
        assert _escape_html("<script>") == "&lt;script&gt;"

    def test_escapes_quotes(self):
        escaped = _escape_html("a \"b\" 'c'")
        assert '"' not in escaped
        assert "'" not in escaped

    def test_empty_is_empty(self):
        assert _escape_html("") == ""

    def test_label_text_is_escaped_in_output(self, generator):
        generator.add_label(Label(code="<script>alert(1)</script>"))
        html = generate_html(generator)
        assert "<script>alert(1)</script>" not in html
        assert "&lt;script&gt;" in html


class TestOutputDirectory:
    """Writing into a directory that does not exist yet should just work."""

    def test_html_creates_missing_parent(self, generator, tmp_path):
        target = tmp_path / "reports" / "2026" / "labels.html"
        generate_html(generator, target)
        assert target.exists()

    def test_docx_creates_missing_parent(self, generator, tmp_path):
        target = tmp_path / "nested" / "deeper" / "labels.docx"
        generate_docx(generator, target)
        assert target.exists()


class TestFontFamilyIsNotInjectable:
    """font_family lands in a CSS declaration, so it is restricted."""

    def test_css_injection_rejected(self):
        with pytest.raises(ValueError, match="font_family must match"):
            LabelConfig(font_family="Arial; } body{display:none} .x{color:red")

    def test_ordinary_font_name_accepted(self):
        assert LabelConfig(font_family="Times New Roman").font_family == "Times New Roman"

    def test_configured_font_reaches_the_stylesheet(self, generator):
        generator.config.font_family = "Courier New"
        assert "font-family: Courier New" in generate_html(generator)
