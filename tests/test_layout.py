"""Tests for the shared label body layout."""

from entomology_labels.label_generator import Label, LabelConfig
from entomology_labels.layout import (
    INFO_FONT_SCALE,
    STYLE_CODE,
    STYLE_DATE,
    STYLE_INFO,
    STYLE_LOCATION,
    STYLE_SPACER,
    render_label_lines,
)


def _label(**kwargs) -> Label:
    base = {
        "location_line1": "Norway, Vestland,",
        "location_line2": "Bergen, Floyen",
        "code": "N1",
        "date": "20.viii.2026",
    }
    base.update(kwargs)
    return Label(**base)


class TestRenderLabelLines:
    """The layout is the single source of truth for label composition."""

    def test_five_field_label_keeps_its_historical_shape(self):
        """Guards the refactor: HTML, DOCX and the GUI all depend on this order."""
        lines = render_label_lines(_label(), LabelConfig())

        assert [(line.text, line.style) for line in lines] == [
            ("Norway, Vestland,", STYLE_LOCATION),
            ("Bergen, Floyen", STYLE_LOCATION),
            ("", STYLE_SPACER),
            ("N1", STYLE_CODE),
            ("20.viii.2026", STYLE_DATE),
        ]

    def test_additional_info_appends_a_scaled_italic_line(self):
        lines = render_label_lines(_label(additional_info="leg. F. Mensa"), LabelConfig())

        assert len(lines) == 6
        info = lines[-1]
        assert info.text == "leg. F. Mensa"
        assert info.style == STYLE_INFO
        assert info.italic is True
        assert info.scale == INFO_FONT_SCALE

    def test_core_lines_are_emitted_even_when_empty(self):
        """Every label must occupy the same height or the printed grid shifts."""
        lines = render_label_lines(Label(), LabelConfig())

        assert len(lines) == 5
        assert all(line.text == "" for line in lines)

    def test_lines_default_to_upright_full_size(self):
        for line in render_label_lines(_label(), LabelConfig()):
            assert line.italic is False
            assert line.scale == 1.0
