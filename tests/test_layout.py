"""Tests for the shared label body layout."""

import pytest

from entomology_labels.label_generator import Label, LabelConfig
from entomology_labels.layout import (
    INFO_FONT_SCALE,
    KIND_BOTH,
    KIND_DETERMINATION,
    KIND_LOCALITY,
    STYLE_ATTRIBUTION,
    STYLE_CODE,
    STYLE_COORDS,
    STYLE_DATE,
    STYLE_ELEVATION,
    STYLE_INFO,
    STYLE_LOCATION,
    STYLE_SPACER,
    build_sheet,
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


class TestExtendedFields:
    """New fields appear only when filled, so old labels render unchanged."""

    def test_coordinates_and_elevation_sit_below_the_locality(self):
        label = _label(coordinates="60.3965N 5.3531E", elevation="310 m")

        styles = [line.style for line in render_label_lines(label, LabelConfig())]

        assert styles == [
            STYLE_LOCATION,
            STYLE_LOCATION,
            STYLE_COORDS,
            STYLE_ELEVATION,
            STYLE_SPACER,
            STYLE_CODE,
            STYLE_DATE,
        ]

    def test_the_collector_is_prefixed_with_leg(self):
        lines = render_label_lines(_label(collector="F. Mensa"), LabelConfig())

        assert lines[-1].text == "leg. F. Mensa"
        assert lines[-1].style == STYLE_ATTRIBUTION

    def test_species_is_italicised(self):
        lines = render_label_lines(_label(species="Myrmica sp."), LabelConfig())

        assert lines[0].text == "Myrmica sp."
        assert lines[0].italic is True

    def test_empty_new_fields_add_no_lines(self):
        assert len(render_label_lines(_label(), LabelConfig())) == 5


class TestDeterminationLabel:
    def test_carries_the_taxon_and_the_determiner(self):
        label = _label(species="Myrmica sp.", determiner="F. Mensa", render_as="determination")

        lines = render_label_lines(label, LabelConfig())

        assert lines[0].text == "Myrmica sp."
        assert lines[0].italic is True
        assert lines[1].text == "det. F. Mensa"

    def test_leaves_a_det_line_to_fill_in_by_hand(self):
        """An undetermined specimen still gets a label to write on."""
        label = _label(species="Myrmica sp.", render_as="determination")

        assert render_label_lines(label, LabelConfig())[1].text == "det."

    def test_repeats_the_code_so_a_loose_label_can_be_matched_back(self):
        label = _label(species="Myrmica sp.", render_as="determination")

        assert any(line.text == "N1" for line in render_label_lines(label, LabelConfig()))


class TestBuildSheet:
    def test_locality_only_is_the_default_shape(self):
        sheet = build_sheet([_label()], KIND_LOCALITY)

        assert len(sheet) == 1
        assert sheet[0].render_as == KIND_LOCALITY

    def test_both_pairs_each_specimen(self):
        sheet = build_sheet([_label(code="N1"), _label(code="N2")], KIND_BOTH)

        assert [(item.code, item.render_as) for item in sheet] == [
            ("N1", KIND_LOCALITY),
            ("N1", KIND_DETERMINATION),
            ("N2", KIND_LOCALITY),
            ("N2", KIND_DETERMINATION),
        ]

    def test_both_does_not_print_the_taxon_twice_on_one_pin(self):
        sheet = build_sheet([_label(species="Myrmica sp.")], KIND_BOTH)

        assert sheet[0].species == ""
        assert sheet[1].species == "Myrmica sp."

    def test_the_original_labels_are_not_modified(self):
        original = _label(species="Myrmica sp.")

        build_sheet([original], KIND_BOTH)

        assert original.species == "Myrmica sp."
        assert original.render_as == KIND_LOCALITY

    def test_an_unknown_kind_is_rejected(self):
        with pytest.raises(ValueError, match="label kind"):
            build_sheet([_label()], "sideways")
