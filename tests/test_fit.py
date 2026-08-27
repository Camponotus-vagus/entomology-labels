"""Tests for the layout fit checks."""

from entomology_labels.fit import (
    KIND_GRID_OVERFLOW,
    KIND_LABEL_HEIGHT,
    KIND_ORIENTATION_MISMATCH,
    KIND_TEXT_OVERFLOW,
    check_fit,
    check_label_fit,
    check_page_fit,
    estimate_label_height_mm,
    estimate_text_width_mm,
    usable_label_width_mm,
)
from entomology_labels.label_generator import Label, LabelConfig, LabelGenerator


def _kinds(warnings):
    return {w.kind for w in warnings}


class TestPageFit:
    """Grid arithmetic is exact, so these warnings are never false alarms."""

    def test_a_layout_that_fits_reports_nothing(self):
        config = LabelConfig(
            labels_per_row=10, labels_per_column=13, label_width_mm=29.0, label_height_mm=13.0
        )

        assert check_page_fit(config) == []

    def test_too_many_columns_for_the_page_width(self):
        config = LabelConfig(labels_per_row=20, label_width_mm=29.0, page_width_mm=297.0)

        warnings = check_page_fit(config)

        assert KIND_GRID_OVERFLOW in _kinds(warnings)
        assert "580.0mm" in warnings[0].message

    def test_too_many_rows_for_the_page_height(self):
        config = LabelConfig(labels_per_column=40, label_height_mm=13.0, page_height_mm=210.0)

        assert KIND_GRID_OVERFLOW in _kinds(check_page_fit(config))

    def test_margins_count_against_the_available_width(self):
        fits = LabelConfig(labels_per_row=10, label_width_mm=29.0, page_width_mm=297.0)
        assert check_page_fit(fits) == []

        with_margins = LabelConfig(
            labels_per_row=10,
            label_width_mm=29.0,
            page_width_mm=297.0,
            margin_left_mm=10.0,
            margin_right_mm=10.0,
        )
        assert KIND_GRID_OVERFLOW in _kinds(check_page_fit(with_margins))

    def test_orientation_disagreeing_with_page_dimensions(self):
        config = LabelConfig(page_width_mm=297.0, page_height_mm=210.0, orientation="portrait")

        assert KIND_ORIENTATION_MISMATCH in _kinds(check_page_fit(config))


class TestTextFit:
    """Width is estimated, so the wording hedges but the check still earns its keep."""

    def test_short_text_on_a_wide_label_is_fine(self):
        config = LabelConfig(label_width_mm=29.0, font_size_pt=6.0)

        assert check_label_fit([Label(code="N1")], config) == []

    def test_an_overlong_locality_is_reported(self):
        config = LabelConfig(label_width_mm=29.0, font_size_pt=6.0)
        label = Label(location_line2="Giustino (TN), Vedretta d'Amola")

        warnings = check_label_fit([label], config)

        assert KIND_TEXT_OVERFLOW in _kinds(warnings)
        assert "Vedretta d'Amola" in warnings[0].message

    def test_the_same_text_is_only_reported_once(self):
        """Duplicated labels would otherwise produce one warning per copy."""
        config = LabelConfig(label_width_mm=15.0, font_size_pt=6.0)
        labels = [Label(location_line1="Norway, Vestland, Bergen") for _ in range(20)]

        warnings = check_label_fit(labels, config)

        assert len([w for w in warnings if w.kind == KIND_TEXT_OVERFLOW]) == 1

    def test_reporting_is_capped(self):
        config = LabelConfig(label_width_mm=10.0, font_size_pt=6.0)
        labels = [Label(location_line1=f"A very long locality number {i}") for i in range(50)]

        assert len(check_label_fit(labels, config, max_reported=3)) == 3

    def test_padding_is_subtracted_from_the_usable_width(self):
        assert usable_label_width_mm(LabelConfig(label_width_mm=29.0)) == 27.0

    def test_width_estimate_scales_with_length_and_size(self):
        assert estimate_text_width_mm("", 6.0) == 0.0
        assert estimate_text_width_mm("AAAA", 6.0) > estimate_text_width_mm("AA", 6.0)
        assert estimate_text_width_mm("AAAA", 12.0) > estimate_text_width_mm("AAAA", 6.0)


class TestLabelHeight:
    """The failure that loses data quietly: an extra line falls off the label."""

    def test_five_lines_fit_the_default_label(self):
        config = LabelConfig(label_height_mm=13.0, font_size_pt=6.0)

        assert check_label_fit([Label(code="N1", date="20.viii.2026")], config) == []

    def test_an_additional_info_line_can_overflow_a_short_label(self):
        config = LabelConfig(label_height_mm=13.0, font_size_pt=6.0)
        label = Label(code="N1", date="20.viii.2026", additional_info="under bark")

        warnings = check_label_fit([label], config)

        assert KIND_LABEL_HEIGHT in _kinds(warnings)

    def test_height_estimate_grows_with_an_extra_line(self):
        config = LabelConfig(font_size_pt=6.0)
        without = estimate_label_height_mm(Label(code="N1"), config)
        with_info = estimate_label_height_mm(Label(code="N1", additional_info="x"), config)

        assert with_info > without

    def test_height_is_only_reported_once_per_run(self):
        config = LabelConfig(label_height_mm=13.0, font_size_pt=6.0)
        labels = [Label(code=f"N{i}", additional_info="under bark") for i in range(10)]

        warnings = [w for w in check_label_fit(labels, config) if w.kind == KIND_LABEL_HEIGHT]

        assert len(warnings) == 1


class TestCheckFit:
    def test_combines_page_and_label_checks(self):
        config = LabelConfig(labels_per_row=20, label_width_mm=29.0, page_width_mm=297.0)
        generator = LabelGenerator(config)
        generator.add_label(Label(location_line2="Giustino (TN), Vedretta d'Amola"))

        kinds = _kinds(check_fit(generator))

        assert KIND_GRID_OVERFLOW in kinds
        assert KIND_TEXT_OVERFLOW in kinds

    def test_a_sound_layout_is_silent(self):
        generator = LabelGenerator(
            LabelConfig(
                label_width_mm=45.0, label_height_mm=18.0, labels_per_row=6, labels_per_column=11
            )
        )
        generator.add_label(Label(location_line1="Norway,", code="N1", date="20.viii.2026"))

        assert check_fit(generator) == []
