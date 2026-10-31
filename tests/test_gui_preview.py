"""Tests for the visual preview's geometry.

The preview exists to be checked before card stock goes through a printer, so
it has to draw a label at the size that label will actually be printed at.
"""

import os

import pytest

from entomology_labels.config import PREVIEW_SCALE_FACTOR
from entomology_labels.label_generator import Label, LabelConfig
from entomology_labels.presets import get_preset

pytest.importorskip("tkinter", reason="tkinter is not installed in this environment")

from entomology_labels.gui import _preview_font_px  # noqa: E402


#: Rendered line height of a small font, near enough for the sizing logic.
def _fake_measure(px: int) -> int:
    return px + 2


class TestPreviewFontSize:
    """Sizing logic, exercised with injected metrics so it needs no display."""

    def test_steps_down_until_a_line_fits_the_printed_height(self):
        config = LabelConfig(font_size_pt=6.0)

        px = _preview_font_px(config, PREVIEW_SCALE_FACTOR, measure=_fake_measure)

        target = config.font_size_pt * config.line_spacing * (25.4 / 72.0) * PREVIEW_SCALE_FACTOR
        assert _fake_measure(px) <= target

    def test_a_smaller_label_font_gives_a_smaller_preview_font(self):
        big = _preview_font_px(LabelConfig(font_size_pt=6.0), 3.5, measure=_fake_measure)
        small = _preview_font_px(LabelConfig(font_size_pt=4.0), 3.5, measure=_fake_measure)

        assert small < big

    def test_a_larger_scale_gives_a_larger_preview_font(self):
        near = _preview_font_px(LabelConfig(), 7.0, measure=_fake_measure)
        far = _preview_font_px(LabelConfig(), 3.5, measure=_fake_measure)

        assert near > far

    def test_never_returns_a_size_below_one(self):
        """A tiny label must still produce a usable font size, not 0 or -1."""
        px = _preview_font_px(LabelConfig(font_size_pt=1.0), 0.1, measure=_fake_measure)

        assert px >= 1

    def test_line_spacing_is_taken_into_account(self):
        tight = _preview_font_px(LabelConfig(line_spacing=1.0), 3.5, measure=_fake_measure)
        loose = _preview_font_px(LabelConfig(line_spacing=2.0), 3.5, measure=_fake_measure)

        assert loose >= tight

    def test_a_font_server_failure_does_not_raise(self):
        import tkinter as tk

        def broken(px):
            raise tk.TclError("no font server")

        assert _preview_font_px(LabelConfig(), 3.5, measure=broken) >= 1


needs_display = pytest.mark.skipif(
    not os.environ.get("DISPLAY"),
    reason="needs an X display; run under xvfb-run",
)


@needs_display
class TestPreviewFrameGeometry:
    """What the frames actually measure once Tk has laid them out."""

    @staticmethod
    def _frame_sizes(config):
        from entomology_labels.gui import EntomologyLabelsGUI

        app = EntomologyLabelsGUI()
        try:
            app.generator.config = config
            app.generator.add_labels([Label(location_line1="Norway, Vestland,", code="N1")] * 12)
            app._update_preview()
            app.root.update_idletasks()
            app.root.update()
            frames = [w for w in app.preview_sheet.winfo_children() if w.winfo_class() == "Frame"]
            return [(f.winfo_width(), f.winfo_height()) for f in frames]
        finally:
            app.root.destroy()

    def test_every_label_is_drawn_at_the_configured_size(self):
        config = LabelConfig()
        sizes = self._frame_sizes(config)

        expected = (
            round(config.label_width_mm * PREVIEW_SCALE_FACTOR),
            round(config.label_height_mm * PREVIEW_SCALE_FACTOR),
        )
        assert sizes, "the preview drew nothing"
        assert set(sizes) == {expected}

    def test_the_page_has_exactly_one_label_size(self):
        """Frames sized by their text make the grid ragged and the columns drift."""
        assert len(set(self._frame_sizes(LabelConfig()))) == 1

    def test_a_wide_flat_label_is_not_drawn_square(self):
        """29x13mm is 2.2:1; the old preview drew it at about 1.1:1."""
        config = LabelConfig(label_width_mm=29.0, label_height_mm=13.0)
        width, height = self._frame_sizes(config)[0]

        assert width / height == pytest.approx(29.0 / 13.0, rel=0.1)

    def test_a_smaller_preset_draws_smaller_boxes(self):
        default = self._frame_sizes(LabelConfig())[0]
        museum = self._frame_sizes(LabelConfig(**get_preset("museum")))[0]

        assert museum[0] < default[0]
        assert museum[1] < default[1]


@needs_display
class TestGuiBehaviour:
    """Form, list and configuration handlers, driven without a person."""

    @pytest.fixture
    def app(self, monkeypatch):
        from tkinter import messagebox

        from entomology_labels.gui import EntomologyLabelsGUI

        self.messages = []
        for name in ("showerror", "showwarning", "showinfo"):
            monkeypatch.setattr(
                messagebox, name, lambda *a, _n=name, **k: self.messages.append((_n, a))
            )
        app = EntomologyLabelsGUI()
        yield app
        app.root.destroy()

    def _full_label(self):
        return Label(
            location_line1="Norway,",
            code="N1",
            species="Myrmica sp.",
            collector="F. Mensa",
            coordinates="60.39N 5.32E",
            render_as="determination",
        )

    def test_duplicate_keeps_every_field(self, app):
        app.generator.add_labels([self._full_label()])
        app._update_labels_tree()
        app.labels_tree.selection_set("0")
        app._duplicate_selected_label()
        copy = app.generator.labels[1]
        assert (copy.species, copy.collector, copy.render_as) == (
            "Myrmica sp.",
            "F. Mensa",
            "determination",
        )

    def test_editing_keeps_fields_the_form_does_not_show(self, app):
        app.generator.add_labels([self._full_label()])
        app._update_labels_tree()
        app.labels_tree.selection_set("0")
        app._edit_selected_label()
        app.entry_vars["code"].set("N2")
        app._add_label()
        assert len(app.generator.labels) == 1
        edited = app.generator.labels[0]
        assert (edited.code, edited.collector, edited.coordinates) == (
            "N2",
            "F. Mensa",
            "60.39N 5.32E",
        )

    @pytest.mark.parametrize("quantity", ["abc", "0", "5000000"])
    def test_a_bad_quantity_adds_nothing(self, app, quantity):
        app.entry_vars["code"].set("N1")
        app.entry_vars["quantity"].set(quantity)
        app._add_label()
        assert app.generator.labels == []
        assert self.messages and self.messages[-1][0] == "showerror"

    def test_applying_the_form_keeps_settings_it_has_no_field_for(self, app):
        from dataclasses import replace

        app.generator.config = replace(
            app.generator.config, text_overflow="clip", spacer_line=False, label_padding_mm=0.4
        )
        app.config_vars["page_width_mm"].set("210")
        app.config_vars["page_height_mm"].set("297")
        app.config_vars["labels_per_row"].set("7")
        assert app._config_form_differs()
        assert app._apply_config()
        config = app.generator.config
        assert (config.text_overflow, config.spacer_line, config.label_padding_mm) == (
            "clip",
            False,
            0.4,
        )
        assert config.orientation == "portrait"
        assert not app._config_form_differs()

    def test_preview_wraps_like_the_printed_page(self, app):
        from dataclasses import replace

        app.generator.config = replace(
            app.generator.config, margin_left_mm=10.0, margin_right_mm=10.0
        )
        app.generator.add_labels([Label(code=f"N{i}") for i in range(12)])
        app._update_preview()
        app.root.update_idletasks()
        columns = {w.grid_info()["column"] for w in app.preview_sheet.winfo_children()}
        assert max(columns) + 1 == 9
        margin_px = int(10.0 * PREVIEW_SCALE_FACTOR)
        assert str(app.preview_sheet.grid_info()["padx"]) in (
            str(margin_px),
            f"{margin_px} {margin_px}",
        )
