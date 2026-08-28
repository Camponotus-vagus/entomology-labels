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
            frames = [w for w in app.paper_frame.winfo_children() if w.winfo_class() == "Frame"]
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
