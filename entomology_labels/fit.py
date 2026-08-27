"""
Checks that label content will actually fit on the page it is printed on.

Both failures this catches are silent: an overlong locality is trimmed with an
ellipsis by the browser, and a grid wider than the page simply reflows. Either
way the first sign of trouble is a printed sheet, by which point the paper and
the toner are already spent.
"""

import logging
from dataclasses import dataclass
from typing import TYPE_CHECKING, List, Optional, Sequence

from .config import LABEL_PADDING_MM, TEXT_WIDTH_WARN_RATIO
from .layout import STYLE_SPACER, render_label_lines

if TYPE_CHECKING:  # pragma: no cover - typing only
    from .label_generator import Label, LabelConfig, LabelGenerator

logger = logging.getLogger(__name__)

#: Millimetres per point.
MM_PER_POINT = 25.4 / 72.0

#: Mean glyph advance as a fraction of the font size, for mixed-case Latin text
#: in the sans-serif faces this tool defaults to. Measuring exactly would mean
#: parsing font files; that is a large dependency to sharpen a warning, so this
#: errs slightly wide and the messages say "may be" rather than "will be".
AVG_CHAR_WIDTH_RATIO = 0.55

#: Tolerance for floating-point noise when comparing millimetre dimensions.
DIMENSION_EPSILON_MM = 0.5

KIND_GRID_OVERFLOW = "grid_overflow"
KIND_TEXT_OVERFLOW = "text_overflow"
KIND_ORIENTATION_MISMATCH = "orientation_mismatch"
KIND_LABEL_HEIGHT = "label_height"


@dataclass(frozen=True)
class FitWarning:
    """One way in which the configured layout may not print correctly."""

    kind: str
    message: str
    label_index: Optional[int] = None

    def __str__(self) -> str:
        if self.label_index is None:
            return self.message
        return f"label {self.label_index + 1}: {self.message}"


def estimate_text_width_mm(text: str, font_size_pt: float) -> float:
    """Estimate the rendered width of a single line of text.

    Args:
        text: The text to measure
        font_size_pt: Font size in points

    Returns:
        Approximate width in millimetres
    """
    return len(text) * font_size_pt * AVG_CHAR_WIDTH_RATIO * MM_PER_POINT


def usable_label_width_mm(config: "LabelConfig") -> float:
    """Return the width available for text inside a label, after padding."""
    return max(0.0, config.label_width_mm - 2 * LABEL_PADDING_MM)


def estimate_label_height_mm(label: "Label", config: "LabelConfig") -> float:
    """Estimate how tall a label's body will render, padding included.

    Args:
        label: The label to measure
        config: Layout configuration

    Returns:
        Approximate height in millimetres
    """
    total_pt = sum(
        config.font_size_pt * line.scale * config.line_spacing
        for line in render_label_lines(label, config)
    )
    return total_pt * MM_PER_POINT + 2 * LABEL_PADDING_MM


def check_page_fit(config: "LabelConfig") -> List[FitWarning]:
    """Check the label grid against the page it is meant to print on.

    Unlike the text check this is exact arithmetic, so these warnings are
    always real.

    Args:
        config: Layout configuration to check

    Returns:
        Warnings, most serious first; empty when the layout fits
    """
    warnings: List[FitWarning] = []

    needed_width = (
        config.labels_per_row * config.label_width_mm
        + config.margin_left_mm
        + config.margin_right_mm
    )
    if needed_width > config.page_width_mm + DIMENSION_EPSILON_MM:
        warnings.append(
            FitWarning(
                KIND_GRID_OVERFLOW,
                f"{config.labels_per_row} labels of {config.label_width_mm}mm plus margins "
                f"need {needed_width:.1f}mm, but the page is only "
                f"{config.page_width_mm:.1f}mm wide; columns will wrap",
            )
        )

    needed_height = (
        config.labels_per_column * config.label_height_mm
        + config.margin_top_mm
        + config.margin_bottom_mm
    )
    if needed_height > config.page_height_mm + DIMENSION_EPSILON_MM:
        warnings.append(
            FitWarning(
                KIND_GRID_OVERFLOW,
                f"{config.labels_per_column} labels of {config.label_height_mm}mm plus margins "
                f"need {needed_height:.1f}mm, but the page is only "
                f"{config.page_height_mm:.1f}mm tall; rows will be cut off",
            )
        )

    page_is_landscape = config.page_width_mm > config.page_height_mm
    if page_is_landscape != (config.orientation == "landscape"):
        warnings.append(
            FitWarning(
                KIND_ORIENTATION_MISMATCH,
                f"orientation is '{config.orientation}' but the page is "
                f"{config.page_width_mm:.0f}x{config.page_height_mm:.0f}mm, which is "
                f"{'landscape' if page_is_landscape else 'portrait'}",
            )
        )

    return warnings


def check_label_fit(
    labels: Sequence["Label"],
    config: "LabelConfig",
    *,
    max_reported: int = 10,
) -> List[FitWarning]:
    """Check whether any label's text is too wide for the label.

    Args:
        labels: The labels to check
        config: Layout configuration
        max_reported: Stop after this many distinct warnings

    Returns:
        Warnings about lines likely to be clipped
    """
    available = usable_label_width_mm(config)
    if available <= 0:
        return []

    warnings: List[FitWarning] = []
    seen = set()
    reported_height = False

    for index, label in enumerate(labels):
        if label is None or label.is_empty():
            continue

        # Vertical overflow is the failure that bites hardest: an extra line
        # pushes the label past its box and the browser simply cuts it off,
        # so the collector's note disappears from the printed sheet.
        height = estimate_label_height_mm(label, config)
        if not reported_height and height > config.label_height_mm + DIMENSION_EPSILON_MM:
            reported_height = True
            line_count = len(render_label_lines(label, config))
            warnings.append(
                FitWarning(
                    KIND_LABEL_HEIGHT,
                    f"{line_count} lines at {config.font_size_pt}pt need about "
                    f"{height:.1f}mm, but the label is only "
                    f"{config.label_height_mm:.1f}mm tall; lower lines will be cut off",
                    label_index=index,
                )
            )
            if len(warnings) >= max_reported:
                return warnings

        for line in render_label_lines(label, config):
            if line.style == STYLE_SPACER or not line.text:
                continue
            width = estimate_text_width_mm(line.text, config.font_size_pt * line.scale)
            if width <= available * TEXT_WIDTH_WARN_RATIO:
                continue
            if line.text in seen:
                continue
            seen.add(line.text)
            verb = "will likely be" if width > available else "may be"
            warnings.append(
                FitWarning(
                    KIND_TEXT_OVERFLOW,
                    f'"{line.text}" is about {width:.1f}mm wide and {verb} clipped '
                    f"at {available:.1f}mm",
                    label_index=index,
                )
            )
            if len(warnings) >= max_reported:
                return warnings

    return warnings


def check_fit(generator: "LabelGenerator", *, max_reported: int = 10) -> List[FitWarning]:
    """Run every fit check against a generator's configuration and labels.

    Args:
        generator: The generator about to produce output
        max_reported: Cap on the number of text warnings reported

    Returns:
        All warnings, page-level ones first
    """
    warnings = check_page_fit(generator.config)
    warnings.extend(check_label_fit(generator.labels, generator.config, max_reported=max_reported))
    return warnings
