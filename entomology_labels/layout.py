"""
Label body layout — the single source of truth for what a label contains.

The HTML writer, the DOCX writer and the GUI preview all render the same
ordered list of :class:`LabelLine` values, so a change to label composition is
made here once instead of in three places that must be kept in lockstep.
"""

from dataclasses import dataclass
from typing import TYPE_CHECKING, List

if TYPE_CHECKING:  # pragma: no cover - import cycle guard, typing only
    from .label_generator import Label, LabelConfig

# Style names a renderer may receive. Each backend maps these to its own
# formatting: HTML to a CSS class, DOCX to run properties, the GUI to a font.
STYLE_LOCATION = "location"
STYLE_SPACER = "spacer"
STYLE_CODE = "code"
STYLE_DATE = "date"
STYLE_INFO = "info"

#: Scale applied to ``font_size_pt`` for the additional-info line.
INFO_FONT_SCALE = 0.9


@dataclass(frozen=True)
class LabelLine:
    """One rendered line of a label body.

    Attributes:
        text: The line's text. May be empty for a spacer.
        style: Which kind of line this is; see the ``STYLE_*`` constants.
        italic: Whether the line is rendered in italics.
        scale: Multiplier applied to the configured font size.
    """

    text: str
    style: str = STYLE_LOCATION
    italic: bool = False
    scale: float = 1.0


def render_label_lines(label: "Label", config: "LabelConfig") -> List[LabelLine]:
    """Return the ordered text lines that make up a label body.

    The two location lines, the spacer, the code and the date are always
    emitted, even when empty, so that every label on a sheet has the same
    height and the grid stays aligned. Optional lines are emitted only when
    they carry text.

    Args:
        label: The label to lay out.
        config: Layout configuration, used for size-dependent decisions.

    Returns:
        The label body as an ordered list of lines.
    """
    lines: List[LabelLine] = [
        LabelLine(label.location_line1, STYLE_LOCATION),
        LabelLine(label.location_line2, STYLE_LOCATION),
        LabelLine("", STYLE_SPACER),
        LabelLine(label.code, STYLE_CODE),
        LabelLine(label.date, STYLE_DATE),
    ]

    if label.additional_info:
        lines.append(
            LabelLine(
                label.additional_info,
                STYLE_INFO,
                italic=True,
                scale=INFO_FONT_SCALE,
            )
        )

    return lines
