"""
Label body layout — the single source of truth for what a label contains.

The HTML writer, the DOCX writer and the GUI preview all render the same
ordered list of :class:`LabelLine` values, so a change to label composition is
made here once instead of in three places that must be kept in lockstep.
"""

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, List, Sequence

if TYPE_CHECKING:  # pragma: no cover - import cycle guard, typing only
    from .label_generator import Label, LabelConfig

# Style names a renderer may receive. Each backend maps these to its own
# formatting: HTML to a CSS class, DOCX to run properties, the GUI to a font.
STYLE_LOCATION = "location"
STYLE_SPACER = "spacer"
STYLE_CODE = "code"
STYLE_DATE = "date"
STYLE_INFO = "info"
STYLE_SPECIES = "species"
STYLE_COORDS = "coords"
STYLE_ELEVATION = "elevation"
STYLE_ATTRIBUTION = "attribution"

#: Scale applied to ``font_size_pt`` for the additional-info line.
INFO_FONT_SCALE = 0.9

# Which label a record renders as. Specimens conventionally carry two: a
# locality label, and a determination label pinned below it.
KIND_LOCALITY = "locality"
KIND_DETERMINATION = "determination"
KIND_BOTH = "both"
LABEL_KINDS = (KIND_LOCALITY, KIND_DETERMINATION, KIND_BOTH)


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

    Which lines appear depends on ``label.render_as``:

    A locality label follows the usual museum ordering -- locality,
    coordinates, elevation, then the specimen code, the date and the
    collector. A determination label carries the taxon and who identified it,
    and is pinned below the locality label on the same specimen.

    The two location lines, the spacer, the code and the date are always
    emitted even when empty, so that every label on a sheet is the same
    height and the printed grid stays aligned. The newer fields appear only
    when they carry text, which is what keeps an existing five-field label
    rendering exactly as it always has.

    Args:
        label: The label to lay out
        config: Layout configuration, for size-dependent decisions

    Returns:
        The label body as an ordered list of lines
    """
    if label.render_as == KIND_DETERMINATION:
        return _determination_lines(label)

    lines: List[LabelLine] = []

    if label.species:
        lines.append(LabelLine(label.species, STYLE_SPECIES, italic=True))

    lines.append(LabelLine(label.location_line1, STYLE_LOCATION))
    lines.append(LabelLine(label.location_line2, STYLE_LOCATION))

    if label.coordinates:
        lines.append(LabelLine(label.coordinates, STYLE_COORDS))
    if label.elevation:
        lines.append(LabelLine(label.elevation, STYLE_ELEVATION))

    lines.append(LabelLine("", STYLE_SPACER))
    lines.append(LabelLine(label.code, STYLE_CODE))
    lines.append(LabelLine(label.date, STYLE_DATE))

    if label.collector:
        lines.append(LabelLine(f"leg. {label.collector}", STYLE_ATTRIBUTION))

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


def _determination_lines(label: "Label") -> List[LabelLine]:
    """Return the body of a determination label.

    Conventionally two lines: the taxon in italics and who identified it.
    The specimen code repeats here when there is one, so that a label
    separated from its specimen can still be matched back.

    Args:
        label: The label whose determination is being rendered

    Returns:
        The determination label body
    """
    lines = [LabelLine(label.species, STYLE_SPECIES, italic=True)]

    if label.determiner:
        lines.append(LabelLine(f"det. {label.determiner}", STYLE_ATTRIBUTION))
    else:
        lines.append(LabelLine("det.", STYLE_ATTRIBUTION))

    if label.code:
        lines.append(LabelLine(label.code, STYLE_CODE))

    return lines


def build_sheet(labels: Sequence["Label"], kind: str) -> List["Label"]:
    """Return the labels to print for a given label kind.

    Args:
        labels: The specimen records
        kind: "locality", "determination", or "both"

    Returns:
        Labels tagged with how each should render. For "both", each
        specimen's pair is adjacent so the two come off the sheet together.

    Raises:
        ValueError: If the kind is not recognised
    """
    if kind not in LABEL_KINDS:
        raise ValueError(f"label kind must be one of {LABEL_KINDS}, got {kind!r}")

    if kind == KIND_LOCALITY:
        return [replace(label, render_as=KIND_LOCALITY) for label in labels]

    if kind == KIND_DETERMINATION:
        return [replace(label, render_as=KIND_DETERMINATION) for label in labels]

    # The taxon belongs on the determination label; repeating it on the
    # locality label would print the same name twice on one pin.
    sheet: List["Label"] = []
    for label in labels:
        sheet.append(replace(label, render_as=KIND_LOCALITY, species=""))
        sheet.append(replace(label, render_as=KIND_DETERMINATION))
    return sheet
