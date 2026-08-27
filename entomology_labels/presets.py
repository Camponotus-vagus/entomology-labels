"""
Named label geometries.

Published guidance from entomological collections converges on a small label
in a small type size: a maximum width around 0.7 in (17.8 mm) with 0.6 in
(15.2 mm) preferred, and 4 pt lettering, occasionally up to 6.

That is considerably tighter than this tool's historical defaults, and the
difference is worth real money in card stock: a narrower label at 4 pt fits
several times as many to an A4 sheet. The old geometry is kept as a named
preset so sheets printed under it stay reproducible.

On typefaces, the sources disagree in a useful way. Some prefer a compact
serif such as Times, whose letterforms stay distinguishable when reduced;
others prefer Helvetica, on the grounds that serifs break up at very small
sizes. Both warn specifically against Arial, where several characters become
hard to tell apart, and against condensed faces, which buy width at the cost
of legibility. Where more room is needed, the advice is to abbreviate and drop
punctuation rather than squeeze the letterforms.
"""

from typing import Any, Dict, List

#: Millimetres per point.
MM_PER_POINT = 25.4 / 72.0

#: A4 in landscape, in millimetres.
A4_LANDSCAPE = (297.0, 210.0)

#: US Letter in landscape, in millimetres.
LETTER_LANDSCAPE = (279.4, 215.9)


def _preset(
    label_width_mm: float,
    lines: int,
    font_size_pt: float,
    page: tuple,
    *,
    font_family: str = "Times New Roman",
    padding_mm: float = 1.0,
    spacer_line: bool = True,
    **extra: Any,
) -> Dict[str, Any]:
    """Build a preset, sizing the label to the text it must hold.

    Height is derived rather than chosen: a label has to fit the lines it
    will carry, and at 4pt a millimetre of padding on each side is a quarter
    of the whole label. Sizing by arithmetic is what keeps the fit checker
    quiet on the tool's own presets.

    Args:
        label_width_mm: Label width
        lines: How many text lines the label must hold
        font_size_pt: Font size
        page: Page width and height in millimetres
        font_family: Typeface
        padding_mm: Padding inside each label
        spacer_line: Whether a blank line separates locality from code
        **extra: Any further LabelConfig settings

    Returns:
        Settings suitable for LabelConfig
    """
    page_width_mm, page_height_mm = page
    label_height_mm = round(lines * font_size_pt * MM_PER_POINT + 2 * padding_mm + 0.2, 1)

    values = {
        "labels_per_row": int(page_width_mm // label_width_mm),
        "labels_per_column": int(page_height_mm // label_height_mm),
        "label_width_mm": label_width_mm,
        "label_height_mm": label_height_mm,
        "page_width_mm": page_width_mm,
        "page_height_mm": page_height_mm,
        "font_size_pt": font_size_pt,
        "font_family": font_family,
        "label_padding_mm": padding_mm,
        "spacer_line": spacer_line,
        "orientation": "landscape" if page_width_mm > page_height_mm else "portrait",
    }
    values.update(extra)
    return values


PRESETS: Dict[str, Dict[str, Any]] = {
    # Five lines: two of locality, the code, the date and the collector --
    # the classic locality label the published guidance describes.
    "museum": _preset(15.2, 5, 4.0, A4_LANDSCAPE, padding_mm=0.4, spacer_line=False),
    # Eight lines, so coordinates and elevation fit alongside the rest.
    "museum-wide": _preset(17.8, 8, 4.0, A4_LANDSCAPE, padding_mm=0.4, spacer_line=False),
    "readable": _preset(24.0, 8, 5.0, A4_LANDSCAPE, padding_mm=0.6),
    "legacy": _preset(29.0, 5, 6.0, A4_LANDSCAPE, font_family="Arial", label_height_mm=13.0),
    "letter": _preset(15.2, 5, 4.0, LETTER_LANDSCAPE, padding_mm=0.4, spacer_line=False),
}

#: One line of explanation each, for `--help` and the `presets` command.
DESCRIPTIONS: Dict[str, str] = {
    "museum": "4pt, the size most collections recommend; fits a 5-line locality label",
    "museum-wide": "4pt and a little wider; room for coordinates and elevation too",
    "readable": "5pt, easier on the eye at the cost of sheet yield",
    "legacy": "6pt in Arial, this tool's original geometry",
    "letter": "as 'museum', on US Letter paper",
}


def preset_names() -> List[str]:
    """Return the available preset names."""
    return list(PRESETS)


def get_preset(name: str) -> Dict[str, Any]:
    """Return a preset's settings.

    Args:
        name: The preset name

    Returns:
        Settings suitable for LabelConfig

    Raises:
        ValueError: If the name is not a known preset
    """
    if name not in PRESETS:
        raise ValueError(f"unknown preset {name!r}; choose from {', '.join(PRESETS)}")
    return dict(PRESETS[name])


def labels_per_page(name: str) -> int:
    """Return how many labels a preset fits on one sheet."""
    preset = get_preset(name)
    return preset["labels_per_row"] * preset["labels_per_column"]
