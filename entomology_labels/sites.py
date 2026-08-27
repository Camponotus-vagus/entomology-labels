"""
Collection sites, declared once and referenced by many specimen rows.

A locality is a property of the place, not of the beetle. Repeating
"Norway, Vestland, / Bergen, Fløyen" on all forty rows from one morning is
both tedious and fragile: correcting the elevation later means editing forty
rows and hoping none was missed.

A file may instead declare sites under a ``sites:`` key and have rows refer to
them by id. Files without that key are untouched by any of this and load
exactly as before.
"""

import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Mapping, Optional

from .config import (
    ELEVATION_M_MAX,
    ELEVATION_M_MIN,
    LATITUDE_MAX,
    LATITUDE_MIN,
    LONGITUDE_MAX,
    LONGITUDE_MIN,
    MAX_SITES,
    SITE_ID_PATTERN,
)
from .dates import normalize_date
from .label_generator import Label, _sanitize_string

logger = logging.getLogger(__name__)

_SITE_ID_RE = re.compile(SITE_ID_PATTERN)

#: Site fields that map straight onto a Label field of the same name.
_PASSTHROUGH_FIELDS = ("location_line1", "location_line2", "date", "collector")


def _coerce_float(value: Any, name: str, low: float, high: float) -> Optional[float]:
    """Parse and bounds-check an optional numeric site field.

    Args:
        value: The raw value from the file
        name: Field name, for the error message
        low: Smallest permitted value
        high: Largest permitted value

    Returns:
        The value as a float, or None if it was absent

    Raises:
        ValueError: If the value is not a number or is out of range
    """
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{name} must be a number, got {value!r}")
    if not (low <= number <= high):
        raise ValueError(f"{name} must be between {low} and {high}, got {number}")
    return number


@dataclass
class Site:
    """One place where specimens were collected."""

    site_id: str
    location_line1: str = ""
    location_line2: str = ""
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    elevation_m: Optional[float] = None
    date: str = ""
    collector: str = ""
    habitat: str = ""

    def __post_init__(self):
        if not _SITE_ID_RE.match(self.site_id):
            raise ValueError(
                f"site id {self.site_id!r} must match {SITE_ID_PATTERN} "
                f"(letters, digits, dash and underscore)"
            )
        for name in ("location_line1", "location_line2", "date", "collector", "habitat"):
            setattr(self, name, _sanitize_string(getattr(self, name)))

        self.latitude = _coerce_float(self.latitude, "latitude", LATITUDE_MIN, LATITUDE_MAX)
        self.longitude = _coerce_float(self.longitude, "longitude", LONGITUDE_MIN, LONGITUDE_MAX)
        self.elevation_m = _coerce_float(
            self.elevation_m, "elevation_m", ELEVATION_M_MIN, ELEVATION_M_MAX
        )

        if self.date:
            self.date = normalize_date(self.date)

    def format_coordinates(self) -> str:
        """Return the coordinates as they should print, or an empty string.

        Four decimal places is about 11 m, which is honest for a handheld
        GPS; more would claim a precision the reading does not have.
        """
        if self.latitude is None or self.longitude is None:
            return ""
        lat_ref = "N" if self.latitude >= 0 else "S"
        lon_ref = "E" if self.longitude >= 0 else "W"
        return f"{abs(self.latitude):.4f}{lat_ref} {abs(self.longitude):.4f}{lon_ref}"

    def format_elevation(self) -> str:
        """Return the elevation as it should print, or an empty string."""
        if self.elevation_m is None:
            return ""
        return f"{round(self.elevation_m / 10) * 10:.0f} m"

    def to_label_fields(self) -> Dict[str, str]:
        """Return this site's contribution to a label."""
        values = {name: getattr(self, name) for name in _PASSTHROUGH_FIELDS}
        values["coordinates"] = self.format_coordinates()
        values["elevation"] = self.format_elevation()
        return {k: v for k, v in values.items() if v}


@dataclass
class SiteFile:
    """A parsed input file that declares sites and defaults."""

    sites: Dict[str, Site] = field(default_factory=dict)
    defaults: Dict[str, Any] = field(default_factory=dict)


def has_sites(data: Any) -> bool:
    """Return whether a loaded file uses the site-registry form.

    Presence of a ``sites:`` key is the discriminator. Requiring an explicit
    version key instead would break every file already in use, for nothing.

    Args:
        data: The parsed contents of an input file

    Returns:
        True if the file declares sites or defaults
    """
    return isinstance(data, dict) and bool(data.get("sites") or data.get("defaults"))


def parse_sites(raw: Any) -> Dict[str, Site]:
    """Parse a ``sites:`` block into Site records.

    Args:
        raw: The value of the file's ``sites`` key

    Returns:
        Sites by id; empty when the block is absent

    Raises:
        ValueError: If the block is malformed, too large, or a site is invalid
    """
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ValueError("'sites' must be a mapping of site id to site details")
    if len(raw) > MAX_SITES:
        raise ValueError(f"too many sites ({len(raw)}); the limit is {MAX_SITES}")

    sites: Dict[str, Site] = {}
    for site_id, details in raw.items():
        if not isinstance(details, dict):
            raise ValueError(f"site {site_id!r} must be a mapping of field to value")

        values = dict(details)
        # Coordinates may be nested, which reads better in YAML.
        coordinates = values.pop("coordinates", None)
        if isinstance(coordinates, dict):
            values.setdefault("latitude", coordinates.get("lat", coordinates.get("latitude")))
            values.setdefault("longitude", coordinates.get("lon", coordinates.get("longitude")))

        known = {k: v for k, v in values.items() if k in Site.__dataclass_fields__}
        unknown = set(values) - set(known)
        if unknown:
            logger.warning(f"site {site_id!r}: ignoring unknown fields {sorted(unknown)}")

        try:
            sites[str(site_id)] = Site(site_id=str(site_id), **known)
        except ValueError as e:
            raise ValueError(f"site {site_id!r}: {e}")

    return sites


def resolve_row(
    row: Mapping[str, Any],
    sites: Mapping[str, Site],
    defaults: Optional[Mapping[str, Any]] = None,
) -> Dict[str, Any]:
    """Merge one label row with its site and the file-level defaults.

    Precedence is row, then site, then defaults: a row may always override
    what its site says, which is what makes "same place, different day"
    expressible without declaring a second site.

    Args:
        row: The label row as written in the file
        sites: Sites declared by the file
        defaults: File-level default values

    Returns:
        The row's fields with site and default values filled in

    Raises:
        ValueError: If the row names a site that was not declared
    """
    merged: Dict[str, Any] = dict(defaults or {})

    site_id = row.get("site", row.get("site_id"))
    if site_id is not None:
        site = sites.get(str(site_id))
        if site is None:
            known = ", ".join(sorted(sites)) or "none declared"
            raise ValueError(f"unknown site id {str(site_id)!r} (known sites: {known})")
        merged.update(site.to_label_fields())

    for key, value in row.items():
        if key in ("site", "site_id") or value is None or value == "":
            continue
        merged[key] = value

    return merged


def labels_from_rows(
    rows: Any,
    sites: Mapping[str, Site],
    defaults: Optional[Mapping[str, Any]] = None,
) -> list:
    """Build labels from rows, resolving each against its site.

    Args:
        rows: The file's label rows
        sites: Sites declared by the file
        defaults: File-level default values

    Returns:
        One Label per row, before any count expansion

    Raises:
        ValueError: If a row is malformed or names an undeclared site
    """
    labels = []
    for index, row in enumerate(rows or []):
        if not isinstance(row, dict):
            raise ValueError(f"label {index + 1} must be a mapping of field to value")
        try:
            labels.append(Label.from_dict(resolve_row(row, sites, defaults)))
        except ValueError as e:
            raise ValueError(f"label {index + 1}: {e}")
    return labels


def load_site_registry(path) -> Dict[str, Site]:
    """Load a standalone site registry from a YAML or JSON file.

    CSV and Excel have nowhere to put a sites block, so a tabular file
    carries a `site` column and the sites themselves live here.

    Args:
        path: Path to the registry file

    Returns:
        Sites by id

    Raises:
        ValueError: If the file cannot be read or declares no sites
    """
    from pathlib import Path as _Path

    registry_path = _Path(path).expanduser()
    text = registry_path.read_text(encoding="utf-8")

    if registry_path.suffix.lower() in (".yaml", ".yml"):
        try:
            import yaml
        except ImportError:
            raise ValueError("PyYAML is required to read a YAML site registry")
        data = yaml.safe_load(text)
    else:
        import json

        data = json.loads(text)

    if not isinstance(data, dict):
        raise ValueError(f"{registry_path} must contain a mapping with a 'sites' key")

    # Accept either a file with a sites: block or a bare mapping of sites.
    raw = data.get("sites", data)
    sites = parse_sites(raw)
    if not sites:
        raise ValueError(f"{registry_path} declares no sites")
    return sites
