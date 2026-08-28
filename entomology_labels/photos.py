"""
Collection dates and coordinates read from field photographs.

A camera that geotags its shots already records, for every specimen
photographed in the field, the two things hardest to reconstruct afterwards:
when it was collected and where. Retyping them from a notebook is both slower
and less accurate than reading the file.

This module reads that metadata, groups photographs into collection events,
and writes a draft site registry. It deliberately stops there: EXIF cannot
supply a place name, a specimen code or a determination, so a person always
edits the draft before labels are printed.
"""

import logging
import math
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

from .config import (
    ELEVATION_M_MAX,
    ELEVATION_M_MIN,
    LATITUDE_MAX,
    LATITUDE_MIN,
    LONGITUDE_MAX,
    LONGITUDE_MIN,
    MAX_EXIF_STRING_LEN,
    MAX_PHOTO_SIZE_BYTES,
    MAX_PHOTOS_PER_SCAN,
    PHOTO_EXTENSIONS,
)
from .dates import to_roman_date
from .label_generator import _sanitize_string

logger = logging.getLogger(__name__)

#: Mean Earth radius in metres, for the haversine distance.
EARTH_RADIUS_M = 6371000.0

#: A new collection event starts after this long without a photograph.
DEFAULT_TIME_GAP_MINUTES = 90.0

#: A new collection event starts beyond this distance from the running centre.
DEFAULT_RADIUS_M = 250.0

#: EXIF timestamps outside this range are a broken tag, not a collection date.
MIN_PHOTO_YEAR = 1990
MAX_PHOTO_YEAR = 2200

_EXIF_DATE_TAGS = ("EXIF DateTimeOriginal", "EXIF DateTimeDigitized", "Image DateTime")


class PhotoReadError(Exception):
    """Raised when a file cannot be read as a photograph with metadata."""


@dataclass(frozen=True)
class PhotoRecord:
    """What one photograph contributes to a collection event."""

    source: str
    taken_at: Optional[datetime] = None
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    elevation_m: Optional[float] = None
    camera: str = ""

    @property
    def has_position(self) -> bool:
        return self.latitude is not None and self.longitude is not None


@dataclass
class CollectionEvent:
    """A group of photographs taken at one place on one occasion."""

    event_id: str
    photos: List[PhotoRecord] = field(default_factory=list)

    @property
    def dated_photos(self) -> List[PhotoRecord]:
        return [p for p in self.photos if p.taken_at is not None]

    @property
    def positioned_photos(self) -> List[PhotoRecord]:
        return [p for p in self.photos if p.has_position]

    @property
    def date(self) -> Optional[datetime]:
        """The earliest timestamp in the group."""
        dated = self.dated_photos
        return min(p.taken_at for p in dated) if dated else None

    @property
    def centroid(self) -> Optional[Tuple[float, float]]:
        """Mean position of the photographs that have one."""
        located = self.positioned_photos
        if not located:
            return None
        return (
            sum(p.latitude for p in located) / len(located),
            sum(p.longitude for p in located) / len(located),
        )

    @property
    def spread_m(self) -> float:
        """Distance from the centre to the furthest photograph."""
        centre = self.centroid
        if centre is None:
            return 0.0
        return max(
            (
                haversine_m(centre[0], centre[1], p.latitude, p.longitude)
                for p in self.positioned_photos
            ),
            default=0.0,
        )

    @property
    def elevation_m(self) -> Optional[float]:
        """Median elevation, which resists the drift a barometer shows."""
        values = sorted(p.elevation_m for p in self.photos if p.elevation_m is not None)
        if not values:
            return None
        middle = len(values) // 2
        if len(values) % 2:
            return values[middle]
        return (values[middle - 1] + values[middle]) / 2

    @property
    def elevation_range_m(self) -> Optional[Tuple[float, float]]:
        """Lowest and highest elevation recorded, or None."""
        values = [p.elevation_m for p in self.photos if p.elevation_m is not None]
        return (min(values), max(values)) if values else None


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Great-circle distance between two points, in metres.

    Args:
        lat1: First latitude in decimal degrees
        lon1: First longitude in decimal degrees
        lat2: Second latitude in decimal degrees
        lon2: Second longitude in decimal degrees

    Returns:
        Distance in metres
    """
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    return 2 * EARTH_RADIUS_M * math.asin(math.sqrt(a))


def _ratio_to_float(value: Any) -> float:
    """Convert an EXIF rational to a float.

    Args:
        value: An exifread Ratio, or anything float() accepts

    Returns:
        The value as a float

    Raises:
        ValueError: If it cannot be converted, including a zero denominator
    """
    numerator = getattr(value, "numerator", None)
    denominator = getattr(value, "denominator", None)
    if numerator is not None and denominator is not None:
        if denominator == 0:
            raise ValueError("rational with a zero denominator")
        return float(numerator) / float(denominator)
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValueError(f"not a number: {value!r}")


def dms_to_decimal(dms: Sequence[Any], ref: str) -> float:
    """Convert EXIF degrees/minutes/seconds plus a hemisphere ref to degrees.

    Args:
        dms: Three rationals: degrees, minutes, seconds
        ref: "N", "S", "E" or "W"

    Returns:
        Signed decimal degrees, negative for south and west

    Raises:
        ValueError: If the components are unusable
    """
    if len(dms) != 3:
        raise ValueError(f"expected three components, got {len(dms)}")

    degrees, minutes, seconds = (_ratio_to_float(part) for part in dms)
    # Exactly 60 is allowed. Writers that encode a decimal degree as rationals
    # land on it through binary rounding -- 41.8 degrees becomes 41 deg 47' 60"
    # because 0.8 * 60 is 47.999... -- and the arithmetic below carries it into
    # the next minute correctly. Anything above 60 is genuinely malformed.
    if not (0 <= minutes <= 60 and 0 <= seconds <= 60):
        raise ValueError(f"minutes/seconds out of range: {minutes}, {seconds}")

    value = degrees + minutes / 60.0 + seconds / 3600.0
    if str(ref).strip().upper().startswith(("S", "W")):
        value = -value
    return value


def _clean_string(value: Any) -> str:
    """Sanitize and cap a free-text EXIF value.

    Camera-written strings reach a YAML file and then a printed label, and
    nothing stops a crafted file carrying megabytes of control characters.
    """
    text = _sanitize_string(str(value)) if value is not None else ""
    return text[:MAX_EXIF_STRING_LEN].strip()


def _parse_exif_datetime(raw: str) -> Optional[datetime]:
    """Parse an EXIF timestamp, rejecting implausible years."""
    for fmt in ("%Y:%m:%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y:%m:%d"):
        try:
            parsed = datetime.strptime(raw.strip(), fmt)
        except ValueError:
            continue
        if MIN_PHOTO_YEAR <= parsed.year <= MAX_PHOTO_YEAR:
            return parsed
        return None
    return None


def _bounded(value: float, low: float, high: float, name: str) -> Optional[float]:
    """Return the value if it is in range, else None with a warning."""
    if low <= value <= high:
        return value
    logger.warning(f"ignoring out-of-range {name}: {value}")
    return None


def _tags_to_photo(tags: Mapping[str, Any], source: str) -> PhotoRecord:
    """Build a PhotoRecord from a tag mapping.

    Kept separate from any file handling so that the conversion, bounds
    checks and hemisphere handling can be tested against plain dictionaries.

    Args:
        tags: EXIF tags, as exifread returns them
        source: Name to record for this photograph

    Returns:
        The record; fields the tags do not support are left as None
    """
    taken_at = None
    for tag in _EXIF_DATE_TAGS:
        if tag in tags:
            taken_at = _parse_exif_datetime(str(tags[tag]))
            if taken_at is not None:
                break

    latitude = longitude = elevation = None

    if "GPS GPSLatitude" in tags and "GPS GPSLongitude" in tags:
        try:
            latitude = dms_to_decimal(
                tags["GPS GPSLatitude"].values, str(tags.get("GPS GPSLatitudeRef", "N"))
            )
            longitude = dms_to_decimal(
                tags["GPS GPSLongitude"].values, str(tags.get("GPS GPSLongitudeRef", "E"))
            )
        except (ValueError, AttributeError, TypeError) as e:
            logger.warning(f"{source}: unusable GPS coordinates ({e})")
            latitude = longitude = None
        else:
            latitude = _bounded(latitude, LATITUDE_MIN, LATITUDE_MAX, "latitude")
            longitude = _bounded(longitude, LONGITUDE_MIN, LONGITUDE_MAX, "longitude")
            if latitude is None or longitude is None:
                latitude = longitude = None

    if "GPS GPSAltitude" in tags:
        try:
            values = tags["GPS GPSAltitude"].values
            raw = values[0] if isinstance(values, (list, tuple)) else values
            elevation = _ratio_to_float(raw)
            # A ref of 1 means the altitude is measured below sea level.
            if str(tags.get("GPS GPSAltitudeRef", "0")).strip() in ("1", "Below Sea Level"):
                elevation = -elevation
            elevation = _bounded(elevation, ELEVATION_M_MIN, ELEVATION_M_MAX, "elevation")
        except (ValueError, AttributeError, TypeError, IndexError) as e:
            logger.warning(f"{source}: unusable GPS altitude ({e})")
            elevation = None

    return PhotoRecord(
        source=source,
        taken_at=taken_at,
        latitude=latitude,
        longitude=longitude,
        elevation_m=elevation,
        camera=_clean_string(tags.get("Image Model", "")),
    )


def read_photo_tags(path: Path) -> Dict[str, Any]:
    """Read a photograph's EXIF tags.

    Only metadata is parsed; no pixel data is decoded, so the image codecs
    are never invoked. MakerNote and thumbnail parsing are skipped as well --
    they are the bulk of the parser's attack surface and carry nothing this
    tool needs.

    Args:
        path: Path to the photograph

    Returns:
        The tags found, possibly empty

    Raises:
        PhotoReadError: If the file is too large or cannot be parsed
    """
    try:
        import exifread
    except ImportError:
        raise PhotoReadError(
            "exifread is required to read photo metadata. "
            "Install with: pip install 'entomology-labels[photos]'"
        )

    size = path.stat().st_size
    if size > MAX_PHOTO_SIZE_BYTES:
        raise PhotoReadError(f"file is too large ({size} bytes): {path}")
    if size == 0:
        raise PhotoReadError(f"file is empty: {path}")

    try:
        with path.open("rb") as handle:
            return exifread.process_file(handle, details=False, truncate_tags=True)
    except Exception as e:
        # A malformed file must not abort a scan of hundreds of others.
        raise PhotoReadError(f"could not read metadata from {path.name}: {e}")


def read_photo_meta(path: Path, *, full_path: bool = False) -> PhotoRecord:
    """Read one photograph's collection metadata.

    Args:
        path: Path to the photograph
        full_path: Record the whole path rather than just the file name

    Returns:
        What the file's metadata supports

    Raises:
        PhotoReadError: If the file cannot be read
    """
    tags = read_photo_tags(path)
    record = _tags_to_photo(tags, str(path) if full_path else path.name)

    # exifread reports an unrecognised file by returning nothing rather than
    # raising. A record with neither a date nor a position contributes
    # nothing to a collection event, so treat it as unreadable and let the
    # caller count it as skipped instead of silently carrying it along.
    if record.taken_at is None and not record.has_position:
        raise PhotoReadError(f"no usable date or position in {path.name}")

    return record


def find_photos(roots: Sequence[Path], *, limit: int = MAX_PHOTOS_PER_SCAN) -> List[Path]:
    """Collect photograph paths under the given files and directories.

    Symlinks are not followed, so a loop cannot turn a scan into an
    unbounded walk.

    Args:
        roots: Files or directories to scan
        limit: Stop after this many files

    Returns:
        Photograph paths, sorted by name so a scan is reproducible
    """
    found: List[Path] = []

    for root in roots:
        root = Path(root)
        if root.is_file():
            if root.suffix.lower() in PHOTO_EXTENSIONS:
                found.append(root)
            continue

        for directory, subdirs, names in os.walk(root, followlinks=False):
            subdirs.sort()
            for name in sorted(names):
                if Path(name).suffix.lower() in PHOTO_EXTENSIONS:
                    found.append(Path(directory) / name)
                    if len(found) >= limit:
                        logger.warning(f"stopping at {limit} photos; some were not scanned")
                        return sorted(found)

    return sorted(found)


def read_photos(paths: Sequence[Path], *, full_path: bool = False) -> Tuple[List[PhotoRecord], int]:
    """Read metadata from many photographs, skipping ones that fail.

    Args:
        paths: Photograph paths
        full_path: Record whole paths rather than file names

    Returns:
        The records read, and how many files were skipped
    """
    records: List[PhotoRecord] = []
    skipped = 0

    for path in paths:
        try:
            records.append(read_photo_meta(path, full_path=full_path))
        except (PhotoReadError, OSError) as e:
            logger.warning(str(e))
            skipped += 1

    return records, skipped


def cluster_photos(
    photos: Sequence[PhotoRecord],
    *,
    time_gap_minutes: float = DEFAULT_TIME_GAP_MINUTES,
    radius_m: float = DEFAULT_RADIUS_M,
) -> List[CollectionEvent]:
    """Group photographs into collection events.

    Photographs are taken in order of time; a new event begins when the gap
    since the previous photograph is too long, or when the camera has moved
    too far from the running centre of the current event. A photograph with
    no position joins whatever event it falls between in time, which matters
    in practice: a receiver that has not yet acquired a fix still produces
    usable frames.

    Args:
        photos: The photographs to group
        time_gap_minutes: Gap that starts a new event
        radius_m: Distance from the event centre that starts a new event

    Returns:
        Events in chronological order
    """
    dated = sorted((p for p in photos if p.taken_at is not None), key=lambda p: p.taken_at)
    undated = [p for p in photos if p.taken_at is None]

    events: List[CollectionEvent] = []
    current: Optional[CollectionEvent] = None
    previous_time: Optional[datetime] = None

    for photo in dated:
        start_new = current is None

        if not start_new:
            gap = (photo.taken_at - previous_time).total_seconds() / 60.0
            if gap > time_gap_minutes:
                start_new = True
            elif photo.has_position:
                centre = current.centroid
                if centre is not None:
                    distance = haversine_m(centre[0], centre[1], photo.latitude, photo.longitude)
                    if distance > radius_m:
                        start_new = True

        if start_new:
            current = CollectionEvent(event_id=f"S{len(events) + 1}")
            events.append(current)

        current.photos.append(photo)
        previous_time = photo.taken_at

    # Undated photographs cannot be placed in time; attach them to the first
    # event rather than discarding them, and say so in the draft.
    if undated and events:
        events[0].photos.extend(undated)
    elif undated:
        events.append(CollectionEvent(event_id="S1", photos=list(undated)))

    for index, event in enumerate(events, start=1):
        event.event_id = _event_id(event, index)

    return events


def _event_id(event: CollectionEvent, index: int) -> str:
    """Build a readable, stable id for an event."""
    when = event.date
    if when is None:
        return f"SITE{index}"
    return f"S{when:%Y%m%d}{chr(ord('A') + (index - 1) % 26)}"


def scan_photos(
    roots: Sequence[Path],
    *,
    time_gap_minutes: float = DEFAULT_TIME_GAP_MINUTES,
    radius_m: float = DEFAULT_RADIUS_M,
    full_path: bool = False,
) -> Tuple[List[CollectionEvent], int, int]:
    """Scan photographs and group them into collection events.

    Args:
        roots: Files or directories to scan
        time_gap_minutes: Gap that starts a new event
        radius_m: Distance that starts a new event
        full_path: Record whole paths rather than file names

    Returns:
        The events found, how many photographs were read, and how many were skipped
    """
    paths = find_photos(roots)
    records, skipped = read_photos(paths, full_path=full_path)
    events = cluster_photos(records, time_gap_minutes=time_gap_minutes, radius_m=radius_m)
    return events, len(records), skipped


def event_to_site(event: CollectionEvent, *, decimals: int = 4) -> Dict[str, Any]:
    """Render an event as a site registry entry.

    The locality lines are left empty on purpose: EXIF gives a position, not
    a place name, and inventing one would put a guess on a museum label.

    Args:
        event: The collection event
        decimals: Decimal places for the coordinates

    Returns:
        A mapping suitable for the site registry
    """
    site: Dict[str, Any] = {"location_line1": "", "location_line2": ""}

    centre = event.centroid
    if centre is not None:
        site["coordinates"] = {
            "lat": round(centre[0], decimals),
            "lon": round(centre[1], decimals),
        }

    elevation = event.elevation_m
    if elevation is not None:
        site["elevation_m"] = round(elevation)

    when = event.date
    if when is not None:
        site["date"] = to_roman_date(when)

    return site
