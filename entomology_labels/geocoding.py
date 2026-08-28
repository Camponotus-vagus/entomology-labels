"""
Suggesting locality names for coordinates.

EXIF gives a position, not a place name, and the place name is the one thing
on a collection label that cannot be derived from the photograph. A reverse
geocoder can propose one.

Three things shape how this is used, and all three are deliberate:

It is opt-in. Looking a site up transmits a collecting locality to a third
party, and precise localities for rare or protected species are often withheld
on purpose. That is the collector's decision to make each time, not a default
to discover afterwards.

It suggests, never fills. A proposal is written into the draft as a comment
beside the field, which stays empty. A wrong toponym silently inherited onto a
museum label is worse than a blank one somebody notices.

It is per site, not per photograph. Forty frames from one morning are one
lookup, which is ordinary use of a free service rather than bulk querying.
"""

import json
import logging
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from .config import (
    GEOCODE_CACHE_PRECISION,
    GEOCODE_MIN_INTERVAL_S,
    GEOCODE_TIMEOUT_S,
    MAX_EXIF_STRING_LEN,
)
from .label_generator import _sanitize_string

logger = logging.getLogger(__name__)


class GeocodeUnavailable(Exception):
    """Raised when the service could not be reached or replied unusably.

    Distinct from a service that answered and had nothing: an outage must
    not be cached, or a later attempt would never be made.
    """


#: Identifies this tool to the geocoding service, as its usage policy requires.
USER_AGENT = "entomology-labels/1.2 (+https://github.com/Camponotus-vagus/entomology-labels)"

#: Attribution that must be shown when OSM data is used.
OSM_ATTRIBUTION = "Locality suggestions © OpenStreetMap contributors (ODbL)"

NOMINATIM_URL = "https://nominatim.openstreetmap.org/reverse"
GEONAMES_URL = "http://api.geonames.org/findNearbyJSON"

#: Address components to try, in order, for the country/region line.
_REGION_KEYS = ("state", "province", "region", "county", "state_district")

#: Address components naming the settlement a site sits in.
_SETTLEMENT_KEYS = ("municipality", "city", "town", "village")

#: Address components naming a place *within* a settlement, most specific
#: first. Settlement names are deliberately absent: they are read separately,
#: and including them here meant that a site with no finer feature resolved to
#: the settlement twice and printed as the bare municipality.
#:
#: The order is from field results rather than intuition. `suburb` used to
#: precede `neighbourhood`, which picked the port district "Skuteviken" for a
#: site 300 m up the Floyen hillside while the same response carried
#: `neighbourhood: Skansemyren`. `locality` was missing entirely, which is the
#: key Nominatim uses for named but unpopulated places -- the usual case for a
#: mountain collecting site. `natural` and `peak` are rare in practice (absent
#: from twelve queries across Norway and the Alps) but are the most specific
#: names when they do appear.
_PLACE_KEYS = (
    "natural",
    "peak",
    "farm",
    "isolated_dwelling",
    "locality",
    "hamlet",
    "neighbourhood",
    "quarter",
    "suburb",
    "city_district",
    "park",
)

#: Nominatim zoom for reverse lookups. 14 returns the whole municipality for a
#: rural site; 16 resolves the named feature within it.
_REVERSE_ZOOM = 16


@dataclass(frozen=True)
class PlaceSuggestion:
    """A proposed locality for a set of coordinates."""

    location_line1: str = ""
    location_line2: str = ""
    source: str = ""

    def is_empty(self) -> bool:
        return not (self.location_line1 or self.location_line2)


def _clean(value: Any) -> str:
    """Sanitize and cap a value returned by a remote service."""
    if value is None:
        return ""
    return _sanitize_string(str(value))[:MAX_EXIF_STRING_LEN].strip()


class _RateLimiter:
    """Enforces a minimum interval between requests."""

    def __init__(self, min_interval_s: float):
        self._min_interval_s = min_interval_s
        self._last_call = 0.0

    def wait(self) -> None:
        elapsed = time.monotonic() - self._last_call
        if 0 < elapsed < self._min_interval_s:
            time.sleep(self._min_interval_s - elapsed)
        self._last_call = time.monotonic()


class NominatimGeocoder:
    """Reverse geocoding through OpenStreetMap's Nominatim service.

    Free and needs no API key. Its usage policy requires an identifying
    User-Agent, at most one request per second, and that OSM be credited
    wherever the results are shown; all three are honoured here.
    """

    name = "nominatim"
    attribution = OSM_ATTRIBUTION

    def __init__(self, *, timeout_s: float = GEOCODE_TIMEOUT_S):
        self._timeout_s = timeout_s
        self._limiter = _RateLimiter(GEOCODE_MIN_INTERVAL_S)

    def reverse(self, latitude: float, longitude: float) -> Optional[PlaceSuggestion]:
        """Look up a place name for a position.

        Args:
            latitude: Decimal degrees
            longitude: Decimal degrees

        Returns:
            A suggestion, or None if the service had nothing or was unreachable
        """
        query = urllib.parse.urlencode(
            {
                "format": "jsonv2",
                "lat": f"{latitude:.5f}",
                "lon": f"{longitude:.5f}",
                "zoom": str(_REVERSE_ZOOM),
                "addressdetails": "1",
            }
        )
        payload = _fetch_json(f"{NOMINATIM_URL}?{query}", self._timeout_s, self._limiter)
        if not payload:
            return None

        address = payload.get("address") or {}
        country = _clean(address.get("country"))
        region = next((_clean(address[k]) for k in _REGION_KEYS if address.get(k)), "")
        place = next((_clean(address[k]) for k in _PLACE_KEYS if address.get(k)), "")
        settlement = next((_clean(address[k]) for k in _SETTLEMENT_KEYS if address.get(k)), "")

        line1 = ", ".join(part for part in (country, region) if part)
        if line1:
            line1 += ","
        line2 = ", ".join(dict.fromkeys(part for part in (settlement, place) if part))

        suggestion = PlaceSuggestion(line1, line2, source=self.name)
        return None if suggestion.is_empty() else suggestion


class GeoNamesGeocoder:
    """Reverse geocoding through GeoNames.

    Needs a free registered username. Often better than Nominatim on named
    natural features, which is what an alpine or forest collecting site
    usually is.
    """

    name = "geonames"
    attribution = "Locality suggestions from GeoNames (CC BY)"

    def __init__(self, username: str, *, timeout_s: float = GEOCODE_TIMEOUT_S):
        if not username:
            raise ValueError("GeoNames requires a username; register free at geonames.org")
        self._username = username
        self._timeout_s = timeout_s
        self._limiter = _RateLimiter(GEOCODE_MIN_INTERVAL_S)

    def reverse(self, latitude: float, longitude: float) -> Optional[PlaceSuggestion]:
        """Look up a place name for a position."""
        query = urllib.parse.urlencode(
            {
                "lat": f"{latitude:.5f}",
                "lng": f"{longitude:.5f}",
                "username": self._username,
                "style": "FULL",
            }
        )
        payload = _fetch_json(f"{GEONAMES_URL}?{query}", self._timeout_s, self._limiter)
        if not payload:
            return None

        entries = payload.get("geonames") or []
        if not entries:
            return None

        entry = entries[0]
        country = _clean(entry.get("countryName"))
        region = _clean(entry.get("adminName1"))
        place = _clean(entry.get("name"))

        line1 = ", ".join(part for part in (country, region) if part)
        if line1:
            line1 += ","

        suggestion = PlaceSuggestion(line1, place, source=self.name)
        return None if suggestion.is_empty() else suggestion


def _fetch_json(url: str, timeout_s: float, limiter: _RateLimiter) -> Optional[Dict[str, Any]]:
    """Fetch and decode a JSON response, returning None on any failure.

    A lookup is a convenience: nothing about a failed one should stop a scan
    that has already read the photographs. Transport failures are raised
    rather than returned so that an outage is never mistaken for -- and
    cached as -- an answer.

    Args:
        url: The request URL
        timeout_s: Socket timeout
        limiter: Rate limiter to respect before the call

    Returns:
        The decoded response, or None if the service answered with nothing

    Raises:
        GeocodeUnavailable: If the service could not be reached or replied
            unusably, which the caller must not cache
    """
    limiter.wait()
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

    try:
        with urllib.request.urlopen(request, timeout=timeout_s) as response:
            body = response.read()
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, ValueError) as e:
        raise GeocodeUnavailable(f"could not reach the locality service: {e}")

    try:
        payload = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        raise GeocodeUnavailable(f"the locality service returned unusable data: {e}")

    return payload if isinstance(payload, dict) else None


class CachingGeocoder:
    """Wraps a geocoder with an on-disk cache keyed by rounded coordinates.

    Re-running a scan is common while a draft is being edited, and it should
    not mean querying a free service again for answers already received.
    """

    def __init__(self, inner, cache_path: Optional[Path] = None):
        self._inner = inner
        self._cache_path = Path(cache_path) if cache_path else None
        self._cache: Dict[str, Any] = {}
        self._load()

    @property
    def name(self) -> str:
        return self._inner.name

    @property
    def attribution(self) -> str:
        return self._inner.attribution

    def _key(self, latitude: float, longitude: float) -> str:
        return (
            f"{self._inner.name}:"
            f"{round(latitude, GEOCODE_CACHE_PRECISION)},"
            f"{round(longitude, GEOCODE_CACHE_PRECISION)}"
        )

    def _load(self) -> None:
        if not self._cache_path or not self._cache_path.is_file():
            return
        try:
            self._cache = json.loads(self._cache_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            logger.debug("locality cache unreadable; starting a new one")
            self._cache = {}

    def _save(self) -> None:
        if not self._cache_path:
            return
        try:
            self._cache_path.parent.mkdir(parents=True, exist_ok=True)
            self._cache_path.write_text(json.dumps(self._cache, indent=2), encoding="utf-8")
        except OSError as e:
            logger.debug(f"could not write the locality cache: {e}")

    def reverse(self, latitude: float, longitude: float) -> Optional[PlaceSuggestion]:
        """Look up a place name, using the cache where possible."""
        key = self._key(latitude, longitude)
        if key in self._cache:
            cached = self._cache[key]
            if cached is None:
                return None
            return PlaceSuggestion(**cached)

        try:
            suggestion = self._inner.reverse(latitude, longitude)
        except GeocodeUnavailable as e:
            # Not cached: the service may be reachable next time, and caching
            # an outage would mean never asking again.
            logger.warning(f"{e}; leaving the locality blank")
            return None

        self._cache[key] = None if suggestion is None else suggestion.__dict__.copy()
        self._save()
        return suggestion


def build_geocoder(
    name: str,
    *,
    username: str = "",
    cache_path: Optional[Path] = None,
):
    """Construct a geocoder by name.

    Args:
        name: "nominatim" or "geonames"
        username: GeoNames username, required for that backend
        cache_path: Where to cache answers

    Returns:
        A geocoder with caching applied

    Raises:
        ValueError: If the name is unknown or a required setting is missing
    """
    if name == "nominatim":
        inner = NominatimGeocoder()
    elif name == "geonames":
        inner = GeoNamesGeocoder(username)
    else:
        raise ValueError(f"unknown geocoder {name!r}; use 'nominatim' or 'geonames'")

    return CachingGeocoder(inner, cache_path)


def suggest_for_events(events, geocoder) -> List[Tuple[str, Optional[PlaceSuggestion]]]:
    """Look up a locality suggestion for each event that has a position.

    Args:
        events: Collection events
        geocoder: The geocoder to ask

    Returns:
        One (event id, suggestion) pair per event; the suggestion may be None
    """
    results = []
    for event in events:
        centre = event.centroid
        if centre is None:
            results.append((event.event_id, None))
            continue
        try:
            results.append((event.event_id, geocoder.reverse(centre[0], centre[1])))
        except GeocodeUnavailable as e:
            logger.warning(f"{e}; leaving the locality blank")
            results.append((event.event_id, None))
    return results
