"""Tests for locality suggestions."""

import json
from unittest import mock

import pytest

from entomology_labels.geocoding import (
    OSM_ATTRIBUTION,
    USER_AGENT,
    CachingGeocoder,
    GeocodeUnavailable,
    GeoNamesGeocoder,
    NominatimGeocoder,
    PlaceSuggestion,
    build_geocoder,
    suggest_for_events,
)
from entomology_labels.photos import CollectionEvent, PhotoRecord

# A Nominatim reply of the shape returned for a hillside above Bergen.
NOMINATIM_BERGEN = {
    "display_name": "Fløyen, Bergen, Vestland, Norway",
    "address": {
        "natural": "Fløyen",
        "city": "Bergen",
        "municipality": "Bergen",
        "state": "Vestland",
        "country": "Norway",
        "country_code": "no",
    },
}

GEONAMES_BERGEN = {
    "geonames": [{"name": "Fløyen", "adminName1": "Vestland", "countryName": "Norway", "fcl": "T"}]
}


def _urlopen_returning(payload):
    """Patch urlopen to return a canned JSON body."""
    response = mock.MagicMock()
    response.read.return_value = json.dumps(payload).encode("utf-8")
    response.__enter__.return_value = response
    return mock.patch("urllib.request.urlopen", return_value=response)


def _urlopen_raising(exc):
    return mock.patch("urllib.request.urlopen", side_effect=exc)


class TestNominatim:
    def test_builds_both_locality_lines(self):
        with _urlopen_returning(NOMINATIM_BERGEN):
            suggestion = NominatimGeocoder().reverse(60.3964, 5.3531)

        assert suggestion.location_line1 == "Norway, Vestland,"
        assert suggestion.location_line2 == "Bergen, Fløyen"

    def test_identifies_itself_as_the_usage_policy_requires(self):
        with _urlopen_returning(NOMINATIM_BERGEN) as opener:
            NominatimGeocoder().reverse(60.3964, 5.3531)

        request = opener.call_args[0][0]
        assert request.get_header("User-agent") == USER_AGENT

    def test_an_empty_address_yields_no_suggestion(self):
        with _urlopen_returning({"address": {}}):
            assert NominatimGeocoder().reverse(0.0, 0.0) is None

    def test_a_network_failure_is_distinguishable_from_no_answer(self):
        """An outage must not be mistaken for the service having nothing."""
        with _urlopen_raising(OSError("no route to host")):
            with pytest.raises(GeocodeUnavailable):
                NominatimGeocoder().reverse(60.0, 5.0)

    def test_unparseable_data_is_treated_as_an_outage(self):
        response = mock.MagicMock()
        response.read.return_value = b"<html>not json</html>"
        response.__enter__.return_value = response
        with mock.patch("urllib.request.urlopen", return_value=response):
            with pytest.raises(GeocodeUnavailable):
                NominatimGeocoder().reverse(60.0, 5.0)

    def test_hostile_text_in_a_response_is_sanitized(self):
        payload = {"address": {"country": "Norway\x00\x07", "city": "B" * 5000}}
        with _urlopen_returning(payload):
            suggestion = NominatimGeocoder().reverse(60.0, 5.0)

        assert "\x00" not in suggestion.location_line1
        assert len(suggestion.location_line2) <= 200


class TestGeoNames:
    def test_builds_a_suggestion(self):
        with _urlopen_returning(GEONAMES_BERGEN):
            suggestion = GeoNamesGeocoder("someuser").reverse(60.3964, 5.3531)

        assert suggestion.location_line1 == "Norway, Vestland,"
        assert suggestion.location_line2 == "Fløyen"

    def test_a_username_is_required(self):
        with pytest.raises(ValueError, match="username"):
            GeoNamesGeocoder("")

    def test_no_results_yields_no_suggestion(self):
        with _urlopen_returning({"geonames": []}):
            assert GeoNamesGeocoder("someuser").reverse(0.0, 0.0) is None


class TestCaching:
    def test_a_repeated_lookup_does_not_hit_the_service(self, tmp_path):
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.attribution = "stub"
        inner.reverse.return_value = PlaceSuggestion("Norway,", "Bergen", "stub")
        geocoder = CachingGeocoder(inner, tmp_path / "cache.json")

        geocoder.reverse(60.3964, 5.3531)
        geocoder.reverse(60.3964, 5.3531)

        assert inner.reverse.call_count == 1

    def test_the_cache_survives_a_new_process(self, tmp_path):
        cache = tmp_path / "cache.json"
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.reverse.return_value = PlaceSuggestion("Norway,", "Bergen", "stub")
        CachingGeocoder(inner, cache).reverse(60.3964, 5.3531)

        second = mock.MagicMock()
        second.name = "stub"
        second.cache_fingerprint = "aaaa1111"
        result = CachingGeocoder(second, cache).reverse(60.3964, 5.3531)

        assert second.reverse.call_count == 0
        assert result.location_line2 == "Bergen"

    def test_a_negative_result_is_cached_too(self, tmp_path):
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.reverse.return_value = None
        geocoder = CachingGeocoder(inner, tmp_path / "cache.json")

        assert geocoder.reverse(0.0, 0.0) is None
        assert geocoder.reverse(0.0, 0.0) is None
        assert inner.reverse.call_count == 1

    def test_a_corrupt_cache_is_ignored_rather_than_fatal(self, tmp_path):
        cache = tmp_path / "cache.json"
        cache.write_text("{not json", encoding="utf-8")
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.reverse.return_value = PlaceSuggestion("Norway,", "Bergen", "stub")

        assert CachingGeocoder(inner, cache).reverse(60.0, 5.0) is not None


class TestBuildGeocoder:
    def test_nominatim_needs_nothing(self, tmp_path):
        assert build_geocoder("nominatim", cache_path=tmp_path / "c.json").name == "nominatim"

    def test_credits_openstreetmap(self, tmp_path):
        geocoder = build_geocoder("nominatim", cache_path=tmp_path / "c.json")

        assert geocoder.attribution == OSM_ATTRIBUTION

    def test_an_unknown_name_is_rejected(self):
        with pytest.raises(ValueError, match="unknown geocoder"):
            build_geocoder("googlemaps")


class TestSuggestForEvents:
    def test_one_lookup_per_site_not_per_photo(self):
        """Forty frames from one morning must be a single request."""
        event = CollectionEvent(
            "S1",
            [PhotoRecord(f"p{i}.ORF", latitude=60.3964, longitude=5.3531) for i in range(40)],
        )
        geocoder = mock.MagicMock()
        geocoder.reverse.return_value = PlaceSuggestion("Norway,", "Bergen", "stub")

        suggest_for_events([event], geocoder)

        assert geocoder.reverse.call_count == 1

    def test_an_event_without_a_position_is_skipped(self):
        event = CollectionEvent("S1", [PhotoRecord("p.ORF")])
        geocoder = mock.MagicMock()

        results = suggest_for_events([event], geocoder)

        assert results == [("S1", None)]
        assert geocoder.reverse.call_count == 0


class TestOutagesAreNotCached:
    """An outage cached as 'no result' would mean never asking again."""

    def test_a_failure_is_reported_as_no_suggestion(self, tmp_path):
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.reverse.side_effect = GeocodeUnavailable("network down")

        assert CachingGeocoder(inner, tmp_path / "cache.json").reverse(60.0, 5.0) is None

    def test_a_later_attempt_still_reaches_the_service(self, tmp_path):
        cache = tmp_path / "cache.json"
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.reverse.side_effect = [
            GeocodeUnavailable("network down"),
            PlaceSuggestion("Norway,", "Bergen", "stub"),
        ]
        geocoder = CachingGeocoder(inner, cache)

        assert geocoder.reverse(60.0, 5.0) is None
        assert geocoder.reverse(60.0, 5.0).location_line2 == "Bergen"

    def test_nothing_is_written_to_the_cache_file(self, tmp_path):
        cache = tmp_path / "cache.json"
        inner = mock.MagicMock()
        inner.name = "stub"
        inner.cache_fingerprint = "aaaa1111"
        inner.reverse.side_effect = GeocodeUnavailable("network down")

        CachingGeocoder(inner, cache).reverse(60.0, 5.0)

        assert not cache.exists()

    def test_a_scan_continues_when_the_service_is_down(self):
        event = CollectionEvent("S1", [PhotoRecord("p.ORF", latitude=60.0, longitude=5.0)])
        geocoder = mock.MagicMock()
        geocoder.reverse.side_effect = GeocodeUnavailable("network down")

        assert suggest_for_events([event], geocoder) == [("S1", None)]


# Response shapes observed in the field for two real collecting sites above
# Bergen. Both produced a wrong or useless suggestion before the key ordering
# was corrected.
FLOYEN_HILLSIDE = {
    "address": {
        "suburb": "Skuteviken",  # port district, sea level, 1.7 km west
        "neighbourhood": "Skansemyren",  # the hillside itself
        "city": "Bergen",
        "municipality": "Bergen",
        "state": "Vestland",
        "country": "Norway",
    }
}

ULRIKEN_SLOPE = {
    "address": {
        "locality": "Ulriken",  # named but unpopulated: the usual mountain case
        "city": "Bergen",
        "municipality": "Bergen",
        "state": "Vestland",
        "country": "Norway",
    }
}


class TestPlaceKeyPriority:
    """Regressions from a field run against live Nominatim."""

    def test_the_hillside_beats_the_port_district(self):
        with _urlopen_returning(FLOYEN_HILLSIDE):
            suggestion = NominatimGeocoder().reverse(60.3964, 5.3531)

        assert suggestion.location_line2 == "Bergen, Skansemyren"
        assert "Skuteviken" not in suggestion.location_line2

    def test_a_named_unpopulated_place_beats_the_municipality(self):
        """Without `locality` this returned the bare city for a 3 m site."""
        with _urlopen_returning(ULRIKEN_SLOPE):
            suggestion = NominatimGeocoder().reverse(60.3733, 5.3908)

        assert suggestion.location_line2 == "Bergen, Ulriken"

    def test_a_site_with_no_finer_feature_still_names_the_settlement(self):
        with _urlopen_returning({"address": {"city": "Bergen", "country": "Norway"}}):
            suggestion = NominatimGeocoder().reverse(60.0, 5.0)

        assert suggestion.location_line2 == "Bergen"

    def test_the_settlement_is_never_repeated(self):
        """Settlement keys used to sit in both lists and print twice."""
        with _urlopen_returning({"address": {"village": "Bondone", "country": "Italy"}}):
            suggestion = NominatimGeocoder().reverse(45.9, 10.9)

        assert suggestion.location_line2 == "Bondone"

    def test_rural_keys_are_recognised(self):
        with _urlopen_returning({"address": {"farm": "Malga Tovel", "municipality": "Ville"}}):
            suggestion = NominatimGeocoder().reverse(46.2, 10.9)

        assert suggestion.location_line2 == "Ville, Malga Tovel"

    def test_a_natural_feature_wins_when_present(self):
        payload = {
            "address": {
                "natural": "Vedretta d'Amola",
                "suburb": "somewhere",
                "municipality": "Giustino",
            }
        }
        with _urlopen_returning(payload):
            suggestion = NominatimGeocoder().reverse(46.1, 10.6)

        assert suggestion.location_line2 == "Giustino, Vedretta d'Amola"

    def test_place_and_settlement_keys_do_not_overlap(self):
        """The overlap is what made a settlement resolve to itself and vanish."""
        from entomology_labels.geocoding import _PLACE_KEYS, _SETTLEMENT_KEYS

        assert not set(_PLACE_KEYS) & set(_SETTLEMENT_KEYS)


class TestReverseZoom:
    def test_the_request_asks_for_feature_level_detail(self):
        """Zoom 14 returns the municipality for a rural site."""
        from entomology_labels.geocoding import _REVERSE_ZOOM

        with _urlopen_returning(ULRIKEN_SLOPE) as opener:
            NominatimGeocoder().reverse(60.3733, 5.3908)

        assert f"zoom={_REVERSE_ZOOM}" in opener.call_args[0][0].full_url
        assert _REVERSE_ZOOM >= 16


class TestCacheInvalidation:
    """A cached answer must not outlive the settings that produced it.

    The cache stores the derived suggestion, not the raw response, so both the
    reverse-lookup zoom and the address keys the answer is read from decide
    what a hit means. Neither was in the key, so a live run after the key
    ordering changed silently returned the previous answers.
    """

    @staticmethod
    def _stub(fingerprint, suggestion=None):
        inner = mock.MagicMock()
        inner.name = "nominatim"
        inner.cache_fingerprint = fingerprint
        inner.reverse.return_value = suggestion or PlaceSuggestion("Norway,", "Bergen", "stub")
        return inner

    def test_an_entry_from_older_settings_is_not_served(self, tmp_path):
        """The exact failure seen in the field, in miniature."""
        cache = tmp_path / "cache.json"
        old = self._stub("old00000", PlaceSuggestion("Norway,", "Bergen, Skuteviken", "stub"))
        CachingGeocoder(old, cache).reverse(60.3964, 5.3531)

        new = self._stub("new11111", PlaceSuggestion("Norway,", "Bergen, Skansemyren", "stub"))
        result = CachingGeocoder(new, cache).reverse(60.3964, 5.3531)

        assert new.reverse.call_count == 1, "should have re-queried, not served the old answer"
        assert result.location_line2 == "Bergen, Skansemyren"

    def test_superseded_entries_are_dropped_from_the_file(self, tmp_path):
        cache = tmp_path / "cache.json"
        CachingGeocoder(self._stub("old00000"), cache).reverse(60.0, 5.0)

        CachingGeocoder(self._stub("new11111"), cache).reverse(60.0, 5.0)

        keys = json.loads(cache.read_text(encoding="utf-8"))
        assert not any("old00000" in key for key in keys)

    def test_another_backend_is_left_alone(self, tmp_path):
        """One file holds both backends; a blanket sweep would delete the other."""
        cache = tmp_path / "cache.json"
        cache.write_text(
            json.dumps(
                {
                    "geonames:ffff9999:60.0,5.0": {
                        "location_line1": "Norway,",
                        "location_line2": "Ulriken",
                        "source": "geonames",
                    }
                }
            ),
            encoding="utf-8",
        )

        CachingGeocoder(self._stub("new11111"), cache).reverse(60.0, 5.0)

        assert "geonames:ffff9999:60.0,5.0" in json.loads(cache.read_text(encoding="utf-8"))

    def test_the_same_settings_still_hit_the_cache(self, tmp_path):
        """Invalidation must not defeat caching for an unchanged configuration."""
        cache = tmp_path / "cache.json"
        CachingGeocoder(self._stub("same0000"), cache).reverse(60.0, 5.0)

        second = self._stub("same0000")
        CachingGeocoder(second, cache).reverse(60.0, 5.0)

        assert second.reverse.call_count == 0


class TestFingerprint:
    """What the fingerprint is derived from, rather than its literal value."""

    def test_the_zoom_is_part_of_it(self, monkeypatch):
        before = NominatimGeocoder().cache_fingerprint
        monkeypatch.setattr("entomology_labels.geocoding._REVERSE_ZOOM", 14)

        assert NominatimGeocoder().cache_fingerprint != before

    def test_the_place_key_order_is_part_of_it(self, monkeypatch):
        """Reordering changes the answer without changing the request."""
        from entomology_labels.geocoding import _PLACE_KEYS

        before = NominatimGeocoder().cache_fingerprint
        monkeypatch.setattr("entomology_labels.geocoding._PLACE_KEYS", tuple(reversed(_PLACE_KEYS)))

        assert NominatimGeocoder().cache_fingerprint != before

    def test_it_is_stable_across_instances(self):
        assert NominatimGeocoder().cache_fingerprint == NominatimGeocoder().cache_fingerprint

    def test_the_backends_do_not_collide(self):
        assert NominatimGeocoder().cache_fingerprint != GeoNamesGeocoder("u").cache_fingerprint

    def test_it_is_short_enough_to_read_in_a_key(self):
        assert len(NominatimGeocoder().cache_fingerprint) == 8
