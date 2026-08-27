"""Tests for the collection site registry."""

import json
from pathlib import Path

import pytest

from entomology_labels.input_handlers import load_data
from entomology_labels.sites import (
    Site,
    has_sites,
    labels_from_rows,
    load_site_registry,
    parse_sites,
    resolve_row,
)

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"

SITES = {
    "FLOYEN": Site(
        "FLOYEN",
        location_line1="Norway, Vestland,",
        location_line2="Bergen, Fløyen",
        latitude=60.3964,
        longitude=5.3531,
        elevation_m=293,
        date="2026-08-20",
        collector="F. Mensa",
    )
}


class TestSite:
    def test_coordinates_print_to_four_decimals(self):
        """More would claim a precision a handheld GPS does not have."""
        assert SITES["FLOYEN"].format_coordinates() == "60.3964N 5.3531E"

    def test_southern_and_western_hemispheres_get_the_right_refs(self):
        site = Site("CPT", latitude=-33.9249, longitude=-18.4241)

        assert site.format_coordinates() == "33.9249S 18.4241W"

    def test_elevation_rounds_to_the_nearest_ten_metres(self):
        """Barometric altitude drifts tens of metres while standing still."""
        assert Site("X", elevation_m=293.4).format_elevation() == "290 m"
        assert Site("X", elevation_m=356.7).format_elevation() == "360 m"

    def test_missing_coordinates_produce_no_text(self):
        assert Site("X").format_coordinates() == ""
        assert Site("X", latitude=60.0).format_coordinates() == ""
        assert Site("X").format_elevation() == ""

    def test_iso_dates_become_label_dates(self):
        assert SITES["FLOYEN"].date == "20.viii.2026"

    @pytest.mark.parametrize(
        "kwargs,message",
        [
            ({"latitude": 91}, "latitude"),
            ({"latitude": -91}, "latitude"),
            ({"longitude": 181}, "longitude"),
            ({"elevation_m": 20000}, "elevation_m"),
            ({"latitude": "north"}, "must be a number"),
        ],
    )
    def test_out_of_range_values_are_rejected(self, kwargs, message):
        with pytest.raises(ValueError, match=message):
            Site("X", **kwargs)

    @pytest.mark.parametrize("site_id", ["has space", "with/slash", "", "x" * 33, "quote'"])
    def test_bad_site_ids_are_rejected(self, site_id):
        with pytest.raises(ValueError, match="site id"):
            Site(site_id)


class TestParseSites:
    def test_nested_coordinates_are_accepted(self):
        sites = parse_sites({"A": {"coordinates": {"lat": 60.0, "lon": 5.0}}})

        assert sites["A"].latitude == 60.0
        assert sites["A"].longitude == 5.0

    def test_absent_block_gives_no_sites(self):
        assert parse_sites(None) == {}

    def test_a_non_mapping_block_is_rejected(self):
        with pytest.raises(ValueError, match="must be a mapping"):
            parse_sites(["FLOYEN"])

    def test_unknown_site_fields_are_ignored_not_fatal(self):
        sites = parse_sites({"A": {"location_line1": "Norway,", "weather": "rain"}})

        assert sites["A"].location_line1 == "Norway,"


class TestResolution:
    def test_a_row_inherits_its_site(self):
        resolved = resolve_row({"site": "FLOYEN", "code": "N1"}, SITES)

        assert resolved["location_line2"] == "Bergen, Fløyen"
        assert resolved["coordinates"] == "60.3964N 5.3531E"
        assert resolved["code"] == "N1"

    def test_a_row_overrides_its_site(self):
        resolved = resolve_row({"site": "FLOYEN", "date": "26.viii.2026"}, SITES)

        assert resolved["date"] == "26.viii.2026"

    def test_a_site_overrides_the_file_defaults(self):
        resolved = resolve_row({"site": "FLOYEN"}, SITES, {"collector": "Someone Else"})

        assert resolved["collector"] == "F. Mensa"

    def test_defaults_fill_what_neither_row_nor_site_supplies(self):
        resolved = resolve_row({"site": "FLOYEN"}, SITES, {"determiner": "F. Mensa"})

        assert resolved["determiner"] == "F. Mensa"

    def test_a_row_with_no_site_still_works(self):
        resolved = resolve_row({"location_line1": "Italy,", "code": "V1"}, SITES)

        assert resolved["location_line1"] == "Italy,"

    def test_an_unknown_site_names_itself_and_the_known_ones(self):
        """A silent empty label would be far worse than an error."""
        with pytest.raises(ValueError, match="unknown site id 'NOPE'.*FLOYEN"):
            resolve_row({"site": "NOPE"}, SITES)

    def test_the_failing_row_is_identified(self):
        rows = [{"site": "FLOYEN"}, {"site": "MISSING"}]

        with pytest.raises(ValueError, match="label 2"):
            labels_from_rows(rows, SITES)


class TestDiscrimination:
    """Presence of a sites key is the discriminator, not a version field."""

    def test_a_file_with_sites_is_recognised(self):
        assert has_sites({"sites": {"A": {}}, "labels": []})

    def test_a_file_with_only_defaults_is_recognised(self):
        assert has_sites({"defaults": {"collector": "F. Mensa"}, "labels": []})

    def test_a_plain_label_file_is_not(self):
        assert not has_sites({"labels": [{"code": "N1"}]})

    def test_a_bare_list_is_not(self):
        assert not has_sites([{"code": "N1"}])


class TestBackwardCompatibility:
    def test_existing_flat_files_are_unaffected(self, tmp_path):
        flat = tmp_path / "labels.json"
        flat.write_text(
            json.dumps({"labels": [{"location_line1": "Norway,", "code": "N1", "count": 3}]}),
            encoding="utf-8",
        )

        labels = load_data(flat)

        assert len(labels) == 3
        assert labels[0].location_line1 == "Norway,"
        assert labels[0].coordinates == ""

    def test_the_shipped_example_loads(self):
        labels = load_data(EXAMPLES / "example_sites.yaml")

        assert len(labels) == 19
        assert {label.location_line2 for label in labels} >= {"Bergen, Fløyen"}

    def test_defaults_can_supply_a_count(self, tmp_path):
        path = tmp_path / "s.json"
        path.write_text(
            json.dumps(
                {
                    "defaults": {"count": 2},
                    "sites": {"A": {"location_line1": "Norway,"}},
                    "labels": [
                        {"site": "A", "code": "N1"},
                        {"site": "A", "code": "N2", "count": 1},
                    ],
                }
            ),
            encoding="utf-8",
        )

        codes = [label.code for label in load_data(path)]

        assert codes == ["N1", "N1", "N2"]


class TestStandaloneRegistry:
    def test_a_registry_file_loads(self, tmp_path):
        path = tmp_path / "sites.json"
        path.write_text(json.dumps({"sites": {"A": {"location_line1": "Norway,"}}}), "utf-8")

        assert "A" in load_site_registry(path)

    def test_a_bare_mapping_of_sites_is_accepted(self, tmp_path):
        path = tmp_path / "sites.json"
        path.write_text(json.dumps({"A": {"location_line1": "Norway,"}}), "utf-8")

        assert "A" in load_site_registry(path)

    def test_an_empty_registry_is_an_error(self, tmp_path):
        path = tmp_path / "sites.json"
        path.write_text(json.dumps({"sites": {}}), "utf-8")

        with pytest.raises(ValueError, match="no sites"):
            load_site_registry(path)
