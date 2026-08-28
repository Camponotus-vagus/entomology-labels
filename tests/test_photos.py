"""Tests for reading collection metadata from photographs."""

from datetime import datetime

import pytest

from entomology_labels.config import MAX_EXIF_STRING_LEN
from entomology_labels.photos import (
    DEFAULT_RADIUS_M,
    PhotoReadError,
    PhotoRecord,
    _tags_to_photo,
    cluster_photos,
    dms_to_decimal,
    event_to_site,
    find_photos,
    haversine_m,
    read_photo_meta,
    scan_photos,
)

exifread = pytest.importorskip("exifread")


class _Ratio:
    """Stands in for an exifread Ratio."""

    def __init__(self, numerator, denominator=1):
        self.numerator = numerator
        self.denominator = denominator


class _Tag:
    def __init__(self, values):
        self.values = values

    def __str__(self):
        return str(self.values)


def _dms(degrees, minutes, seconds):
    return [_Ratio(degrees), _Ratio(minutes), _Ratio(int(seconds * 100), 100)]


def _tags(**overrides):
    tags = {
        "EXIF DateTimeOriginal": "2026:08:20 14:19:27",
        "GPS GPSLatitude": _Tag(_dms(60, 23, 47.26)),
        "GPS GPSLatitudeRef": "N",
        "GPS GPSLongitude": _Tag(_dms(5, 21, 11.32)),
        "GPS GPSLongitudeRef": "E",
        "GPS GPSAltitude": _Tag(_Ratio(3122, 10)),
        "GPS GPSAltitudeRef": "0",
        "Image Model": "TG-7",
    }
    tags.update(overrides)
    return {k: v for k, v in tags.items() if v is not None}


class TestDmsToDecimal:
    def test_converts_degrees_minutes_seconds(self):
        assert dms_to_decimal(_dms(60, 23, 47.26), "N") == pytest.approx(60.39646, abs=1e-5)

    def test_south_is_negative(self):
        assert dms_to_decimal(_dms(33, 55, 29.6), "S") == pytest.approx(-33.92489, abs=1e-5)

    def test_west_is_negative(self):
        assert dms_to_decimal(_dms(18, 25, 26.8), "W") == pytest.approx(-18.42411, abs=1e-5)

    def test_a_zero_denominator_is_rejected(self):
        with pytest.raises(ValueError, match="zero denominator"):
            dms_to_decimal([_Ratio(60, 0), _Ratio(0), _Ratio(0)], "N")

    def test_wrong_component_count_is_rejected(self):
        with pytest.raises(ValueError, match="three components"):
            dms_to_decimal([_Ratio(60), _Ratio(23)], "N")

    def test_out_of_range_minutes_are_rejected(self):
        with pytest.raises(ValueError, match="out of range"):
            dms_to_decimal([_Ratio(60), _Ratio(99), _Ratio(0)], "N")


class TestTagsToPhoto:
    """The conversion is tested against plain dicts, so most cases need no file."""

    def test_reads_date_position_and_elevation(self):
        record = _tags_to_photo(_tags(), "P8203037.ORF")

        assert record.taken_at == datetime(2026, 8, 20, 14, 19, 27)
        assert record.latitude == pytest.approx(60.39646, abs=1e-5)
        assert record.longitude == pytest.approx(5.35314, abs=1e-5)
        assert record.elevation_m == pytest.approx(312.2)
        assert record.camera == "TG-7"

    def test_a_photo_with_no_gps_still_yields_its_date(self):
        """The receiver had not acquired a fix for the first frames of a walk."""
        record = _tags_to_photo(
            _tags(**{"GPS GPSLatitude": None, "GPS GPSLongitude": None}), "x.ORF"
        )

        assert record.taken_at is not None
        assert record.has_position is False

    def test_altitude_ref_one_means_below_sea_level(self):
        record = _tags_to_photo(_tags(**{"GPS GPSAltitudeRef": "1"}), "x.ORF")

        assert record.elevation_m == pytest.approx(-312.2)

    def test_an_out_of_range_latitude_is_discarded(self):
        record = _tags_to_photo(_tags(**{"GPS GPSLatitude": _Tag(_dms(200, 0, 0))}), "x.ORF")

        assert record.has_position is False

    def test_unusable_coordinates_do_not_raise(self):
        record = _tags_to_photo(_tags(**{"GPS GPSLatitude": _Tag("garbage")}), "x.ORF")

        assert record.has_position is False

    def test_an_implausible_year_is_discarded(self):
        record = _tags_to_photo(_tags(**{"EXIF DateTimeOriginal": "1899:01:01 00:00:00"}), "x")

        assert record.taken_at is None

    def test_camera_text_is_truncated(self):
        """Camera-written strings reach a YAML file and then a printed label."""
        record = _tags_to_photo(_tags(**{"Image Model": "A" * 5000}), "x.ORF")

        assert len(record.camera) <= MAX_EXIF_STRING_LEN

    def test_control_characters_are_stripped_from_camera_text(self):
        record = _tags_to_photo(_tags(**{"Image Model": "TG\x00-\x077"}), "x.ORF")

        assert "\x00" not in record.camera


class TestHaversine:
    def test_zero_distance(self):
        assert haversine_m(60.0, 5.0, 60.0, 5.0) == pytest.approx(0.0)

    def test_a_known_separation(self):
        """The owner's two Bergen sites are about 3.3 km apart."""
        distance = haversine_m(60.3964, 5.3531, 60.3733, 5.3908)

        assert distance == pytest.approx(3300, rel=0.05)

    def test_is_symmetric(self):
        assert haversine_m(60.0, 5.0, 61.0, 6.0) == pytest.approx(haversine_m(61.0, 6.0, 60.0, 5.0))


def _photo(day, hour, minute, lat=60.3964, lon=5.3531, elev=300.0):
    return PhotoRecord(
        source=f"{day}-{hour}{minute}.ORF",
        taken_at=datetime(2026, 8, day, hour, minute),
        latitude=lat,
        longitude=lon,
        elevation_m=elev,
    )


class TestClustering:
    def test_one_morning_at_one_place_is_one_event(self):
        photos = [_photo(20, 14, m) for m in (9, 14, 19, 25, 29)]

        assert len(cluster_photos(photos)) == 1

    def test_separate_days_are_separate_events(self):
        photos = [_photo(20, 14, 19), _photo(25, 14, 24)]

        assert len(cluster_photos(photos)) == 2

    def test_a_long_gap_starts_a_new_event(self):
        photos = [_photo(20, 9, 0), _photo(20, 15, 0)]

        assert len(cluster_photos(photos, time_gap_minutes=90)) == 2

    def test_moving_far_enough_starts_a_new_event(self):
        photos = [_photo(20, 14, 0), _photo(20, 14, 10, lat=60.5000, lon=5.5000)]

        assert len(cluster_photos(photos, radius_m=DEFAULT_RADIUS_M)) == 2

    def test_moving_within_the_radius_does_not(self):
        photos = [_photo(20, 14, 0), _photo(20, 14, 10, lat=60.3974)]

        assert len(cluster_photos(photos, radius_m=DEFAULT_RADIUS_M)) == 1

    def test_photos_without_gps_are_not_discarded(self):
        photos = [
            _photo(20, 14, 0),
            PhotoRecord("nogps.ORF", taken_at=datetime(2026, 8, 20, 14, 5)),
        ]

        events = cluster_photos(photos)

        assert len(events) == 1
        assert len(events[0].photos) == 2

    def test_undated_photos_are_kept(self):
        events = cluster_photos([_photo(20, 14, 0), PhotoRecord("undated.ORF")])

        assert sum(len(e.photos) for e in events) == 2

    def test_no_photos_gives_no_events(self):
        assert cluster_photos([]) == []

    def test_elevation_uses_the_median_not_the_mean(self):
        """One bad barometric reading should not move the site's elevation."""
        photos = [
            _photo(20, 14, 0, elev=300),
            _photo(20, 14, 1, elev=305),
            _photo(20, 14, 2, elev=9000),
        ]

        assert cluster_photos(photos)[0].elevation_m == 305

    def test_the_elevation_range_is_reported(self):
        photos = [_photo(20, 14, 0, elev=270), _photo(20, 14, 1, elev=329)]

        assert cluster_photos(photos)[0].elevation_range_m == (270, 329)


class TestRealWorldShape:
    """Reproduces the shape of the owner's two Bergen collecting days."""

    DAY1 = [
        (60, 23, 48.64, 5, 21, 9.14, 328.7, 14, 14),
        (60, 23, 47.20, 5, 21, 11.36, 312.7, 14, 17),
        (60, 23, 47.26, 5, 21, 11.32, 312.2, 14, 19),
        (60, 23, 43.37, 5, 21, 13.01, 269.7, 14, 46),
    ]
    DAY2 = [
        (60, 22, 23.88, 5, 23, 26.95, 389.4, 14, 6),
        (60, 22, 23.76, 5, 23, 26.92, 353.0, 14, 24),
        (60, 22, 23.75, 5, 23, 26.90, 356.2, 14, 29),
    ]

    def _records(self):
        records = []
        for day, points in ((20, self.DAY1), (25, self.DAY2)):
            for d1, m1, s1, d2, m2, s2, alt, hour, minute in points:
                records.append(
                    PhotoRecord(
                        source=f"P8{day}.ORF",
                        taken_at=datetime(2026, 8, day, hour, minute),
                        latitude=dms_to_decimal(_dms(d1, m1, s1), "N"),
                        longitude=dms_to_decimal(_dms(d2, m2, s2), "E"),
                        elevation_m=alt,
                    )
                )
        return records

    def test_two_days_produce_exactly_two_sites(self):
        assert len(cluster_photos(self._records())) == 2

    def test_the_centroids_match_the_hand_computed_values(self):
        first, second = cluster_photos(self._records())

        assert first.centroid[0] == pytest.approx(60.3963, abs=5e-4)
        assert first.centroid[1] == pytest.approx(5.3531, abs=5e-4)
        assert second.centroid[0] == pytest.approx(60.3733, abs=5e-4)
        assert second.centroid[1] == pytest.approx(5.3908, abs=5e-4)

    def test_the_default_radius_never_merges_sites_3km_apart(self):
        first, second = cluster_photos(self._records())
        separation = haversine_m(*first.centroid, *second.centroid)

        assert separation > DEFAULT_RADIUS_M * 10

    def test_a_site_becomes_a_registry_entry_with_a_blank_locality(self):
        site = event_to_site(cluster_photos(self._records())[0])

        assert site["date"] == "20.viii.2026"
        assert site["location_line1"] == "", "EXIF cannot supply a place name"
        assert site["coordinates"]["lat"] == pytest.approx(60.3963, abs=5e-4)


class TestFileReading:
    """The only tests that need a real file; they prove the ORF path works."""

    def test_orf_magic_is_read(self, tiny_orf):
        record = read_photo_meta(tiny_orf)

        assert record.taken_at == datetime(2026, 8, 20, 14, 19, 27)
        assert record.latitude == pytest.approx(60.39646, abs=1e-4)
        assert record.camera == "TG-7"

    def test_plain_tiff_magic_is_read_too(self, tmp_path, exif_builder, tiff_magic):
        path = tmp_path / "photo.tif"
        path.write_bytes(exif_builder(magic=tiff_magic))

        assert read_photo_meta(path).taken_at is not None

    def test_the_fixture_is_tiny(self, tiny_orf):
        """Small enough to build in code rather than commit a raw frame."""
        assert tiny_orf.stat().st_size < 2048

    def test_an_empty_file_is_rejected(self, tmp_path):
        path = tmp_path / "empty.ORF"
        path.write_bytes(b"")

        with pytest.raises(PhotoReadError, match="empty"):
            read_photo_meta(path)

    def test_an_unreadable_file_raises_rather_than_returning_nonsense(self, tmp_path):
        path = tmp_path / "junk.ORF"
        path.write_bytes(b"this is not a photograph at all")

        with pytest.raises(PhotoReadError):
            read_photo_meta(path)

    def test_only_the_file_name_is_recorded_by_default(self, tiny_orf):
        """A shared draft should not carry someone's home directory."""
        assert read_photo_meta(tiny_orf).source == tiny_orf.name
        assert read_photo_meta(tiny_orf, full_path=True).source == str(tiny_orf)


class TestDirectoryScanning:
    def test_finds_photos_and_ignores_other_files(self, tmp_path, exif_builder, tiff_magic):
        (tmp_path / "a.ORF").write_bytes(exif_builder())
        (tmp_path / "b.jpg").write_bytes(exif_builder(magic=tiff_magic))
        (tmp_path / "notes.txt").write_text("not a photo", encoding="utf-8")

        found = find_photos([tmp_path])

        assert {p.name for p in found} == {"a.ORF", "b.jpg"}

    def test_the_scan_limit_is_honoured(self, tmp_path, exif_builder):
        for i in range(10):
            (tmp_path / f"p{i}.ORF").write_bytes(exif_builder())

        assert len(find_photos([tmp_path], limit=4)) == 4

    def test_a_broken_file_is_skipped_not_fatal(self, tmp_path, exif_builder):
        (tmp_path / "good.ORF").write_bytes(exif_builder())
        (tmp_path / "broken.ORF").write_bytes(b"nope")

        events, read, skipped = scan_photos([tmp_path])

        assert read == 1
        assert skipped == 1
        assert len(events) == 1

    def test_scanning_is_reproducible(self, tmp_path, exif_builder):
        for i in range(5):
            (tmp_path / f"p{i}.ORF").write_bytes(exif_builder())

        assert find_photos([tmp_path]) == find_photos([tmp_path])


class TestSixtySecondRounding:
    """Writers encode a decimal degree as rationals and land on exactly 60.

    41.8 degrees becomes 41 deg 47' 60" because 0.8 * 60 is 47.999... in
    binary. Rejecting that discarded the position with only a warning.
    """

    def test_sixty_seconds_carries_into_the_next_minute(self):
        assert dms_to_decimal([_Ratio(41), _Ratio(47), _Ratio(60)], "N") == pytest.approx(41.8)

    def test_sixty_minutes_carries_into_the_next_degree(self):
        assert dms_to_decimal([_Ratio(41), _Ratio(60), _Ratio(0)], "N") == pytest.approx(42.0)

    def test_it_agrees_with_the_unrounded_form(self):
        rounded = dms_to_decimal([_Ratio(41), _Ratio(47), _Ratio(60)], "N")
        exact = dms_to_decimal([_Ratio(41), _Ratio(48), _Ratio(0)], "N")

        assert rounded == pytest.approx(exact)

    def test_the_southern_hemisphere_still_negates(self):
        assert dms_to_decimal([_Ratio(41), _Ratio(47), _Ratio(60)], "S") == pytest.approx(-41.8)

    @pytest.mark.parametrize("minutes,seconds", [(61, 0), (0, 61), (100, 100), (-1, 0)])
    def test_genuinely_malformed_values_are_still_rejected(self, minutes, seconds):
        with pytest.raises(ValueError, match="out of range"):
            dms_to_decimal([_Ratio(41), _Ratio(minutes), _Ratio(seconds)], "N")

    def test_a_rounded_coordinate_survives_the_round_trip(self):
        """One decimal place is where this bites: hand-copied waypoints."""
        record = _tags_to_photo(
            _tags(**{"GPS GPSLatitude": _Tag([_Ratio(60), _Ratio(23), _Ratio(60)])}), "x.ORF"
        )

        assert record.has_position
        assert record.latitude == pytest.approx(60.4)
