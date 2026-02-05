"""Tests for the filters module."""

import pytest
from gbif_downloader.filters import FilterConfig, RecordFilter, filter_records
from gbif_downloader.api import OccurrenceRecord


class TestFilterConfig:
    """Tests for FilterConfig."""

    def test_valid_genus_config(self):
        """Test creating a valid config with genus."""
        config = FilterConfig(genus="Nebria")
        assert config.genus == "Nebria"
        assert config.year_start == 1800
        assert config.uncertainty_max == 1000

    def test_valid_family_config(self):
        """Test creating a valid config with family."""
        config = FilterConfig(family="Carabidae")
        assert config.family == "Carabidae"
        assert config.genus is None

    def test_missing_taxonomy_raises_error(self):
        """Test that missing genus and family raises ValueError."""
        with pytest.raises(ValueError, match="Either genus or family"):
            FilterConfig()

    def test_invalid_year_start_raises_error(self):
        """Test that invalid year raises ValueError."""
        with pytest.raises(ValueError, match="must be >= 1700"):
            FilterConfig(genus="Nebria", year_start=1500)

    def test_future_year_raises_error(self):
        """Test that future year raises ValueError."""
        with pytest.raises(ValueError, match="cannot be in the future"):
            FilterConfig(genus="Nebria", year_end=3000)

    def test_year_end_before_start_raises_error(self):
        """Test that year_end < year_start raises ValueError."""
        with pytest.raises(ValueError, match="cannot be after"):
            FilterConfig(genus="Nebria", year_start=2000, year_end=1990)

    def test_species_list_cleaned(self):
        """Test that species list is cleaned and lowercased."""
        config = FilterConfig(
            genus="Nebria",
            species_list=["Germarii", " castanea ", "", "ALPINA"]
        )
        assert config.species_list == ["germarii", "castanea", "alpina"]

    def test_countries_uppercased(self):
        """Test that country codes are uppercased."""
        config = FilterConfig(genus="Nebria", countries=["it", "Ch", "AT"])
        assert config.countries == ["IT", "CH", "AT"]

    def test_from_dict_nested(self):
        """Test creating config from nested dictionary."""
        data = {
            "taxonomy": {
                "genus": "Nebria",
                "species": ["germarii"],
            },
            "filters": {
                "year_start": 1900,
                "uncertainty_max": 500,
            }
        }
        config = FilterConfig.from_dict(data)
        assert config.genus == "Nebria"
        assert config.species_list == ["germarii"]
        assert config.year_start == 1900
        assert config.uncertainty_max == 500

    def test_to_dict(self):
        """Test converting config to dictionary."""
        config = FilterConfig(genus="Nebria", year_start=1900)
        data = config.to_dict()
        assert data["taxonomy"]["genus"] == "Nebria"
        assert data["filters"]["year_start"] == 1900


class TestRecordFilter:
    """Tests for RecordFilter."""

    @pytest.fixture
    def sample_record(self):
        """Create a sample occurrence record."""
        return OccurrenceRecord(
            key=12345,
            year=2020,
            event_date="2020-06-15",
            latitude=46.5,
            longitude=11.2,
            coordinate_uncertainty=50.0,
            elevation=1500.0,
            locality="Alps",
            genus="Nebria",
            species="Nebria germarii",
            scientific_name="Nebria germarii Heer, 1837",
            specific_epithet="germarii",
            institution_code="MZUF",
            catalog_number="123",
            recorded_by="J. Doe",
            country="Italy",
            state_province="Trentino",
            basis_of_record="PRESERVED_SPECIMEN",
        )

    @pytest.fixture
    def default_filter(self):
        """Create a default filter config."""
        return RecordFilter(FilterConfig(genus="Nebria"))

    def test_keep_valid_record(self, sample_record, default_filter):
        """Test that a valid record is kept."""
        result = default_filter.apply(sample_record)
        assert result.keep is True

    def test_filter_missing_year(self, sample_record):
        """Test filtering records without year."""
        sample_record.year = None
        filter_obj = RecordFilter(FilterConfig(genus="Nebria", require_year=True))
        result = filter_obj.apply(sample_record)
        assert result.keep is False
        assert result.reason == "missing_year"

    def test_keep_missing_year_when_not_required(self, sample_record):
        """Test keeping records without year when not required."""
        sample_record.year = None
        filter_obj = RecordFilter(FilterConfig(genus="Nebria", require_year=False))
        result = filter_obj.apply(sample_record)
        assert result.keep is True

    def test_filter_missing_elevation(self, sample_record):
        """Test filtering records without elevation."""
        sample_record.elevation = None
        filter_obj = RecordFilter(FilterConfig(genus="Nebria", require_elevation=True))
        result = filter_obj.apply(sample_record)
        assert result.keep is False
        assert result.reason == "missing_elevation"

    def test_filter_uncertainty_exceeded(self, sample_record):
        """Test filtering records with high uncertainty."""
        sample_record.coordinate_uncertainty = 5000.0
        filter_obj = RecordFilter(FilterConfig(genus="Nebria", uncertainty_max=1000))
        result = filter_obj.apply(sample_record)
        assert result.keep is False
        assert result.reason == "uncertainty_exceeded"

    def test_keep_uncertainty_at_limit(self, sample_record):
        """Test keeping records with uncertainty exactly at limit."""
        sample_record.coordinate_uncertainty = 1000.0
        filter_obj = RecordFilter(FilterConfig(genus="Nebria", uncertainty_max=1000))
        result = filter_obj.apply(sample_record)
        assert result.keep is True

    def test_unknown_uncertainty_kept(self, sample_record):
        """Test keeping records with unknown uncertainty."""
        sample_record.coordinate_uncertainty = None
        filter_obj = RecordFilter(
            FilterConfig(genus="Nebria", keep_unknown_uncertainty=True)
        )
        result = filter_obj.apply(sample_record)
        assert result.keep is True
        assert result.uncertainty_status == "unknown"

    def test_unknown_uncertainty_dropped(self, sample_record):
        """Test dropping records with unknown uncertainty."""
        sample_record.coordinate_uncertainty = None
        filter_obj = RecordFilter(
            FilterConfig(genus="Nebria", keep_unknown_uncertainty=False)
        )
        result = filter_obj.apply(sample_record)
        assert result.keep is False
        assert result.reason == "uncertainty_unknown"

    def test_filter_year_too_old(self, sample_record):
        """Test filtering records from before year range."""
        sample_record.year = 1750
        filter_obj = RecordFilter(FilterConfig(genus="Nebria", year_start=1800))
        result = filter_obj.apply(sample_record)
        assert result.keep is False
        assert result.reason == "year_too_old"

    def test_filter_species_not_matched(self, sample_record):
        """Test filtering when species doesn't match."""
        sample_record.specific_epithet = "alpina"
        filter_obj = RecordFilter(
            FilterConfig(genus="Nebria", species_list=["germarii", "castanea"])
        )
        result = filter_obj.apply(sample_record)
        assert result.keep is False
        assert result.reason == "species_not_matched"

    def test_filter_species_matched(self, sample_record):
        """Test keeping when species matches."""
        filter_obj = RecordFilter(
            FilterConfig(genus="Nebria", species_list=["germarii", "castanea"])
        )
        result = filter_obj.apply(sample_record)
        assert result.keep is True

    def test_deduplication(self, sample_record, default_filter):
        """Test that duplicate records are filtered."""
        result1 = default_filter.apply(sample_record)
        result2 = default_filter.apply(sample_record)
        assert result1.keep is True
        assert result2.keep is False
        assert result2.reason == "duplicate"

    def test_deduplication_disabled(self, sample_record):
        """Test deduplication can be disabled."""
        filter_obj = RecordFilter(
            FilterConfig(genus="Nebria", deduplicate=False)
        )
        result1 = filter_obj.apply(sample_record)
        result2 = filter_obj.apply(sample_record)
        assert result1.keep is True
        assert result2.keep is True

    def test_reset_clears_seen_keys(self, sample_record, default_filter):
        """Test that reset clears deduplication state."""
        default_filter.apply(sample_record)
        assert default_filter.seen_count == 1
        default_filter.reset()
        assert default_filter.seen_count == 0


class TestFilterRecords:
    """Tests for the filter_records function."""

    def test_filter_records_with_stats(self):
        """Test filtering a list of records and getting stats."""
        records = [
            OccurrenceRecord(
                key=1, year=2020, latitude=46.0, longitude=11.0,
                coordinate_uncertainty=50.0, elevation=1000.0,
                locality=None, genus="Nebria", species=None,
                scientific_name="Nebria germarii", specific_epithet="germarii",
                institution_code=None, catalog_number=None, recorded_by=None,
                country=None, state_province=None, basis_of_record=None,
                event_date=None,
            ),
            OccurrenceRecord(
                key=2, year=None, latitude=46.0, longitude=11.0,  # Missing year
                coordinate_uncertainty=50.0, elevation=1000.0,
                locality=None, genus="Nebria", species=None,
                scientific_name="Nebria germarii", specific_epithet="germarii",
                institution_code=None, catalog_number=None, recorded_by=None,
                country=None, state_province=None, basis_of_record=None,
                event_date=None,
            ),
            OccurrenceRecord(
                key=3, year=2020, latitude=46.0, longitude=11.0,
                coordinate_uncertainty=5000.0, elevation=1000.0,  # High uncertainty
                locality=None, genus="Nebria", species=None,
                scientific_name="Nebria germarii", specific_epithet="germarii",
                institution_code=None, catalog_number=None, recorded_by=None,
                country=None, state_province=None, basis_of_record=None,
                event_date=None,
            ),
        ]

        config = FilterConfig(genus="Nebria", require_year=True, uncertainty_max=1000)
        filtered, stats = filter_records(records, config)

        assert len(filtered) == 1
        assert filtered[0].key == 1
        assert stats["total"] == 3
        assert stats["kept"] == 1
        assert stats["missing_year"] == 1
        assert stats["uncertainty_exceeded"] == 1
