"""Tests for the API module."""

import pytest
from unittest.mock import Mock, patch
from gbif_downloader.api import (
    GBIFClient,
    TaxonMatch,
    OccurrenceRecord,
    TaxonNotFoundError,
    APIError,
)


class TestTaxonMatch:
    """Tests for TaxonMatch class."""

    def test_from_api_response(self):
        """Test creating TaxonMatch from API response."""
        data = {
            "usageKey": 1035566,
            "scientificName": "Nebria Latreille, 1802",
            "canonicalName": "Nebria",
            "rank": "GENUS",
            "status": "ACCEPTED",
            "confidence": 100,
            "matchType": "EXACT",
            "kingdom": "Animalia",
            "family": "Carabidae",
            "genus": "Nebria",
        }
        match = TaxonMatch.from_api_response(data)

        assert match.usage_key == 1035566
        assert match.canonical_name == "Nebria"
        assert match.rank == "GENUS"
        assert match.match_type == "EXACT"
        assert match.kingdom == "Animalia"

    def test_from_api_response_missing_fields(self):
        """Test handling missing optional fields."""
        data = {
            "usageKey": 123,
            "matchType": "NONE",
        }
        match = TaxonMatch.from_api_response(data)

        assert match.usage_key == 123
        assert match.canonical_name == ""
        assert match.rank == ""


class TestOccurrenceRecord:
    """Tests for OccurrenceRecord class."""

    def test_from_api_response(self):
        """Test creating OccurrenceRecord from API response."""
        data = {
            "key": 12345,
            "year": 2020,
            "eventDate": "2020-06-15",
            "decimalLatitude": 46.5,
            "decimalLongitude": 11.2,
            "coordinateUncertaintyInMeters": 50.0,
            "elevation": 1500,
            "locality": "Alps",
            "genus": "Nebria",
            "species": "Nebria germarii",
            "scientificName": "Nebria germarii Heer, 1837",
            "specificEpithet": "germarii",
            "institutionCode": "MZUF",
            "catalogNumber": "123",
            "recordedBy": "J. Doe",
            "country": "Italy",
            "stateProvince": "Trentino",
            "basisOfRecord": "PRESERVED_SPECIMEN",
        }
        record = OccurrenceRecord.from_api_response(data)

        assert record.key == 12345
        assert record.year == 2020
        assert record.latitude == 46.5
        assert record.longitude == 11.2
        assert record.coordinate_uncertainty == 50.0
        assert record.elevation == 1500
        assert record.genus == "Nebria"

    def test_gbif_url(self):
        """Test GBIF URL generation."""
        record = OccurrenceRecord(
            key=12345, year=2020, event_date=None,
            latitude=46.5, longitude=11.2,
            coordinate_uncertainty=50.0, elevation=1500,
            locality=None, genus="Nebria", species=None,
            scientific_name=None, specific_epithet=None,
            institution_code=None, catalog_number=None,
            recorded_by=None, country=None, state_province=None,
            basis_of_record=None,
        )
        assert record.gbif_url == "https://www.gbif.org/occurrence/12345"

    def test_to_dict(self):
        """Test conversion to dictionary."""
        record = OccurrenceRecord(
            key=12345, year=2020, event_date="2020-06-15",
            latitude=46.5, longitude=11.2,
            coordinate_uncertainty=50.0, elevation=1500,
            locality="Alps", genus="Nebria", species="Nebria germarii",
            scientific_name="Nebria germarii", specific_epithet="germarii",
            institution_code="MZUF", catalog_number="123",
            recorded_by="J. Doe", country="Italy", state_province="Trentino",
            basis_of_record="PRESERVED_SPECIMEN",
        )
        data = record.to_dict()

        assert data["Year"] == 2020
        assert data["Latitude"] == 46.5
        assert data["Longitude"] == 11.2
        assert data["Genus"] == "Nebria"
        assert "Link" in data


class TestGBIFClient:
    """Tests for GBIFClient class."""

    @pytest.fixture
    def client(self):
        """Create a GBIFClient instance."""
        return GBIFClient()

    def test_initialization(self, client):
        """Test client initialization."""
        assert client.timeout == (10, 30)
        assert client.max_retries == 3
        assert client.page_size == 300

    def test_page_size_capped(self):
        """Test that page size is capped at 300."""
        client = GBIFClient(page_size=500)
        assert client.page_size == 300

    @patch.object(GBIFClient, '_make_request')
    def test_match_taxon_success(self, mock_request, client):
        """Test successful taxon matching."""
        mock_request.return_value = {
            "usageKey": 1035566,
            "scientificName": "Nebria Latreille, 1802",
            "canonicalName": "Nebria",
            "rank": "GENUS",
            "status": "ACCEPTED",
            "confidence": 100,
            "matchType": "EXACT",
        }

        taxon = client.match_taxon("Nebria", rank="GENUS")

        assert taxon.usage_key == 1035566
        assert taxon.canonical_name == "Nebria"
        assert taxon.rank == "GENUS"

    @patch.object(GBIFClient, '_make_request')
    def test_match_taxon_not_found(self, mock_request, client):
        """Test taxon not found raises error."""
        mock_request.return_value = {
            "matchType": "NONE",
        }

        with pytest.raises(TaxonNotFoundError):
            client.match_taxon("InvalidGenus")

    @patch.object(GBIFClient, '_make_request')
    def test_match_taxon_wrong_rank(self, mock_request, client):
        """Test matching wrong rank raises error with strict mode."""
        # API returns a CLASS when looking for a GENUS
        mock_request.return_value = {
            "usageKey": 216,
            "scientificName": "Insecta",
            "canonicalName": "Insecta",
            "rank": "CLASS",
            "status": "ACCEPTED",
            "confidence": 100,
            "matchType": "HIGHERRANK",
        }

        with pytest.raises(TaxonNotFoundError, match="matched as CLASS"):
            client.match_taxon("InvalidGenusXYZ", rank="GENUS", strict=True)

    @patch.object(GBIFClient, '_make_request')
    def test_match_taxon_name_mismatch(self, mock_request, client):
        """Test matching different name raises error."""
        mock_request.return_value = {
            "usageKey": 1035566,
            "scientificName": "Nebria Latreille, 1802",
            "canonicalName": "Nebria",  # Different from "Nebra"
            "rank": "GENUS",
            "status": "ACCEPTED",
            "confidence": 90,
            "matchType": "FUZZY",
        }

        with pytest.raises(TaxonNotFoundError, match="may not be what you intended"):
            client.match_taxon("Nebra", rank="GENUS", strict=True)

    @patch.object(GBIFClient, '_make_request')
    def test_count_occurrences(self, mock_request, client):
        """Test counting occurrences."""
        mock_request.return_value = {"count": 39355}

        count = client.count_occurrences(1035566)

        assert count == 39355
        mock_request.assert_called_once()

    @patch.object(GBIFClient, '_make_request')
    def test_iter_occurrences(self, mock_request, client):
        """Test iterating over occurrences."""
        mock_request.side_effect = [
            {
                "count": 2,
                "results": [
                    {"key": 1, "year": 2020, "decimalLatitude": 46.0, "decimalLongitude": 11.0},
                    {"key": 2, "year": 2020, "decimalLatitude": 47.0, "decimalLongitude": 12.0},
                ],
                "endOfRecords": True,
            }
        ]

        records = list(client.iter_occurrences(1035566))

        assert len(records) == 2
        assert records[0].key == 1
        assert records[1].key == 2

    def test_context_manager(self, client):
        """Test using client as context manager."""
        with GBIFClient() as c:
            assert isinstance(c, GBIFClient)
        # Session should be closed after exiting


class TestGBIFClientIntegration:
    """Integration tests that actually call the GBIF API.

    These tests are marked with 'integration' and can be skipped
    in CI environments without network access.
    """

    @pytest.mark.integration
    def test_real_taxon_match(self):
        """Test matching a real taxon against GBIF API."""
        with GBIFClient() as client:
            taxon = client.match_taxon("Nebria", rank="GENUS")
            assert taxon.canonical_name == "Nebria"
            assert taxon.rank == "GENUS"
            assert taxon.usage_key > 0

    @pytest.mark.integration
    def test_real_count(self):
        """Test counting real occurrences."""
        with GBIFClient() as client:
            taxon = client.match_taxon("Nebria", rank="GENUS")
            count = client.count_occurrences(taxon.usage_key)
            assert count > 1000  # Nebria has many records
