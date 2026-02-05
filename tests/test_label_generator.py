"""Tests for the label generator module."""

import pytest
from entomology_labels.label_generator import Label, LabelConfig, LabelGenerator


class TestLabel:
    """Tests for the Label class."""

    def test_label_creation(self):
        """Test creating a basic label."""
        label = Label(
            location_line1="Italia, Trentino Alto Adige,",
            location_line2="Giustino (TN), Vedretta d'Amola",
            code="N1",
            date="15.vi.2024"
        )
        assert label.location_line1 == "Italia, Trentino Alto Adige,"
        assert label.location_line2 == "Giustino (TN), Vedretta d'Amola"
        assert label.code == "N1"
        assert label.date == "15.vi.2024"

    def test_label_is_empty(self):
        """Test the is_empty method."""
        empty_label = Label()
        assert empty_label.is_empty()

        non_empty = Label(code="N1")
        assert not non_empty.is_empty()

    def test_label_to_dict(self):
        """Test converting label to dictionary."""
        label = Label(
            location_line1="Italia",
            location_line2="Milano",
            code="X1",
            date="01.i.2024"
        )
        d = label.to_dict()
        assert d["location_line1"] == "Italia"
        assert d["code"] == "X1"

    def test_label_from_dict(self):
        """Test creating label from dictionary."""
        data = {
            "location_line1": "Italia",
            "location_line2": "Roma",
            "code": "R1",
            "date": "10.iii.2024"
        }
        label = Label.from_dict(data)
        assert label.location_line1 == "Italia"
        assert label.code == "R1"

    def test_label_from_dict_alternative_keys(self):
        """Test creating label from dictionary with alternative keys."""
        data = {
            "location1": "Francia",
            "location2": "Parigi",
            "specimen_code": "P1",
            "collection_date": "05.v.2024"
        }
        label = Label.from_dict(data)
        assert label.location_line1 == "Francia"
        assert label.code == "P1"


class TestLabelConfig:
    """Tests for the LabelConfig class."""

    def test_default_config(self):
        """Test default configuration values (A4 Landscape)."""
        config = LabelConfig()
        assert config.labels_per_row == 10
        assert config.labels_per_column == 13
        assert config.labels_per_page == 130
        assert config.page_width_mm == 297.0  # A4 landscape width
        assert config.page_height_mm == 210.0  # A4 landscape height
        assert config.orientation == "landscape"
        assert config.is_landscape is True

    def test_custom_config(self):
        """Test custom configuration."""
        config = LabelConfig(labels_per_row=12, labels_per_column=15)
        assert config.labels_per_row == 12
        assert config.labels_per_column == 15
        assert config.labels_per_page == 180

    def test_portrait_orientation(self):
        """Test portrait orientation configuration."""
        config = LabelConfig(
            orientation="portrait",
            page_width_mm=210.0,
            page_height_mm=297.0
        )
        assert config.is_landscape is False
        assert config.page_width_mm == 210.0
        assert config.page_height_mm == 297.0

    def test_config_to_dict(self):
        """Test converting config to dictionary."""
        config = LabelConfig()
        d = config.to_dict()
        assert "labels_per_row" in d
        assert d["labels_per_row"] == 10

    def test_config_from_dict(self):
        """Test creating config from dictionary."""
        data = {"labels_per_row": 8, "labels_per_column": 10}
        config = LabelConfig.from_dict(data)
        assert config.labels_per_row == 8
        assert config.labels_per_column == 10


class TestLabelGenerator:
    """Tests for the LabelGenerator class."""

    def test_generator_creation(self):
        """Test creating a label generator."""
        generator = LabelGenerator()
        assert generator.total_labels == 0
        assert generator.total_pages == 0

    def test_add_label(self):
        """Test adding a single label."""
        generator = LabelGenerator()
        label = Label(code="N1")
        generator.add_label(label)
        assert generator.total_labels == 1

    def test_add_labels(self):
        """Test adding multiple labels."""
        generator = LabelGenerator()
        labels = [Label(code=f"N{i}") for i in range(5)]
        generator.add_labels(labels)
        assert generator.total_labels == 5

    def test_clear_labels(self):
        """Test clearing all labels."""
        generator = LabelGenerator()
        generator.add_labels([Label(code=f"N{i}") for i in range(10)])
        assert generator.total_labels == 10
        generator.clear_labels()
        assert generator.total_labels == 0

    def test_total_pages_calculation(self):
        """Test page count calculation."""
        config = LabelConfig(labels_per_row=10, labels_per_column=10)  # 100 per page
        generator = LabelGenerator(config)

        # Add 250 labels -> should be 3 pages
        labels = [Label(code=f"X{i}") for i in range(250)]
        generator.add_labels(labels)
        assert generator.total_pages == 3

    def test_get_labels_for_page(self):
        """Test getting labels for a specific page."""
        config = LabelConfig(labels_per_row=2, labels_per_column=2)  # 4 per page
        generator = LabelGenerator(config)
        labels = [Label(code=f"L{i}") for i in range(10)]
        generator.add_labels(labels)

        page0_labels = generator.get_labels_for_page(0)
        assert len(page0_labels) == 4
        assert page0_labels[0].code == "L0"

        page2_labels = generator.get_labels_for_page(2)
        assert len(page2_labels) == 2

    def test_get_labels_grid(self):
        """Test getting labels as a grid."""
        config = LabelConfig(labels_per_row=2, labels_per_column=2)
        generator = LabelGenerator(config)
        labels = [Label(code=f"G{i}") for i in range(4)]
        generator.add_labels(labels)

        grid = generator.get_labels_grid(0)
        assert len(grid) == 2  # 2 rows
        assert len(grid[0]) == 2  # 2 columns
        assert grid[0][0].code == "G0"
        assert grid[1][1].code == "G3"

    def test_generate_sequential_labels(self):
        """Test generating sequential labels."""
        generator = LabelGenerator()
        labels = generator.generate_sequential_labels(
            location_line1="Italia",
            location_line2="Milano",
            code_prefix="M",
            start_number=1,
            end_number=5,
            date="01.i.2024"
        )
        assert len(labels) == 5
        assert labels[0].code == "M1"
        assert labels[4].code == "M5"
        assert all(l.location_line1 == "Italia" for l in labels)
