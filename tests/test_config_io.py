"""Tests for loading and saving layout configuration."""

import json

import pytest

from entomology_labels.config import DEFAULT_CONFIG_PATH
from entomology_labels.config_io import (
    find_config_file,
    load_config_file,
    resolve_config,
    save_config_file,
    user_config_path,
)
from entomology_labels.label_generator import LabelConfig


@pytest.fixture
def config_file(tmp_path):
    path = tmp_path / "layout.json"
    save_config_file(LabelConfig(label_width_mm=45.0, font_size_pt=5.0), path)
    return path


class TestRoundTrip:
    def test_a_saved_configuration_loads_back_unchanged(self, tmp_path):
        original = LabelConfig(
            labels_per_row=6, label_width_mm=45.0, font_size_pt=5.0, text_overflow="clip"
        )
        path = save_config_file(original, tmp_path / "layout.json")

        assert load_config_file(path).to_dict() == original.to_dict()

    def test_missing_parent_directories_are_created(self, tmp_path):
        path = save_config_file(LabelConfig(), tmp_path / "a" / "b" / "layout.json")

        assert path.is_file()


class TestLoadRejectsBadInput:
    """A config file is untrusted input like any other."""

    def test_a_missing_file_is_an_error(self, tmp_path):
        with pytest.raises(ValueError, match="not found"):
            load_config_file(tmp_path / "absent.json")

    def test_malformed_json_is_an_error(self, tmp_path):
        path = tmp_path / "layout.json"
        path.write_text("{not json", encoding="utf-8")

        with pytest.raises(ValueError, match="not valid JSON"):
            load_config_file(path)

    def test_a_json_array_is_rejected(self, tmp_path):
        path = tmp_path / "layout.json"
        path.write_text("[1, 2, 3]", encoding="utf-8")

        with pytest.raises(ValueError, match="JSON object"):
            load_config_file(path)

    def test_out_of_bounds_settings_are_rejected(self, tmp_path):
        path = tmp_path / "layout.json"
        path.write_text(json.dumps({"labels_per_row": 100000}), encoding="utf-8")

        with pytest.raises(ValueError, match="Invalid configuration"):
            load_config_file(path)

    def test_css_injection_through_font_family_is_rejected(self, tmp_path):
        path = tmp_path / "layout.json"
        path.write_text(
            json.dumps({"font_family": "Arial; } body { display: none; "}), encoding="utf-8"
        )

        with pytest.raises(ValueError, match="Invalid configuration"):
            load_config_file(path)

    def test_unknown_keys_are_ignored_rather_than_crashing(self, tmp_path):
        path = tmp_path / "layout.json"
        path.write_text(json.dumps({"font_size_pt": 5.0, "nonsense": True}), encoding="utf-8")

        assert load_config_file(path).font_size_pt == 5.0


class TestPrecedence:
    def test_no_file_and_no_overrides_gives_the_defaults(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "xdg"))

        assert resolve_config().to_dict() == LabelConfig().to_dict()

    def test_an_explicit_file_beats_the_defaults(self, config_file):
        assert resolve_config(config_file).label_width_mm == 45.0

    def test_overrides_beat_the_file(self, config_file):
        config = resolve_config(config_file, {"font_size_pt": 9.0})

        assert config.font_size_pt == 9.0
        assert config.label_width_mm == 45.0, "unset settings must still come from the file"

    def test_none_overrides_do_not_mask_the_file(self, config_file):
        """click cannot tell an unset option from one typed at its default."""
        config = resolve_config(config_file, {"font_size_pt": None, "label_width_mm": None})

        assert config.font_size_pt == 5.0
        assert config.label_width_mm == 45.0

    def test_a_working_directory_file_is_found(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "xdg"))
        save_config_file(LabelConfig(font_size_pt=4.0), tmp_path / DEFAULT_CONFIG_PATH)

        assert resolve_config().font_size_pt == 4.0

    def test_a_named_but_missing_file_is_an_error(self, tmp_path):
        """Silently ignoring it would apply settings the user did not ask for."""
        with pytest.raises(ValueError, match="not found"):
            find_config_file(tmp_path / "absent.json")


class TestUserConfigPath:
    def test_respects_xdg_config_home(self, monkeypatch, tmp_path):
        monkeypatch.setattr("os.name", "posix")
        monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))

        path = user_config_path()

        assert path.parent.name == "entomology-labels"
        assert path.name == DEFAULT_CONFIG_PATH

    def test_the_dead_constant_is_now_the_filename(self):
        assert user_config_path().name == DEFAULT_CONFIG_PATH
