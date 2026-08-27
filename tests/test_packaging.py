"""Checks on packaging metadata that CI would not otherwise catch."""

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
PACKAGE = ROOT / "entomology_labels"
RELEASE_WORKFLOW = ROOT / ".github" / "workflows" / "release.yml"


def _package_modules():
    """Every importable module in the package."""
    modules = set()
    for path in PACKAGE.glob("*.py"):
        if path.stem != "__init__":
            modules.add(path.stem)
    for path in PACKAGE.iterdir():
        if path.is_dir() and (path / "__init__.py").exists():
            modules.add(path.name)
    return modules


class TestFrozenBinaryImports:
    """A module missing from the workflow breaks the released binaries only.

    PyInstaller does not follow every import, so each module is named
    explicitly. Nothing in the normal test run exercises the frozen build, so
    an omission would surface as a ModuleNotFoundError for whoever downloaded
    the release -- which is exactly the audience least able to diagnose it.
    """

    def test_the_workflow_exists(self):
        assert RELEASE_WORKFLOW.is_file()

    @pytest.mark.parametrize("module", sorted(_package_modules()))
    def test_every_module_is_declared_as_a_hidden_import(self, module):
        text = RELEASE_WORKFLOW.read_text(encoding="utf-8")

        assert re.search(rf"--hidden-import entomology_labels\.{module}\b(?!_)", text), (
            f"entomology_labels.{module} is missing from release.yml; "
            f"the frozen binaries will fail to import it at runtime"
        )

    def test_every_build_command_declares_the_same_modules(self):
        """Two binaries on three platforms; all of them need the full list."""
        text = RELEASE_WORKFLOW.read_text(encoding="utf-8")
        # Anchored on a word boundary: "config" is a prefix of "config_io".
        counts = {
            module: len(re.findall(rf"--hidden-import entomology_labels\.{module}\b(?!_)", text))
            for module in _package_modules()
        }

        assert len(set(counts.values())) == 1, f"uneven hidden-import declarations: {counts}"


class TestVersionConsistency:
    """Three places used to disagree about the version number."""

    def test_the_cli_does_not_hardcode_a_version(self):
        source = (PACKAGE / "cli.py").read_text(encoding="utf-8")

        assert not re.search(
            r'version="\d+\.\d+\.\d+"', source
        ), "the CLI should report the installed version, not a literal"

    def test_the_package_version_is_importable(self):
        import entomology_labels

        assert re.match(r"^\d+\.\d+", entomology_labels.__version__)


class TestOptionalDependencies:
    def test_every_extra_is_included_in_all(self):
        text = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
        block = re.search(r"\[project\.optional-dependencies\](.*?)(?=\n\[)", text, re.S).group(1)

        extras = dict(re.findall(r"^(\w+) = \[(.*?)\]", block, re.S | re.M))
        all_deps = extras.pop("all", "")
        extras.pop("dev", None)

        for name, deps in extras.items():
            for package in re.findall(r'"([A-Za-z0-9_.-]+)', deps):
                assert package in all_deps, f"{package} (from [{name}]) is missing from [all]"


class TestPresetsFitTheirOwnLabels:
    """A preset the tool ships should not trip the tool's own fit checker."""

    @pytest.mark.parametrize("name", ["museum", "museum-wide", "readable", "legacy", "letter"])
    def test_the_grid_fits_the_page(self, name):
        from entomology_labels.fit import check_page_fit
        from entomology_labels.label_generator import LabelConfig
        from entomology_labels.presets import get_preset

        assert check_page_fit(LabelConfig(**get_preset(name))) == []

    @pytest.mark.parametrize("name", ["museum", "letter"])
    def test_a_five_line_locality_label_fits(self, name):
        """These presets are sized for the classic five-line label."""
        from entomology_labels.fit import KIND_LABEL_HEIGHT, check_label_fit
        from entomology_labels.label_generator import Label, LabelConfig
        from entomology_labels.presets import get_preset

        config = LabelConfig(**get_preset(name))
        label = Label(
            location_line1="Norway,",
            location_line2="Bergen",
            code="N1",
            date="20.viii.2026",
            collector="F. Mensa",
        )

        warnings = [w for w in check_label_fit([label], config) if w.kind == KIND_LABEL_HEIGHT]

        assert warnings == []

    def test_museum_wide_also_fits_coordinates_and_elevation(self):
        from entomology_labels.fit import KIND_LABEL_HEIGHT, check_label_fit
        from entomology_labels.label_generator import Label, LabelConfig
        from entomology_labels.presets import get_preset

        config = LabelConfig(**get_preset("museum-wide"))
        label = Label(
            location_line1="Norway,",
            location_line2="Bergen",
            coordinates="60.3964N 5.3531E",
            elevation="290 m",
            code="N1",
            date="20.viii.2026",
            collector="F. Mensa",
        )

        warnings = [w for w in check_label_fit([label], config) if w.kind == KIND_LABEL_HEIGHT]

        assert warnings == []

    def test_the_recommended_size_beats_the_original_on_yield(self):
        """The whole point of following the guidance is more labels per sheet."""
        from entomology_labels.presets import labels_per_page

        assert labels_per_page("museum") > 2 * labels_per_page("legacy")
