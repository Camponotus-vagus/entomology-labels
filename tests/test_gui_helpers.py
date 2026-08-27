"""Tests for GUI helpers that do not need a display."""

from collections import Counter

import pytest

from entomology_labels.label_generator import Label

pytest.importorskip("tkinter", reason="tkinter is not installed in this environment")

from entomology_labels.gui import EntomologyLabelsGUI  # noqa: E402


class TestLabelKey:
    """Backs the Copies column, which used to be hardcoded to 1."""

    def test_identical_labels_share_a_key(self):
        first = Label(location_line1="Norway,", code="N1", date="20.viii.2026")
        second = Label(location_line1="Norway,", code="N1", date="20.viii.2026")

        assert EntomologyLabelsGUI._label_key(first) == EntomologyLabelsGUI._label_key(second)

    def test_labels_differing_anywhere_do_not(self):
        base = Label(location_line1="Norway,", code="N1")

        for changed in (
            Label(location_line1="Norway,", code="N2"),
            Label(location_line1="Italy,", code="N1"),
            Label(location_line1="Norway,", code="N1", collector="F. Mensa"),
        ):
            assert EntomologyLabelsGUI._label_key(base) != EntomologyLabelsGUI._label_key(changed)

    def test_the_key_is_hashable(self):
        """It is used as a Counter key, so this is load-bearing."""
        assert isinstance(hash(EntomologyLabelsGUI._label_key(Label(code="N1"))), int)

    def test_counting_recovers_the_requested_quantity(self):
        """Quantity is expanded at add-time, so it can only be counted back."""
        from entomology_labels.label_generator import expand_label

        labels = expand_label(Label(code="N1"), 5) + expand_label(Label(code="N2"), 3)

        counts = Counter(EntomologyLabelsGUI._label_key(label) for label in labels)

        assert sorted(counts.values()) == [3, 5]
