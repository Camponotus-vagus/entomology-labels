# Palette's journal

Critical UX and accessibility learnings for this app only.

## 2026-10-31 - The data-entry pane clips anything fixed-width
**Learning:** At the default 1100x800 window the "Add Single Label" pane of the
Tkinter GUI is only about 200 px wide. Single-line text there is cut off
without an ellipsis (the "Clear Form" button already shows "Clear For"), so a
hint or label that does not fit is silently half-shown.
**Action:** Give helper text in that pane a `wraplength` (about 180 px) and
check it in a screenshot at the default size. Drive and screenshot the GUI
headlessly with `xvfb-run` and a Python that has tkinter (system python3.12
with python3-tk here); `tests/test_gui_preview.py` has a fixture that stubs
the message boxes.
