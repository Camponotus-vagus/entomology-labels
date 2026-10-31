# Palette's journal

Critical UX and accessibility learnings for this app only.

## 2026-10-31 - In ttk panes, clipping is silent and comes from the neighbour
**Learning:** A ttk.PanedWindow shares any shortfall between its panes, and
Tk clips a widget without an ellipsis or warning. The "Add Single Label"
form was cut to "Clear For" not by its own layout but by the list pane: its
button bar had been packed after a side=LEFT tree, so it sat beside the list
instead of under it and demanded 1095 px of a 1058 px window. Labels whose
text grows at run time (the "Showing 1-500 of 100000" counter) change the
requested width after the first layout.
**Action:** When something is cut off, compare `winfo_reqwidth()` with
`winfo_width()` for both panes, not just the clipped one. Pack bottom bars
with `side=BOTTOM, before=<main widget>`, and put text that grows on its own
row. `test_no_button_on_the_data_tab_is_cut_off` checks this at the minimum
and default window sizes. Drive and screenshot the GUI headlessly with
`xvfb-run` and a Python that has tkinter (system python3.12 with python3-tk
here); `tests/test_gui_preview.py` has a fixture that stubs the message boxes.
