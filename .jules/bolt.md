# Bolt's journal

Critical performance learnings for this codebase only.

## 2026-10-31 - In python-docx, formatting costs more than content
**Learning:** Word export was about 230x slower than HTML (36 s against
0.16 s for 20,000 labels). The time was not in writing text: every
python-docx property setter (`run.font.size`, `run.font.name`,
`paragraph_format.space_before`, and so on) walks the element's children
to insert in schema order, and the export set five of them on every line of
every label. All of them repeated values the Normal style already carried.
Also, `table_row.cells` rebuilds its list on each access, so indexing it
per cell made each row O(columns^2).
**Action:** In `generate_docx`, put shared formatting on a style and set
per-run properties only where a line differs from it. Prove "same output" by
rendering before and after with LibreOffice
(`soffice --convert-to pdf`, `pdftoppm`, `compare -metric AE`). Timing alone
does not show that nothing changed.
