# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project uses
[semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `photos scan` and `photos probe`: read collection dates and GPS from field
  photographs, group them into collection sites, and write a draft registry.
  Raw formats (ORF, NEF, CR2, DNG) are read directly.
- Opt-in `--geocode` suggests place names from OpenStreetMap or GeoNames, as
  comments for you to confirm rather than values filled in automatically.
- Collection sites: declare a locality once under `sites:` and reference it by
  id from many rows, with a `defaults:` block. Existing flat files are
  unaffected.
- Species, coordinates, elevation, collector and determiner fields, and
  `--label-kind` to print locality labels, determination labels, or both.
- Label geometry presets following published museum guidance. `museum` fits 475
  labels to an A4 sheet against the original geometry's 160.
- `--config` and `--save-config` for reusable layouts, wiring up a config path
  that had been declared but never used.
- Date parsing: ISO, EXIF, day-first and Roman-numeral forms in;
  Roman numerals out. `--normalize-dates` opts in to rewriting.
- Fit checking before output, with `--strict-fit` to make warnings fatal.
- `layout.render_label_lines()`: one place that defines what a label contains.
  The HTML, DOCX and GUI-preview renderers all read from it.
- International example data covering Norwegian, Italian and Spanish localities.

### Changed
- Overlong label text now wraps instead of being trimmed with an ellipsis, so
  locality data is not discarded silently. `--text-overflow clip` restores the
  old behaviour.
- `generate` with no layout flags now matches `LabelConfig` and the GUI. It
  previously used a different, undocumented set of defaults.

### Fixed
- Duplicated labels no longer risk losing fields. The `count` expansion was
  copy-pasted in five places, each naming the fields by hand; all five now go
  through `expand_label()`.
- The GUI preview shows the additional-info line, which it previously omitted
  while the printed output included it.
- `--version` reports the installed version instead of a hardcoded `1.0.0` that
  disagreed with the packaging metadata.
- The GUI label list showed a hardcoded "1" for every row's quantity; it now
  counts identical labels and the column is named "Copies".
- English column headings are matched before Italian ones, so a file with a
  `location` column no longer loses to `località1`.
- Spreadsheet cells are read as text. A numeric column with one blank cell
  printed `5.0` for code 5 and `310.0` for an elevation; leading zeros were
  dropped; `NA` (Namibia) was read as a missing value. Duplicate headings
  (`Code` and `code`) are now an error instead of a garbled label.
- Counts must be whole numbers: `2.7` is rejected instead of printing 2, and
  `inf` no longer crashes the loader.
- A plain-text label file containing a colon (a time, `12:30`) is no longer
  mistaken for the key/value format, which lost every field. Key/value blocks
  separated by a whitespace-only line no longer merge, and a UTF-8 byte-order
  mark no longer hides the first key. JSON and YAML accept the mark too.
- Layout settings are type-checked: a configuration file with `"10"` for a
  count, a zero or negative page size, zero line spacing or a non-numeric size
  is rejected. Such values previously reached the generated stylesheet, so an
  edited configuration file could inject script into the HTML.
- Word output matches the HTML geometry: columns were 24.6 mm wide whatever
  the label width, rows grew with their text, blank lines took Word's 11 pt
  default and label padding was ignored, so lower rows drifted off the label
  stock. Table properties are written in the order Word's schema requires.
- In HTML and PDF a blank line keeps its height, so labels in a row line up
  whether or not a locality line is empty. The text-overflow mode now applies
  to every line, not only the locality, and the last page no longer forces a
  trailing page break.
- Line breaks and tabs inside a field print as spaces in every output, and
  characters that made the HTML or Word write fail are dropped.
- `photos scan`: a day with more than 26 sites no longer reuses ids
  (`...Z`, then `...AA`), which silently overwrote sites in the draft. A site
  on the antimeridian is centred there, not at longitude 0. A 0,0 GPS reading
  (no fix) is treated as no position. Undated photographs no longer move a
  site's coordinates. A folder given together with its subfolder no longer
  counts photographs twice. A missing exifread is reported once, up front.
- `--geocode`: GeoNames is queried over HTTPS. Its quota and bad-username
  replies are no longer cached as "nothing here", and a refusal (HTTP 429/403
  or a GeoNames error) stops the remaining lookups instead of repeating them
  against the usage policy. A truncated reply or a malformed cache file no
  longer aborts the scan.
- `generate` refuses to write over its input (a `.docx` table could be
  replaced by the label sheet), warns when rows name a site but no `--sites`
  file is given, and reports a missing optional dependency without a
  traceback. `sequence` reads `label_config.json` like `generate` does.
  `photos scan` no longer accepts a `.json` draft it would write as YAML.
  `python -m entomology_labels.cli photos` works.
- GUI: Duplicate and Edit keep species, coordinates, collector and the label
  kind, which both silently dropped. An out-of-range or non-numeric quantity
  is an error instead of a silent 1 or a frozen window. The About box reports
  the installed version.
- Release workflow: the manual-run tag is no longer interpolated into a shell
  script, releases are built from the tagged commit, a tag that does not match
  the package version fails the run, and only the release job can write.
- CSV and tab-separated text are read with the standard library, so the bare
  install reads them as the README says; pandas was required before. An
  unterminated quote is reported with its line instead of swallowing the rest
  of the file, and Excel's Windows-1252 CSV export is accepted. `.xls`
  workbooks work: the `excel` extra now includes `xlrd`.
- `--label-kind both` keeps each locality/determination pair on one sheet,
  leaving a blank slot when a pair would start on the last slot of a page.
- GUI: the preview draws the page at its real size with its margins, and wraps
  labels the way the printed page does, instead of an idealised grid. Applying
  the Configuration tab no longer resets text overflow, label padding and the
  spacer line, derives the orientation from the page size, and warns when the
  grid does not fit. Exporting asks to apply unapplied settings first and
  writes from a copy of the labels, so editing during a long PDF export cannot
  change the file being written.
- Release workflow: the test suite gates publishing, a missing PyPI token fails
  the run instead of skipping the upload, a failed GUI build is no longer
  hidden, and actions are pinned by commit (kept current by Dependabot).
- CI runs the GUI tests under a virtual display instead of skipping them.

## [1.2.0]

### Added
- Bounds on untrusted input and a documented security review (`SECURITY_AUDIT.md`).

### Fixed
- File loading for several input formats.
