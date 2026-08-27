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

## [1.2.0]

### Added
- Bounds on untrusted input and a documented security review (`SECURITY_AUDIT.md`).

### Fixed
- File loading for several input formats.
