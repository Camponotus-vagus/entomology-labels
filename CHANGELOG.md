# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project uses
[semantic versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- `layout.render_label_lines()`: one place that defines what a label contains.
  The HTML, DOCX and GUI-preview renderers all read from it.
- International example data covering Norwegian, Italian and Spanish localities.

### Fixed
- Duplicated labels no longer risk losing fields. The `count` expansion was
  copy-pasted in five places, each naming the fields by hand; all five now go
  through `expand_label()`.
- The GUI preview shows the additional-info line, which it previously omitted
  while the printed output included it.
- `--version` reports the installed version instead of a hardcoded `1.0.0` that
  disagreed with the packaging metadata.

## [1.2.0]

### Added
- Bounds on untrusted input and a documented security review (`SECURITY_AUDIT.md`).

### Fixed
- File loading for several input formats.
