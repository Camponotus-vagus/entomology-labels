# Contributing

Bug reports, localities that render badly, and pull requests are all welcome.

## Setting up

```bash
git clone https://github.com/Camponotus-vagus/entomology-labels.git
cd entomology-labels
python -m venv .venv && source .venv/bin/activate
pip install -e '.[all]'
pip install pytest black isort flake8
```

`weasyprint` needs system libraries for PDF output. If it will not install, skip
the `pdf` extra — everything else works without it, and you can print the HTML
from a browser instead.

## Before opening a pull request

```bash
pytest
black entomology_labels tests
isort entomology_labels tests
flake8 entomology_labels tests
```

CI runs the same four commands on Python 3.9 through 3.12.

## Things worth knowing

**Label composition lives in one place.** `entomology_labels/layout.py` defines
the ordered lines that make up a label; the HTML, DOCX and GUI-preview renderers
all consume it. Add a field there, not in the renderers.

**Input parsing is bounded on purpose.** Label data often comes from files you
did not write. `SECURITY_AUDIT.md` explains the limits in `config.py` and why
they exist; keep new input paths behind the same checks.

**Italian column aliases are a compatibility contract.** `località1`, `codice`,
`data`, `quantità` and friends are accepted alongside their English names so
that existing spreadsheets keep working. Do not remove them.

**Roman-numeral months are deliberate.** `15.vi.2024` is the international
entomological date convention, not a locale quirk. It stays as the output format.

## Reporting a label that renders wrong

Include the input file (or a couple of rows from it), the command you ran, and
the layout settings. A photograph of the printed sheet next to a ruler is more
useful than it sounds.
