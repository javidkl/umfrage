# umfrage

**umfrage** is a command-line tool for creating, distributing, and collecting
Excel-based questionnaires. Survey forms are defined in human-readable YAML
files and compiled into protected `.xlsx` files that can be emailed to
respondents. Returned files are validated and aggregated into a single result
spreadsheet.

Licensed under the [Apache License 2.0](LICENSE).

---

## Features

- **Config-driven**: questionnaires are fully described in YAML — no code required
- **Protected Excel forms**: question cells are locked; only answer cells and
  respondent fields are editable; optional password protection
- **Data validation**: dropdown lists for scale (1–N) and yes/no answers
- **Completeness check**: `umfrage validate` catches errors before generation
- **Config-free collection**: companion `*_metadata.yaml` embeds the full model
  so `umfrage collect` works without the original config file
- **Multi-questionnaire folders**: different questionnaires can coexist in the
  same response folder; one result file is produced per questionnaire
- **LLM authoring guide**: `docs/llm_guide.md` and `docs/questionnaire.schema.json`
  enable AI-assisted questionnaire authoring with IDE validation
- **Full test suite**: 115 unit tests via pytest

---

## Installation

**Requirements:** Python 3.10+

```bash
# From the project root (inside the venv):
pip install -e .

# Including dev/test extras:
pip install -e ".[dev]"
```

Verify the installation:

```bash
umfrage --version
umfrage --help
```

---

## Quick Start

### 1. Author a questionnaire

Copy `config/questionnaire_sample.yaml` to `config/questionnaire.yaml`
(the working copy is gitignored) and edit it:

```bash
cp config/questionnaire_sample.yaml config/questionnaire.yaml
# Edit config/questionnaire.yaml in your editor
```

Add the following as the first line for IDE autocompletion (requires the
[YAML extension](https://marketplace.visualstudio.com/items?itemName=redhat.vscode-yaml)):

```yaml
# yaml-language-server: $schema=../docs/questionnaire.schema.json
```

See [docs/llm_guide.md](docs/llm_guide.md) for a full authoring guide
(especially useful when generating questionnaires with an AI assistant).

### 2. Validate the config

```bash
umfrage validate config/questionnaire.yaml
# [OK] 'config/questionnaire.yaml' is valid — 3 section(s), 9 question(s).
```

### 3. Generate the Excel file

```bash
umfrage generate config/questionnaire.yaml --metadata-file
# [OK] Questionnaire generated: ./annual-cooperation-survey-2024_questionnaire.xlsx
# [OK] Metadata file written:   ./annual-cooperation-survey-2024_metadata.yaml
```

The `--metadata-file` flag writes a `*_metadata.yaml` companion that embeds
the full questionnaire model, so `umfrage collect` can run without the
original `questionnaire.yaml`.

Specify an output directory with `--output-dir`:

```bash
umfrage generate config/questionnaire.yaml --output-dir out/ --metadata-file
```

### 4. Distribute

Send `*_questionnaire.xlsx` to each institution. Keep the `*_metadata.yaml`
in your responses folder.

### 5. Collect and aggregate

Place all returned `.xlsx` files in a single folder (together with the
`*_metadata.yaml`), then run:

```bash
umfrage collect responses/
# [OK] 'Annual Cooperation Survey 2024': 5/5 valid → responses/results_annual-...xlsx
```

The result file has:
- Row 1: questionnaire title
- Row 2: collection date and organizer info
- Row 4: column headers (Section | Q-ID | Question | Scale/Comment | Institution A | …)
- Subsequent rows: one per question; institution answers as columns
- Missing required answers are highlighted in the configured warning color

---

## CLI Reference

### `umfrage validate CONFIG`

Validate a questionnaire YAML config for syntax and completeness.

| Option | Description |
|---|---|
| `--style STYLE` | Path to `style.yaml` (optional) |

Exits with code **0** on success, **1** on error.

---

### `umfrage generate CONFIG`

Generate a protected Excel questionnaire from a YAML config.

| Option | Description |
|---|---|
| `--output-dir DIR` | Output directory (default: current directory) |
| `--style STYLE` | Path to `style.yaml` |
| `--metadata-file` | Write a `*_metadata.yaml` companion file |

---

### `umfrage collect RESPONSES_DIR`

Collect and aggregate returned files into result spreadsheets.

| Option | Description |
|---|---|
| `--config CONFIG` | Path to questionnaire YAML (optional if `*_metadata.yaml` present) |
| `--style STYLE` | Path to `style.yaml` |
| `--output-dir DIR` | Output directory (default: same as `RESPONSES_DIR`) |

Multiple questionnaires in one folder are handled automatically — one
`results_*.xlsx` is produced per questionnaire group found.

---

## Config File Reference

### questionnaire.yaml

See `config/questionnaire_sample.yaml` for a fully annotated template and
`docs/questionnaire.schema.json` for the JSON Schema.

**Top-level structure:**

```yaml
title: "Survey Title"
version: "1.0"

organizer:
  name: "Dr. Jane Smith"
  institution: "Research Institute"
  email: "j.smith@example.org"
  phone: "+1 555-0100"          # optional

respondent_fields:
  - label: "Name"
  - label: "Institution"
  - label: "Email"
    required: false             # default: true

sections:
  - title: "Section Name"
    questions:
      - id: "S1.Q1"             # unique, slug-safe
        text: "Question text"
        answer:
          type: scale           # scale | yes_no | freetext
          min_value: 1          # required for scale
          max_value: 5          # required for scale
          description: "1=poor, 5=excellent"  # optional hint
        comment: "Additional context"         # optional
        required: true          # default: true
```

**Answer types:**

| Type | Description | Excel behavior |
|---|---|---|
| `scale` | Integer in `[min_value, max_value]` | Whole-number data validation |
| `yes_no` | "Yes" or "No" | Dropdown list |
| `freetext` | Any text | No validation, open cell |

**Question ID rules:** alphanumeric, dots, hyphens, underscores; must start and
end with a letter or digit; globally unique across all sections.

---

### style.yaml

See `config/style.yaml` for a fully annotated template.

Key settings:

| Section | Controls |
|---|---|
| `header` | Title row colors and font |
| `section_header` | Section band row styling |
| `question_row` | Question cell styling and alternate row color |
| `answer_cell` | Answer cell styling (always white/unlocked) |
| `respondent_header` | "RESPONDENT INFORMATION" label row |
| `result_header` | Column headers in the result spreadsheet |
| `warning_color` | Background for missing required answers in results |
| `column_widths` | Character widths for ID, text, answer, comment columns |
| `protection_password` | Optional worksheet password (null = no password) |

All color values are 6-digit hex codes **without** a `#`.

---

## Project Structure

```
umfrage/
├── umfrage/
│   ├── cli.py            CLI entry point (validate, generate, collect)
│   ├── models.py         Pydantic domain models
│   ├── config_loader.py  YAML loading and Pydantic validation
│   ├── checker.py        Completeness checks (9 rules)
│   ├── generator.py      Excel questionnaire generation
│   ├── validator.py      Returned file validation
│   ├── collector.py      Multi-questionnaire aggregation
│   └── styles.py         openpyxl styling helpers
├── config/
│   ├── questionnaire_sample.yaml  Annotated template (tracked)
│   └── style.yaml                 Appearance config (tracked)
├── docs/
│   ├── llm_guide.md               AI/LLM authoring guide
│   └── questionnaire.schema.json  JSON Schema for IDE validation
├── tests/                         pytest test suite (115 tests)
├── examples/
│   └── questionnaire_example.yaml
├── pyproject.toml
├── LICENSE                        Apache 2.0
└── CHANGELOG.md
```

---

## Development

Run the test suite:

```bash
pytest tests/ -v
```

With coverage:

```bash
pytest tests/ --cov=umfrage --cov-report=term-missing
```

---

## Contributing

1. Fork the repository.
2. Create a feature branch.
3. Add tests for new functionality.
4. Ensure `pytest tests/` passes without failures.
5. Run `umfrage validate config/questionnaire_sample.yaml` to confirm examples still work.
6. Open a pull request.

---

## License

Apache License 2.0 — see [LICENSE](LICENSE) for details.
