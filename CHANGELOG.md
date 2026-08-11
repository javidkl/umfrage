# Changelog

All notable changes to **umfrage** will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [0.1.0] – 2026-08-11

### Added
- `umfrage validate` command: syntax and completeness check for questionnaire YAML configs
- `umfrage generate` command: creates a protected Excel questionnaire from a YAML config
  - Worksheet protection with optional password
  - Data validation dropdowns for scale (1–N) and yes/no answer types
  - Hidden `_meta` sheet storing structural metadata for later validation
  - Optional `--metadata-file` flag: writes a `*_metadata.yaml` companion file
    embedding the full questionnaire model (enables config-free collection)
- `umfrage collect` command: aggregates returned Excel files into a single results spreadsheet
  - Automatic grouping of files by questionnaire identity (config hash)
  - Supports multiple different questionnaires in one folder
  - Config auto-discovery from `*_metadata.yaml` files; `--config` always optional
  - Warning color highlighting for missing or invalid answers in the result file
- YAML-based questionnaire config (`config/questionnaire_sample.yaml`)
- YAML-based style/appearance config (`config/style.yaml`) with optional protection password
- JSON Schema for questionnaire YAML (`docs/questionnaire.schema.json`)
- LLM/AI authoring guide (`docs/llm_guide.md`)
- Full unit test suite (`tests/`) with pytest
- Apache 2.0 license
