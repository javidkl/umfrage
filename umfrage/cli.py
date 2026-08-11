"""Command-line interface for the umfrage questionnaire tool.

Commands
--------
``umfrage validate <config.yaml>``
    Check a questionnaire YAML for syntax and completeness without generating
    any files. Exits with code 1 if errors are found.

``umfrage generate <config.yaml>``
    Validate the config and generate a protected ``.xlsx`` questionnaire file.
    Optionally write a companion ``*_metadata.yaml`` (``--metadata-file``).

``umfrage collect <responses-dir>``
    Scan a folder for returned ``.xlsx`` files, group them by questionnaire
    identity, validate each, and aggregate into one ``results_*.xlsx`` per
    questionnaire found.
"""

from __future__ import annotations

import sys
from pathlib import Path

import click

from umfrage.checker import check_questionnaire
from umfrage.collector import collect_all
from umfrage.config_loader import ConfigError, load_questionnaire, load_style
from umfrage.generator import generate_questionnaire, write_metadata_file
from umfrage.models import StyleConfig

# Resolved relative to the current working directory when no --style is given.
_DEFAULT_STYLE_PATH = Path("config/style.yaml")


def _load_style_or_default(style_path: Path | None) -> StyleConfig:
    """Load a StyleConfig from *style_path*, from the default path, or return defaults."""
    if style_path is not None:
        return load_style(style_path)
    if _DEFAULT_STYLE_PATH.exists():
        return load_style(_DEFAULT_STYLE_PATH)
    return StyleConfig()


@click.group()
@click.version_option(package_name="umfrage")
def main() -> None:
    """umfrage — Excel-based questionnaire generation and collection tool.

    Typical workflow:

    \b
      1. Author a questionnaire YAML (use config/questionnaire_sample.yaml
         as a template or docs/llm_guide.md as an LLM prompt).
      2. Validate:   umfrage validate config/questionnaire.yaml
      3. Generate:   umfrage generate config/questionnaire.yaml --metadata-file
      4. Distribute the .xlsx to respondents.
      5. Collect:    umfrage collect responses/
    """


# ── validate ──────────────────────────────────────────────────────────────────

@main.command("validate")
@click.argument(
    "config",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option(
    "--style",
    "style_path",
    metavar="STYLE",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help="Path to style.yaml. If omitted, config/style.yaml is tried, then defaults.",
)
def cmd_validate(config: Path, style_path: Path | None) -> None:
    """Validate a questionnaire YAML config for syntax and completeness.

    CONFIG is the path to the questionnaire YAML file to check.

    Runs Pydantic structural validation followed by nine completeness checks
    (unique IDs, slug-safe IDs, scale min/max, email format, etc.).
    Exits with code 1 if any errors are found; warnings are printed but do
    not cause a non-zero exit.
    """
    try:
        questionnaire = load_questionnaire(config)
    except ConfigError as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)

    if style_path:
        try:
            load_style(style_path)
            click.echo(f"  Style config OK: {style_path}")
        except ConfigError as exc:
            click.echo(f"[ERROR] Style config: {exc}", err=True)
            sys.exit(1)

    check_result = check_questionnaire(questionnaire)

    for warning in check_result.warnings:
        click.echo(f"[WARNING] {warning}")

    if not check_result.is_valid:
        click.echo(
            f"\n[FAIL] {len(check_result.errors)} error(s) found in '{config}':",
            err=True,
        )
        for error in check_result.errors:
            click.echo(f"  • {error}", err=True)
        sys.exit(1)

    sections = len(questionnaire.sections)
    questions = len(questionnaire.all_questions())
    click.echo(
        f"[OK] '{config}' is valid — "
        f"{sections} section(s), {questions} question(s)."
    )


# ── generate ──────────────────────────────────────────────────────────────────

@main.command("generate")
@click.argument(
    "config",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option(
    "--output-dir", "-o",
    metavar="DIR",
    type=click.Path(file_okay=False, path_type=Path),
    default=None,
    help="Directory where the .xlsx file is written. Defaults to the current directory.",
)
@click.option(
    "--style",
    "style_path",
    metavar="STYLE",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help="Path to style.yaml. If omitted, config/style.yaml is tried, then defaults.",
)
@click.option(
    "--metadata-file",
    is_flag=True,
    default=False,
    help=(
        "Write a companion *_metadata.yaml file alongside the questionnaire. "
        "The file embeds the full questionnaire model so that 'umfrage collect' "
        "can be run without the original questionnaire.yaml."
    ),
)
def cmd_generate(
    config: Path,
    output_dir: Path | None,
    style_path: Path | None,
    metadata_file: bool,
) -> None:
    """Generate a protected Excel questionnaire file from a YAML config.

    CONFIG is the path to the questionnaire YAML file.

    The config is validated before generation. Use 'umfrage validate' to check
    the config without generating any files.

    The output filename is derived from the questionnaire title:
    ``{slug}_questionnaire.xlsx``.
    """
    try:
        questionnaire = load_questionnaire(config)
    except ConfigError as exc:
        click.echo(f"[ERROR] {exc}", err=True)
        sys.exit(1)

    check_result = check_questionnaire(questionnaire)
    if not check_result.is_valid:
        click.echo(
            "[ERROR] Config validation failed. Fix the following errors before generating:",
            err=True,
        )
        for error in check_result.errors:
            click.echo(f"  • {error}", err=True)
        sys.exit(1)

    for warning in check_result.warnings:
        click.echo(f"[WARNING] {warning}")

    style = _load_style_or_default(style_path)
    out_dir = output_dir or Path(".")
    qid = questionnaire.questionnaire_id()
    xlsx_path = out_dir / f"{qid}_questionnaire.xlsx"

    try:
        generate_questionnaire(questionnaire, style, xlsx_path)
    except Exception as exc:
        click.echo(f"[ERROR] Failed to generate Excel file: {exc}", err=True)
        sys.exit(1)

    click.echo(f"[OK] Questionnaire generated: {xlsx_path}")

    if metadata_file:
        meta_path = out_dir / f"{qid}_metadata.yaml"
        try:
            write_metadata_file(questionnaire, meta_path)
            click.echo(f"[OK] Metadata file written:  {meta_path}")
        except Exception as exc:
            click.echo(f"[WARNING] Could not write metadata file: {exc}")


# ── collect ───────────────────────────────────────────────────────────────────

@main.command("collect")
@click.argument(
    "responses_dir",
    metavar="RESPONSES_DIR",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
@click.option(
    "--config", "config_path",
    metavar="CONFIG",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help=(
        "Path to questionnaire YAML. Optional if *_metadata.yaml files are "
        "present in RESPONSES_DIR. If provided, overrides metadata discovery "
        "for all questionnaire groups."
    ),
)
@click.option(
    "--style",
    "style_path",
    metavar="STYLE",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    default=None,
    help="Path to style.yaml. If omitted, config/style.yaml is tried, then defaults.",
)
@click.option(
    "--output-dir", "-o",
    metavar="DIR",
    type=click.Path(file_okay=False, path_type=Path),
    default=None,
    help=(
        "Directory where result files are written. "
        "Defaults to RESPONSES_DIR."
    ),
)
def cmd_collect(
    responses_dir: Path,
    config_path: Path | None,
    style_path: Path | None,
    output_dir: Path | None,
) -> None:
    """Collect and aggregate returned questionnaire files into result spreadsheets.

    RESPONSES_DIR is the folder containing the returned .xlsx files.

    Response files are automatically grouped by questionnaire identity
    (config hash stored in each file's hidden _meta sheet). One
    ``results_*.xlsx`` file is produced per questionnaire group found.

    Config discovery order:

    \b
      1. --config flag (overrides everything)
      2. *_metadata.yaml files in RESPONSES_DIR (auto-discovered by hash)
    """
    config_override = None
    if config_path:
        try:
            config_override = load_questionnaire(config_path)
        except ConfigError as exc:
            click.echo(f"[ERROR] {exc}", err=True)
            sys.exit(1)

    style = _load_style_or_default(style_path)
    out_dir = output_dir or responses_dir

    try:
        summaries = collect_all(responses_dir, style, out_dir, config_override)
    except Exception as exc:
        click.echo(f"[ERROR] Collection failed unexpectedly: {exc}", err=True)
        sys.exit(1)

    if not summaries:
        click.echo(
            "[WARNING] No processable questionnaire groups found.\n"
            "  • Ensure response files are .xlsx format.\n"
            "  • Ensure *_metadata.yaml files are present, or pass --config."
        )
        return

    for s in summaries:
        status = "OK" if s.skipped_count == 0 else "WARN"
        parts = [
            f"[{status}] '{s.questionnaire_title}': "
            f"{s.valid_count}/{s.total_files} valid"
        ]
        if s.skipped_count:
            parts.append(f", {s.skipped_count} skipped")
        if s.output_path:
            parts.append(f" → {s.output_path}")
        click.echo("".join(parts))

        for skipped_path, errors in s.skipped_files:
            click.echo(f"       Skipped: {skipped_path.name}")
            for error in errors:
                click.echo(f"         • {error}")
