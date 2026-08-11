"""YAML configuration loader for the umfrage questionnaire tool.

Provides safe YAML parsing followed by Pydantic model validation,
surfacing all errors through a single ``ConfigError`` exception type.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml
from pydantic import ValidationError

from umfrage.models import Questionnaire, StyleConfig


class ConfigError(ValueError):
    """Raised when a configuration file cannot be loaded or is structurally invalid."""


def load_questionnaire(path: Path) -> Questionnaire:
    """Load and validate a questionnaire YAML configuration file.

    Args:
        path: Path to the questionnaire YAML file.

    Returns:
        A validated :class:`~umfrage.models.Questionnaire` instance.

    Raises:
        ConfigError: If the file cannot be read, parsed, or validated.
    """
    raw = _load_yaml(path)
    try:
        return Questionnaire.model_validate(raw)
    except ValidationError as exc:
        raise ConfigError(
            f"Questionnaire config '{path}' is invalid:\n{_format_validation_error(exc)}"
        ) from exc


def load_style(path: Path) -> StyleConfig:
    """Load and validate a style YAML configuration file.

    Args:
        path: Path to the style YAML file.

    Returns:
        A validated :class:`~umfrage.models.StyleConfig` instance.

    Raises:
        ConfigError: If the file cannot be read, parsed, or validated.
    """
    raw = _load_yaml(path)
    try:
        return StyleConfig.model_validate(raw)
    except ValidationError as exc:
        raise ConfigError(
            f"Style config '{path}' is invalid:\n{_format_validation_error(exc)}"
        ) from exc


def questionnaire_from_dict(data: dict[str, Any]) -> Questionnaire:
    """Reconstruct a Questionnaire from a plain dict (e.g. from a metadata YAML).

    Args:
        data: Dictionary matching the Questionnaire model structure.

    Returns:
        A validated :class:`~umfrage.models.Questionnaire` instance.

    Raises:
        ConfigError: If the dict cannot be validated.
    """
    try:
        return Questionnaire.model_validate(data)
    except ValidationError as exc:
        raise ConfigError(
            f"Embedded questionnaire data is invalid:\n{_format_validation_error(exc)}"
        ) from exc


# ── Internal helpers ──────────────────────────────────────────────────────────

def _load_yaml(path: Path) -> dict[str, Any]:
    """Read a YAML file and return the top-level mapping."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ConfigError(f"Cannot read config file '{path}': {exc}") from exc
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ConfigError(f"YAML syntax error in '{path}': {exc}") from exc
    if not isinstance(data, dict):
        raise ConfigError(
            f"Config file '{path}' must be a YAML mapping at the top level, "
            f"got {type(data).__name__}."
        )
    return data


def _format_validation_error(exc: ValidationError) -> str:
    """Format a Pydantic ValidationError into a human-readable string."""
    lines = []
    for error in exc.errors():
        loc = " -> ".join(str(p) for p in error["loc"])
        lines.append(f"  [{loc}] {error['msg']}")
    return "\n".join(lines)
