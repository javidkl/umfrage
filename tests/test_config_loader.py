"""Tests for umfrage.config_loader."""

from __future__ import annotations

from pathlib import Path

import pytest

from umfrage.config_loader import ConfigError, load_questionnaire, load_style
from umfrage.models import AnswerType


VALID_QUESTIONNAIRE_YAML = """\
title: "Sample Survey"
version: "1.0"
organizer:
  name: "Alice"
  institution: "Org"
  email: "alice@example.org"
respondent_fields:
  - label: "Name"
  - label: "Institution"
sections:
  - title: "Section 1"
    questions:
      - id: "S1.Q1"
        text: "Rate this"
        answer:
          type: scale
          min_value: 1
          max_value: 5
"""

VALID_STYLE_YAML = """\
header:
  background_color: "003366"
  font_color: "FFFFFF"
  font_size: 14
  bold: true
"""


class TestLoadQuestionnaire:
    def test_valid_yaml_parses_correctly(self, tmp_path: Path) -> None:
        config = tmp_path / "q.yaml"
        config.write_text(VALID_QUESTIONNAIRE_YAML)
        q = load_questionnaire(config)
        assert q.title == "Sample Survey"
        assert q.version == "1.0"
        assert len(q.sections) == 1
        assert q.sections[0].questions[0].answer.type == AnswerType.SCALE

    def test_missing_file_raises_config_error(self) -> None:
        with pytest.raises(ConfigError, match="Cannot read"):
            load_questionnaire(Path("/nonexistent/does_not_exist.yaml"))

    def test_invalid_yaml_syntax_raises_config_error(self, tmp_path: Path) -> None:
        bad = tmp_path / "bad.yaml"
        bad.write_text("title: [unclosed bracket\n")
        with pytest.raises(ConfigError, match="YAML syntax error"):
            load_questionnaire(bad)

    def test_non_mapping_yaml_raises_config_error(self, tmp_path: Path) -> None:
        bad = tmp_path / "bad.yaml"
        bad.write_text("- item1\n- item2\n")
        with pytest.raises(ConfigError, match="must be a YAML mapping"):
            load_questionnaire(bad)

    def test_missing_title_raises_config_error(self, tmp_path: Path) -> None:
        yaml_text = VALID_QUESTIONNAIRE_YAML.replace('title: "Sample Survey"\n', "")
        bad = tmp_path / "bad.yaml"
        bad.write_text(yaml_text)
        with pytest.raises(ConfigError, match="invalid"):
            load_questionnaire(bad)

    def test_missing_organizer_raises_config_error(self, tmp_path: Path) -> None:
        lines = [
            ln for ln in VALID_QUESTIONNAIRE_YAML.splitlines()
            if not ln.startswith("organizer") and not ln.startswith("  name")
            and not ln.startswith("  institution") and not ln.startswith("  email")
        ]
        bad = tmp_path / "bad.yaml"
        bad.write_text("\n".join(lines))
        with pytest.raises(ConfigError):
            load_questionnaire(bad)

    def test_unknown_answer_type_raises_config_error(self, tmp_path: Path) -> None:
        bad_yaml = VALID_QUESTIONNAIRE_YAML.replace("type: scale", "type: unknown_type")
        bad = tmp_path / "bad.yaml"
        bad.write_text(bad_yaml)
        with pytest.raises(ConfigError):
            load_questionnaire(bad)

    def test_all_answer_types_accepted(self, tmp_path: Path) -> None:
        for answer_type in ("scale", "yes_no", "freetext"):
            extra = ""
            if answer_type == "scale":
                extra = "\n          min_value: 1\n          max_value: 5"
            yaml_text = VALID_QUESTIONNAIRE_YAML.replace(
                "          type: scale\n          min_value: 1\n          max_value: 5",
                f"          type: {answer_type}{extra}",
            )
            config = tmp_path / f"q_{answer_type}.yaml"
            config.write_text(yaml_text)
            q = load_questionnaire(config)
            assert q.sections[0].questions[0].answer.type.value == answer_type

    def test_optional_phone_field(self, tmp_path: Path) -> None:
        yaml_with_phone = VALID_QUESTIONNAIRE_YAML.replace(
            "  email: \"alice@example.org\"",
            "  email: \"alice@example.org\"\n  phone: \"+1 555-0100\"",
        )
        config = tmp_path / "q.yaml"
        config.write_text(yaml_with_phone)
        q = load_questionnaire(config)
        assert q.organizer.phone == "+1 555-0100"

    def test_optional_required_field_default(self, tmp_path: Path) -> None:
        config = tmp_path / "q.yaml"
        config.write_text(VALID_QUESTIONNAIRE_YAML)
        q = load_questionnaire(config)
        assert q.sections[0].questions[0].required is True

    def test_language_field_accepted(self, tmp_path: Path) -> None:
        yaml_with_lang = VALID_QUESTIONNAIRE_YAML + "language: \"de\"\n"
        config = tmp_path / "q.yaml"
        config.write_text(yaml_with_lang)
        q = load_questionnaire(config)
        assert q.language == "de"

    def test_language_defaults_to_english(self, tmp_path: Path) -> None:
        config = tmp_path / "q.yaml"
        config.write_text(VALID_QUESTIONNAIRE_YAML)
        q = load_questionnaire(config)
        assert q.language == "en"


class TestLoadStyle:
    def test_valid_style_yaml(self, tmp_path: Path) -> None:
        style_file = tmp_path / "style.yaml"
        style_file.write_text(VALID_STYLE_YAML)
        style = load_style(style_file)
        assert style.header.background_color == "003366"
        assert style.header.font_size == 14
        assert style.header.bold is True

    def test_font_size_too_large_raises(self, tmp_path: Path) -> None:
        bad = tmp_path / "style.yaml"
        bad.write_text("header:\n  font_size: 999\n")
        with pytest.raises(ConfigError):
            load_style(bad)

    def test_empty_style_yaml_returns_defaults(self, tmp_path: Path) -> None:
        empty = tmp_path / "style.yaml"
        empty.write_text("{}\n")
        style = load_style(empty)
        assert style.header.background_color == "1F4E79"

    def test_protection_password_accepted(self, tmp_path: Path) -> None:
        style_file = tmp_path / "style.yaml"
        style_file.write_text("protection_password: \"secret123\"\n")
        style = load_style(style_file)
        assert style.protection_password == "secret123"

    def test_null_protection_password(self, tmp_path: Path) -> None:
        style_file = tmp_path / "style.yaml"
        style_file.write_text("protection_password: null\n")
        style = load_style(style_file)
        assert style.protection_password is None
