"""Domain models for the umfrage questionnaire tool.

All models use Pydantic v2 for automatic validation on construction.
"""

from __future__ import annotations

import hashlib
import json
import re
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class AnswerType(str, Enum):
    """Supported answer types for questionnaire questions."""

    SCALE = "scale"
    YES_NO = "yes_no"
    FREETEXT = "freetext"
    CHOICES = "choices"


class AnswerConfig(BaseModel):
    """Configuration for a question's expected answer."""

    type: AnswerType
    min_value: int | None = Field(
        default=None,
        description="Lower bound (inclusive) for SCALE type.",
    )
    max_value: int | None = Field(
        default=None,
        description="Upper bound (inclusive) for SCALE type.",
    )
    description: str | None = Field(
        default=None,
        description="Optional hint text shown in the Excel comment/scale column.",
    )
    choices: list[str] | None = Field(
        default=None,
        description=(
            "Inline list of options for CHOICES type. "
            "Provide either this or choices_ref, not both."
        ),
    )
    choices_ref: str | None = Field(
        default=None,
        description=(
            "Name of a list defined in the questionnaire-level choice_lists dict. "
            "Provide either this or choices, not both."
        ),
    )
    show_choices_in_comment: bool = Field(
        default=True,
        description=(
            "CHOICES type only. When false, the options are not listed in the "
            "Scale/Comment column. Useful for long lists where the column would "
            "become too wide. The description field (if set) is still shown."
        ),
    )


class Question(BaseModel):
    """A single question within a section."""

    id: str = Field(
        description=(
            "Unique question identifier, slug-safe "
            "(letters, digits, dots, hyphens, underscores only). "
            "Example: 'S1.Q1', 'general.q01'."
        )
    )
    text: str = Field(description="The question text shown to respondents.")
    answer: AnswerConfig
    comment: str | None = Field(
        default=None,
        description="Additional context or scale-label hint shown in the Excel comment column.",
    )
    required: bool = Field(
        default=True,
        description="Whether an answer is mandatory. Empty required answers fail validation.",
    )


class Section(BaseModel):
    """A group of related questions forming a chapter in the questionnaire."""

    title: str
    questions: list[Question] = Field(min_length=1)


class OrganizerInfo(BaseModel):
    """Information about the questionnaire organizer."""

    name: str
    institution: str
    email: str
    phone: str | None = None


class RespondentField(BaseModel):
    """A metadata field that each respondent must fill in (e.g. name, institution)."""

    label: str
    required: bool = True


class Questionnaire(BaseModel):
    """Root model representing a complete questionnaire configuration."""

    title: str
    version: str = "1.0"
    language: str = Field(
        default="en",
        description=(
            "Language code for UI labels in the generated Excel file. "
            "Must match a file in umfrage/i18n/ (e.g. 'en', 'de'). "
            "Run 'umfrage validate' to check availability. "
            "Add a YAML file to umfrage/i18n/ to register a new language."
        ),
    )
    choice_lists: dict[str, list[str]] = Field(
        default_factory=dict,
        description=(
            "Named reusable choice lists. Keys are arbitrary identifiers; values are "
            "the list of option strings. Reference a list from a question with "
            "answer.choices_ref: <name>. May also be defined inline per question with "
            "answer.choices: [...]."
        ),
    )
    organizer: OrganizerInfo
    respondent_fields: list[RespondentField] = Field(min_length=1)
    sections: list[Section] = Field(min_length=1)

    def resolved_choices(self, answer_config: AnswerConfig) -> list[str] | None:
        """Return the resolved choice list for a CHOICES answer config.

        Inline ``choices`` takes precedence over ``choices_ref``.
        Returns ``None`` when neither is set or the ref key is unknown.
        """
        if answer_config.choices:
            return answer_config.choices
        if answer_config.choices_ref:
            return self.choice_lists.get(answer_config.choices_ref)
        return None

    def questionnaire_id(self) -> str:
        """Return a URL/filename-safe identifier derived from the title (max 50 chars)."""
        slug = re.sub(r"[^a-zA-Z0-9]+", "-", self.title.lower()).strip("-")
        return slug[:50]

    def all_questions(self) -> list[Question]:
        """Return all questions across all sections in definition order."""
        return [q for section in self.sections for q in section.questions]

    def config_hash(self) -> str:
        """Return a SHA-256 hex digest of the canonical serialized model.

        Used to identify whether two Excel files were generated from the same
        questionnaire configuration and to match responses to their metadata YAML.
        """
        data = self.model_dump_json(exclude_none=False)
        return hashlib.sha256(data.encode()).hexdigest()


# ── Style models ──────────────────────────────────────────────────────────────

class CellStyle(BaseModel):
    """Visual style settings for a category of Excel cells."""

    background_color: str = Field(
        default="FFFFFF",
        description="Six-digit hex color code without '#', e.g. 'FFFFFF'.",
    )
    font_color: str = Field(
        default="000000",
        description="Six-digit hex color code without '#', e.g. '000000'.",
    )
    font_size: int = Field(default=10, ge=6, le=72)
    bold: bool = False
    italic: bool = False
    alternate_color: str | None = Field(
        default=None,
        description="Optional alternate background color for odd question rows (striping).",
    )


class ColumnWidths(BaseModel):
    """Width (in character units) for each column of the questionnaire sheet."""

    question_id: int = Field(default=10, description="Column A: question ID.")
    question_text: int = Field(default=55, description="Column B: question text.")
    answer: int = Field(default=20, description="Column C: answer input cell.")
    comment: int = Field(default=35, description="Column D: scale labels / comments.")


class StyleConfig(BaseModel):
    """Root model for the Excel appearance configuration (style.yaml)."""

    header: CellStyle = Field(
        default_factory=lambda: CellStyle(
            background_color="1F4E79", font_color="FFFFFF", font_size=14, bold=True
        )
    )
    section_header: CellStyle = Field(
        default_factory=lambda: CellStyle(
            background_color="2E75B6", font_color="FFFFFF", font_size=11, bold=True
        )
    )
    question_row: CellStyle = Field(
        default_factory=lambda: CellStyle(
            background_color="D6E4F0", alternate_color="EEF4FA"
        )
    )
    answer_cell: CellStyle = Field(default_factory=CellStyle)
    respondent_header: CellStyle = Field(
        default_factory=lambda: CellStyle(
            background_color="375623", font_color="FFFFFF", bold=True
        )
    )
    result_header: CellStyle = Field(
        default_factory=lambda: CellStyle(
            background_color="375623", font_color="FFFFFF", bold=True
        )
    )
    warning_color: str = Field(
        default="FFCCCC",
        description="Background hex color for missing/invalid answer cells in the result file.",
    )
    missing_answer_marker: str = Field(
        default="XXXXX",
        description=(
            "Placeholder text written into result cells for missing answers when a "
            "file that failed validation is force-included during collection."
        ),
    )
    column_widths: ColumnWidths = Field(default_factory=ColumnWidths)
    protection_password: str | None = Field(
        default=None,
        description=(
            "Optional password for worksheet protection. "
            "If null (default), the sheet is protected without a password "
            "(respondents can still unprotect via the Excel UI, which is acceptable "
            "since the goal is guidance, not strict enforcement)."
        ),
    )
    show_footer: bool = Field(
        default=True,
        description=(
            "When true (default), a footer row is appended to the questionnaire sheet "
            "containing the tool name and a link to https://github.com/scinnod/umfrage. "
            "Set to false to omit the footer entirely."
        ),
    )
