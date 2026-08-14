"""Shared pytest fixtures for the umfrage test suite."""

from __future__ import annotations

from pathlib import Path

import pytest

from umfrage.models import (
    AnswerConfig,
    AnswerType,
    OrganizerInfo,
    Question,
    Questionnaire,
    RespondentField,
    Section,
    StyleConfig,
)


@pytest.fixture()
def sample_questionnaire() -> Questionnaire:
    """Minimal but complete questionnaire covering all four answer types."""
    return Questionnaire(
        title="Test Survey 2024",
        version="1.0",
        choice_lists={
            "satisfaction": ["Poor", "Fair", "Good", "Excellent"],
        },
        organizer=OrganizerInfo(
            name="Dr. Test Author",
            institution="Test Institute",
            email="test@example.org",
        ),
        respondent_fields=[
            RespondentField(label="Name"),
            RespondentField(label="Institution"),
            RespondentField(label="Email", required=False),
        ],
        sections=[
            Section(
                title="General",
                questions=[
                    Question(
                        id="G.Q1",
                        text="How satisfied are you overall?",
                        answer=AnswerConfig(
                            type=AnswerType.SCALE, min_value=1, max_value=5
                        ),
                        comment="1 = not satisfied, 5 = very satisfied",
                    ),
                    Question(
                        id="G.Q2",
                        text="Would you participate again?",
                        answer=AnswerConfig(type=AnswerType.YES_NO),
                        required=True,
                    ),
                    Question(
                        id="G.Q3",
                        text="Please share any additional comments.",
                        answer=AnswerConfig(type=AnswerType.FREETEXT),
                        required=False,
                    ),
                    Question(
                        id="G.Q4",
                        text="How would you rate the overall quality?",
                        answer=AnswerConfig(
                            type=AnswerType.CHOICES,
                            choices_ref="satisfaction",
                        ),
                        required=True,
                    ),
                ],
            ),
            Section(
                title="Technical Quality",
                questions=[
                    Question(
                        id="T.Q1",
                        text="Rate the technical quality (1–10).",
                        answer=AnswerConfig(
                            type=AnswerType.SCALE, min_value=1, max_value=10
                        ),
                    ),
                ],
            ),
        ],
    )


@pytest.fixture()
def sample_style() -> StyleConfig:
    """Default StyleConfig with all factory defaults."""
    return StyleConfig()


@pytest.fixture()
def generated_xlsx(
    tmp_path: Path,
    sample_questionnaire: Questionnaire,
    sample_style: StyleConfig,
) -> Path:
    """Generate a fresh questionnaire xlsx in tmp_path and return its path."""
    from umfrage.generator import generate_questionnaire

    out = tmp_path / "test_questionnaire.xlsx"
    generate_questionnaire(sample_questionnaire, sample_style, out)
    return out
