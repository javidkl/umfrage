"""Tests for umfrage.checker."""

from __future__ import annotations

import pytest

from umfrage.checker import CheckResult, check_questionnaire
from umfrage.models import (
    AnswerConfig,
    AnswerType,
    OrganizerInfo,
    Question,
    Questionnaire,
    RespondentField,
    Section,
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_scale_q(qid: str, min_v: int | None = 1, max_v: int | None = 5) -> Question:
    return Question(
        id=qid,
        text=f"Question {qid}",
        answer=AnswerConfig(type=AnswerType.SCALE, min_value=min_v, max_value=max_v),
    )


def _make_yesno_q(qid: str) -> Question:
    return Question(
        id=qid,
        text=f"Question {qid}",
        answer=AnswerConfig(type=AnswerType.YES_NO),
    )


def _make_freetext_q(qid: str) -> Question:
    return Question(
        id=qid,
        text=f"Question {qid}",
        answer=AnswerConfig(type=AnswerType.FREETEXT),
    )


def _base_questionnaire(**overrides) -> Questionnaire:
    defaults: dict = dict(
        title="Test Survey",
        organizer=OrganizerInfo(
            name="Alice", institution="Org", email="alice@example.org"
        ),
        respondent_fields=[RespondentField(label="Name")],
        sections=[
            Section(title="S1", questions=[_make_scale_q("S1.Q1")])
        ],
    )
    defaults.update(overrides)
    return Questionnaire(**defaults)


# ── Tests ─────────────────────────────────────────────────────────────────────

class TestValidQuestionnaire:
    def test_passes_with_all_types(self, sample_questionnaire: Questionnaire) -> None:
        result = check_questionnaire(sample_questionnaire)
        assert result.is_valid
        assert result.errors == []

    def test_passes_minimal_questionnaire(self) -> None:
        q = _base_questionnaire()
        result = check_questionnaire(q)
        assert result.is_valid


class TestDuplicateIds:
    def test_duplicate_in_same_section(self) -> None:
        q = _base_questionnaire(
            sections=[
                Section(
                    title="S1",
                    questions=[
                        _make_scale_q("Q1"),
                        _make_yesno_q("Q1"),  # duplicate
                    ],
                )
            ]
        )
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("Duplicate" in e and "Q1" in e for e in result.errors)

    def test_duplicate_across_sections(self) -> None:
        q = _base_questionnaire(
            sections=[
                Section(title="S1", questions=[_make_scale_q("SHARED.ID")]),
                Section(title="S2", questions=[_make_yesno_q("SHARED.ID")]),
            ]
        )
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("Duplicate" in e for e in result.errors)

    def test_unique_ids_pass(self) -> None:
        q = _base_questionnaire(
            sections=[
                Section(
                    title="S1",
                    questions=[_make_scale_q("Q1"), _make_yesno_q("Q2")],
                )
            ]
        )
        assert check_questionnaire(q).is_valid


class TestSlugSafeIds:
    @pytest.mark.parametrize("bad_id", [
        "Q 1",         # space
        "Q!1",         # exclamation
        "Q/1",         # slash
        " Q1",         # leading space
        "Q1 ",         # trailing space
        ".Q1",         # leading dot
        "Q1.",         # trailing dot
        "-Q1",         # leading hyphen
        "Q1-",         # trailing hyphen
    ])
    def test_invalid_id_formats(self, bad_id: str) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_yesno_q(bad_id)])]
        )
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("slug-safe" in e for e in result.errors)

    @pytest.mark.parametrize("good_id", [
        "Q1", "S1.Q1", "general.q01", "tech-1", "A_B_C", "a1b2c3",
        "S1.Q1.sub", "AB-CD",
    ])
    def test_valid_id_formats(self, good_id: str) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_yesno_q(good_id)])]
        )
        result = check_questionnaire(q)
        # Only slug errors — should not error on slug for valid IDs
        slug_errors = [e for e in result.errors if "slug-safe" in e]
        assert slug_errors == []


class TestScaleConstraints:
    def test_scale_missing_min_value(self) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_scale_q("Q1", min_v=None, max_v=5)])]
        )
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("min_value" in e for e in result.errors)

    def test_scale_missing_max_value(self) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_scale_q("Q1", min_v=1, max_v=None)])]
        )
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("max_value" in e for e in result.errors)

    def test_scale_min_equals_max(self) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_scale_q("Q1", min_v=5, max_v=5)])]
        )
        result = check_questionnaire(q)
        assert not result.is_valid

    def test_scale_min_greater_than_max(self) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_scale_q("Q1", min_v=10, max_v=1)])]
        )
        result = check_questionnaire(q)
        assert not result.is_valid

    def test_valid_scale_range(self) -> None:
        q = _base_questionnaire(
            sections=[Section(title="S1", questions=[_make_scale_q("Q1", min_v=0, max_v=100)])]
        )
        assert check_questionnaire(q).is_valid


class TestAnswerTypeWarnings:
    def test_yesno_with_min_max_gives_warning_not_error(self) -> None:
        q = _base_questionnaire(
            sections=[
                Section(
                    title="S1",
                    questions=[
                        Question(
                            id="Q1",
                            text="Yes/no?",
                            answer=AnswerConfig(
                                type=AnswerType.YES_NO, min_value=1, max_value=5
                            ),
                        )
                    ],
                )
            ]
        )
        result = check_questionnaire(q)
        assert result.is_valid
        assert any("YES_NO" in w for w in result.warnings)

    def test_freetext_with_min_max_gives_warning_not_error(self) -> None:
        q = _base_questionnaire(
            sections=[
                Section(
                    title="S1",
                    questions=[
                        Question(
                            id="Q1",
                            text="Text?",
                            answer=AnswerConfig(
                                type=AnswerType.FREETEXT, min_value=0, max_value=500
                            ),
                        )
                    ],
                )
            ]
        )
        result = check_questionnaire(q)
        assert result.is_valid
        assert any("FREETEXT" in w for w in result.warnings)


class TestOrganizerEmail:
    def test_invalid_email_format(self) -> None:
        q = _base_questionnaire(
            organizer=OrganizerInfo(
                name="A", institution="B", email="not-an-email"
            )
        )
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("email" in e.lower() for e in result.errors)

    def test_valid_email_passes(self) -> None:
        q = _base_questionnaire(
            organizer=OrganizerInfo(
                name="A", institution="B", email="user@domain.org"
            )
        )
        assert check_questionnaire(q).is_valid


class TestRespondentFields:
    def test_empty_respondent_fields_list(self) -> None:
        q = _base_questionnaire()
        q.respondent_fields = []  # bypass Pydantic min_length for checker test
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any("respondent_fields" in e for e in result.errors)


class TestLanguageCheck:
    def test_unknown_language_gives_error(self) -> None:
        q = _base_questionnaire()
        q.language = "fr"
        result = check_questionnaire(q)
        assert not result.is_valid
        assert any(
            "fr" in e and ("language" in e.lower() or "Language" in e)
            for e in result.errors
        )

    def test_english_language_passes(self) -> None:
        q = _base_questionnaire()
        q.language = "en"
        assert check_questionnaire(q).is_valid

    def test_german_language_passes(self) -> None:
        q = _base_questionnaire()
        q.language = "de"
        assert check_questionnaire(q).is_valid

    def test_default_language_is_english_and_passes(self) -> None:
        q = _base_questionnaire()
        assert q.language == "en"
        assert check_questionnaire(q).is_valid
