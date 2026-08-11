"""Tests for umfrage.translator."""

from __future__ import annotations

import pytest

from umfrage.translator import Translator, TranslationError, list_languages


class TestListLanguages:
    def test_returns_at_least_en_and_de(self) -> None:
        langs = list_languages()
        assert "en" in langs
        assert "de" in langs

    def test_returns_sorted_list(self) -> None:
        langs = list_languages()
        assert langs == sorted(langs)

    def test_returns_list_type(self) -> None:
        assert isinstance(list_languages(), list)


class TestTranslatorCreation:
    def test_creates_english_translator(self) -> None:
        t = Translator("en")
        assert t.language == "en"

    def test_creates_german_translator(self) -> None:
        t = Translator("de")
        assert t.language == "de"

    def test_unknown_language_raises(self) -> None:
        with pytest.raises(TranslationError, match="not available"):
            Translator("fr")

    def test_empty_language_raises(self) -> None:
        with pytest.raises(TranslationError):
            Translator("")

    def test_default_language_is_english(self) -> None:
        t = Translator()
        assert t.language == "en"


class TestTranslatorLookup:
    def test_english_column_headers(self) -> None:
        t = Translator("en")
        assert t.t("col_id") == "ID"
        assert t.t("col_question") == "Question"
        assert t.t("col_answer") == "Answer"
        assert t.t("col_scale_comment") == "Scale / Comment"

    def test_german_column_headers(self) -> None:
        t = Translator("de")
        assert t.t("col_id") == "ID"
        assert t.t("col_question") == "Frage"
        assert t.t("col_answer") == "Antwort"
        assert t.t("col_scale_comment") == "Skala / Kommentar"

    def test_english_respondent_header(self) -> None:
        t = Translator("en")
        assert t.t("label_respondent_information") == "RESPONDENT INFORMATION"

    def test_german_respondent_header(self) -> None:
        t = Translator("de")
        label = t.t("label_respondent_information")
        assert "ANGABEN" in label

    def test_missing_key_raises(self) -> None:
        t = Translator("en")
        with pytest.raises(TranslationError, match="not found"):
            t.t("nonexistent_key_xyz")

    def test_parametric_hint_scale(self) -> None:
        t_en = Translator("en")
        result = t_en.t("hint_scale", min=1, max=5)
        assert "1" in result and "5" in result

        t_de = Translator("de")
        result_de = t_de.t("hint_scale", min=1, max=10)
        assert "1" in result_de and "10" in result_de

    def test_parametric_dv_scale_error(self) -> None:
        t = Translator("en")
        msg = t.t("dv_scale_error", min=1, max=5)
        assert "1" in msg and "5" in msg

    def test_parametric_dv_yesno_error(self) -> None:
        t = Translator("de")
        yes, no = t.yes_no_values()
        msg = t.t("dv_yesno_error", yes=yes, no=no)
        assert "Ja" in msg and "Nein" in msg


class TestYesNoHelpers:
    def test_english_yes_no_values(self) -> None:
        t = Translator("en")
        yes, no = t.yes_no_values()
        assert yes == "Yes"
        assert no == "No"

    def test_german_yes_no_values(self) -> None:
        t = Translator("de")
        yes, no = t.yes_no_values()
        assert yes == "Ja"
        assert no == "Nein"

    def test_english_yes_no_formula(self) -> None:
        t = Translator("en")
        assert t.yes_no_formula() == '"Yes,No"'

    def test_german_yes_no_formula(self) -> None:
        t = Translator("de")
        assert t.yes_no_formula() == '"Ja,Nein"'


class TestAllLanguagesComplete:
    """Verify every available language file contains all required keys."""

    _REQUIRED_KEYS = [
        "label_respondent_information", "label_organizer", "label_institution",
        "label_email", "label_phone",
        "col_id", "col_question", "col_answer", "col_scale_comment",
        "hint_scale", "hint_yesno", "hint_freetext",
        "yes", "no",
        "dv_error_title", "dv_scale_error", "dv_yesno_error",
        "result_title_suffix", "result_collected", "result_organizer",
        "result_col_section", "result_col_qid", "result_col_question",
        "result_col_scale_comment",
    ]

    @pytest.mark.parametrize("lang", list_languages())
    def test_language_has_all_required_keys(self, lang: str) -> None:
        t = Translator(lang)
        for key in self._REQUIRED_KEYS:
            assert t.t(key) is not None, (
                f"Key '{key}' missing or None in language '{lang}'"
            )
