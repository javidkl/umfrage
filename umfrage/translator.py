"""Internationalisation (i18n) support for umfrage.

Translations are stored in YAML files under ``umfrage/i18n/``, one file per
language code (e.g. ``en.yaml``, ``de.yaml``).  To add a new language, create
a new file containing **every key** from ``en.yaml`` with translated values —
no Python changes required.

The ``language`` field in a questionnaire YAML file selects which translations
are used for all static UI labels in the generated Excel files.

Usage
-----
::

    from umfrage.translator import Translator, list_languages

    print(list_languages())              # ['de', 'en']

    t = Translator("de")
    t.t("col_question")                  # "Frage"
    t.t("hint_scale", min=1, max=5)      # "Skala: 1–5"
    t.yes_no_values()                    # ("Ja", "Nein")
    t.yes_no_formula()                   # '"Ja,Nein"'  (Excel list-validation)

Adding a new language
---------------------
1. Copy ``umfrage/i18n/en.yaml`` to ``umfrage/i18n/<code>.yaml``.
2. Translate every value (keep all keys identical).
3. Run ``umfrage validate`` — the checker will now accept the new code.
"""

from __future__ import annotations

from pathlib import Path

import yaml

_I18N_DIR = Path(__file__).parent / "i18n"


class TranslationError(ValueError):
    """Raised when a language is unavailable or a required key is missing."""


def list_languages() -> list[str]:
    """Return a sorted list of all available language codes.

    The list is derived from the ``*.yaml`` files present in
    ``umfrage/i18n/``.  Adding a new file there automatically registers the
    new language without any code changes.

    Returns:
        Sorted list of language code strings, e.g. ``['de', 'en']``.
    """
    return sorted(p.stem for p in _I18N_DIR.glob("*.yaml"))


class Translator:
    """Provides translated UI strings for a specific language.

    Args:
        language: A language code matching one of the YAML files in
            ``umfrage/i18n/`` (e.g. ``"en"``, ``"de"``).  Call
            :func:`list_languages` to discover available codes.

    Raises:
        TranslationError: If *language* is not available.
    """

    def __init__(self, language: str = "en") -> None:
        available = list_languages()
        if language not in available:
            raise TranslationError(
                f"Language '{language}' is not available. "
                f"Available languages: {', '.join(available)}."
            )
        lang_file = _I18N_DIR / f"{language}.yaml"
        with lang_file.open(encoding="utf-8") as fh:
            self._strings: dict[str, str] = yaml.safe_load(fh) or {}
        self.language = language

    def t(self, key: str, **kwargs: object) -> str:
        """Return the translated string for *key*.

        If *kwargs* are provided the value is formatted with
        ``str.format(**kwargs)`` before returning, enabling parametric
        strings such as ``hint_scale`` (``"Scale: {min}–{max}"``).

        Args:
            key: A translation key defined in the language YAML file.
            **kwargs: Named format arguments substituted into the string.

        Returns:
            The translated (and optionally formatted) string.

        Raises:
            TranslationError: If *key* is not present in the language file.
        """
        if key not in self._strings:
            raise TranslationError(
                f"Translation key '{key}' not found for language "
                f"'{self.language}'."
            )
        value = str(self._strings[key])
        return value.format(**kwargs) if kwargs else value

    def yes_no_values(self) -> tuple[str, str]:
        """Return the ``(yes_string, no_string)`` display pair for this language.

        These strings are used both for the Excel dropdown and for validating
        respondent input (case-insensitive comparison is applied by the
        validator).

        Example: English → ``("Yes", "No")``, German → ``("Ja", "Nein")``.
        """
        return self.t("yes"), self.t("no")

    def yes_no_formula(self) -> str:
        """Return the Excel data-validation list formula string for yes/no.

        The returned string is suitable as the ``formula1`` argument for an
        openpyxl ``DataValidation(type="list")``.

        Example: English → ``'"Yes,No"'``, German → ``'"Ja,Nein"'``.
        """
        yes, no = self.yes_no_values()
        return f'"{yes},{no}"'
