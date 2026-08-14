"""Completeness and consistency checks for questionnaire configurations.

These checks go beyond Pydantic's structural validation to catch logical
errors such as duplicate question IDs, scale misconfiguration, or invalid
email addresses. They are run automatically by ``umfrage generate`` and can
be invoked explicitly via ``umfrage validate``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from umfrage.models import AnswerType, Questionnaire
from umfrage.translator import list_languages

# Question IDs must start and end with an alphanumeric character and may
# contain alphanumerics, dots, hyphens, and underscores in between.
_SLUG_RE = re.compile(r"^[a-zA-Z0-9]([a-zA-Z0-9._-]*[a-zA-Z0-9])?$")

# Minimal email sanity check (not RFC 5321 full validation, but catches obvious typos).
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


@dataclass
class CheckResult:
    """Outcome of running completeness checks on a questionnaire."""

    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        """True when no errors were found (warnings do not affect validity)."""
        return len(self.errors) == 0

    def add_error(self, message: str) -> None:
        self.errors.append(message)

    def add_warning(self, message: str) -> None:
        self.warnings.append(message)


def check_questionnaire(q: Questionnaire) -> CheckResult:
    """Run all completeness and consistency checks on a questionnaire model.

    Checks performed:

    1. At least one section exists.
    2. Each section has at least one question (Pydantic enforces this via
       ``min_length=1``, but the check is repeated here for clarity).
    3. All question IDs are unique across the entire questionnaire.
    4. Each question ID is slug-safe (alphanumeric, dots, hyphens, underscores).
    5. SCALE answers have ``min_value`` **and** ``max_value`` set, and ``min < max``.
    6. YES_NO answers have no ``min_value``/``max_value`` (warning if present).
    7. FREETEXT answers have no ``min_value``/``max_value`` (warning if present).
    8. ``respondent_fields`` list is non-empty.
    9. Organizer email passes a basic format check.
    10. Language code is available.
    11. CHOICES answers have exactly one of ``choices`` or ``choices_ref`` set.
    12. ``choices_ref`` references a key that exists in ``q.choice_lists``.
    13. Resolved choices list has at least 2 items.
    14. Comma-joined choices fit within Excel's 255-char DV formula limit.
    15. Choices list contains no duplicate values (case-insensitive, warning only).
    16. CHOICES answers have no ``min_value``/``max_value`` (warning if present).

    Args:
        q: A Pydantic-validated :class:`~umfrage.models.Questionnaire`.

    Returns:
        A :class:`CheckResult` with ``errors`` (block generation) and
        ``warnings`` (informational, do not block generation).
    """
    result = CheckResult()

    # Check 8 — respondent_fields non-empty (Pydantic min_length=1 catches it
    # structurally, but the field may be bypassed via direct model mutation).
    if not q.respondent_fields:
        result.add_error("respondent_fields must contain at least one field.")

    # Check 9 — organizer email format
    if not _EMAIL_RE.match(q.organizer.email):
        result.add_error(
            f"Organizer email '{q.organizer.email}' does not look like a valid "
            "email address."
        )

    # Check 10 — language availability
    available_languages = list_languages()
    if q.language not in available_languages:
        result.add_error(
            f"Language '{q.language}' is not available. "
            f"Available languages: {', '.join(available_languages)}. "
            "To add a new language, create umfrage/i18n/<code>.yaml with all "
            "required keys (copy umfrage/i18n/en.yaml as a template)."
        )

    # Check 1 — at least one section (Pydantic enforces min_length=1)
    if not q.sections:
        result.add_error("Questionnaire must have at least one section.")
        return result  # Cannot proceed further without sections

    seen_ids: set[str] = set()

    for section_idx, section in enumerate(q.sections, start=1):
        section_label = f"Section {section_idx} ('{section.title}')"

        # Check 2 — each section has questions
        if not section.questions:
            result.add_error(f"{section_label}: must contain at least one question.")
            continue

        for q_obj in section.questions:
            qid = q_obj.id
            ans = q_obj.answer

            # Check 3 — unique IDs
            if qid in seen_ids:
                result.add_error(
                    f"Duplicate question ID '{qid}' found in {section_label}."
                )
            else:
                seen_ids.add(qid)

            # Check 4 — slug-safe IDs
            if not _SLUG_RE.match(qid):
                result.add_error(
                    f"Question ID '{qid}' in {section_label} is not slug-safe. "
                    "Use only letters, digits, dots, hyphens, and underscores "
                    "(must start and end with a letter or digit)."
                )

            # Checks 5, 6, 7 — answer type consistency
            if ans.type == AnswerType.SCALE:
                if ans.min_value is None or ans.max_value is None:
                    result.add_error(
                        f"Question '{qid}' (SCALE): both min_value and max_value "
                        "must be set."
                    )
                elif ans.min_value >= ans.max_value:
                    result.add_error(
                        f"Question '{qid}' (SCALE): min_value ({ans.min_value}) "
                        f"must be strictly less than max_value ({ans.max_value})."
                    )

            elif ans.type == AnswerType.YES_NO:
                if ans.min_value is not None or ans.max_value is not None:
                    result.add_warning(
                        f"Question '{qid}' (YES_NO): min_value/max_value are "
                        "ignored for the YES_NO type and will be excluded from "
                        "the generated Excel file."
                    )

            elif ans.type == AnswerType.FREETEXT:
                if ans.min_value is not None or ans.max_value is not None:
                    result.add_warning(
                        f"Question '{qid}' (FREETEXT): min_value/max_value are "
                        "ignored for the FREETEXT type and will be excluded from "
                        "the generated Excel file."
                    )

            elif ans.type == AnswerType.CHOICES:
                # Check 16 — min/max irrelevant for CHOICES
                if ans.min_value is not None or ans.max_value is not None:
                    result.add_warning(
                        f"Question '{qid}' (CHOICES): min_value/max_value are "
                        "ignored for the CHOICES type."
                    )

                # Check 11 — exactly one of choices / choices_ref must be set
                has_inline = bool(ans.choices)
                has_ref = bool(ans.choices_ref)
                if not has_inline and not has_ref:
                    result.add_error(
                        f"Question '{qid}' (CHOICES): either 'choices' (inline list) "
                        "or 'choices_ref' (named list from choice_lists) must be set."
                    )
                    continue
                if has_inline and has_ref:
                    result.add_error(
                        f"Question '{qid}' (CHOICES): set either 'choices' or "
                        "'choices_ref', not both."
                    )
                    continue

                # Check 12 — choices_ref must resolve
                if has_ref:
                    if ans.choices_ref not in q.choice_lists:
                        result.add_error(
                            f"Question '{qid}' (CHOICES): choices_ref "
                            f"'{ans.choices_ref}' is not defined in choice_lists. "
                            f"Available keys: {sorted(q.choice_lists.keys()) or '(none)'}."
                        )
                        continue

                resolved = q.resolved_choices(ans)
                if resolved is None:
                    continue  # already reported above

                # Check 13 — at least 2 choices
                if len(resolved) < 2:
                    result.add_error(
                        f"Question '{qid}' (CHOICES): at least 2 options are required, "
                        f"got {len(resolved)}."
                    )

                # Check 14 — Excel DV formula length limit (255 chars for the CSV string)
                formula_str = ",".join(resolved)
                if len(formula_str) > 255:
                    result.add_error(
                        f"Question '{qid}' (CHOICES): the comma-joined choices string "
                        f"is {len(formula_str)} characters, exceeding Excel's 255-character "
                        "data-validation formula limit. Shorten the option labels or "
                        "reduce the number of options."
                    )

                # Check 15 — no duplicate values (case-insensitive, warning only)
                seen_lower: set[str] = set()
                dupes: list[str] = []
                for opt in resolved:
                    key = opt.strip().lower()
                    if key in seen_lower:
                        dupes.append(opt)
                    seen_lower.add(key)
                if dupes:
                    result.add_warning(
                        f"Question '{qid}' (CHOICES): duplicate option(s) detected "
                        f"(case-insensitive): {dupes}."
                    )

    return result
