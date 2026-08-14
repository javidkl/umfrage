# LLM / AI Authoring Guide for umfrage Questionnaire YAML

This guide is intended for AI language models (and human authors) generating or
editing questionnaire configuration files for the **umfrage** CLI tool. It
describes every field, all constraints, common patterns, and pitfalls to avoid.

> **Quick start for LLMs:** read this file in full before generating a
> questionnaire YAML. Then validate with `umfrage validate <file.yaml>` before
> generating the Excel file.

---

## 1. Purpose

A questionnaire YAML file fully describes a survey that `umfrage generate`
turns into a protected Excel file (`.xlsx`). When respondents return filled
copies, `umfrage collect` aggregates their answers into a single result
spreadsheet.

---

## 2. File Header (IDE autocomplete)

Place this comment as the very first line to enable schema-based autocompletion
and inline validation in VS Code (requires the YAML extension):

```yaml
# yaml-language-server: $schema=../docs/questionnaire.schema.json
```

---

## 3. Complete Field Reference

### 3.1 Top-Level Fields

| Field              | Type             | Required | Default | Notes |
|--------------------|------------------|----------|---------|-------|
| `title`            | string           | ✅       | —       | Shown as the large header in the Excel file. Also used as the file-name slug. |
| `version`          | string           | ❌       | `"1.0"` | Increment when changing questions after distribution. |
| `language`         | string           | ❌       | `"en"`  | Language code for all static UI labels (see §3.7). |
| `choice_lists`     | object           | ❌       | `{}`    | Named reusable option lists for `choices` questions (see §3.8). |
| `organizer`        | object           | ✅       | —       | See §3.2 |
| `respondent_fields`| array of objects | ✅       | —       | At least one required. See §3.3 |
| `sections`         | array of objects | ✅       | —       | At least one required. See §3.4 |

**No extra top-level keys are allowed.** Do not add `description`, `date`,
`notes`, or any other field not listed above — the schema will reject them.

---

### 3.2 `organizer`

```yaml
organizer:
  name: "Dr. Jane Smith"          # required: full name
  institution: "Research Inst."   # required: organization name
  email: "j.smith@example.org"    # required: valid email format
  phone: "+1 555-0100"            # optional
```

- `email` must be a valid email address (validated by `umfrage validate`).
- `phone` is optional; omit or set to `null` if unavailable.

---

### 3.3 `respondent_fields`

Each respondent fills these in when they receive the Excel file. The values
appear as column headers in the result spreadsheet.

```yaml
respondent_fields:
  - label: "Name"          # required — displayed next to the input cell
    required: true         # optional, default: true
  - label: "Institution"
  - label: "Email"
    required: false        # optional field — empty is acceptable
```

**Rules:**
- Must have **at least one** item.
- `label` must be a non-empty string.
- `required` defaults to `true`; set to `false` for optional fields.
- The label value is used by the collector to find the "institution" column in
  the result. Include a field named `"Institution"` (or containing the word
  "institution") for best results.

---

### 3.4 `sections`

```yaml
sections:
  - title: "General Information"  # required, shown as a section header row
    questions: [...]               # required, at least 1 item
```

- Must have **at least one** section.
- Each section must have **at least one** question.
- Section titles do not need to be unique, but descriptive titles help respondents.

---

### 3.5 Questions

```yaml
questions:
  - id: "S1.Q1"                     # required, unique slug
    text: "How satisfied are you?"  # required, full question text
    answer:                          # required, see §4
      type: scale
      min_value: 1
      max_value: 5
    comment: "1 = not satisfied, 5 = very satisfied"  # optional
    required: true                   # optional, default: true
```

| Field     | Type    | Required | Notes |
|-----------|---------|----------|-------|
| `id`      | string  | ✅       | Must be **globally unique** across all sections. Slug-safe (see §3.6). |
| `text`    | string  | ✅       | The question asked of the respondent. |
| `answer`  | object  | ✅       | See §4. |
| `comment` | string  | ❌       | Extra context / scale labels. For `scale`, if omitted the range is shown automatically. |
| `required`| bool    | ❌       | Default `true`. Non-required questions may be left blank. |

---

### 3.6 Question ID Naming Conventions

IDs must be **globally unique** within a questionnaire and **slug-safe**:

- Allowed characters: `a–z`, `A–Z`, `0–9`, `.` (dot), `-` (hyphen), `_` (underscore)
- Must start **and** end with a letter or digit
- No spaces, no special characters

**Recommended patterns:**

| Pattern      | Example       | Good for |
|--------------|---------------|----------|
| `S{n}.Q{n}`  | `S1.Q1`       | Simple numbered sections |
| `section.q{n}` | `general.q01` | Named sections |
| `{topic}-{n}` | `tech-1`      | Thematic grouping |

---

### 3.7 `language` — UI Language

Controls the language of **all static labels** in the generated Excel file:
column headers (`ID`, `Question`, `Answer`, `Scale / Comment`), section labels
(`RESPONDENT INFORMATION`), dropdown options (`Yes`/`No` ↔ `Ja`/`Nein`), and
result-sheet headers.

```yaml
language: "de"   # optional; default: "en"
```

| Code | Language | Yes / No |
|------|----------|----------|
| `en` | English  | Yes / No |
| `de` | German   | Ja / Nein |

**Adding a new language:** copy `umfrage/i18n/en.yaml` to
`umfrage/i18n/<code>.yaml` and translate every value. The new code becomes
immediately available — no Python changes required.

**Validation:** `umfrage validate` checks that the specified language code
exists. An unknown code blocks Excel generation.

---

### 3.8 `choice_lists` — Named Reusable Option Lists

Define option lists once at the top level and reference them in any number of
`choices`-type questions with `choices_ref: <name>`. This avoids repeating the
same list in multiple places and ensures a consistent dropdown across questions.

```yaml
choice_lists:
  frequency:            # key: any identifier you choose
    - "Never"
    - "Rarely"
    - "Sometimes"
    - "Often"
    - "Always"
  satisfaction_level:
    - "Very dissatisfied"
    - "Dissatisfied"
    - "Neutral"
    - "Satisfied"
    - "Very satisfied"
```

**Rules:**
- Each list must have **at least 2** options.
- The comma-joined option string must not exceed **255 characters** (Excel data-
  validation formula limit). Shorten labels or reduce the number of options if needed.
- Option strings must be **unique** within a list (case-insensitive; duplicates
  produce a warning).
- Omit `choice_lists` entirely if all choices questions use inline `choices: [...]`.

---

## 4. Answer Type Reference

### 4.1 `scale` — Numeric Scale

Respondents enter an integer within a defined range. The Excel cell gets a data
validation rule enforcing the range.

```yaml
answer:
  type: scale
  min_value: 1         # required — inclusive lower bound
  max_value: 5         # required — inclusive upper bound, must be > min_value
  description: "1 = poor, 5 = excellent"  # optional hint
```

**Common ranges:** `1–3`, `1–5`, `1–7`, `1–10`, `0–100`.

**Rules:**
- Both `min_value` and `max_value` are **required**.
- `min_value` must be **strictly less than** `max_value`.
- Values must be integers (not floats).

---

### 4.2 `yes_no` — Binary Choice

Respondents select "Yes" or "No" from a dropdown. The validator accepts
any capitalization (`yes`, `Yes`, `YES`).

```yaml
answer:
  type: yes_no
  # description is optional
  description: "Select Yes if applicable"
```

**Rules:**
- Do **not** set `min_value` or `max_value` (a warning will be raised if you do).
- No additional configuration required.

---

### 4.3 `freetext` — Open-Ended Text

Respondents type any text. No data validation is applied. The cell is unlocked
and accepts any value.

```yaml
answer:
  type: freetext
  # description is optional
  description: "Please describe in 2–3 sentences"
```

**Rules:**
- Do **not** set `min_value` or `max_value` (a warning will be raised if you do).
- The `required` flag still applies — a blank freetext answer for a required
  question will fail validation during collection.

---

### 4.4 `choices` — Fixed Option Dropdown

Respondents select one option from a predefined list. The cell shows an Excel
dropdown. The Scale/Comment column lists the options as a hint (suppressable).
Validation accepts the answer case-insensitively.

**Option A — inline list (one-off):**
```yaml
answer:
  type: choices
  choices:
    - "Email"
    - "Video call"
    - "In-person meeting"
    - "Other"
```

**Option B — named list (reusable):**
```yaml
answer:
  type: choices
  choices_ref: frequency    # references choice_lists.frequency
```

**Suppressing the option list from the comment column (no `comment` set):**
```yaml
answer:
  type: choices
  choices_ref: satisfaction_level
  show_choices_in_comment: false  # options not listed in column D
  description: "Select one"        # optional replacement hint
```

> **Note:** `show_choices_in_comment` only takes effect when **no** `comment` is set on the
> question. An explicit `comment` is always a complete override (same as for `scale`,
> `yes_no`, and `freetext`) and suppresses the auto-generated options hint entirely,
> regardless of this flag. In that case the flag is redundant and can be omitted.

**Rules:**
- Provide **either** `choices` or `choices_ref`, **not both**.
- `choices_ref` must reference a key defined in `choice_lists`.
- At least **2** options are required.
- Comma-joined options must not exceed **255 characters**.
- Do **not** set `min_value` or `max_value` (a warning will be raised).
- `show_choices_in_comment` defaults to `true`; set to `false` to suppress the
  auto-generated option list when no `comment` is set and the list would make
  column D too wide. Rarely needed; omit in most cases.
- The `required` flag applies normally.

---

## 5. Validation Rules (mirrors `umfrage validate`)

The following errors **block** file generation:

1. No sections defined.
2. A section has no questions.
3. Duplicate question ID across any sections.
4. Question ID is not slug-safe.
5. `scale` answer missing `min_value` or `max_value`.
6. `scale` answer where `min_value >= max_value`.
7. `respondent_fields` is empty.
8. Organizer `email` is not a valid email address.
9. `choices` answer has neither `choices` nor `choices_ref`.
10. `choices` answer has both `choices` and `choices_ref`.
11. `choices_ref` references an undefined key in `choice_lists`.
12. Resolved choices list has fewer than 2 items.
13. Comma-joined choices exceed 255 characters.

The following produce **warnings** (generation proceeds):

- `yes_no`, `freetext`, or `choices` answer has `min_value`/`max_value` set.
- Duplicate option strings (case-insensitive) in a choices list.

---

## 6. Full Worked Example

```yaml
# yaml-language-server: $schema=../docs/questionnaire.schema.json

title: "Annual Cooperation Survey 2024"
version: "1.0"

organizer:
  name: "Dr. Jane Smith"
  institution: "Research Institute XYZ"
  email: "j.smith@xyz.org"
  phone: "+49 89 12345-678"

respondent_fields:
  - label: "Name"
  - label: "Institution"
  - label: "Email"
    required: false

choice_lists:
  satisfaction_level:
    - "Very dissatisfied"
    - "Dissatisfied"
    - "Neutral"
    - "Satisfied"
    - "Very satisfied"
  frequency:
    - "Never"
    - "Rarely"
    - "Sometimes"
    - "Often"
    - "Always"

sections:
  - title: "General Satisfaction"
    questions:
      - id: "GEN.Q1"
        text: "How satisfied are you with the overall cooperation?"
        answer:
          type: scale
          min_value: 1
          max_value: 5
          description: "1 = not satisfied, 5 = very satisfied"
        required: true

      - id: "GEN.Q2"
        text: "Would you participate in this survey again next year?"
        answer:
          type: yes_no
        required: true

      - id: "GEN.Q3"
        text: "Please share any additional comments on general cooperation."
        answer:
          type: freetext
        required: false

      - id: "GEN.Q4"
        text: "How would you describe the overall quality of the cooperation?"
        answer:
          type: choices
          choices_ref: satisfaction_level
        required: true

      - id: "GEN.Q5"
        text: "How often did you exchange information with the project team?"
        answer:
          type: choices
          choices_ref: frequency
          show_choices_in_comment: false
          description: "Select frequency"
        required: true

  - title: "Technical Quality"
    questions:
      - id: "TECH.Q1"
        text: "How would you rate the technical quality of the provided tools?"
        answer:
          type: scale
          min_value: 1
          max_value: 10
          description: "1 = very poor, 10 = excellent"
        required: true

      - id: "TECH.Q2"
        text: "Did you experience any technical issues during the project?"
        answer:
          type: yes_no
        required: true

      - id: "TECH.Q3"
        text: "If yes, please describe the issues you encountered."
        answer:
          type: freetext
        required: false
        comment: "Only fill in if you answered 'Yes' to TECH.Q2"
```

---

## 7. Tips for LLMs

Follow these rules strictly when generating or editing questionnaire YAML files:

1. **Always generate unique IDs.** Before outputting the YAML, verify that no
   two questions share the same `id`. If sections exist already, check all of them.

2. **Always set `min_value` and `max_value` for `scale` questions.** Omitting
   either causes a validation error and blocks Excel generation.

3. **Never add extra top-level keys.** The schema uses `additionalProperties: false`.
   Keys like `description`, `date`, `notes`, `tags` at the top level will cause
   validation errors. Use `comment` inside questions instead.

4. **Do not set `min_value`/`max_value` for `yes_no` or `freetext` questions.**
   They are ignored and generate warnings. Clean output is preferred.

5. **Respect the slug pattern for IDs.** The regex is:
   `^[a-zA-Z0-9]([a-zA-Z0-9._-]*[a-zA-Z0-9])?$`
   Spaces, parentheses, slashes, and other special characters are not allowed.

6. **Include an `"Institution"` respondent field** (exact spelling preferred) so
   the collector can auto-detect the institution name for the result column header.

7. **Use `required: false`** on open-ended or follow-up questions that are
   genuinely optional so respondents are not penalized for skipping them.

8. **For `choices` type, provide either `choices` or `choices_ref`, not both.**
   If reusing the same list in multiple questions, define it in `choice_lists`
   and use `choices_ref`. Use inline `choices: [...]` for one-off lists.

9. **`show_choices_in_comment: false`** suppresses the auto-generated option list
   in the Scale/Comment column (column D). Only relevant when the list is long
   (say, more than 5–6 options) **and no `comment` is set** on the question.
   When a `comment` is present it already overrides the auto-generated hint
   entirely, making this flag redundant. The dropdown is never affected.

10. **Use `language: "de"`** (or another available code) when the questionnaire
    is intended for German-speaking respondents. This translates all static
    labels, column headers, and the yes/no dropdown values (`Ja`/`Nein`).

11. **Validate before generating.** Always run `umfrage validate <file.yaml>`
    after generating a config to catch any issues before distributing.

---

## 8. IDE Integration

With the [YAML extension for VS Code](https://marketplace.visualstudio.com/items?itemName=redhat.vscode-yaml)
installed, add this as the first line of your YAML file:

```yaml
# yaml-language-server: $schema=../docs/questionnaire.schema.json
```

This enables:
- Inline error highlighting for invalid fields
- Autocomplete for field names and enum values (`scale`, `yes_no`, `freetext`)
- Hover documentation for every field
