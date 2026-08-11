"""Tests for umfrage.generator."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
import yaml
from openpyxl import load_workbook

from umfrage.generator import generate_questionnaire, write_metadata_file
from umfrage.models import StyleConfig


class TestGeneratedFile:
    def test_output_file_exists(self, generated_xlsx: Path) -> None:
        assert generated_xlsx.exists()
        assert generated_xlsx.suffix == ".xlsx"

    def test_output_file_is_valid_xlsx(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        assert wb is not None

    def test_questionnaire_sheet_present(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        assert "Questionnaire" in wb.sheetnames

    def test_meta_sheet_present(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        assert "_meta" in wb.sheetnames

    def test_meta_sheet_is_hidden(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        assert wb["_meta"].sheet_state == "hidden"

    def test_output_dir_created_if_missing(
        self, tmp_path: Path, sample_questionnaire, sample_style
    ) -> None:
        nested = tmp_path / "a" / "b" / "c"
        out = nested / "q.xlsx"
        generate_questionnaire(sample_questionnaire, sample_style, out)
        assert out.exists()


class TestMetaSheet:
    def _read_meta(self, path: Path) -> dict:
        wb = load_workbook(path, data_only=True)
        ws = wb["_meta"]
        return {
            row[0]: row[1]
            for row in ws.iter_rows(max_col=2, values_only=True)
            if row[0]
        }

    def test_meta_has_questionnaire_id(
        self, generated_xlsx: Path, sample_questionnaire
    ) -> None:
        meta = self._read_meta(generated_xlsx)
        assert meta["questionnaire_id"] == sample_questionnaire.questionnaire_id()

    def test_meta_has_config_hash(
        self, generated_xlsx: Path, sample_questionnaire
    ) -> None:
        meta = self._read_meta(generated_xlsx)
        assert meta["config_hash"] == sample_questionnaire.config_hash()

    def test_meta_question_ids_match(
        self, generated_xlsx: Path, sample_questionnaire
    ) -> None:
        meta = self._read_meta(generated_xlsx)
        stored_ids = json.loads(meta["question_ids"])
        expected_ids = [q.id for q in sample_questionnaire.all_questions()]
        assert set(stored_ids) == set(expected_ids)
        assert stored_ids == expected_ids  # order preserved

    def test_meta_has_embedded_model(self, generated_xlsx: Path) -> None:
        meta = self._read_meta(generated_xlsx)
        assert "questionnaire_json" in meta
        model_data = json.loads(meta["questionnaire_json"])
        assert "title" in model_data
        assert "sections" in model_data

    def test_meta_has_version(
        self, generated_xlsx: Path, sample_questionnaire
    ) -> None:
        meta = self._read_meta(generated_xlsx)
        assert meta["version"] == sample_questionnaire.version


class TestSheetProtection:
    def test_worksheet_protection_enabled(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        ws = wb["Questionnaire"]
        assert ws.protection.sheet is True

    def test_answer_cells_are_unlocked(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        ws = wb["Questionnaire"]
        unlocked = [
            ws.cell(row=r, column=3).protection.locked
            for r in range(1, ws.max_row + 1)
        ]
        # At least one answer cell (column C) must be explicitly unlocked
        assert False in unlocked, (
            "Expected at least one unlocked cell in column C (answer column)"
        )

    def test_title_row_is_locked(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        ws = wb["Questionnaire"]
        assert ws.cell(row=1, column=1).protection.locked is True

    def test_password_protection_applied(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        style = StyleConfig(protection_password="test_secret")
        out = tmp_path / "protected.xlsx"
        generate_questionnaire(sample_questionnaire, style, out)
        wb = load_workbook(out)
        ws = wb["Questionnaire"]
        # openpyxl stores a hashed password; check it is set (not None/empty)
        assert ws.protection.sheet is True
        assert ws.protection.password is not None
        assert ws.protection.password != ""


class TestDataValidation:
    def test_scale_cells_have_whole_number_validation(
        self, generated_xlsx: Path
    ) -> None:
        wb = load_workbook(generated_xlsx)
        ws = wb["Questionnaire"]
        dv_types = [dv.type for dv in ws.data_validations.dataValidation]
        assert "whole" in dv_types

    def test_yesno_cells_have_list_validation(self, generated_xlsx: Path) -> None:
        wb = load_workbook(generated_xlsx)
        ws = wb["Questionnaire"]
        dv_types = [dv.type for dv in ws.data_validations.dataValidation]
        assert "list" in dv_types


class TestSheetContent:
    def test_all_question_ids_present_in_column_a(
        self, generated_xlsx: Path, sample_questionnaire
    ) -> None:
        wb = load_workbook(generated_xlsx, data_only=True)
        ws = wb["Questionnaire"]
        col_a = {ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)}
        for q in sample_questionnaire.all_questions():
            assert q.id in col_a, f"Question ID '{q.id}' not found in column A"

    def test_title_in_row_one(self, generated_xlsx: Path, sample_questionnaire) -> None:
        wb = load_workbook(generated_xlsx, data_only=True)
        ws = wb["Questionnaire"]
        assert ws.cell(row=1, column=1).value == sample_questionnaire.title

    def test_respondent_fields_present(
        self, generated_xlsx: Path, sample_questionnaire
    ) -> None:
        wb = load_workbook(generated_xlsx, data_only=True)
        ws = wb["Questionnaire"]
        col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        for rf in sample_questionnaire.respondent_fields:
            assert (rf.label + ":") in col_a, (
                f"Respondent field label '{rf.label}:' not found in column A"
            )


class TestWriteMetadataFile:
    def test_metadata_file_created(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        meta_path = tmp_path / "q_metadata.yaml"
        write_metadata_file(sample_questionnaire, meta_path)
        assert meta_path.exists()

    def test_metadata_contains_hash(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        meta_path = tmp_path / "q_metadata.yaml"
        write_metadata_file(sample_questionnaire, meta_path)
        data = yaml.safe_load(meta_path.read_text())
        assert data["config_hash"] == sample_questionnaire.config_hash()

    def test_metadata_contains_questionnaire_id(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        meta_path = tmp_path / "q_metadata.yaml"
        write_metadata_file(sample_questionnaire, meta_path)
        data = yaml.safe_load(meta_path.read_text())
        assert data["questionnaire_id"] == sample_questionnaire.questionnaire_id()

    def test_metadata_embeds_full_model(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        meta_path = tmp_path / "q_metadata.yaml"
        write_metadata_file(sample_questionnaire, meta_path)
        data = yaml.safe_load(meta_path.read_text())
        assert "questionnaire" in data
        assert data["questionnaire"]["title"] == sample_questionnaire.title
        assert "sections" in data["questionnaire"]

    def test_metadata_question_ids_ordered(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        meta_path = tmp_path / "q_metadata.yaml"
        write_metadata_file(sample_questionnaire, meta_path)
        data = yaml.safe_load(meta_path.read_text())
        expected = [q.id for q in sample_questionnaire.all_questions()]
        assert data["question_ids"] == expected

    def test_metadata_creates_parent_dirs(
        self, tmp_path: Path, sample_questionnaire
    ) -> None:
        meta_path = tmp_path / "nested" / "dir" / "q_metadata.yaml"
        write_metadata_file(sample_questionnaire, meta_path)
        assert meta_path.exists()
