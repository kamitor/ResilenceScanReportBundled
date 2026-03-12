"""
test_csv_validation.py — tests for DataCleaningValidator and clean_and_fix().

Covers:
- Missing required columns → ValueError / (False, error message)
- All required columns present → passes column validation
- Missing score columns → warning only, not error
- clean_and_fix() returns False when file not found
- clean_and_fix() returns False when required columns missing
- clean_and_fix() succeeds with minimal valid CSV
- SCORE_COLUMNS and REQUIRED_COLUMNS imported from utils.constants
"""

import pathlib
import sys

import pandas as pd
import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from clean_data import DataCleaningValidator, clean_and_fix  # noqa: E402
import clean_data  # noqa: E402
from utils.constants import REQUIRED_COLUMNS, SCORE_COLUMNS  # noqa: E402


# ---------------------------------------------------------------------------
# DataCleaningValidator.validate_columns
# ---------------------------------------------------------------------------


def test_validate_columns_all_present():
    """No exception when all required + score columns are present."""
    validator = DataCleaningValidator()
    cols = REQUIRED_COLUMNS + SCORE_COLUMNS
    df = pd.DataFrame(columns=cols)
    # Should not raise
    validator.validate_columns(df)


def test_validate_columns_missing_required_raises():
    """ValueError raised when a required column is absent."""
    validator = DataCleaningValidator()
    # Only one of the three required columns present
    df = pd.DataFrame(columns=["company_name", "up__r"])
    with pytest.raises(ValueError, match="Missing required columns"):
        validator.validate_columns(df)


def test_validate_columns_missing_all_required_raises():
    """ValueError raised when no required columns present."""
    validator = DataCleaningValidator()
    df = pd.DataFrame(columns=["submitdate", "reportsent"])
    with pytest.raises(ValueError):
        validator.validate_columns(df)


def test_validate_columns_missing_scores_warns_not_raises():
    """Missing score columns produce a WARNING log entry, not an exception."""
    validator = DataCleaningValidator()
    # All required cols, no score cols
    df = pd.DataFrame(columns=REQUIRED_COLUMNS)
    validator.validate_columns(df)  # must not raise
    # Verify a warning was logged about missing score columns
    assert any("score" in w["message"].lower() for w in validator.warnings)


def test_validate_columns_case_insensitive():
    """Column validation is case-insensitive."""
    validator = DataCleaningValidator()
    # Uppercase versions of required columns
    df = pd.DataFrame(columns=[c.upper() for c in REQUIRED_COLUMNS] + SCORE_COLUMNS)
    validator.validate_columns(df)  # must not raise


# ---------------------------------------------------------------------------
# clean_and_fix() integration
# ---------------------------------------------------------------------------


def test_clean_and_fix_file_not_found(tmp_path, monkeypatch):
    """clean_and_fix() returns (False, ...) when CSV does not exist."""
    monkeypatch.setattr(clean_data, "DATA_DIR", tmp_path)
    monkeypatch.setattr(clean_data, "INPUT_PATH", tmp_path / "cleaned_master.csv")
    monkeypatch.setattr(clean_data, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(
        clean_data, "VALIDATION_LOG", tmp_path / "cleaning_validation_log.json"
    )
    monkeypatch.setattr(clean_data, "CLEANING_REPORT", tmp_path / "cleaning_report.txt")
    monkeypatch.setattr(
        clean_data, "REPLACEMENT_LOG", tmp_path / "value_replacements_log.csv"
    )

    success, msg = clean_and_fix()
    assert success is False
    assert "not found" in msg.lower() or "convert" in msg.lower()


def test_clean_and_fix_missing_required_columns(tmp_path, monkeypatch):
    """clean_and_fix() returns (False, ...) when CSV is missing required columns."""
    csv_path = tmp_path / "cleaned_master.csv"
    # Write CSV without required columns
    pd.DataFrame({"submitdate": ["2023-01-01"], "up__r": [3.5]}).to_csv(
        csv_path, index=False
    )

    monkeypatch.setattr(clean_data, "DATA_DIR", tmp_path)
    monkeypatch.setattr(clean_data, "INPUT_PATH", csv_path)
    monkeypatch.setattr(clean_data, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(
        clean_data, "VALIDATION_LOG", tmp_path / "cleaning_validation_log.json"
    )
    monkeypatch.setattr(clean_data, "CLEANING_REPORT", tmp_path / "cleaning_report.txt")
    monkeypatch.setattr(
        clean_data, "REPLACEMENT_LOG", tmp_path / "value_replacements_log.csv"
    )

    success, msg = clean_and_fix()
    assert success is False
    assert "missing" in msg.lower()


def _make_valid_csv(path):
    """Write a minimal valid CSV to path."""
    data = {
        "company_name": ["Acme Corp", "Beta BV"],
        "name": ["Alice Smith", "Bob Jones"],
        "email_address": ["alice@example.com", "bob@example.com"],
    }
    for col in SCORE_COLUMNS:
        data[col] = [3.0, 4.0]
    pd.DataFrame(data).to_csv(path, index=False)


def test_clean_and_fix_valid_csv_succeeds(tmp_path, monkeypatch):
    """clean_and_fix() returns (True, ...) for a well-formed CSV."""
    csv_path = tmp_path / "cleaned_master.csv"
    _make_valid_csv(csv_path)

    monkeypatch.setattr(clean_data, "DATA_DIR", tmp_path)
    monkeypatch.setattr(clean_data, "INPUT_PATH", csv_path)
    monkeypatch.setattr(clean_data, "BACKUP_DIR", tmp_path / "backups")
    monkeypatch.setattr(
        clean_data, "VALIDATION_LOG", tmp_path / "cleaning_validation_log.json"
    )
    monkeypatch.setattr(clean_data, "CLEANING_REPORT", tmp_path / "cleaning_report.txt")
    monkeypatch.setattr(
        clean_data, "REPLACEMENT_LOG", tmp_path / "value_replacements_log.csv"
    )

    success, msg = clean_and_fix()
    assert success is True


# ---------------------------------------------------------------------------
# utils.constants integrity
# ---------------------------------------------------------------------------


def test_score_columns_count():
    """SCORE_COLUMNS must have exactly 15 entries (3 layers × 5 dimensions)."""
    assert len(SCORE_COLUMNS) == 15


def test_score_columns_pattern():
    """Each SCORE_COLUMNS entry matches ^(up|in|do)__(r|c|f|v|a)$."""
    import re

    pat = re.compile(r"^(up|in|do)__(r|c|f|v|a)$")
    for col in SCORE_COLUMNS:
        assert pat.match(col), f"Unexpected column name: {col!r}"


def test_required_columns_content():
    """REQUIRED_COLUMNS contains the three expected column names."""
    assert set(REQUIRED_COLUMNS) == {"company_name", "name", "email_address"}
