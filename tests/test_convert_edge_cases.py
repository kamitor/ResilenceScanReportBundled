"""
test_convert_edge_cases.py — additional edge-case coverage for convert_data.py.

Focuses on areas not covered by test_convert_formats.py:
- _normalize_col() completeness
- _apply_col_aliases() alias table completeness
- _find_source_file() priority ordering
- _upsert_with_existing() matching logic
- convert_and_save() column normalisation pipeline
"""

import pathlib
import sys

import pandas as pd
import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import convert_data  # noqa: E402
from convert_data import (  # noqa: E402
    _COL_ALIASES,
    _SUPPORTED_EXTENSIONS,
    _apply_col_aliases,
    _find_source_file,
    _normalize_col,
    _upsert_with_existing,
    convert_and_save,
)


# ---------------------------------------------------------------------------
# _normalize_col — comprehensive table
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "raw, expected",
    [
        # _normalize_col: lower+strip → drop [^a-z0-9 ] → replace spaces with _
        # Underscores in source are dropped (they are non-alphanumeric, non-space).
        ("company_name", "companyname"),  # underscore removed
        ("Company Name", "company_name"),  # space → underscore
        ("COMPANY_NAME", "companyname"),  # underscore removed + lowercase
        ("company-name", "companyname"),  # hyphen removed
        ("E-mail address", "email_address"),  # hyphen removed, space → _
        ("E-mail:", "email"),  # hyphen + colon removed
        ("submitdate", "submitdate"),
        ("SubmitDate", "submitdate"),
        ("up__r", "upr"),  # underscores removed → "upr"
        ("Up - R", "up__r"),  # " - " → "  " → "__" after replace
        ("reportsent", "reportsent"),
        ("name", "name"),
        ("  name  ", "name"),
        ("123abc", "123abc"),
    ],
)
def test_normalize_col_parametrize(raw, expected):
    assert _normalize_col(raw) == expected


# ---------------------------------------------------------------------------
# _COL_ALIASES — alias table sanity
# ---------------------------------------------------------------------------


def test_alias_date_maps_to_submitdate():
    assert _COL_ALIASES.get("date") == "submitdate"


def test_alias_email_maps_to_email_address():
    assert _COL_ALIASES.get("email") == "email_address"


def test_alias_companyname_maps_to_company_name():
    assert _COL_ALIASES.get("companyname") == "company_name"


def test_apply_aliases_renames_date(tmp_path):
    """A DataFrame with a 'date' column gets it renamed to 'submitdate'."""
    df = pd.DataFrame([{"date": "2024-01-01", "name": "Alice"}])
    result = _apply_col_aliases(df)
    assert "submitdate" in result.columns
    assert "date" not in result.columns


def test_apply_aliases_skips_if_target_exists():
    """If the target column already exists, no rename occurs (avoids duplicates)."""
    df = pd.DataFrame([{"email": "a@b.com", "email_address": "c@d.com"}])
    result = _apply_col_aliases(df)
    # Both should survive untouched (no rename since target exists)
    assert "email_address" in result.columns


def test_apply_aliases_no_false_renames():
    """Columns not in _COL_ALIASES are not affected."""
    df = pd.DataFrame([{"company_name": "Acme", "name": "Alice", "foo": "bar"}])
    result = _apply_col_aliases(df)
    assert "foo" in result.columns
    assert "company_name" in result.columns


# ---------------------------------------------------------------------------
# _find_source_file() — priority ordering
# ---------------------------------------------------------------------------


def test_find_source_file_prefers_xlsx_over_csv(tmp_path):
    """xlsx is preferred over csv when both are present."""
    (tmp_path / "data.csv").write_text("col1\nval1", encoding="utf-8")
    (tmp_path / "data.xlsx").write_bytes(b"")  # dummy, just needs to exist
    result = _find_source_file(tmp_path)
    assert result is not None
    assert result.suffix == ".xlsx"


def test_find_source_file_returns_none_when_empty(tmp_path):
    """Returns None when the directory contains no supported files."""
    assert _find_source_file(tmp_path) is None


def test_find_source_file_returns_path(tmp_path):
    """Returns a pathlib.Path instance."""
    (tmp_path / "data.csv").write_text("col1\nval1", encoding="utf-8")
    result = _find_source_file(tmp_path)
    assert isinstance(result, pathlib.Path)


def test_find_source_file_finds_json(tmp_path):
    """Returns a .json file when present and no higher-priority type exists."""
    (tmp_path / "data.json").write_text("[]", encoding="utf-8")
    result = _find_source_file(tmp_path)
    assert result is not None
    assert result.suffix == ".json"


def test_find_source_file_finds_jsonl(tmp_path):
    """Returns a .jsonl file when present and no higher-priority type exists."""
    (tmp_path / "data.jsonl").write_text('{"x": 1}\n', encoding="utf-8")
    result = _find_source_file(tmp_path)
    assert result is not None
    assert result.suffix == ".jsonl"


def test_all_extensions_recognised():
    """Every extension in _SUPPORTED_EXTENSIONS is non-empty and starts with '.'."""
    for ext in _SUPPORTED_EXTENSIONS:
        assert ext.startswith("."), f"Extension '{ext}' must start with '.'"
        assert len(ext) > 1


# ---------------------------------------------------------------------------
# _upsert_with_existing() — key matching via email vs name+company
# ---------------------------------------------------------------------------


def test_upsert_matches_by_email(tmp_path):
    """Upsert deduplicates using email_address as the matching key."""
    existing = tmp_path / "existing.csv"
    existing.write_text(
        "company_name,name,email_address,reportsent\nOld Co,Old Name,alice@x.com,True\n",
        encoding="utf-8",
    )
    new_df = pd.DataFrame(
        [
            {
                "company_name": "New Co",
                "name": "New Name",
                "email_address": "alice@x.com",  # same email → counts as same record
            }
        ]
    )
    result = _upsert_with_existing(new_df, existing)
    # Record should appear once; reportsent should be True (restored from old)
    assert len(result) == 1
    assert result.iloc[0]["reportsent"] == True  # noqa: E712


def test_upsert_matches_by_name_company_when_no_email(tmp_path):
    """Upsert falls back to name+company_name key when email_address is absent."""
    existing = tmp_path / "existing.csv"
    existing.write_text(
        "company_name,name,reportsent\nAcme,Alice,True\n",
        encoding="utf-8",
    )
    new_df = pd.DataFrame([{"company_name": "Acme", "name": "Alice"}])
    result = _upsert_with_existing(new_df, existing)
    assert len(result) == 1
    assert result.iloc[0]["reportsent"] == True  # noqa: E712


def test_upsert_retains_unmatched_old_rows(tmp_path):
    """Old records not present in new_df are appended at the end."""
    existing = tmp_path / "existing.csv"
    existing.write_text(
        "company_name,name,email_address\nAcme,Alice,a@x.com\nBeta,Bob,b@x.com\n",
        encoding="utf-8",
    )
    new_df = pd.DataFrame(
        [{"company_name": "Acme", "name": "Alice", "email_address": "a@x.com"}]
    )
    result = _upsert_with_existing(new_df, existing)
    assert len(result) == 2
    companies = list(result["company_name"])
    assert "Beta" in companies


# ---------------------------------------------------------------------------
# Full pipeline: column normalisation + alias application
# ---------------------------------------------------------------------------


def test_convert_and_save_normalises_column_names(monkeypatch, tmp_path):
    """Column names from the source file are normalised to snake_case."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = tmp_path / "input.csv"
    src.write_text(
        "Company Name,Name,E-mail address\nAcme,Alice,a@x.com\n",
        encoding="utf-8",
    )
    result = convert_and_save(path=src)
    assert result is True
    df = pd.read_csv(output)
    assert "company_name" in df.columns
    assert "name" in df.columns
    assert "email_address" in df.columns


def test_convert_and_save_applies_col_aliases(monkeypatch, tmp_path):
    """Source columns named 'email' and 'date' are aliased to the CSV convention."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = tmp_path / "input.csv"
    src.write_text(
        "company_name,name,email,date\nAcme,Alice,a@x.com,2024-01-01\n",
        encoding="utf-8",
    )
    convert_and_save(path=src)
    df = pd.read_csv(output)
    assert "email_address" in df.columns
    assert "submitdate" in df.columns


def test_convert_and_save_drops_unnamed_columns(monkeypatch, tmp_path):
    """Unnamed artifact columns (unnamed_0, unnamed_1, …) are dropped."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = tmp_path / "input.csv"
    # pandas names extra blank columns as "Unnamed: N" which normalises to "unnamed_N"
    src.write_text(
        "company_name,name,,\nAcme,Alice,,\n",
        encoding="utf-8",
    )
    convert_and_save(path=src)
    df = pd.read_csv(output)
    for col in df.columns:
        assert not col.startswith("unnamed_"), f"unnamed column not dropped: {col}"


def test_convert_and_save_utf8_output(monkeypatch, tmp_path):
    """Output CSV is written as UTF-8 (can be re-read without errors)."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = tmp_path / "input.csv"
    src.write_text(
        "company_name,name,email_address\nBédrijf,Ångström,a@x.com\n",
        encoding="utf-8",
    )
    convert_and_save(path=src)
    df = pd.read_csv(output, encoding="utf-8")
    assert "Bédrijf" in df["company_name"].values
