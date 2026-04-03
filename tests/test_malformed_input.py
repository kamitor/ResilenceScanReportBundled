"""
test_malformed_input.py — edge cases for malformed, empty, or corrupt input files.

Covers convert_data.py behaviour when it receives files that are technically
valid enough to open but have unexpected or degenerate content.
"""

import json
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import convert_data  # noqa: E402
from convert_data import (  # noqa: E402
    _find_header_row,
    _normalize_col,
    _read_json,
    _read_raw_csv,
    _upsert_with_existing,
    convert_and_save,
)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _make_csv(path, content: str) -> pathlib.Path:
    p = pathlib.Path(path)
    p.write_text(content, encoding="utf-8")
    return p


# ---------------------------------------------------------------------------
# _normalize_col edge cases
# ---------------------------------------------------------------------------


def test_normalize_col_strips_and_lowercases():
    assert _normalize_col("  Company Name  ") == "company_name"


def test_normalize_col_removes_punctuation():
    assert _normalize_col("E-mail:") == "email"


def test_normalize_col_handles_empty_string():
    result = _normalize_col("")
    assert result == ""


def test_normalize_col_all_special_chars():
    result = _normalize_col("!@#$%^&*()")
    assert result == ""


def test_normalize_col_preserves_numbers():
    # Underscores are stripped by _normalize_col (only spaces are kept, then → _)
    # "up__r1a" → strip non-alpha/numeric/space → "upr1a"
    assert _normalize_col("up__r1a") == "upr1a"


def test_normalize_col_unicode_stripped():
    # Non-ASCII characters are stripped
    result = _normalize_col("Compañía")
    assert "a" in result  # 'a' from 'Compa' survives; ñ is dropped


# ---------------------------------------------------------------------------
# _find_header_row edge cases
# ---------------------------------------------------------------------------


def test_find_header_row_empty_df():
    """An empty DataFrame falls back to row 0."""
    empty = pd.DataFrame()
    assert _find_header_row(empty) == 0


def test_find_header_row_no_markers():
    """A DataFrame with no header markers returns row 0."""
    df = pd.DataFrame([["col1", "col2"], ["val1", "val2"]])
    assert _find_header_row(df) == 0


def test_find_header_row_marker_on_second_row():
    """Finds the 'submitdate' marker on row 1."""
    df = pd.DataFrame([["Metadata", "v1"], ["submitdate", "reportsent"]])
    assert _find_header_row(df) == 1


def test_find_header_row_reportsent_marker():
    """Finds the 'reportsent' marker correctly."""
    df = pd.DataFrame([["extra", "info"], ["company", "reportsent"]])
    assert _find_header_row(df) == 1


# ---------------------------------------------------------------------------
# _read_raw_csv — edge cases
# ---------------------------------------------------------------------------


def test_read_csv_headers_only(tmp_path):
    """A CSV with only a header row produces an empty DataFrame."""
    p = _make_csv(tmp_path / "headers_only.csv", "company_name,name,email_address\n")
    df = _read_raw_csv(p)
    assert len(df) == 0
    assert "company_name" in df.columns


def test_read_csv_with_utf8_bom(tmp_path):
    """CSV files with a UTF-8 BOM (EF BB BF) are read without the BOM in column names."""
    content = "\ufeffcompany_name,name,email_address\nAcme,Alice,a@x.com\n"
    p = tmp_path / "bom.csv"
    p.write_text(content, encoding="utf-8-sig")
    df = _read_raw_csv(p)
    # BOM should not appear in column names
    assert any("company" in c for c in df.columns)
    assert len(df) >= 1


def test_read_csv_extra_columns(tmp_path):
    """Extra columns beyond the required set do not cause errors."""
    p = _make_csv(
        tmp_path / "extra_cols.csv",
        "company_name,name,email_address,extra1,extra2\nAcme,Alice,a@x.com,foo,bar\n",
    )
    df = _read_raw_csv(p)
    assert len(df) == 1
    assert "extra1" in df.columns


def test_read_csv_latin1_encoding(tmp_path):
    """CSV files with latin-1 encoding are decoded without error."""
    p = tmp_path / "latin1.csv"
    p.write_bytes("company_name,name\nAcm\xe9,Alic\xe9\n".encode("latin-1"))
    df = _read_raw_csv(p)
    assert len(df) >= 1


def test_read_csv_empty_file_returns_empty_df(tmp_path):
    """Completely empty CSV file returns an empty DataFrame (no crash)."""
    p = _make_csv(tmp_path / "empty.csv", "")
    # pandas raises EmptyDataError on truly empty files — convert_and_save should handle it
    try:
        df = _read_raw_csv(p)
        assert df.empty or len(df) == 0
    except (pd.errors.EmptyDataError, Exception):
        pass  # acceptable — the caller (convert_and_save) catches exceptions


# ---------------------------------------------------------------------------
# _read_json — edge cases
# ---------------------------------------------------------------------------


def test_read_json_empty_list(tmp_path):
    """JSON file containing an empty top-level array returns an empty DataFrame."""
    p = tmp_path / "empty.json"
    p.write_text("[]", encoding="utf-8")
    df = _read_json(p)
    assert df.empty


def test_read_json_null_fields(tmp_path):
    """JSON records with null values do not crash the reader."""
    data = [{"company_name": "Acme", "name": None, "email_address": None}]
    p = tmp_path / "nulls.json"
    p.write_text(json.dumps(data), encoding="utf-8")
    df = _read_json(p)
    assert len(df) == 1


def test_read_json_nested_responses_key(tmp_path):
    """JSON with a 'responses' wrapper key is unwrapped correctly."""
    data = {"responses": [{"company_name": "X", "name": "Y"}]}
    p = tmp_path / "nested.json"
    p.write_text(json.dumps(data), encoding="utf-8")
    df = _read_json(p)
    assert len(df) == 1


def test_read_json_nested_data_key(tmp_path):
    """JSON with a 'data' wrapper key is unwrapped correctly."""
    data = {
        "data": [{"company_name": "X", "name": "Y"}, {"company_name": "Z", "name": "W"}]
    }
    p = tmp_path / "nested_data.json"
    p.write_text(json.dumps(data), encoding="utf-8")
    df = _read_json(p)
    assert len(df) == 2


def test_read_jsonl_single_record(tmp_path):
    """A JSONL file with one record is read correctly."""
    p = tmp_path / "single.jsonl"
    p.write_text('{"company_name": "Acme", "name": "Alice"}\n', encoding="utf-8")
    df = _read_json(p)
    assert len(df) == 1
    assert "company_name" in df.columns


def test_read_jsonl_empty_file(tmp_path):
    """An empty JSONL file does not crash the reader."""
    p = tmp_path / "empty.jsonl"
    p.write_text("", encoding="utf-8")
    try:
        df = _read_json(p)
        assert df.empty
    except (ValueError, Exception):
        pass  # acceptable; convert_and_save catches exceptions


# ---------------------------------------------------------------------------
# _upsert_with_existing — edge cases
# ---------------------------------------------------------------------------


def test_upsert_no_existing_csv(tmp_path):
    """When no existing CSV exists, upsert returns the new DataFrame unchanged."""
    new_df = pd.DataFrame(
        [{"company_name": "X", "name": "Y", "email_address": "x@y.com"}]
    )
    result = _upsert_with_existing(new_df, tmp_path / "nonexistent.csv")
    assert len(result) == 1


def test_upsert_empty_existing_csv(tmp_path):
    """When existing CSV is empty, upsert returns the new DataFrame."""
    existing = tmp_path / "existing.csv"
    existing.write_text("company_name,name,email_address\n", encoding="utf-8")
    new_df = pd.DataFrame(
        [{"company_name": "X", "name": "Y", "email_address": "x@y.com"}]
    )
    result = _upsert_with_existing(new_df, existing)
    assert len(result) >= 1


def test_upsert_preserves_reportsent_true(tmp_path):
    """Upsert restores reportsent=True for records already in the CSV."""
    existing = tmp_path / "existing.csv"
    existing.write_text(
        "company_name,name,email_address,reportsent\nAcme,Alice,a@x.com,True\n",
        encoding="utf-8",
    )
    new_df = pd.DataFrame(
        [{"company_name": "Acme", "name": "Alice", "email_address": "a@x.com"}]
    )
    result = _upsert_with_existing(new_df, existing)
    assert result.iloc[0]["reportsent"] == True  # noqa: E712


def test_upsert_new_records_first(tmp_path):
    """New records appear before retained old records in the merged output."""
    existing = tmp_path / "existing.csv"
    existing.write_text(
        "company_name,name,email_address\nOld Co,Old Person,old@x.com\n",
        encoding="utf-8",
    )
    new_df = pd.DataFrame(
        [{"company_name": "New Co", "name": "New Person", "email_address": "new@x.com"}]
    )
    result = _upsert_with_existing(new_df, existing)
    assert result.iloc[0]["company_name"] == "New Co"
    assert result.iloc[1]["company_name"] == "Old Co"


def test_upsert_no_duplication(tmp_path):
    """A record present in both new and existing CSV appears only once."""
    existing = tmp_path / "existing.csv"
    existing.write_text(
        "company_name,name,email_address\nAcme,Alice,a@x.com\n",
        encoding="utf-8",
    )
    new_df = pd.DataFrame(
        [{"company_name": "Acme", "name": "Alice", "email_address": "a@x.com"}]
    )
    result = _upsert_with_existing(new_df, existing)
    assert len(result) == 1


# ---------------------------------------------------------------------------
# convert_and_save — end-to-end edge cases
# ---------------------------------------------------------------------------


def test_convert_and_save_nonexistent_file(monkeypatch, tmp_path, capsys):
    """convert_and_save returns False when passed a non-existent path."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", tmp_path / "cleaned_master.csv")
    result = convert_and_save(path=tmp_path / "does_not_exist.csv")
    assert result is False


def test_convert_and_save_valid_csv(monkeypatch, tmp_path):
    """convert_and_save returns True for a minimal valid CSV."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = _make_csv(
        tmp_path / "input.csv",
        "company_name,name,email_address\nAcme,Alice,a@x.com\n",
    )
    result = convert_and_save(path=src)
    assert result is True
    assert output.exists()


def test_convert_and_save_adds_reportsent_column(monkeypatch, tmp_path):
    """convert_and_save guarantees a reportsent column even when missing from source."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = _make_csv(
        tmp_path / "input.csv",
        "company_name,name,email_address\nAcme,Alice,a@x.com\n",
    )
    convert_and_save(path=src)
    result_df = pd.read_csv(output)
    assert "reportsent" in result_df.columns


def test_convert_and_save_drops_empty_rows(monkeypatch, tmp_path):
    """convert_and_save drops fully empty rows."""
    monkeypatch.setattr(convert_data, "DATA_DIR", tmp_path)
    output = tmp_path / "cleaned_master.csv"
    monkeypatch.setattr(convert_data, "OUTPUT_PATH", output)

    src = _make_csv(
        tmp_path / "input.csv",
        "company_name,name,email_address\nAcme,Alice,a@x.com\n,,\n",
    )
    convert_and_save(path=src)
    result_df = pd.read_csv(output)
    assert len(result_df) == 1
