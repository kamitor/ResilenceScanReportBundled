"""
convert_data.py — converts the master database file to cleaned_master.csv.

Supported input formats: .xlsx, .xlsm, .xls (including SpreadsheetML), .ods, .xml, .json, .jsonl, .csv, .tsv
Called by the GUI's "Convert Data" button via convert_and_save() -> bool.
Also runnable standalone: python convert_data.py
"""

import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

from utils.path_utils import get_user_base_dir

_user_base = get_user_base_dir()

DATA_DIR = _user_base / "data"
OUTPUT_PATH = DATA_DIR / "cleaned_master.csv"

_SUPPORTED_EXTENSIONS = (
    ".xlsx",
    ".xlsm",
    ".xls",
    ".ods",
    ".xml",
    ".json",
    ".jsonl",
    ".csv",
    ".tsv",
)
# Columns whose presence marks the real header row
_HEADER_MARKERS = {"submitdate", "reportsent"}

# SpreadsheetML 2003 namespace (used by Excel when saving as "XML Spreadsheet")
_SPREADSHEETML_NS = "urn:schemas-microsoft-com:office:spreadsheet"

# Column name aliases: source-format names → cleaned_master.csv convention.
# Applied after _normalize_col() so keys must already be in normalised form.
# "company-name" loses its hyphen → "companyname"; "Company name:" keeps the
# space → "company_name" — so we alias companyname → company_name for sources
# that use the hyphenated header (e.g. SpreadsheetML exports from LimeSurvey).
_COL_ALIASES: dict[str, str] = {
    "date": "submitdate",
    "email": "email_address",
    "companyname": "company_name",
}


def _normalize_col(name: str) -> str:
    """Normalize a column name to cleaned_master.csv convention.

    Removes non-alphanumeric characters (keeping spaces), then replaces
    spaces with underscores.  Examples:
      'Name:'          -> 'name'
      'E-mail address' -> 'email_address'
      'Up - R1a'       -> 'up__r1a'
      '# competitors'  -> '_competitors'
    """
    name = str(name).lower().strip()
    name = re.sub(r"[^a-z0-9 ]", "", name)  # drop non-alphanumeric, keep spaces
    name = name.replace(" ", "_")
    return name


def _find_source_file(data_dir: Path) -> Path | None:
    """Return the first supported file in data_dir.

    Priority: .xlsx > .xls > .ods > .xml > .csv > .tsv
    """
    for ext in _SUPPORTED_EXTENSIONS:
        matches = sorted(data_dir.glob(f"*{ext}"))
        if matches:
            return matches[0]
    return None


def _find_header_row(raw_df: pd.DataFrame) -> int:
    """Return the row index that contains the real column header.

    Scans rows for one that contains a known header marker ('submitdate' or
    'reportsent').  Falls back to 0 if not found.
    """
    for i, row in raw_df.iterrows():
        vals = {str(v).lower().strip() for v in row if pd.notna(v)}
        if vals & _HEADER_MARKERS:
            return int(i)
    return 0


def _header_skiprows(path: Path, sheet: str | int) -> int:
    """Return the number of rows to skip so the real column header is row 0.

    Thin wrapper around _find_header_row for Excel files.
    """
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=10)
    return _find_header_row(raw)


def _is_spreadsheetml(path: Path) -> bool:
    """Return True if path is a SpreadsheetML 2003 XML file disguised as .xls."""
    if path.suffix.lower() != ".xls":
        return False
    try:
        with path.open("rb") as f:
            chunk = f.read(512)
        return b"urn:schemas-microsoft-com:office:spreadsheet" in chunk
    except OSError:
        return False


def _read_spreadsheetml(path: Path) -> pd.DataFrame:
    """Read an Excel XML Spreadsheet (SpreadsheetML 2003) file.

    These files have a .xls extension but are plain XML.  pandas cannot read
    them with the xlrd/openpyxl engines; we parse them directly with
    ElementTree.
    """
    ns = f"{{{_SPREADSHEETML_NS}}}"
    tree = ET.parse(path)
    root = tree.getroot()
    rows = root.findall(f".//{ns}Row")
    if not rows:
        raise ValueError(f"No rows found in SpreadsheetML file: {path}")

    def _cell_text(cell: ET.Element) -> str:
        data = cell.find(f"{ns}Data")
        return data.text or "" if data is not None else ""

    headers = [_cell_text(c) for c in rows[0].findall(f"{ns}Cell")]
    records = []
    for row in rows[1:]:
        cells = row.findall(f"{ns}Cell")
        record: dict[str, str | None] = {h: None for h in headers if h is not None}
        for i, cell in enumerate(cells):
            if i < len(headers) and headers[i] is not None:
                record[headers[i]] = _cell_text(cell)
        records.append(record)
    return pd.DataFrame(records)


def _read_excel(path: Path) -> pd.DataFrame:
    """Read an Excel file, auto-detecting sheet name and header row.

    Delegates to _read_spreadsheetml() for .xls files that are actually
    SpreadsheetML 2003 XML (e.g. exported by LimeSurvey or similar tools).
    """
    if _is_spreadsheetml(path):
        return _read_spreadsheetml(path)
    with pd.ExcelFile(path) as xl:
        sheet = "MasterData" if "MasterData" in xl.sheet_names else xl.sheet_names[0]
    skip = _header_skiprows(path, sheet)
    return pd.read_excel(path, sheet_name=sheet, skiprows=skip)


def _read_ods(path: Path) -> pd.DataFrame:
    """Read an ODS file using the odf engine, auto-detecting the header row."""
    with pd.ExcelFile(path, engine="odf") as xl:
        sheet = "MasterData" if "MasterData" in xl.sheet_names else xl.sheet_names[0]
        raw = xl.parse(sheet, header=None, nrows=10)
        skip = _find_header_row(raw)
        return xl.parse(sheet, skiprows=skip)


def _read_xml(path: Path) -> pd.DataFrame:
    """Read a tabular XML file into a DataFrame.

    Tries three strategies in order:
      1. pd.read_xml(path) — simple flat XML
      2. pd.read_xml(path, xpath='.//row') — LimeSurvey-style <row> elements
      3. ElementTree fallback — finds the most common repeating child tag and
         builds a DataFrame from its child element text + attributes
    Raises ValueError if all strategies fail.
    """
    # Strategy 1: simple flat XML (use etree parser; lxml may not be installed)
    try:
        df = pd.read_xml(path, parser="etree")
        if not df.empty:
            return df
    except Exception:
        pass

    # Strategy 2: LimeSurvey-style <row> elements anywhere in the tree
    try:
        df = pd.read_xml(path, xpath=".//row", parser="etree")
        if not df.empty:
            return df
    except Exception:
        pass

    # Strategy 3: ElementTree fallback — find repeating child elements at any level
    try:
        tree = ET.parse(path)
        root = tree.getroot()

        # Walk the tree to find the deepest level with multiple siblings of the
        # same tag (likely the row-level elements).
        def _find_rows(node: ET.Element) -> list[ET.Element]:
            """Return the largest list of same-tag siblings found anywhere."""
            best: list[ET.Element] = []
            tag_groups: dict[str, list[ET.Element]] = {}
            for child in node:
                tag_groups.setdefault(child.tag, []).append(child)
            for group in tag_groups.values():
                if len(group) > len(best):
                    best = group
            # Recurse into children
            for child in node:
                candidate = _find_rows(child)
                if len(candidate) > len(best):
                    best = candidate
            return best

        rows = _find_rows(root)
        if not rows:
            raise ValueError("No repeating elements found")

        records = []
        for elem in rows:
            record: dict[str, str | None] = {}
            # Include attributes
            record.update(elem.attrib)
            # Include child element text
            for child in elem:
                record[child.tag] = child.text
            records.append(record)

        if not records:
            raise ValueError("No records extracted from XML")

        return pd.DataFrame(records)
    except (ET.ParseError, ValueError) as exc:
        raise ValueError(f"All XML read strategies failed for {path}: {exc}") from exc


def _csv_header_skip(path: Path, encoding: str) -> int:
    """Return the number of lines to skip to reach the real header row.

    Scans the first 20 lines of the file for one whose comma/tab-split tokens
    contain a known header marker ('submitdate' or 'reportsent').  Falls back
    to 0 if no marker is found.
    """
    with path.open(encoding=encoding, errors="replace") as fh:
        for i, line in enumerate(fh):
            if i >= 20:
                break
            tokens = {t.strip().lower() for t in line.replace("\t", ",").split(",")}
            if tokens & _HEADER_MARKERS:
                return i
    return 0


def _read_raw_csv(path: Path) -> pd.DataFrame:
    """Read a CSV or TSV file, auto-detecting the header row.

    Tries utf-8 encoding first, falls back to latin-1.
    Detects tab separator for .tsv files.
    Skips leading metadata rows using _csv_header_skip.
    """
    sep = "\t" if path.suffix.lower() == ".tsv" else ","

    for enc in ("utf-8", "latin-1"):
        try:
            skip = _csv_header_skip(path, enc)
            df = pd.read_csv(
                path, sep=sep, skiprows=skip, encoding=enc, on_bad_lines="skip"
            )
            return df
        except UnicodeDecodeError:
            continue
        except Exception:
            raise

    raise ValueError(
        f"Could not read CSV/TSV file with utf-8 or latin-1 encoding: {path}"
    )


def _read_json(path: Path) -> pd.DataFrame:
    """Read a JSON or JSON Lines file into a DataFrame.

    Tries four strategies in order:
      1. JSON Lines (.jsonl or explicit lines=True) — one JSON object per line.
      2. JSON top-level array: [{...}, {...}, ...]
      3. JSON object with a known data-wrapper key
         ("data", "responses", "records", "rows", "results", "items").
      4. First list-valued key found in a top-level dict.
    Raises ValueError if no strategy succeeds.
    """
    ext = path.suffix.lower()

    # Strategy 1: JSON Lines
    if ext == ".jsonl":
        df = pd.read_json(path, lines=True, dtype=str)
        if not df.empty:
            return df

    # Parse as full JSON document
    with path.open(encoding="utf-8") as fh:
        data = json.load(fh)

    # Strategy 2: top-level list
    if isinstance(data, list):
        return pd.DataFrame(data)

    if isinstance(data, dict):
        # Strategy 3: known wrapper key
        for key in ("data", "responses", "records", "rows", "results", "items"):
            if key in data and isinstance(data[key], list):
                return pd.DataFrame(data[key])
        # Strategy 4: first list-valued key
        for val in data.values():
            if isinstance(val, list) and val:
                return pd.DataFrame(val)
        # Single-record dict
        return pd.DataFrame([data])

    raise ValueError(f"Cannot convert JSON structure to DataFrame: {path}")


def _read_source(path: Path) -> pd.DataFrame:
    """Dispatch to the correct reader based on file extension."""
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return _read_excel(path)
    elif ext == ".ods":
        return _read_ods(path)
    elif ext == ".xml":
        return _read_xml(path)
    elif ext in (".json", ".jsonl"):
        return _read_json(path)
    elif ext in (".csv", ".tsv"):
        return _read_raw_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def _apply_col_aliases(df: pd.DataFrame) -> pd.DataFrame:
    """Rename columns that differ between source formats and cleaned_master.csv convention.

    Applies _COL_ALIASES after _normalize_col() has already been run.
    Only renames if the target column does not already exist.
    """
    renames = {
        src: dst
        for src, dst in _COL_ALIASES.items()
        if src in df.columns and dst not in df.columns
    }
    return df.rename(columns=renames)


def _upsert_with_existing(new_df: pd.DataFrame, existing_path: Path) -> pd.DataFrame:
    """Merge new_df on top of the existing CSV (upsert, new records first).

    All rows from new_df appear first (in their original order), followed by
    rows from the existing CSV that are not matched by any row in new_df.
    This ensures newly imported records are always at the top of the database.

    reportsent values from the existing CSV are restored for any record in
    new_df that was already present, so email-send tracking is never lost.

    Matching key priority: email_address > name+company_name.

    If no existing CSV is present the function returns new_df unchanged.
    """
    if not existing_path.exists():
        return new_df
    try:
        old_df = pd.read_csv(existing_path, low_memory=False, encoding="utf-8")
    except Exception:
        return new_df
    if old_df.empty:
        return new_df

    def _match_key(df: pd.DataFrame) -> pd.Series:
        if "email_address" in df.columns:
            return df["email_address"].astype(str).str.lower().str.strip()
        if {"name", "company_name"} <= set(df.columns):
            return (
                df["name"].astype(str).str.lower().str.strip()
                + "|"
                + df["company_name"].astype(str).str.lower().str.strip()
            )
        return pd.Series([""] * len(df), dtype=str)

    new_keys = _match_key(new_df)
    old_keys = _match_key(old_df)

    # Restore reportsent for records in new_df that already existed in the CSV
    if "reportsent" in old_df.columns:
        new_df = new_df.copy()
        if "reportsent" not in new_df.columns:
            new_df["reportsent"] = False
        key_to_sent = dict(zip(old_keys, old_df["reportsent"]))
        restored = new_keys.map(key_to_sent)
        new_df["reportsent"] = restored.fillna(new_df["reportsent"]).astype(bool)

    # Append old rows that are not represented in new_df
    new_key_set = set(new_keys.tolist())
    old_remainder = old_df[~old_keys.isin(new_key_set)].copy()

    merged = pd.concat([new_df, old_remainder], ignore_index=True)
    print(
        f"[INFO] Upsert: {len(new_df)} from new file"
        f" + {len(old_remainder)} retained from existing CSV"
        f" = {len(merged)} total"
    )
    return merged


def convert_and_save(path: Path | None = None) -> bool:
    """Convert a supported file to cleaned_master.csv.

    If ``path`` is provided, it is used directly.  Otherwise, the first
    supported file found in DATA_DIR is used (.xlsx/.xls preferred).

    If cleaned_master.csv already exists the new records are merged on top
    (upsert mode): newly imported rows appear first, existing rows not present
    in the new file are retained afterwards.  reportsent (email-send tracking)
    is always preserved for records that already existed.

    Returns True on success, False on failure.
    """
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if path is not None:
        source_file = path
    else:
        source_file = _find_source_file(DATA_DIR)

    if source_file is None:
        print(f"[ERROR] No supported file found in {DATA_DIR}")
        return False

    print(f"[INFO] Reading: {source_file.name}")
    try:
        df = _read_source(source_file)
    except Exception as e:
        print(f"[ERROR] Cannot read file: {e}")
        return False

    # Drop fully-empty rows and columns
    df = df.dropna(how="all").reset_index(drop=True)
    df = df.dropna(axis=1, how="all")

    # Normalize column names to snake_case CSV convention
    df.columns = [_normalize_col(c) for c in df.columns]

    # Drop unnamed artifact columns produced by Excel/ODS formatting
    df = df.loc[:, ~df.columns.str.fullmatch(r"unnamed_\d+")]

    # Map source-specific column names to cleaned_master.csv convention
    df = _apply_col_aliases(df)

    print(f"[INFO] {len(df)} rows, {len(df.columns)} columns after normalization")

    # Upsert: merge new records on top of any existing CSV
    df = _upsert_with_existing(df, OUTPUT_PATH)

    # Guarantee reportsent column exists and defaults to False
    if "reportsent" not in df.columns:
        df.insert(1, "reportsent", False)

    try:
        df.to_csv(OUTPUT_PATH, index=False, encoding="utf-8")
        print(f"[OK] Saved {len(df)} rows to {OUTPUT_PATH}")
        return True
    except Exception as e:
        print(f"[ERROR] Failed to write CSV: {e}")
        return False


if __name__ == "__main__":
    ok = convert_and_save()
    exit(0 if ok else 1)
