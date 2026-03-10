"""
convert_data.py — converts the master database file to cleaned_master.csv.

Supported input formats: .xlsx, .xls, .ods, .xml, .csv, .tsv
Called by the GUI's "Convert Data" button via convert_and_save() -> bool.
Also runnable standalone: python convert_data.py
"""

import os
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Path resolution — same strategy as clean_data.py
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    if sys.platform == "win32":
        _user_base = Path(os.environ.get("APPDATA", Path.home())) / "ResilienceScan"
    else:
        _user_base = Path.home() / ".local" / "share" / "resiliencescan"
else:
    _user_base = Path(__file__).resolve().parent

DATA_DIR = _user_base / "data"
OUTPUT_PATH = DATA_DIR / "cleaned_master.csv"

_SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".ods", ".xml", ".csv", ".tsv")
# Columns whose presence marks the real header row
_HEADER_MARKERS = {"submitdate", "reportsent"}


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


def _read_excel(path: Path) -> pd.DataFrame:
    """Read an Excel file, auto-detecting sheet name and header row."""
    xl = pd.ExcelFile(path)
    sheet = "MasterData" if "MasterData" in xl.sheet_names else xl.sheet_names[0]
    skip = _header_skiprows(path, sheet)
    return pd.read_excel(path, sheet_name=sheet, skiprows=skip)


def _read_ods(path: Path) -> pd.DataFrame:
    """Read an ODS file using the odf engine, auto-detecting the header row."""
    xl = pd.ExcelFile(path, engine="odf")
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
    except Exception as exc:
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


def _read_source(path: Path) -> pd.DataFrame:
    """Dispatch to the correct reader based on file extension."""
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        return _read_excel(path)
    elif ext == ".ods":
        return _read_ods(path)
    elif ext == ".xml":
        return _read_xml(path)
    elif ext in (".csv", ".tsv"):
        return _read_raw_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def _preserve_reportsent(df: pd.DataFrame, old_csv: Path) -> pd.DataFrame:
    """Override reportsent values with those from the existing CSV.

    The app's CSV is the authoritative source for email-send tracking state;
    the source file may have been re-exported with stale values.  Matching is
    attempted in order: 'hash', 'email_address', then 'name'+'company_name'.
    """
    if "reportsent" not in df.columns or not old_csv.exists():
        return df
    try:
        old = pd.read_csv(old_csv, low_memory=False)
    except Exception:
        return df
    if "reportsent" not in old.columns:
        return df

    for key in ("hash", "email_address"):
        if key in old.columns and key in df.columns:
            mapping = old.dropna(subset=[key]).set_index(key)["reportsent"].to_dict()
            filled = df[key].map(mapping).fillna(df["reportsent"])
            df["reportsent"] = filled.infer_objects(copy=False)
            return df

    if {"name", "company_name"} <= (set(old.columns) & set(df.columns)):
        old_key = old["name"].astype(str) + "|" + old["company_name"].astype(str)
        new_key = df["name"].astype(str) + "|" + df["company_name"].astype(str)
        mapping = dict(zip(old_key, old["reportsent"]))
        filled = new_key.map(mapping).fillna(df["reportsent"])
        df["reportsent"] = filled.infer_objects(copy=False)
    return df


def convert_and_save(path: Path | None = None) -> bool:
    """Convert a supported file to cleaned_master.csv.

    If ``path`` is provided, it is used directly.  Otherwise, the first
    supported file found in DATA_DIR is used (.xlsx/.xls preferred).

    Preserves the reportsent column from any existing CSV so that email
    send-tracking state is not lost when the source is refreshed.

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

    print(f"[INFO] {len(df)} rows, {len(df.columns)} columns after normalization")

    # Restore per-row send-tracking state from any existing CSV
    df = _preserve_reportsent(df, OUTPUT_PATH)

    # Guarantee reportsent column exists and defaults to False
    if "reportsent" not in df.columns:
        df.insert(1, "reportsent", False)

    try:
        df.to_csv(OUTPUT_PATH, index=False)
        print(f"[OK] Saved {len(df)} rows to {OUTPUT_PATH}")
        return True
    except Exception as e:
        print(f"[ERROR] Failed to write CSV: {e}")
        return False


if __name__ == "__main__":
    ok = convert_and_save()
    exit(0 if ok else 1)
