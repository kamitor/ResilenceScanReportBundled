"""
test_shared_utils.py — tests for utils/filename_utils.py helpers.

Covers safe_filename() and safe_display_name() edge cases.
"""

import pathlib
import sys

import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.filename_utils import safe_display_name, safe_filename  # noqa: E402


# ---------------------------------------------------------------------------
# safe_filename
# ---------------------------------------------------------------------------


def test_safe_filename_normal():
    assert safe_filename("Acme Corp") == "Acme_Corp"


def test_safe_filename_spaces_become_underscores():
    assert safe_filename("Jan de Vries") == "Jan_de_Vries"


def test_safe_filename_hyphens_preserved():
    assert safe_filename("Alpha-Beta") == "Alpha-Beta"


def test_safe_filename_special_chars_replaced():
    result = safe_filename("Acme/Corp:Ltd")
    assert "/" not in result
    assert ":" not in result


def test_safe_filename_empty_string():
    assert safe_filename("") == "Unknown"


def test_safe_filename_nan():
    import pandas as pd

    assert safe_filename(pd.NA) == "Unknown"
    assert safe_filename(float("nan")) == "Unknown"


def test_safe_filename_alphanumeric_preserved():
    assert safe_filename("ABC123") == "ABC123"


# ---------------------------------------------------------------------------
# safe_display_name
# ---------------------------------------------------------------------------


def test_safe_display_name_normal():
    assert safe_display_name("Acme Corp") == "Acme Corp"


def test_safe_display_name_slash_becomes_dash():
    assert safe_display_name("A/B") == "A-B"


def test_safe_display_name_backslash_becomes_dash():
    assert safe_display_name("A\\B") == "A-B"


def test_safe_display_name_colon_becomes_dash():
    assert safe_display_name("A:B") == "A-B"


def test_safe_display_name_star_removed():
    assert safe_display_name("A*B") == "AB"


def test_safe_display_name_question_removed():
    assert safe_display_name("A?B") == "AB"


def test_safe_display_name_double_quote_to_single():
    assert safe_display_name('A"B') == "A'B"


def test_safe_display_name_angle_brackets():
    assert safe_display_name("A<B>C") == "A(B)C"


def test_safe_display_name_pipe_becomes_dash():
    assert safe_display_name("A|B") == "A-B"


def test_safe_display_name_empty_string():
    assert safe_display_name("") == "Unknown"


def test_safe_display_name_nan():
    import pandas as pd

    assert safe_display_name(pd.NA) == "Unknown"
    assert safe_display_name(float("nan")) == "Unknown"


def test_safe_display_name_strips_whitespace():
    assert safe_display_name("  Acme  ") == "Acme"


def test_safe_display_name_spaces_preserved():
    result = safe_display_name("Jan de Vries")
    assert " " in result


# ---------------------------------------------------------------------------
# round-trip: safe_filename(safe_display_name(x)) is filesystem-safe
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "name",
    [
        "Normal Company",
        "Acme/Corp",
        "A|B:C*D?E",
        'Quotes"Here',
        "<Brackets>",
        "Hyphens-are-fine",
    ],
)
def test_pipeline_safe(name):
    """safe_filename(safe_display_name(x)) produces a valid filename fragment."""
    display = safe_display_name(name)
    filename = safe_filename(display)
    # No path separators allowed in a filename fragment
    assert "/" not in filename
    assert "\\" not in filename
    assert filename != ""
