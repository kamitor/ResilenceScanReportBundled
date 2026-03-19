"""
constants.py — shared project-wide constants.

Import from here instead of redefining in multiple modules.
"""

# Score columns present in cleaned_master.csv (0–5 scale)
SCORE_COLUMNS = [
    "up__r",
    "up__c",
    "up__f",
    "up__v",
    "up__a",
    "in__r",
    "in__c",
    "in__f",
    "in__v",
    "in__a",
    "do__r",
    "do__c",
    "do__f",
    "do__v",
    "do__a",
]

# Columns that must be present for a row to generate a report
REQUIRED_COLUMNS = ["company_name", "name", "email_address"]

# Timeout constants (seconds)
QUARTO_TIMEOUT_SECONDS = 300  # 5 minutes per report render
R_SUBPROCESS_TIMEOUT = 30  # R version / package checks
SMTP_TIMEOUT_SECONDS = 30  # SMTP connection timeout

# Filename fallback when name is missing
UNKNOWN_NAME_PLACEHOLDER = "Unknown"

# Email send mode label shown in logs and summary dialogs
TEST_MODE_LABEL = "[TEST MODE]"
