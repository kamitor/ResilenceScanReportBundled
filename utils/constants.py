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
