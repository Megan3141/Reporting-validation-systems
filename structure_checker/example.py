"""
Example usage for the Google Sheets Structure Checker.
"""

from checker import clean_week_value, make_week_label
from config import BRAND_RULES


SPREADSHEET_ID = "YOUR_GOOGLE_SHEET_ID_HERE"

current_week = 15
previous_week = current_week - 1


for brand_name, rules in BRAND_RULES.items():
    previous_label = make_week_label(previous_week, rules["week_format"])
    current_label = make_week_label(current_week, rules["week_format"])

    print(f"{brand_name}: comparing {previous_label} to {current_label}")
