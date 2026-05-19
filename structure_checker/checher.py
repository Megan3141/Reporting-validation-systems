"""
Google Sheets Structure Checker

Compares weekly reporting blocks in Google Sheets and flags row-count changes
before downstream dashboard automation runs.
"""

import pandas as pd


def clean_week_value(value):
    value = str(value).strip().upper()
    value = value.replace(" ", "")
    value = value.replace("WEEK", "WK")
    return value


def make_week_label(week_num, week_format):
    if week_format in ["WEEK_SPACE", "WK_SPACE", "WK_NOSPACE"]:
        return f"WK{week_num}"
    if week_format == "NUMBER":
        return str(week_num)
    raise ValueError(f"Unknown week_format: {week_format}")
