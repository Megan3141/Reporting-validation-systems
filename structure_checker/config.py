"""
Configuration for Google Sheets structure validation.
"""

BRAND_RULES = {
    "Brand A": {
        "week_format": "WK_SPACE",
        "end_marker": "NPV TOTAL"
    },
    "Brand B": {
        "week_format": "WK_NOSPACE",
        "end_marker": "TOTAL"
    },
    "Brand C": {
        "week_format": "NUMBER",
        "end_marker": "TOTAL"
    }
}
