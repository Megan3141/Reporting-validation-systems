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


def clean_dataframe(df):
    """
    Standardise dataframe structure and clean key columns.
    """

    df = df.iloc[:, :2].copy()

    df.columns = [
        "col_A_week_marker",
        "col_B_metric_name"
    ]

    df["row_index"] = df.index + 1

    df["cleaned_week_marker"] = (
        df["col_A_week_marker"]
        .apply(clean_week_value)
    )

    df["cleaned_metric_name"] = (
        df["col_B_metric_name"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df


def get_row_count(
    df,
    week_label,
    end_marker
):
    """
    Count rows between a week marker and end marker.
    """

    start_rows = df[
        df["cleaned_week_marker"].eq(week_label)
    ]["row_index"].tolist()

    if len(start_rows) == 0:
        return None

    start_row = start_rows[0]

    block = df[
        df["row_index"] > start_row
    ]

    end_rows = block[
        block["cleaned_metric_name"].eq(end_marker)
    ]["row_index"].tolist()

    if len(end_rows) == 0:
        return None

    end_row = end_rows[0]

    return end_row - start_row - 1
