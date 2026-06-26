"""Detect and normalize EBS / ship-to rate sheet layouts."""

from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

EBS_SIGNATURE_COLUMNS = {
    "EBS Customer Ship-to",
    "NSAP Ship-to",
    "Customer Name",
    "Customer Location",
    "Customer Postcode",
    "Carrier",
}

DEFAULT_ORIGIN_COUNTRY = "DE"
DEFAULT_ORIGIN_POSTAL_CODE = "64584"
DEFAULT_DESTINATION_COUNTRY = "DE"
ORIGIN_CITY_FOR_BIEBESHEIM = "BIEBESHEIM"


def is_ebs_format(df: pd.DataFrame) -> bool:
    columns = {str(column).strip() for column in df.columns}
    signature_matches = len(columns & EBS_SIGNATURE_COLUMNS)
    has_rate_column = any(str(column).strip().startswith("Rate ") for column in df.columns)
    return signature_matches >= 4 and has_rate_column


def find_rate_column(columns: Iterable[object]) -> str:
    for column in columns:
        text = str(column).strip()
        if text.startswith("Rate "):
            return text
    raise ValueError("No rate column found (expected a column starting with 'Rate ').")


def prompt_default_value(label: str, default: str) -> str:
    while True:
        answer = input(f"{label} (default: {default}) [yes/value]: ").strip()
        if answer == "" or answer.lower() in {"yes", "y"}:
            return default
        return answer


def _normalize_postal_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def parse_rate_value(value: object) -> float | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None

    text = re.sub(r"[€$£\xa0\s]", "", text)
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
        text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return None


def normalize_ebs_df(df: pd.DataFrame) -> pd.DataFrame:
    rate_column = find_rate_column(df.columns)

    origin_country = prompt_default_value(
        "Is Origin Country",
        DEFAULT_ORIGIN_COUNTRY,
    )
    origin_postal_code = prompt_default_value(
        "Is Origin postal code",
        DEFAULT_ORIGIN_POSTAL_CODE,
    )
    destination_country = prompt_default_value(
        "Is Destination Country",
        DEFAULT_DESTINATION_COUNTRY,
    )

    origin_city = (
        ORIGIN_CITY_FOR_BIEBESHEIM
        if origin_postal_code == DEFAULT_ORIGIN_POSTAL_CODE
        else None
    )

    normalized = pd.DataFrame(
        {
            "Origin Country": origin_country,
            "Origin ZIP Code": origin_postal_code,
            "Destination Country": destination_country,
            "Destination ZIP Code": df["Customer Postcode"].map(_normalize_postal_code),
            "Destination City": df["Customer Location"],
            "Freight Rate": df[rate_column].map(parse_rate_value),
        }
    )

    if origin_city:
        origin_city_idx = normalized.columns.get_loc("Destination Country")
        normalized.insert(origin_city_idx, "Origin City", origin_city)

    return normalized
