"""Detect and normalize HOYER bulk tender price sheets."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from ebs_format import parse_rate_value

SIGNATURE_MARKERS = {
    "lane id",
    "origin country",
    "fsc to use",
    "entity",
}

MODE_INTERMODAL = "Intermodal"
MODE_ROAD = "Road"
ACTIVE_MODES = {MODE_INTERMODAL, MODE_ROAD}

NEW_PRICE_INTERMODAL = "New PRICE Intermodal"
NEW_PRICE_ROAD = "New PRICE Road"
CURRENT_PRICE_INTERMODAL = "Current PRICE Intermodal"
CURRENT_PRICE_ROAD = "Current PRICE Road"

TRANSPORT_COST_INTERMODAL = "Transport cost (Intermodal)"
TRANSPORT_COST_ROAD = "Transport cost (Road)"

INFRA_SOURCE = "Surcharge infraestructure Hoyer"
ETS_SOURCE = "ETS 2026 Hoyer"
TOLL_SOURCE = "2026 Toll BE-NL"


def find_hoyer_bulk_header_row(raw: pd.DataFrame, *, max_scan: int = 30) -> int | None:
    limit = min(len(raw), max_scan)
    for idx in range(limit):
        values = {
            str(value).strip().casefold()
            for value in raw.iloc[idx].tolist()
            if pd.notna(value) and str(value).strip()
        }
        if SIGNATURE_MARKERS.issubset(values):
            return int(idx)
    return None


def is_hoyer_bulk_path(path: Path) -> bool:
    if path.suffix.lower() != ".xlsx":
        return False
    raw = pd.read_excel(path, header=None, nrows=30, engine="openpyxl")
    return find_hoyer_bulk_header_row(raw) is not None


def is_hoyer_bulk_format(df: pd.DataFrame) -> bool:
    columns = {str(column).strip().casefold() for column in df.columns}
    return SIGNATURE_MARKERS.issubset(columns)


def _unique_column_name(name: str, used: dict[str, int]) -> str:
    count = used.get(name, 0)
    used[name] = count + 1
    if count == 0:
        return name
    return f"{name}.{count}"


def _row_text(raw: pd.DataFrame, row_idx: int | None, col_idx: int) -> str:
    if row_idx is None or row_idx < 0:
        return ""
    value = raw.iat[row_idx, col_idx]
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _build_column_names(raw: pd.DataFrame, header_row: int) -> list[str]:
    label_row = header_row - 1 if header_row > 0 else None
    mode_row = header_row - 2 if header_row > 1 else None
    used: dict[str, int] = {}
    names: list[str] = []

    for col_idx in range(raw.shape[1]):
        name = _row_text(raw, header_row, col_idx)
        label = _row_text(raw, label_row, col_idx)
        mode_label = _row_text(raw, mode_row, col_idx).casefold()

        if label in {NEW_PRICE_INTERMODAL, NEW_PRICE_ROAD}:
            name = label
        elif label.casefold() == "current price":
            if mode_label == "im":
                name = CURRENT_PRICE_INTERMODAL
            elif mode_label == "road":
                name = CURRENT_PRICE_ROAD
            else:
                name = label
        elif not name and label:
            name = label

        if not name:
            name = f"Unnamed: {col_idx}"
        names.append(_unique_column_name(name, used))

    return names


def load_hoyer_bulk(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    header_row = find_hoyer_bulk_header_row(raw)
    if header_row is None:
        raise ValueError(f"HOYER bulk tender headers not found in {path.name}")

    columns = _build_column_names(raw, header_row)
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = columns
    data = data.dropna(how="all").reset_index(drop=True)
    return data


def _find_column(columns: pd.Index, *needles: str) -> str | None:
    for column in columns:
        text = str(column).strip().casefold()
        if any(needle in text for needle in needles):
            return column
    return None


def _cell_string(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.casefold() in {"", "nan", "none", "nat"}:
        return ""
    return text


def _normalize_postal_code(value: object) -> str:
    text = _cell_string(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _normalize_mode(value: object) -> str | None:
    text = _cell_string(value)
    if not text:
        return None
    if text.casefold() == MODE_INTERMODAL.casefold():
        return MODE_INTERMODAL
    if text.casefold() == MODE_ROAD.casefold():
        return MODE_ROAD
    return None


def _build_lane_id(
    *,
    lane_value: object,
    origin_country: object,
    origin_postal: object,
    destination_country: object,
    destination_postal: object,
) -> str:
    existing = _cell_string(lane_value)
    if existing:
        return existing

    origin = f"{_cell_string(origin_country)}{_normalize_postal_code(origin_postal)}"
    destination = _cell_string(destination_country)
    dest_postal = _normalize_postal_code(destination_postal)
    parts = [part for part in (origin, destination, dest_postal) if part]
    return "-".join(parts)


def filter_hoyer_bulk_rate_rows(df: pd.DataFrame) -> pd.DataFrame:
    intermodal = df.get(TRANSPORT_COST_INTERMODAL)
    road = df.get(TRANSPORT_COST_ROAD)
    if intermodal is None and road is None:
        return df.iloc[0:0].copy()
    if intermodal is None:
        mask = road.notna()
    elif road is None:
        mask = intermodal.notna()
    else:
        mask = intermodal.notna() | road.notna()
    return df[mask].drop_duplicates(keep="first").reset_index(drop=True)


def normalize_hoyer_bulk_df(df: pd.DataFrame) -> pd.DataFrame:
    # Prefer exact "Lane ID" over "Lane ID Key".
    lane_col = next(
        (column for column in df.columns if str(column).strip().casefold() == "lane id"),
        _find_column(df.columns, "lane id"),
    )
    origin_country_col = _find_column(df.columns, "origin country")
    origin_postal_col = _find_column(df.columns, "origin postal")
    dest_country_col = _find_column(df.columns, "destination country")
    dest_postal_col = _find_column(df.columns, "receiving postal", "destination postal")
    vendor_col = _find_column(df.columns, "entity")
    mode_col = _find_column(df.columns, "fsc to use")
    origin_city_col = _find_column(df.columns, "plant name")
    intermodal_price_col = next(
        (
            column
            for column in df.columns
            if str(column).strip().casefold() == CURRENT_PRICE_INTERMODAL.casefold()
        ),
        _find_column(df.columns, "current price intermodal"),
    )
    road_price_col = next(
        (
            column
            for column in df.columns
            if str(column).strip().casefold() == CURRENT_PRICE_ROAD.casefold()
        ),
        _find_column(df.columns, "current price road"),
    )
    intermodal_cleaning_col = _find_column(df.columns, "intermodal cleaning")
    road_cleaning_col = _find_column(df.columns, "road cleaning")
    ets_col = next(
        (column for column in df.columns if str(column).strip().casefold() == "ets"),
        None,
    )
    infra_col = next(
        (column for column in df.columns if str(column).strip().casefold() == "infra"),
        None,
    )
    toll_col = _find_column(df.columns, "toll")

    required = {
        "Lane ID": lane_col,
        "Origin Country": origin_country_col,
        "Origin Postal Code": origin_postal_col,
        "Destination Country": dest_country_col,
        "Destination Postal Code": dest_postal_col,
        "Vendor Name": vendor_col,
        "Current PRICE Intermodal": intermodal_price_col,
        "Current PRICE Road": road_price_col,
    }
    missing = [name for name, column in required.items() if column is None]
    if missing:
        raise ValueError(
            "HOYER bulk tender sheet is missing expected columns: "
            + ", ".join(missing)
        )

    # Keep priced lanes; exclude blank vendor rows from trailing sheet noise.
    working = df[df[vendor_col].map(_cell_string) != ""].copy().reset_index(drop=True)

    lane_ids = [
        _build_lane_id(
            lane_value=working.at[idx, lane_col],
            origin_country=working.at[idx, origin_country_col],
            origin_postal=working.at[idx, origin_postal_col],
            destination_country=working.at[idx, dest_country_col],
            destination_postal=working.at[idx, dest_postal_col],
        )
        for idx in range(len(working))
    ]

    normalized = pd.DataFrame(
        {
            "Lane": lane_ids,
            "Origin Country": working[origin_country_col].map(_cell_string),
            "Origin ZIP Code": working[origin_postal_col].map(_normalize_postal_code),
            "Destination Country": working[dest_country_col].map(_cell_string),
            "Destination ZIP Code": working[dest_postal_col].map(_normalize_postal_code),
            "Vendor Name": working[vendor_col].map(_cell_string),
            TRANSPORT_COST_INTERMODAL: working[intermodal_price_col].map(parse_rate_value),
            TRANSPORT_COST_ROAD: working[road_price_col].map(parse_rate_value),
        }
    )

    if mode_col:
        normalized.insert(
            normalized.columns.get_loc(TRANSPORT_COST_INTERMODAL),
            "Mode",
            working[mode_col].map(_normalize_mode),
        )

    if origin_city_col:
        normalized.insert(
            normalized.columns.get_loc("Destination Country"),
            "Origin City",
            working[origin_city_col].map(_cell_string),
        )

    if intermodal_cleaning_col:
        normalized["Cleaning costs (Intermodal)"] = working[intermodal_cleaning_col].map(
            parse_rate_value
        )
    if road_cleaning_col:
        normalized["Cleaning costs (Road)"] = working[road_cleaning_col].map(parse_rate_value)

    if ets_col:
        normalized[ETS_SOURCE] = working[ets_col].map(parse_rate_value)
    if infra_col:
        normalized[INFRA_SOURCE] = working[infra_col].map(parse_rate_value)
    if toll_col:
        normalized[TOLL_SOURCE] = working[toll_col].map(parse_rate_value)

    return normalized
