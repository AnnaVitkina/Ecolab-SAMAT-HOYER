"""Parse HOYER Purate rate tables from Azure Document Intelligence JSON exports."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

from ebs_format import parse_rate_value

CARRIER_NAME = "HOYER SVENSKA NSAP"
ORIGIN_COUNTRY = "DE"
ORIGIN_POSTAL_CODE = "64584"
ORIGIN_CITY = "BIEBESHEIM"

MODE_INTERMODAL = "Intermodal"
MODE_ROAD = "Road"

HEADER_RECEIVER = "Receiver/Purate"
HEADER_POSTAL_CODE = "Postal Code"

INFRASTRUCTURE_FEE_COLUMN = "Infrastructure Fee"
MAUT_COLUMN = "MAUT (DE)"
ETS_COLUMN = "ETS"
FREIGHT_RATE_COLUMN = "Freight Rate"

_JSON_HEADER_MARKERS = {
    HEADER_RECEIVER.lower(),
    HEADER_POSTAL_CODE.lower(),
    "discharge place",
    "country",
}


def is_hoyer_json_path(path: Path) -> bool:
    return path.suffix.lower() == ".json"


def _field_string(field: object) -> str | None:
    if not isinstance(field, dict):
        return None
    value = field.get("valueString")
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _is_header_row(row: dict[str, Any]) -> bool:
    receiver = (_field_string(row.get("Receiver")) or "").lower()
    postal = (_field_string(row.get("PostalCode")) or "").lower()
    if receiver in _JSON_HEADER_MARKERS or postal in _JSON_HEADER_MARKERS:
        return True
    if receiver == HEADER_RECEIVER.lower():
        return True
    if postal == HEADER_POSTAL_CODE.lower():
        return True
    main_im = _field_string(row.get("MainCostsIM")) or ""
    main_road = _field_string(row.get("MainCostsRoad")) or ""
    if "(IM)" in main_im or "(Road)" in main_road:
        return True
    return False


def _normalize_postal_code(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _route_key(
    *,
    destination_country: str,
    destination_postal_code: str,
    destination_city: str,
    receiver: str,
) -> tuple[str, str, str, str]:
    return (
        destination_country.casefold(),
        destination_postal_code.casefold(),
        destination_city.casefold(),
        receiver.casefold(),
    )


def _rows_from_main_cost_item(item: dict[str, Any]) -> list[dict[str, object]]:
    row_obj = item.get("valueObject")
    if not isinstance(row_obj, dict):
        return []
    if _is_header_row(row_obj):
        return []

    destination_country = _field_string(row_obj.get("Country")) or ""
    destination_postal_code = _normalize_postal_code(_field_string(row_obj.get("PostalCode")))
    destination_city = _field_string(row_obj.get("DischargePlace")) or ""
    receiver = _field_string(row_obj.get("Receiver")) or ""

    if not destination_country and not destination_postal_code and not destination_city:
        return []

    infrastructure = parse_rate_value(_field_string(row_obj.get("AddedCost1")))
    maut = parse_rate_value(_field_string(row_obj.get("MAUT")))
    ets = parse_rate_value(_field_string(row_obj.get("ETS")))

    im_rate = parse_rate_value(_field_string(row_obj.get("MainCostsIM")))
    road_rate = parse_rate_value(_field_string(row_obj.get("MainCostsRoad")))

    base = {
        "Vendor Name": CARRIER_NAME,
        "Origin Country": ORIGIN_COUNTRY,
        "Origin ZIP Code": ORIGIN_POSTAL_CODE,
        "Origin City": ORIGIN_CITY,
        "Destination Country": destination_country,
        "Destination ZIP Code": destination_postal_code,
        "Destination City": destination_city,
        "Receiver": receiver,
        INFRASTRUCTURE_FEE_COLUMN: infrastructure,
        MAUT_COLUMN: maut,
        ETS_COLUMN: ets,
    }

    rows: list[dict[str, object]] = []
    if im_rate is not None:
        entry = dict(base)
        entry["Mode"] = MODE_INTERMODAL
        entry[FREIGHT_RATE_COLUMN] = im_rate
        rows.append(entry)
    if road_rate is not None:
        entry = dict(base)
        entry["Mode"] = MODE_ROAD
        entry[FREIGHT_RATE_COLUMN] = road_rate
        rows.append(entry)
    return rows


def _extract_main_costs_array(payload: dict[str, Any]) -> list[dict[str, Any]]:
    analyze = payload.get("analyzeResult")
    if not isinstance(analyze, dict):
        raise ValueError("JSON is missing 'analyzeResult'.")

    documents = analyze.get("documents")
    if not isinstance(documents, list) or not documents:
        raise ValueError("JSON analyzeResult has no documents.")

    for document in documents:
        if not isinstance(document, dict):
            continue
        fields = document.get("fields")
        if not isinstance(fields, dict):
            continue
        main_costs = fields.get("MainCosts")
        if not isinstance(main_costs, dict):
            continue
        value_array = main_costs.get("valueArray")
        if isinstance(value_array, list):
            return [item for item in value_array if isinstance(item, dict)]

    raise ValueError("JSON does not contain MainCosts.valueArray.")


def load_hoyer_json(path: Path) -> pd.DataFrame:
    with path.open(encoding="utf-8") as handle:
        payload = json.load(handle)

    if not isinstance(payload, dict):
        raise ValueError("Expected JSON root object.")

    items = _extract_main_costs_array(payload)
    records: list[dict[str, object]] = []
    for item in items:
        records.extend(_rows_from_main_cost_item(item))

    if not records:
        raise ValueError("No rate rows found in JSON MainCosts.")

    df = pd.DataFrame(records)
    return _dedupe_hoyer_json_rows(df)


def _dedupe_hoyer_json_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Drop duplicate mode rows for the same route; keep first occurrence."""
    if df.empty:
        return df

    route_modes: set[tuple[tuple[str, str, str, str], str]] = set()
    keep_indices: list[int] = []

    for idx, row in df.iterrows():
        route = _route_key(
            destination_country=str(row.get("Destination Country", "")),
            destination_postal_code=str(row.get("Destination ZIP Code", "")),
            destination_city=str(row.get("Destination City", "")),
            receiver=str(row.get("Receiver", "")),
        )
        mode = str(row.get("Mode", ""))
        key = (route, mode.casefold())
        if key in route_modes:
            continue
        route_modes.add(key)
        keep_indices.append(int(idx))

    trimmed = df.loc[keep_indices].copy()
    trimmed = trimmed.drop(columns=["Receiver"], errors="ignore")
    return trimmed.reset_index(drop=True)


def is_hoyer_json_df(df: pd.DataFrame) -> bool:
    required = {
        "Origin Country",
        "Origin ZIP Code",
        "Destination Country",
        "Destination ZIP Code",
        "Mode",
        FREIGHT_RATE_COLUMN,
        INFRASTRUCTURE_FEE_COLUMN,
        MAUT_COLUMN,
        ETS_COLUMN,
    }
    columns = set(df.columns)
    return required.issubset(columns) and df.get("Vendor Name", pd.Series(dtype=str)).eq(CARRIER_NAME).all()


def prepare_hoyer_json_df(df: pd.DataFrame) -> pd.DataFrame:
    filtered = df[df[FREIGHT_RATE_COLUMN].notna()].copy()
    before = len(filtered)
    filtered = _dedupe_hoyer_json_rows(filtered)
    removed = before - len(filtered)
    if removed:
        print(f"      Removed {removed} duplicate route/mode row(s).")
    return filtered.reset_index(drop=True)
