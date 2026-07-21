#!/usr/bin/env python3
"""Load, clean, and save rate data from XLSX files."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from ebs_format import is_ebs_format, normalize_ebs_df, parse_rate_value
from hoyer_json_format import is_hoyer_json_path, load_hoyer_json, prepare_hoyer_json_df
from project_paths import CODE_DIR, INPUT_DIR, PROCESSING_DIR

REQUIRED_PREFIX_COLUMNS = [
    "Lane",
    "Origin Country",
    "Origin ZIP Code",
    "Destination Country",
    "Destination ZIP Code",
    "Vendor Name",
]
OPTIONAL_PREFIX_COLUMNS = ["Mode", "Vendor Number"]
PREFIX_COLUMNS = REQUIRED_PREFIX_COLUMNS + OPTIONAL_PREFIX_COLUMNS
EBS_PREFIX_COLUMNS = [
    "Origin Country",
    "Origin ZIP Code",
    "Destination Country",
    "Destination ZIP Code",
]
FREIGHT_RATE_COLUMN = "Freight Rate"
OPTIONAL_SHIPMENT_COLUMNS = ["Origin City", "Destination City"]


def find_input_files() -> list[Path]:
    patterns = ("*.xlsx", "*.json")
    files: list[Path] = []
    for pattern in patterns:
        files.extend(
            path
            for path in INPUT_DIR.glob(pattern)
            if not path.name.startswith("~$")
        )
    if not files:
        for pattern in patterns:
            files.extend(
                path
                for path in CODE_DIR.glob(pattern)
                if not path.name.startswith("~$")
            )
    return sorted(files, key=lambda path: path.name.casefold())


def find_xlsx_files() -> list[Path]:
    return [path for path in find_input_files() if path.suffix.lower() == ".xlsx"]


def prompt_for_file() -> Path:
    files = find_input_files()

    if files:
        print("\nAvailable input files:")
        for index, path in enumerate(files, start=1):
            print(f"  {index}. {path.name}")
        print("  0. Enter a custom file path")
        print()

        while True:
            choice = input("Select file number (or 0 for custom path): ").strip()
            if choice == "0":
                break
            if choice.isdigit() and 1 <= int(choice) <= len(files):
                return files[int(choice) - 1]
            print("Invalid selection. Try again.")
    else:
        print("No input files found in input/. Enter a file path.")

    while True:
        custom_path = input("Enter path to input file (.xlsx or .json): ").strip().strip('"')
        path = Path(custom_path).expanduser().resolve()
        if path.is_file() and path.suffix.lower() in {".xlsx", ".json"}:
            return path
        print("File not found or not a supported input file (.xlsx / .json). Try again.")


def load_source(path: Path) -> pd.DataFrame:
    if is_hoyer_json_path(path):
        return load_hoyer_json(path)
    return load_xlsx(path)


def load_xlsx(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, engine="openpyxl")


def find_freight_rate_column(columns: pd.Index) -> str | None:
    for column in columns:
        text = str(column).strip()
        if text == FREIGHT_RATE_COLUMN or text.startswith(f"{FREIGHT_RATE_COLUMN} "):
            return column
    return None


def normalize_freight_rate_column(df: pd.DataFrame) -> pd.DataFrame:
    freight_column = find_freight_rate_column(df.columns)
    if freight_column is None:
        raise ValueError(
            f"Column '{FREIGHT_RATE_COLUMN}' not found in file "
            "(expected 'Freight Rate' or a column starting with 'Freight Rate ')."
        )

    normalized = df.copy()
    if freight_column != FREIGHT_RATE_COLUMN:
        normalized = normalized.rename(columns={freight_column: FREIGHT_RATE_COLUMN})
    normalized[FREIGHT_RATE_COLUMN] = normalized[FREIGHT_RATE_COLUMN].map(parse_rate_value)
    return normalized


def clean_df(df: pd.DataFrame, *, is_ebs: bool = False) -> pd.DataFrame:
    prefix_columns = EBS_PREFIX_COLUMNS if is_ebs else REQUIRED_PREFIX_COLUMNS
    missing = [col for col in prefix_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    df = normalize_freight_rate_column(df)

    freight_idx = df.columns.get_loc(FREIGHT_RATE_COLUMN)
    rate_columns = list(df.columns[freight_idx:])
    optional_columns = [
        column
        for column in (OPTIONAL_PREFIX_COLUMNS + OPTIONAL_SHIPMENT_COLUMNS)
        if column in df.columns
    ]
    return df[prefix_columns + optional_columns + rate_columns].copy()


def filter_rate_rows(df: pd.DataFrame) -> pd.DataFrame:
    filtered = df[df["Freight Rate"].notna()].copy()
    return filtered.drop_duplicates(keep="first").reset_index(drop=True)


def prepare_df(df: pd.DataFrame, *, source_path: Path | None = None) -> tuple[pd.DataFrame, bool, bool]:
    is_hoyer_json = source_path is not None and is_hoyer_json_path(source_path)
    if is_hoyer_json:
        print("\nDetected HOYER JSON (Purate) layout.")
        df = prepare_hoyer_json_df(df)
        return df, False, True

    is_ebs = is_ebs_format(df)
    if is_ebs:
        print("\nDetected EBS ship-to layout.")
        df = normalize_ebs_df(df)
    df = clean_df(df, is_ebs=is_ebs)
    before_count = len(df)
    df = filter_rate_rows(df)
    removed = before_count - len(df)
    if removed:
        print(f"      Removed {removed} row(s) with empty transport cost or duplicates.")
    return df, is_ebs, False


def save_xlsx(df: pd.DataFrame, source_path: Path) -> Path:
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    output_path = PROCESSING_DIR / f"{source_path.stem}_extracted.xlsx"
    df.to_excel(output_path, index=False, engine="openpyxl")
    return output_path


def main() -> None:
    from pipeline import main as run_pipeline_main

    run_pipeline_main()


if __name__ == "__main__":
    main()
