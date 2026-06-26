#!/usr/bin/env python3
"""End-to-end pipeline: choose file -> extract -> clean -> layout output."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from extract_xlsx import (
    load_xlsx,
    prepare_df,
    prompt_for_file,
    save_xlsx,
)
from format_rate_layout import (
    prompt_split_by_vendor,
    save_layout_xlsx,
    split_df_by_vendor,
)


@dataclass(frozen=True)
class PipelineResult:
    source_path: Path
    extracted_path: Path
    layout_path: Path
    row_count: int
    column_count: int
    format_name: str


def detect_format_name(df: pd.DataFrame) -> str:
    from ebs_format import is_ebs_format

    if is_ebs_format(df):
        return "EBS ship-to"
    return "standard rates"


def run_pipeline(source_path: Path | None = None) -> PipelineResult:
    """Run the full conversion pipeline and return output paths."""
    if source_path is None:
        source_path = prompt_for_file()

    print(f"\n[1/4] Reading: {source_path}")
    raw_df = load_xlsx(source_path)
    format_name = detect_format_name(raw_df)
    print(f"      Loaded {len(raw_df)} rows, {len(raw_df.columns)} columns ({format_name})")

    print("\n[2/4] Preparing data...")
    df, is_ebs = prepare_df(raw_df)
    print(f"      Cleaned to {len(df)} rows, {len(df.columns)} columns")

    print("\n[3/4] Saving extracted data to processing/...")
    extracted_path = save_xlsx(df, source_path)
    print(f"      {extracted_path}")

    print("\n[4/4] Building layout workbook in output/...")
    split_by_vendor = prompt_split_by_vendor(df)
    layout_path = save_layout_xlsx(
        df,
        source_path,
        is_ebs=is_ebs,
        split_by_vendor=split_by_vendor,
    )
    sheet_count = len(split_df_by_vendor(df, split_by_vendor=split_by_vendor))
    print(f"      {layout_path}")
    if sheet_count > 1:
        print(f"      Created {sheet_count} carrier tabs")

    return PipelineResult(
        source_path=source_path,
        extracted_path=extracted_path,
        layout_path=layout_path,
        row_count=len(df),
        column_count=len(df.columns),
        format_name=format_name,
    )


def print_summary(result: PipelineResult) -> None:
    print("\nPipeline complete.")
    print(f"  Source file:      {result.source_path}")
    print(f"  Detected format:  {result.format_name}")
    print(f"  Rows processed:   {result.row_count}")
    print(f"  Columns kept:     {result.column_count}")
    print(f"  Extracted output: {result.extracted_path}")
    print(f"  Layout output:    {result.layout_path}")


def main() -> None:
    result = run_pipeline()
    print_summary(result)


if __name__ == "__main__":
    main()
