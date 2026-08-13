"""Reusable Excel processing functions for the desktop and command-line apps."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pandas as pd


ProgressCallback = Callable[[int, int, Path], None]


def split_excel(
    input_file: str | Path,
    output_dir: str | Path,
    rows_per_file: int = 20_000,
    progress_callback: ProgressCallback | None = None,
) -> list[Path]:
    """Split an Excel workbook into files containing at most rows_per_file rows."""
    input_path = Path(input_file)
    destination = Path(output_dir)

    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if rows_per_file <= 0:
        raise ValueError("Rows per file must be greater than zero.")

    destination.mkdir(parents=True, exist_ok=True)
    dataframe = pd.read_excel(input_path, engine="openpyxl")
    total_rows = len(dataframe)

    if total_rows == 0:
        raise ValueError("The selected workbook has no data rows to split.")

    total_files = (total_rows + rows_per_file - 1) // rows_per_file
    created_files: list[Path] = []

    for file_number, start_row in enumerate(
        range(0, total_rows, rows_per_file), start=1
    ):
        chunk = dataframe.iloc[start_row : start_row + rows_per_file]
        output_file = destination / f"{input_path.stem}_split_{file_number}.xlsx"
        chunk.to_excel(output_file, index=False, engine="openpyxl")
        created_files.append(output_file)
        if progress_callback is not None:
            progress_callback(file_number, total_files, output_file)

    return created_files


def generate_gender_report(dataframe, worksheet, start_row=0, start_col=0):
    """Write the grouped totals section of the CBT report."""
    if len(dataframe.index) < 4 or len(dataframe.columns) < 12:
        raise ValueError(
            "The worksheet does not have the expected CBT layout "
            "(at least 4 rows and 12 columns)."
        )

    row2 = dataframe.iloc[1].tolist()
    row3 = dataframe.iloc[2].tolist()
    numeric_part = dataframe.iloc[3:, 11:]
    summation = numeric_part.apply(pd.to_numeric, errors="coerce").sum().tolist()

    groups, subs, values = [], [], []
    current_group = None
    for index, (group_value, sub_value) in enumerate(
        zip(row2[11:], row3[11:]), start=11
    ):
        if pd.notna(group_value) and str(group_value).strip():
            current_group = str(group_value).strip()
        groups.append(current_group if current_group else "")
        subs.append("" if pd.isna(sub_value) else str(sub_value))
        values.append(summation[index - 11])

    merge_format = worksheet.book.add_format(
        {
            "align": "center",
            "valign": "vcenter",
            "bold": True,
            "text_wrap": True,
        }
    )
    sub_format = worksheet.book.add_format({"align": "center", "bold": True})

    for row_offset, (sub, value) in enumerate(zip(subs, values)):
        row_index = start_row + row_offset
        worksheet.write(row_index, start_col + 1, sub, sub_format)
        worksheet.write(row_index, start_col + 2, value)

    def write_group(first_row, last_row, value):
        if first_row == last_row:
            worksheet.write(first_row, start_col, value, merge_format)
        else:
            worksheet.merge_range(
                first_row, start_col, last_row, start_col, value, merge_format
            )

    group_start = 0
    current_group = groups[0]
    for index, group in enumerate(groups):
        if group != current_group:
            write_group(
                start_row + group_start,
                start_row + index - 1,
                current_group,
            )
            group_start = index
            current_group = group
    write_group(
        start_row + group_start,
        start_row + len(groups) - 1,
        current_group,
    )

    worksheet.set_column(start_col, start_col, 25)
    worksheet.set_column(start_col + 1, start_col + 1, 15)
    worksheet.set_column(start_col + 2, start_col + 2, 12)


def generate_unique_providers(dataframe, worksheet, start_row=0, start_col=4):
    """Write provider and unique qualification counts to the CBT report."""
    providers = dataframe.iloc[3:, 3].dropna().unique().tolist()

    header_format = worksheet.book.add_format({"bold": True, "align": "left"})
    number_format = worksheet.book.add_format({"align": "center"})

    worksheet.write(start_row, start_col, "Unique Providers", header_format)
    worksheet.write(
        start_row,
        start_col + 1,
        "Unique Qualifications Count",
        header_format,
    )

    provider_data = dataframe.iloc[3:, :]
    for index, provider in enumerate(providers, start=1):
        worksheet.write(start_row + index, start_col, provider)
        provider_rows = provider_data[provider_data.iloc[:, 3] == provider]
        unique_qualifications = provider_rows.iloc[:, 5].dropna().unique()
        worksheet.write(
            start_row + index,
            start_col + 1,
            len(unique_qualifications),
            number_format,
        )

    worksheet.set_column(start_col, start_col, 50)
    worksheet.set_column(start_col + 1, start_col + 1, 20)


def consolidate_cbt(
    input_file: str | Path,
    output_dir: str | Path,
    sheet_name: str = "Sheet2",
) -> Path:
    """Generate the consolidated CBT report and return its output path."""
    input_path = Path(input_file)
    destination = Path(output_dir)

    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if not sheet_name.strip():
        raise ValueError("Worksheet name cannot be blank.")

    dataframe = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
    destination.mkdir(parents=True, exist_ok=True)
    output_file = destination / f"consolidated_{input_path.stem}.xlsx"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        worksheet = writer.book.add_worksheet("Report")
        worksheet.book = writer.book
        generate_gender_report(dataframe, worksheet, start_row=0, start_col=0)
        generate_unique_providers(dataframe, worksheet, start_row=0, start_col=4)

    return output_file
