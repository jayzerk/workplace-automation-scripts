"""Command-line entry point for the Excel splitter."""

from pathlib import Path

from excel_tools import split_excel


def main():
    project_dir = Path(__file__).resolve().parent.parent
    input_dir = project_dir / "Input" / "ExcelSplitter"
    output_dir = project_dir / "Output" / "ExcelSplitter"

    filename = input("Enter Excel filename: ").strip()
    rows = input("Enter number of rows per split (default 20000): ").strip()
    rows_per_file = int(rows) if rows.isdigit() else 20_000

    input_file = input_dir / filename
    created_files = split_excel(input_file, output_dir, rows_per_file)
    print(f"Created {len(created_files)} file(s) in {output_dir}")


if __name__ == "__main__":
    main()
