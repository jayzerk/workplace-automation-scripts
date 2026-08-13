"""Command-line entry point for CBT report consolidation."""

from pathlib import Path

from excel_tools import consolidate_cbt


def main():
    project_dir = Path(__file__).resolve().parent.parent
    input_dir = project_dir / "Input" / "CBT"
    output_dir = project_dir / "Output" / "CBT"

    files = sorted(path.name for path in input_dir.glob("*.xlsx"))
    if not files:
        print(f"No .xlsx files found in {input_dir}.")
        return

    print(f"Files available in {input_dir}:")
    for filename in files:
        print(f" - {filename}")

    filename = input("\nEnter Excel file name from the list above: ").strip()
    output_file = consolidate_cbt(input_dir / filename, output_dir)
    print(f"\nConsolidated report created: {output_file}")


if __name__ == "__main__":
    main()
