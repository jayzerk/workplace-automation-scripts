"""Command-line entry point for PWD report normalization."""

import sys
from pathlib import Path

# Support running directly or from repository root
script_dir = Path(__file__).resolve().parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

from pwd_data import format_pwd_report


def main():
    project_dir = script_dir.parent
    input_dir = project_dir / "Input" / "PWD"
    output_dir = project_dir / "Output" / "PWD"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    if len(sys.argv) > 1:
        target_path = Path(sys.argv[1]).resolve()
        out_target = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else output_dir
        sheet = sys.argv[3] if len(sys.argv) > 3 else None
        print(f"Processing: {target_path}")
        output_file = format_pwd_report(target_path, out_target, sheet_name=sheet)
        print(f"Report generated: {output_file}")
        return

    # Check for files in Input/PWD and forAgent
    candidate_files = sorted(input_dir.glob("*.xlsx"))
    agent_dir = project_dir / "forAgent"
    if agent_dir.is_dir():
        candidate_files.extend(sorted(agent_dir.glob("*.xlsx")))

    if not candidate_files:
        path_input = input("Enter path to Excel file: ").strip().strip('"')
        if not path_input:
            print("No file specified.")
            return
        target_path = Path(path_input).resolve()
    else:
        print("Available files:")
        for idx, file_path in enumerate(candidate_files, 1):
            print(f" [{idx}] {file_path.name} ({file_path.parent.name}/)")
        choice = input(f"\nSelect a file [1-{len(candidate_files)}] or enter custom path: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(candidate_files):
            target_path = candidate_files[int(choice) - 1]
        else:
            target_path = Path(choice.strip('"')).resolve()

    sheet = input("Worksheet name (press Enter for default Sheet2/first sheet): ").strip()
    sheet_name = sheet if sheet else None

    def progress(current, total, msg):
        print(f"[{current:3.0f}%] {msg}")

    output_file = format_pwd_report(
        target_path, output_dir, sheet_name=sheet_name, progress_callback=progress
    )
    print(f"\nPWD report successfully created:\n{output_file}")


if __name__ == "__main__":
    main()
