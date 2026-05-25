import os
import pandas as pd

def split_excel(input_file, rows_per_file=20000):
    base_dir = os.path.dirname(os.path.abspath(__file__))

    input_dir = os.path.join(os.path.dirname(base_dir), "Input", "ExcelSplitter")
    output_dir = os.path.join(os.path.dirname(base_dir), "Output", "ExcelSplitter")

    os.makedirs(output_dir, exist_ok=True)

    file_path = os.path.join(input_dir, input_file)

    df = pd.read_excel(file_path, engine="openpyxl")
    total_rows = len(df)

    num_files = (total_rows + rows_per_file - 1) // rows_per_file
    print("\n--- SUMMARY ---")
    print(f"File: {input_file}")
    print(f"Total rows: {total_rows}")
    print(f"Rows per file: {rows_per_file}")
    print(f"Files to be created: {num_files}\n")

    confirm = input("Proceed with splitting? (Y/N): ").strip().lower()
    if confirm != "y":
        print("Operation cancelled.")
        return

    for i in range(0, total_rows, rows_per_file):
        chunk = df.iloc[i:i+rows_per_file]
        output_file = os.path.join(
            output_dir,
            f"{os.path.splitext(input_file)[0]}_split_{i//rows_per_file + 1}.xlsx"
        )
        chunk.to_excel(output_file, index=False, engine="openpyxl")
        print(f"Saved {output_file} with {len(chunk)} rows")

if __name__ == "__main__":
    filename = input("Enter Excel filename: ").strip()
    rows = input("Enter number of rows per split (default 20000): ").strip()
    rows_per_file = int(rows) if rows.isdigit() else 20000

    split_excel(filename, rows_per_file)
