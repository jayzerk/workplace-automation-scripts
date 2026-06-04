import os
import pandas as pd

def generate_gender_report(df, worksheet, start_row, start_col=0):
    row2 = df.iloc[1].tolist()
    row3 = df.iloc[2].tolist()
    numeric_part = df.iloc[3:, 11:]
    summation = numeric_part.apply(pd.to_numeric, errors="coerce").sum().tolist()

    groups, subs, values = [], [], []
    current_group = None
    for i, (group_val, sub_val) in enumerate(zip(row2[11:], row3[11:]), start=11):
        if pd.notna(group_val) and str(group_val).strip() != "":
            current_group = str(group_val).strip()
        groups.append(current_group if current_group else "")
        subs.append(str(sub_val))
        values.append(summation[i-11])

    merge_format = worksheet.book.add_format({
        "align": "center", "valign": "vcenter", "bold": True, "text_wrap": True
    })
    sub_format = worksheet.book.add_format({"align": "center", "bold": True})

    for r, (sub, val) in enumerate(zip(subs, values)):
        row_index = start_row + r
        worksheet.write(row_index, start_col+1, sub, sub_format)
        worksheet.write(row_index, start_col+2, val)

    start = 0
    current = groups[0]
    for idx, grp in enumerate(groups):
        if grp != current:
            worksheet.merge_range(
                start_row+start, start_col, start_row+idx-1, start_col, current, merge_format
            )
            start = idx
            current = grp
    worksheet.merge_range(
        start_row+start, start_col, start_row+len(groups)-1, start_col, current, merge_format
    )

    worksheet.set_column(start_col, start_col, 25)
    worksheet.set_column(start_col+1, start_col+1, 15)
    worksheet.set_column(start_col+2, start_col+2, 12)

    


def generate_unique_providers(df, worksheet, start_row, start_col=4):
    providers = df.iloc[3:, 3].dropna().unique().tolist()

    header_format = worksheet.book.add_format({"bold": True, "align": "left"})
    num_format = worksheet.book.add_format({"align": "center"})

    worksheet.write(start_row, start_col, "Unique Providers", header_format)
    worksheet.write(start_row, start_col+1, "Unique Qualifications Count", header_format)

    for i, provider in enumerate(providers, start=1):
        worksheet.write(start_row + i, start_col, provider)

        provider_rows = df.iloc[3:, :]
        provider_rows = provider_rows[provider_rows.iloc[:, 3] == provider]

        unique_quals = provider_rows.iloc[:, 5].dropna().unique()
        worksheet.write(start_row + i, start_col+1, len(unique_quals), num_format)



def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.abspath(os.path.join(script_dir, "..", "Input", "CBT"))
    output_dir = os.path.abspath(os.path.join(script_dir, "..", "Output", "CBT"))

    if not os.path.exists(input_dir):
        print(f"Folder '{input_dir}' does not exist.")
        return

    files = os.listdir(input_dir)
    if not files:
        print(f"No files found in {input_dir}.")
        return

    print(f"Files available in {input_dir}:")
    for f in files:
        print(f" - {f}")

    file_name = input("\nEnter Excel file name from the list above: ").strip()
    file_path = os.path.join(input_dir, file_name)

    if not os.path.exists(file_path):
        print(f"\nFile '{file_name}' not found in {input_dir}")
        return

    print(f"\nFile found: {file_name}")

    try:
        df = pd.read_excel(file_path, sheet_name="Sheet2", header=None)
        os.makedirs(output_dir, exist_ok=True)
        export_file = os.path.join(output_dir, f"consolidated_{file_name}")

        with pd.ExcelWriter(export_file, engine="xlsxwriter") as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Report")
            worksheet.book = workbook  # attach workbook reference

            generate_gender_report(df, worksheet, start_row=0, start_col=0)
            generate_unique_providers(df, worksheet, start_row=4)

        print(f"\n✅ Consolidated report created: {export_file}")

    except Exception as e:
        print(f"\nError reading 'Sheet2' from {file_name}: {e}")


if __name__ == "__main__":
    main()
