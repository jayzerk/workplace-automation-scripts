# Excel Splitter (Quick Tutorial)

This script splits large Excel files into smaller chunks for easier handling.

---

## Setup
- Install Python 3.9+
- Install required libraries:
  pip install pandas openpyxl

- Folder structure:
  Python/
  ├── Input/ExcelSplitter/ # Put your Excel files here
  ├── Output/ExcelSplitter/ # Split files will be saved here
  └── Scripts/excel-splitter.py # The script

---

## How to Use
1. Place your Excel file in Input/ExcelSplitter
2. Run the script:
   python Scripts/excel-splitter.py
3. Follow the prompts:
   Enter Excel filename: filename.xlsx
   Enter number of rows per split (default 20000): 20000
4. The script shows a summary:
   --- SUMMARY ---
   File: filename.xlsx
   Total rows: 160000
   Rows per file: 20000
   Files to be created: 8
5. Confirm:
   Proceed with splitting? (Y/N): Y
6. Split files will appear in Output/ExcelSplitter:
   filename_split_1.xlsx
   filename_split_2.xlsx
   ...

---

## ✅ Notes
- Only .xlsx files are supported
- Default split size is 20,000 rows per file
- Headers are preserved in all split files
- Large files may take time depending on your system
