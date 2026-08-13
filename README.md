# Excel Tools

I put this project together to make a couple of repetitive Excel tasks easier. Instead of running separate scripts and answering prompts in a command window, you can open one Windows app, choose the tool you need, and fill in a simple form.

The app currently includes:

- **Excel Splitter** — splits a large workbook into smaller files while keeping the column headers.
- **CBT Consolidation** — reads the expected CBT worksheet layout and creates a summary report containing gender totals and unique provider qualification counts.

## Installing the project

You will need a current 64-bit version of Python 3. The project is tested with Python 3.13 on Windows 11.

First, clone the repository and move into its folder:

```powershell
git clone https://github.com/jayzerk/workplace-automation-scripts.git
cd workplace-automation-scripts
```

I recommend using a virtual environment so the project's packages stay separate from your other Python projects:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks the activation script, you can run the virtual environment's Python directly:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

## Running the Windows app

From the project folder, run:

```powershell
python app.py
```

The **Excel Tools** window will open. Choose a tool from the left side, select your input workbook and output folder, enter the requested options, and start the job. Excel processing runs in the background, so the window should remain responsive even with a large file.

### Excel Splitter

1. Select **Excel Splitter**.
2. Browse to an `.xlsx` or `.xlsm` workbook.
3. Choose where the split files should be saved.
4. Enter the maximum number of rows per file. The default is `20,000`.
5. Click **Split Workbook**.

The generated files are named like this:

```text
original_name_split_1.xlsx
original_name_split_2.xlsx
original_name_split_3.xlsx
```

### CBT Consolidation

1. Select **CBT Consolidation**.
2. Browse to the CBT workbook.
3. Choose an output folder.
4. Enter the worksheet name. It defaults to `Sheet2`.
5. Click **Create Report**.

The report will be saved as:

```text
consolidated_original_name.xlsx
```

The CBT tool expects the same worksheet structure used by the original script. If the selected sheet is missing or its rows and columns do not match that structure, the app will show an error instead of creating an incomplete report.

## Building the Windows executable

The repository includes an `ExcelTools.spec` file with the packaging settings already configured. Build the app with:

```powershell
pyinstaller --noconfirm ExcelTools.spec
```

When the build finishes, the executable will be located at:

```text
dist\ExcelTools\ExcelTools.exe
```

This is a folder-based build. If you want to copy the app to another Windows computer, copy the entire `dist\ExcelTools` folder. The other computer does not need Python installed, but the `.exe` will not work if it is separated from the accompanying `_internal` folder.

Windows may show a SmartScreen warning because a locally built executable is not digitally signed. If you built it yourself and trust the source, use **More info → Run anyway**.

## Running the original command-line scripts

The command-line versions are still available:

```powershell
python Scripts\excel-splitter.py
python Scripts\cbt-consolidation.py
```

These scripts use the corresponding folders under `Input` and `Output`, while the desktop app lets you select files and folders from anywhere on your computer.

## Project layout

```text
workplace-automation-scripts/
├── app.py
├── ExcelTools.spec
├── requirements.txt
├── Scripts/
│   ├── excel_tools.py
│   ├── excel-splitter.py
│   └── cbt-consolidation.py
├── Input/
└── Output/
```

## A note about data privacy

Input workbooks, generated reports, build files, and the executable are excluded through `.gitignore`. They should not be included by a normal `git add .` command. It is still a good habit to check `git diff --cached --name-only` before pushing, especially when working with workplace data.

Please do not commit real employee, provider, customer, or other confidential information to a public repository.
