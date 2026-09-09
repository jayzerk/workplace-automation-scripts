# Excel Tools

I put this project together to make a couple of repetitive Excel tasks easier. Instead of running separate scripts and answering prompts in a command window, you can open one Windows app, choose the tool you need, and fill in a simple form.

The app currently includes:

- **Queue Board** — maintains draggable General, Monitored, and Monthly stacks with colored, clickable bricks.
- **Excel Splitter** — splits a large workbook into smaller files while keeping the column headers.
- **CBT Consolidation** — reads the expected CBT worksheet layout and creates a summary report containing gender totals and unique provider qualification counts.
- **Geographic Name Cleaner** — standardizes region, province, and municipality names using canonical Excel files in the `Metadata` folder.

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

The **Work Queue & Excel Tools** window will open. Choose a page from the left side. Excel processing runs in the background, so the window should remain responsive even with a large file.

### Queue Board

The Queue Board is the app's main page. Click **+ Add Brick**, choose General,
Monitored, or Monthly, enter a name, pick a color, and optionally add an approved
HTTPS URL or absolute local file path. Click a linked brick to open it, drag it
to reorder its stack or move it between columns, or right-click it to edit or
delete it. Dragging leaves a drop silhouette while nearby cards ease into
position, and the released card animates smoothly into its saved slot.

Right-click an active brick and choose **Archive** to remove it from the active
board without deleting it. **Show Archive** swaps the whole three-column board
between active and archived General, Monitored, and Monthly stacks. Archived cards
retain their original color beneath a gray overlay and cannot be opened or
dragged. Right-click one to restore, edit, or delete it.

If a saved link no longer passes the current security rules, only that brick is
automatically archived; all other bricks continue loading. Repair or remove its
link through Edit before restoring it.

Google Workspace links hosted on `docs.google.com` are allowed by default,
including Sheets, Docs, Slides, and Forms. Use **Allowed Domains...** to add
other exact domains such as `drive.google.com`, or explicit wildcard entries
such as `*.example.com`. HTTP, embedded URL credentials, nonstandard web ports,
and unlisted domains are blocked. Local executable, script, shortcut, archive,
and macro-capable or active-document file links are also blocked. Modern
`.docx`, `.xlsx`, and `.pptx` files are accepted.

The app makes no background web requests. It asks Windows to open a web link
only when you click a linked brick, after checking the HTTPS address against
your allowlist. Adding a domain means you trust that domain; the destination is
opened in your normal browser and is then subject to the browser's own security.

Use **Launch at sign-in** to explicitly add or remove this app's per-user Windows
startup entry. The setting does not require administrator access and manages
only the `ExcelToolsQueueBoard` entry.

Queue data is saved automatically beside the executable at:

```text
Data\queue_board.json
Data\settings.json
```

The JSON file is written atomically so an interrupted save does not partially
replace the existing queue.

### Excel Splitter

1. Select **Excel Splitter**.
2. Browse to an `.xlsx` workbook. Macro-enabled workbooks are blocked.
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

### Geographic Name Cleaner

Place these workbooks in the `Metadata` folder beside the app:

```text
regions.xlsx
provinces.xlsx
municipalities.xlsx
```

Each metadata workbook must contain a column named `name`. For hierarchy-aware
matching, `regions.xlsx` should also contain `id`, `provinces.xlsx` should contain
`id` and `region_id`, and `municipalities.xlsx` should contain `id` and
`province_id`.

Select an incoming `.xlsx` workbook and click **Fix Geographic Names**.
The cleaner searches the first 20 rows of every worksheet for these headers,
ignoring case and punctuation:

```text
region / regions
province / provinces
municipality / municipalities
```

It creates a corrected copy and a separate match report. The original workbook
is never overwritten. Ambiguous and low-confidence values remain unchanged and
are included in the report for review.

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

Windows may show a SmartScreen warning because a locally built executable is not
digitally signed. Code signing is the proper way to establish publisher identity;
do not create a broad antivirus exclusion for the app. Scan the finished build
and submit it to Microsoft as a false positive if Defender incorrectly detects it.

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
├── Metadata/
├── Data/
└── Output/
```

## A note about data privacy

Input workbooks, generated reports, build files, and the executable are excluded through `.gitignore`. They should not be included by a normal `git add .` command. It is still a good habit to check `git diff --cached --name-only` before pushing, especially when working with workplace data.

Please do not commit real employee, provider, customer, or other confidential information to a public repository.
