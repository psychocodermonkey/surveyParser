# Survey Name Extraction Tool

This tool processes a spreadsheet of survey responses and extracts named mentions from each response.
It generates two outputs:

- A flattened Excel report for easy review
- A SQLite database for querying and analysis

The script expects a specific directory layout and input file structure.

---

## Directory Layout

```text
project-root/
│
├── processSpreadsheet.py
└── data/
    ├── import.xlsx
    ├── exclude_names.txt
    ├── name_mentions.xlsx      (generated output)
    └── name_mentions.sqlite    (generated output)
```

All input and output files live in the **`data/` directory**.

---

## Required Input Files

### `data/import.xlsx`

The spreadsheet containing the survey data.

The script expects the following column order:

| Column | Description       |
| ------ | ----------------- |
| A      | College           |
| B      | Department        |
| C      | Semester          |
| D      | Question          |
| E      | Response / Answer |
| F      | Mentions          |

#### Mentions Column

The **Mentions** column should contain names separated by semicolons.

Example:

``` text
John Smith; Jane Doe; Dr. Robert Brown
```

Each name will be associated with the response in that row.

The script does **not** attempt to detect names in the response text.
It relies entirely on the contents of the **Mentions** column.

---

### `data/exclude_names.txt`

Optional list of entries that should **not** be treated as individual names.

Sometimes the Mentions column may contain phrases, organizations, or generic references instead of actual people.
Items listed in this file will be ignored during processing.

Each entry should appear on its own line:

``` text
company
support team
management
the organization
```

Entries are matched case-insensitively.

This file is useful for filtering out survey phrases or generic terms that appear in the Mentions column but are not actual people.

---

## Generated Output Files

Running the script creates two output files.

### `data/name_mentions.xlsx`

A flattened spreadsheet where each row represents a single name linked to a response.

Example structure:

| Name | College | Department | Response |

If a response mentions three names, three rows will appear in the output.

---

### `data/name_mentions.sqlite`

SQLite database containing structured tables for analysis.

Tables created:

* `responses`
* `names`
* `questions`
* `response_names` (junction table linking names to responses)

This database can be queried using any SQLite client.

---

## Running the Script

From the project root:

```bash
python processSpreadsheet.py
```

The script will:

1. Load `data/import.xlsx`
2. Parse the Mentions column
3. Filter excluded entries
4. Generate the Excel and SQLite outputs

---

## Changing Input or Output File Names

File locations and names are defined at the **top of `processSpreadsheet.py`**.

Example:

```python
DATA_DIR = Path(__file__).parent / "data"
EXCEL_FILE = DATA_DIR / "import.xlsx"
OUTPUT_XLSX = DATA_DIR / "name_mentions.xlsx"
SQLITE_DB = DATA_DIR / "name_mentions.sqlite"
```

To change file names or locations, update these variables.

For example, if the input file is renamed:

```python
EXCEL_FILE = DATA_DIR / "survey_data.xlsx"
```

Or if you want a different output name:

```python
OUTPUT_XLSX = DATA_DIR / "survey_mentions.xlsx"
```

No other changes are required as long as the file paths remain valid.

---

## Notes

- The script assumes the spreadsheet format is consistent.
- The Mentions column must contain semicolon-separated names.
- Non-name entries can be filtered using `exclude_names.txt`.
- If the spreadsheet column order changes, the script will need to be updated.
