# Qualification Report

A Python-based Excel reporting tool that measures qualification coverage for technicians listed in the Dark Universe Line Schedule.

The report matches technicians to IBM Maximo qualification exports by SSO, applies schedule-specific qualification rules, calculates qualification coverage percentages, and generates a formatted Excel workbook with a color-coded dashboard.

## Purpose

The report answers two main questions:

1. What percentage of the applicable technicians hold each qualification?
2. Up to what module level is each technician qualified?

Only technicians listed in the current Line Schedule are included. People appearing in the qualification exports but not in the Line Schedule are ignored.

## Data Sources

### Line Schedule

Place one current Excel workbook in:

```text
Input\Line Schedule
```

The workbook may have any `.xlsx` filename, but the folder must contain only one current workbook.

The report uses only:

- Technician name
- SSO / ID Number
- Worksheet or schedule group

Open positions, section headings, staffing totals, and entries without a valid nine-digit SSO are excluded from qualification calculations. Entries with a name but no valid SSO are sent to the Exceptions sheet.

Expected worksheet names:

```text
1st Shift
2nd Shift
3rd Shift Unchained
3rd Shift Wolf
```

### Qualification Reports

Qualification reports are read from the folder configured in:

```text
Config\settings.json
```

Current source folder:

```text
G:\Dev\Qual_Tracker\Input\Qualification
```

The report expects these 11 unedited Maximo exports:

```text
CUMOD0 Qualification.xlsx
CUMOD1 Qualification.xlsx
CUMOD2 Qualification.xlsx
CUMOD3 Qualification.xlsx
DUMOD0 Qualification.xlsx
DUMOD1 Qualification.xlsx
UCMOD0 Qualification.xlsx
UCMOD1 Qualification.xlsx
UCMOD2 Qualification.xlsx
UCMOD3 Qualification.xlsx
UCMOD4 Qualification.xlsx
```

Each export must contain a worksheet named `List of Labor` with these required columns:

- Labor
- Name
- Effective Date
- Status

Only records with `Status = ACTIVE` count toward qualification coverage.

## Qualification Rules

### Operational: 1st and 2nd Shift

Count all qualifications:

- Darkmoor: DUMOD0 and DUMOD1
- Curse of the Werewolf: CUMOD0 through CUMOD3
- Monsters Unchained: UCMOD0 through UCMOD4

### 3rd Shift Unchained

Count only Monsters Unchained qualifications:

- UCMOD0 through UCMOD4

### 3rd Shift Wolf

Count only Curse of the Werewolf and Darkmoor qualifications:

- CUMOD0 through CUMOD3
- DUMOD0 through DUMOD1

## Matching Logic

- SSO is the primary key.
- Technician names are used for display only.
- Only technicians with valid SSOs in the Line Schedule are included.
- Qualification records for people outside the Line Schedule are ignored.
- If multiple active records exist for the same SSO and qualification, the record with the latest valid effective date is retained and the duplicate is reported.
- Inactive records do not count and are reported as exceptions when they match a listed technician.

## Percentage Logic

Module coverage is calculated as:

```text
Qualified applicable technicians / Total applicable technicians
```

Venue Total is calculated as:

```text
Total achieved module occurrences /
(Applicable technicians x Number of modules for the venue)
```

This is equivalent to averaging the module coverage percentages before display rounding.

## Dashboard Colors

- 90% to 100%: Green
- 80% to 89.9%: Light green
- 70% to 79.9%: Yellow
- 60% to 69.9%: Orange
- Below 60%: Light red

## Generated Workbook

The report creates one workbook in:

```text
Output
```

Filename format:

```text
Dark Universe Qualification Coverage YYYY-MM-DD.xlsx
```

The workbook contains:

### Dashboard

Shows color-coded module percentages, venue totals, and applicable technician populations for:

- 3rd Shift
- Operational
- Total

### Technician Qualifications

Shows one row per technician, including:

- Name
- SSO
- Schedule group
- Qualification indicators
- Qualification effective dates
- Highest contiguous qualification level by venue
- Review flags for qualification gaps

### Venue Summary

Shows the qualified count, applicable population, and percentage for each module and reporting section.

### Exceptions

Shows issues such as:

- Named Line Schedule entries without a valid SSO
- Missing or unrecognized qualification files
- Inactive qualification records
- Invalid effective dates
- Duplicate active qualification records

### Run Summary

Shows the run timestamp, output filename, technician totals, applicable populations, active qualification matches, and exception count.

## Running the Report

Double-click:

```text
Run Qualification Report.bat
```

The runner will:

1. Use the local Python virtual environment.
2. Move existing `.xlsx` workbooks from `Output` into a timestamped folder under `Archive`.
3. Generate the new current workbook in `Output`.
4. Open the `Output` folder after a successful run.
5. Keep all input reports in their original locations.

Close the current Excel report before running again so Windows can move the file into Archive.

## Manual Run

From the project root with the virtual environment activated:

```powershell
python Scripts\report_generator.py
```

To review calculations in the terminal without generating Excel:

```powershell
python Scripts\coverage_processor.py
```

To inspect qualification file recognition and matches:

```powershell
python Scripts\qualification_reader.py
```

To inspect the Line Schedule roster:

```powershell
python Scripts\line_schedule_reader.py
```

## Project Structure

```text
Qualification Report
â”œâ”€â”€ Archive
â”œâ”€â”€ Config
â”‚   â””â”€â”€ settings.json
â”œâ”€â”€ Documentation
â”œâ”€â”€ Input
â”‚   â””â”€â”€ Line Schedule
â”œâ”€â”€ Logs
â”œâ”€â”€ Output
â”œâ”€â”€ Scripts
â”‚   â”œâ”€â”€ coverage_processor.py
â”‚   â”œâ”€â”€ line_schedule_reader.py
â”‚   â”œâ”€â”€ qualification_reader.py
â”‚   â””â”€â”€ report_generator.py
â”œâ”€â”€ Tests
â”œâ”€â”€ .gitignore
â”œâ”€â”€ README.md
â””â”€â”€ Run Qualification Report.bat
```

## Configuration

Current `Config\settings.json` structure:

```json
{
  "line_schedule_folder": "Input\\Line Schedule",
  "qualification_folder": "G:\\Dev\\Qual Tracker\\Input\\Qualification",
  "output_folder": "Output",
  "archive_folder": "Archive",
  "logs_folder": "Logs"
}
```

Update the configured paths if the project or qualification source folder is moved.

## Environment

The project currently uses:

- Python 3.14.8
- pandas 3.0.6
- openpyxl 3.1.5

The virtual environment is stored locally in `.venv` and is excluded from Git.

## Git and Data Protection

The repository excludes operational and employee data, including:

- `.venv`
- Line Schedule workbooks
- Generated output workbooks
- Archived workbooks
- Logs
- Temporary Excel files

Do not commit employee rosters, SSOs, Maximo exports, generated reports, or other operational data to GitHub.

