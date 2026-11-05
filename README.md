# CA1 Top Sheets Automation

Generate one completed Word top sheet and PDF per student from any Excel workbook and compatible DOCX template.

The workbook and template are supplied at runtime, so their filenames and locations are entirely up to the user. Original input files are never changed.

Nothing about the sheet's layout is hard-coded. What gets filled, which workbook column feeds which label, the exam date, the full marks, the fonts and the duplicate rules all live in [`config.json`](config.json). The values that change from exam to exam are asked for when you start a run, so you rarely need to edit a config file by hand. The question rows and their marks are read from the template itself, so a top sheet with more rows or fewer columns works without touching the code.

## Features

- Reads student details and total marks from an XLSX workbook. **The roll number is the primary key.**
- Fills the labels you declare in `config.json` in a DOCX top-sheet template, keeping the template's own font and formatting (or a custom font, your choice).
- Reads the question rows and their `Marks Allotted` from the template's marks table, groups them into sections ("answer any 5 of 1.a to 1.g"), and spreads each student's total across the answered questions so that it adds up exactly.
- Adds a typed student signature in one of five handwriting-style fonts, with the examination date beneath it.
- Asks for the per-exam values (exam date, full marks, paper code, subject teacher, font) on a short settings screen at the start of each run, and can remember your answers.
- Checks the whole workbook before generating anything: duplicate roll numbers, bad marks, missing columns, missing template labels.
- Creates files named `Rollnumber_Name.docx` and `Rollnumber_Name.pdf`, and checks for filename collisions.
- Writes a `run_report.csv` with every student's question-wise marks and status.
- Generates DOCX files in parallel, then exports PDFs with Microsoft Word, or with LibreOffice if Word is not available.

## Requirements

- Python 3.10 or later
- For PDF export, either Microsoft Word on Windows (uses `pywin32`) or LibreOffice on any system. Not needed with `--keep-docx-only`.

Install the Python dependencies:

```powershell
py -m pip install -r requirements.txt
```

## Quick start

After placing or cloning this project on a computer, open PowerShell in the project folder and run:

```powershell
py .\generate_top_sheets.py
```

Two standard file pickers will open in sequence:

1. Choose the student Excel workbook (`.xlsx`).
2. Choose the Word top-sheet template (`.docx`).

Next comes a short settings screen (see [Settings screen](#settings-screen)). Press Enter to accept the values shown, or type a number to change one.

The first time a template is used, the program shows the question sections it found and asks how many questions are to be answered in each (see [Sections](#sections-answer-any-n)). The answers are saved, so later runs with the same template do not ask again.

The generated files are written to `output\docx` and `output\pdf`. Input filenames do not need to match any predefined names.

## Settings screen

At the start of every interactive run the program shows the values that usually change from exam to exam:

```
Settings for this run
  1. Exam date                02-09-2026
  2. Full marks               25
  3. Paper (course) Code      (blank)
  4. Subject Teacher          (blank)
  5. UPID                     (blank)
  6. Font mode                template

Press Enter to continue, or type a number to change it:
```

- Press Enter to carry on, or type a line's number to change it. At a value prompt, Enter keeps the current value and `-` clears it.
- The exam date is checked as `DD-MM-YYYY` (`2-9-2026` becomes `02-09-2026`). Any other format is accepted only after you confirm it.
- If the template states its own "Full Marks: N" text and you have not set a value yourself, that number is offered as the full marks. If you have set a different one, the screen points out the difference. Nothing is ever filled from the template without showing it here.
- Every fixed-value entry in `template_fields` appears on the screen automatically, except derived ones such as `{exam_date}`. Add a field to the config and it shows up here.
- After a change you are asked whether to save it. **Saved answers go to `config.local.json`**, a file beside `config.json` that is ignored by git, so personal values never reach the repository. It overrides `config.json` setting by setting, and it is safe to delete to start again.
- The screen is skipped with `--no-prompt`, with `--check`, and when nobody is there to answer. Those runs use `config.json` and `config.local.json` exactly as they are.
- The shipped `config.json` has no exam date, so the first run always asks. If the date is still empty, the "Date of Examination" field and the date under the signature are left blank, with a warning.

## Configuration (`config.json`)

Every value that may change between exams is here. Settings you leave out use the built-in default, and a misspelt setting name is rejected with an error rather than ignored. `config.local.json` (written by the settings screen, or by hand) uses the same setting names and is checked just as strictly.

| Setting | Meaning |
| --- | --- |
| `exam_date` | Printed in any field whose value is `{exam_date}` and under the signature; empty in the shipped config |
| `full_marks` | Highest total a student can have; totals above it are errors |
| `font_mode` | `"template"` copies the template's own font; `"custom"` uses `custom_font` for everything the program fills in |
| `custom_font` | `name` and optional `size_pt` used in `"custom"` mode |
| `columns` | Workbook headers for `roll_number` (primary key), `name` and `marks_obtained` |
| `absent_markers` | Values in the marks column that mean absent, for example `["Ab"]` |
| `template_fields` | The labels to fill, see below |
| `fixed_values` | Overrides for fixed-value labels, as `{"Subject Teacher:": "Dr. A. Sen"}`. Normally written by the settings screen into `config.local.json` |
| `marks_table` | Header texts of the marks table in the template; `null` switches question-wise marks off |
| `ignore_question_labels` | Marks-table rows with these labels (for example `total`) are not questions |
| `signature` | `enabled`, the label to insert above, the five fonts, and the name and date sizes |
| `duplicates` | `"stop"`, `"keep-first"` or `"keep-last"` for conflicting rows with the same roll number |
| `existing_files` | `"overwrite"` or `"skip"` for files already in the output folder |
| `unreachable_marks` | `"stop"` or `"skip"` for a total that cannot be split under the section rules |
| `full_marks_label` | Label of the "Full Marks" text in the template, used only for a warning if it disagrees with `full_marks`; `null` to skip |
| `mobile_check` | Column and digit count for a warning-only mobile number check; `null` to skip |

### Filling labels

Each entry in `template_fields` links one template label to either a workbook column or a fixed value:

```json
{ "label": "Name of the Student:", "column": "Name" }
{ "label": "Date of Examination:", "value": "{exam_date}" }
{ "label": "Subject Teacher:",    "value": "Dr. A. Sen" }
{ "label": "UPID:",               "value": "" }
```

- Fixed values can be typed on the [settings screen](#settings-screen) instead of editing the config.
- A `column` must match a workbook header exactly. A `value` of `""` leaves that label exactly as the template has it.
- **Nothing is matched by guesswork.** If a declared header or label is missing, the run stops and the message names what was not found and what was found instead. Labels that are not declared are never touched.
- The value replaces everything after the label in its paragraph; the label and anything before it stay.

## Input requirements

### Excel workbook

The selected worksheet must have the headers declared in `config.json` in its first row. With the default config:

| Header | Used for |
| --- | --- |
| `Name` | Student name and typed signature |
| `UNI Roll number` | Roll number (primary key) and output filename |
| `Mobile Number` | Mobile Number field |
| `Marks Obtained` | Total score and question-wise allocation |

`Marks Obtained` must be a whole number from 0 to `full_marks`, or `Ab`. [Example template](https://docs.google.com/spreadsheets/d/1g16tx5055atuTEP42zXqHH92BK4apdje/edit?usp=sharing&ouid=105329508417402305211&rtpof=true&sd=true)

**Roll number as the primary key.** Roll numbers are normalised first (`34900725001`, `34900725001.0` and ` 34900725001 ` are the same student). If a roll number appears more than once:

- rows with identical details are merged into one sheet;
- rows with different details are a conflict. With `"duplicates": "stop"` (the default) the run stops before generating anything and lists every row involved. `"keep-first"` and `"keep-last"` pick one row and warn.

Rows with a blank name or roll number are skipped with a note. Blank values in a column that a label is linked to are reported, and that label is left as the template has it.

### Word template

Every label in `template_fields` must appear in the template exactly once, plus the signature label (`Signature of the student with date` by default). The typed signature and date are inserted above the signature label.

The question-wise marks need a table whose header row contains `Q. No.`, `Marks Allotted` and `Marks Awarded` (set under `marks_table`). Column order and any other columns do not matter, and any number of question rows is fine. Rows whose `Marks Allotted` cell is blank can never receive marks. [Example template](https://docs.google.com/document/d/1MEd09kKXdDmPIVZjeK-MCoThG1fBDorq/edit?usp=sharing&ouid=105329508417402305211&rtpof=true&sd=true)

## Sections ("answer any N")

Exam papers usually list more questions than a student answers. The program groups the question rows into sections and asks how many are answered in each:

```
Section 1: 1.a) to 1.g)  (7 rows; 6 carry marks: 1 each)
Section 2: 2 to 7  (6 rows; 5 carry marks: 5 each)

Answer how many questions from Section 1? [1-6, Enter = 6]: 5
Answer how many questions from Section 2? [1-5, Enter = 5]: 4

A student's total can range from 9 to 25 under these rules (full marks in the config: 25).
```

- Rows with sub-part labels (`1.a)`, `1.b)` ...) form one section, and consecutive plain numbers (`2`, `3`, `4` ...) form another. If that grouping is wrong, type your own row ranges at the prompt, for example `1-7,8-13`.
- In every section **exactly N** questions are answered, each with at least 1 mark and at most its `Marks Allotted`. The other questions stay blank. The marks add up exactly to `Marks Obtained`.
- The reachable range is shown. A total outside it (for example 5 when N = 5 and 4 means at least 9) cannot be split, so the run stops and names the student (`"unreachable_marks": "skip"` skips them instead). `Ab` leaves every question blank.
- The answers are saved in `section_rules.json` (ignored by git), tied to a fingerprint of the template's question rows. A different template asks again; `--ask-sections` forces the questions for the current one. `--answer 5,4` supplies the numbers without prompting.
- The split is varied but repeatable for a given roll number, template and rules.

## Checking before generating

```powershell
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --check
```

`--check` validates the workbook, config, template and section rules, prints every problem and warning, and writes nothing.

## Command-line options (optional)

For repeatable or scripted runs, provide the paths directly instead of using the file pickers:

```powershell
py .\generate_top_sheets.py --workbook "C:\path\to\students.xlsx" --template "C:\path\to\top-sheet-template.docx"
```

By default, generated files are saved alongside the script:

```
output/
  docx/
    34900725001_Durjoy pandit.docx
  pdf/
    34900725001_Durjoy pandit.pdf
  run_report.csv
```

Useful options:

```powershell
# Use a named worksheet
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --sheet "Sheet1"

# Choose an output directory and six DOCX-generation workers
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --output ".\generated" --workers 6

# Create DOCX files only; do not export PDFs
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --keep-docx-only

# Change values for one run without editing config.json
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --exam-date "15-12-2026" --full-marks 30 --font-mode custom --custom-font "Calibri"

# Export PDFs with LibreOffice instead of Word
py .\generate_top_sheets.py --workbook ".\students.xlsx" --template ".\template.docx" --pdf-engine libreoffice
```

| Option | Meaning |
| --- | --- |
| `--config PATH` | Use another config file (default: `config.json` beside the script) |
| `--check` | Validate and report only |
| `--no-prompt` | Skip the settings screen and use the config files as they are |
| `--answer 5,4` | Questions to answer per section, without prompting |
| `--ask-sections` | Ask about the sections again even if rules are saved |
| `--pdf-engine` | `auto` (default), `word` or `libreoffice`; `auto` uses Word when available, otherwise LibreOffice |
| `--soffice PATH` | Location of LibreOffice's `soffice` program, if it is not found automatically |
| `--duplicates` | `stop`, `keep-first` or `keep-last`; overrides `config.duplicates` for one run |
| `--existing` | `overwrite` or `skip`; overrides `config.existing_files` for one run |

## Fonts

- `"font_mode": "template"` (default): the values the program fills in take the exact font, size and bold of the template's own label, and each awarded mark takes the formatting of that row's `Marks Allotted` cell.
- `"font_mode": "custom"`: the same, but with `custom_font` applied to everything the program fills in. The template's own text is never changed.
- The typed signature always uses the handwriting-style fonts. The font is chosen from the roll number, exactly as before.

## Output files and collisions

Files are named `Rollnumber_Name.docx` and `Rollnumber_Name.pdf`. Because the roll number is unique, two students can never share a name; the program still verifies that no two names clean up to the same filename. Files left over from a previous run follow `existing_files`: `"overwrite"` (with a notice) or `"skip"`.

## Tests

```powershell
py -m unittest -v
```

The tests build their own small templates and workbooks, so they need neither your real template nor any student data.

## Notes

- PDF export is intentionally sequential, using one Word instance for the whole batch. Running multiple Microsoft Word automation instances in parallel can cause file locks and orphaned Word processes.
- The five signature fonts are common Windows fonts: Brush Script MT, Freestyle Script, Lucida Handwriting, Mistral, and Viner Hand ITC. Word will substitute a font if one is missing. A PDF made by LibreOffice can look slightly different from one made by Word.
- Input workbooks, templates, and generated files are excluded by `.gitignore` so student data is not committed accidentally.
