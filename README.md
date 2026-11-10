# CA1 Top Sheets Automation

[![tests](https://github.com/Subhrajit-Mandal/Top-Sheet-Filing-V2/actions/workflows/tests.yml/badge.svg)](https://github.com/Subhrajit-Mandal/Top-Sheet-Filing-V2/actions/workflows/tests.yml)

Generate one completed Word top sheet and PDF per student from any Excel workbook and compatible DOCX template.

The workbook and template are supplied at runtime, so their filenames and locations are entirely up to the user. Original input files are never changed.

Nothing about the sheet's layout is hard-coded. What gets filled, which workbook column feeds which label, the exam date, the full marks, the fonts and the duplicate rules all live in [`config.json`](config.json). The values that change from exam to exam are asked for when you start a run, so you rarely need to edit a config file by hand. The question rows and their marks are read from the template itself, so a top sheet with more rows or fewer columns works without touching the code.

## Features

- Works in two ways: **generate** a new sheet for every student from a template, or **complete the sheets students filled in themselves** (they keep their own details and signature; the program adds the marks).
- Reads student details and total marks from an XLSX workbook. **The roll number is the primary key.**
- Fills the labels you declare in `config.json` in a DOCX top-sheet template, keeping the template's own font and formatting (or a custom font, your choice).
- Reads the question rows and their `Marks Allotted` from the template's marks table, groups them into sections ("answer any 5 of 1.a to 1.g"), and spreads each student's total across the answered questions so that it adds up exactly.
- Adds a typed student signature in one of five handwriting-style fonts, with the examination date beneath it.
- Asks for the per-exam values (exam date, full marks, paper code, subject teacher, font) on a short settings screen at the start of each run, and can remember your answers.
- Optionally stamps the official seal, kept as a separate image file, just below the last line of every sheet.
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

On Windows you can instead double-click `run.bat`. It finds Python, installs the required packages the first time, starts the program and keeps the window open at the end so you can read the result.

Two standard pickers will open in sequence:

1. Choose the student Excel workbook (`.xlsx`).
2. The program asks what you want to do: **1** generate new sheets from a template, or **2** complete sheets that students filled in. Choose the Word template (`.docx`) for 1, or the folder of student sheets for 2.

Next comes a short settings screen (see [Settings screen](#settings-screen)). Press Enter to accept the values shown, or type a number to change one.

The first time a template is used, the program shows the question sections it found and asks how many questions are to be answered in each (see [Sections](#sections-answer-any-n)). The answers are saved, so later runs with the same template do not ask again.

The generated files are written to `output\docx` and `output\pdf`. Input filenames do not need to match any predefined names.

### Try it with the examples

The `examples` folder has a sample workbook (fake students) and the template it was made for:

```powershell
py .\generate_top_sheets.py --workbook ".\examples\Template_student_data.xlsx" --template ".\examples\ME_CA1_2026.docx"
```

Enter an exam date on the settings screen, accept the detected sections, and answer `5` and `4` for the number of questions in each. The sample mobile numbers have 9 digits, so you will see one harmless warning about that. Everything is written to `output`.

## Completing sheets that students filled in

Instead of generating a sheet for every student, you can ask students to fill in the Word sheet themselves (name, roll number, mobile number and their own signature) and let the program add the marks:

```powershell
py .\generate_top_sheets.py --workbook ".\students.xlsx" --submissions ".\submissions" --config ".\presets\ca2.json"
```

For every `.docx` in the folder the program:

1. reads the roll number from the sheet (not from the file name) and matches it to the workbook's roll number, ignoring capital letters and stray spaces, so `LME 2026410` finds `LME2026410`;
2. writes the student's total into the marks field and the question-wise marks into the Marks Awarded column, exactly as in the other mode;
3. fills a field the student left blank (for example a missing mobile number) from the workbook, and fills the fixed values (such as the exam date) only where the sheet has nothing;
4. saves it as `Rollnumber_Name.docx` (the name comes from the workbook), plus the PDF.

**What the program never does:** it never overwrites what a student wrote, never adds or changes a signature, and never touches the pictures in a sheet. If a student's name or mobile number differs from the workbook, the sheet's text is kept and the difference is listed so you can look at it.

**Problems never stop the run.** A sheet that cannot be used is skipped, listed by name with the reason, and everything else is completed:

| Reason | Example message |
| --- | --- |
| Not a readable Word file | `could not be opened as a Word document (BadZipFile ...)` |
| A picture is missing inside the file | `could not be opened as a Word document (KeyError: "There is no item named 'word/media/image1.png' in the archive")` |
| Not the expected sheet | `does not match the expected sheet: The template has no text '...'` |
| Empty roll number | `the 'Roll Number:' field is empty` |
| Roll number not in the workbook | `roll number '...' is not in the workbook` |
| Same roll number in two files | `roll ... was submitted more than once (...); none of these files was used` |
| Total outside what the section rules allow | `... marks cannot be split ...` |

The summary at the end also lists the students in the workbook who have **no** submitted sheet, the details that differ from the workbook, and the blank fields that were filled. All of it is in `run_report.csv` too.

**Damaged pictures are recovered.** Some apps save a picture inside the Word file with no checksum (a checksum of zero). The picture is fine, but Python refuses to open the file (`Bad CRC-32 for file 'word/media/image1.jpeg'`). The program rebuilds a clean copy in memory, reading each such picture exactly as it is, completes the sheet normally, and lists it under **Sheets completed after recovering a picture** so you can look at it. JPEGs are checked for a proper start and end, PNGs against their own built-in checksums; the note says whether the picture "looks complete" or "may be damaged". Only pictures are recovered. If a text part of the file has a bad checksum, the file is skipped, so a student's typed details are never silently corrupted.

**A file that cannot be used is not "no submission".** The report marks a student whose file could not be used as `FILE RECEIVED, NOT USED`, with the file name and the reason. For this label only, the roll number is taken from the start of the file name; it is never used to match or fill anything. Students listed under "no sheet" really sent nothing.

**Sheets for students who sent nothing.** Give the standard blank template together with the submissions folder:

```powershell
py .\generate_top_sheets.py --workbook ".\students.xlsx" --submissions ".\submissions" --template ".\template.docx"
```

(An interactive run asks "For students who sent no sheet, make one from a template? [y/N]" and opens a picker.) Every student who sent **no file at all** then gets a sheet made the other way: filled in from the workbook, with the question-wise marks and the **typed handwriting-style signature**, and the seal if you chose one. `run_report.csv` marks each of these rows `GENERATED (typed signature)` with the note "no sheet was submitted", and the summary lists them separately from the sheets students submitted.

- Students whose file was received but could not be used are **not** replaced by a generated sheet, so their real signature is never swapped for a typed one.
- A total too low to split under the section rules is not generated; the student is listed with the reason.
- Use a **blank** template. Anything already in it, such as a picture or a signature, is copied onto every generated sheet.
- If the template does not fit the config (a label is missing, for example), the submitted sheets are still completed, and the students who sent nothing are listed instead, with a warning.

**Where the values are.** Sheets differ in how they lay out a field. In the CA2 sheet each label has its own cell and the value is in the cell to the right, so each field in `template_fields` says `"where": "next-cell"`. Where the value follows the label in the same paragraph (the default, `"after-label"`), leave it out. [`presets/ca2.json`](presets/ca2.json) is ready for the CA2 sheet. If no roll number can be read from any sheet, the program stops with a message pointing at this setting.

**Section rules.** The sections ("answer any N") are asked once per sheet layout and remembered, as before. If the folder holds sheets with different layouts, each layout is asked about separately.

**Privacy and safety.** The submissions folder holds names, phone numbers and signatures, and it is ignored by git. Only `.docx` files are read; `.docm` and other formats are refused, and nothing inside a file is ever run. Word opens each file to export the PDF; for sheets you have not checked, `--pdf-engine libreoffice` avoids opening them in Word.

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

## Official seal (optional)

The college seal is not part of the template. Keep it as a separate image file and let the program stamp it on every sheet it creates, in both modes:

- Just before the sheets are made, an interactive run asks **Stamp the official seal on all N sheet(s)? [y/N]**. Answer `y` and a file picker opens for the seal image. Any other answer, or closing the picker, creates the sheets without a seal.
- For scripted runs give the image with `--seal PATH`. With `--no-prompt` and no `--seal` nothing is stamped and nothing is asked.
- The seal goes at the **very end of the document, towards the bottom right, a small distance below the last line.** The program adds a one-point-high paragraph after the last line and the seal floats from it, so the layout of the pages does not change and the seal follows wherever the document actually ends.
- Use a PNG with a transparent background. A JPG works, but it cannot be transparent, so a white box may show around the seal (the program warns about this). A file over 1.5 MB is also warned about, because a copy goes into every sheet.
- A sheet that already contains a picture or shape whose name or description says "seal" is **not** stamped again and is listed in the summary, so a sheet never gets two seals.
- The summary ends with `Seal: stamped on N of M sheet(s) (file name)`, and `run_report.csv` has a `Seal` column.
- The `seal` setting in `config.json` sets the size and position: `width_cm` (3.3), `gap_cm` below the last line (0.3), `align` (`right`, `left` or `center`) and `offset_cm` inward from that side. **The image file itself is never stored in a config file;** it is chosen for each run.

**Keep the seal file safe.** It must not go into the repository: `seal/` and `seals/` are ignored by git, so keep the file in one of those folders (or outside the project). Only the person entitled to apply the seal should hold it and run the stamping.

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
| `seal` | Size and position of the optional seal: `width_cm`, `gap_cm`, `align`, `offset_cm`. The image is never stored here |
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

- Add `"where": "next-cell"` when the value belongs in the table cell to the right of the label (see [above](#completing-sheets-that-students-filled-in)). The default, `"after-label"`, puts it after the label in the same paragraph.
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

`Marks Obtained` must be a whole number from 0 to `full_marks`, or `Ab`. [Example workbook](examples/Template_student_data.xlsx)

**Roll number as the primary key.** Roll numbers are normalised first (`34900725001`, `34900725001.0` and ` 34900725001 ` are the same student). If a roll number appears more than once:

- rows with identical details are merged into one sheet;
- rows with different details are a conflict. With `"duplicates": "stop"` (the default) the run stops before generating anything and lists every row involved. `"keep-first"` and `"keep-last"` pick one row and warn.

Rows with a blank name or roll number are skipped with a note. Blank values in a column that a label is linked to are reported, and that label is left as the template has it.

### Word template

Every label in `template_fields` must appear in the template exactly once, plus the signature label (`Signature of the student with date` by default). The typed signature and date are inserted above the signature label.

The question-wise marks need a table whose header row contains `Q. No.`, `Marks Allotted` and `Marks Awarded` (set under `marks_table`). Column order and any other columns do not matter, and any number of question rows is fine. Rows whose `Marks Allotted` cell is blank can never receive marks. [Example template](examples/ME_CA1_2026.docx)

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
| `--submissions FOLDER` | Complete the sheets students filled in, instead of generating new ones for everyone |
| `--template FILE` | Generate new sheets from this template. Together with `--submissions` it is used only for students who sent no sheet |
| `--seal PATH` | Image of the official seal to stamp on every sheet. Without it, an interactive run asks whether to stamp one |
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

## Troubleshooting

Problems with the workbook, config or template are all listed together, before any file is generated.

| What you see | What it means and what to do |
| --- | --- |
| `The workbook has no column with the header(s) declared in the config` | A header differs from the config. Headers must match exactly, including spaces and capitals. Fix the workbook header, or the `columns` / `template_fields` entry in `config.json`. The message lists the headers it did find. |
| `No roll number could be read from any of the N sheet(s)` | In submitted sheets the value is usually in the cell next to the label. Set `"where": "next-cell"` on the fields in the config, or use `presets/ca2.json`. |
| `Bad CRC-32 for file 'word/media/...'` | A picture inside the student's file was saved with a wrong checksum. The program recovers it automatically and lists the sheet under "Sheets completed after recovering a picture". Nothing to do except look at the sheet. |
| `FILE RECEIVED, NOT USED` (in the report) | The student sent a file, but it could not be used (see the reason beside it). Ask for a new copy. It is not the same as `NO SUBMISSION`. |
| `could not be opened as a Word document (KeyError: "There is no item named 'word/media/...' in the archive")` | The student's file refers to a picture that is not inside it, usually after another program re-saved or re-zipped the file. The file is skipped. Ask the student to open it in Word, save it again as `.docx`, and resubmit. |
| `'...' could not be read as a picture` | The seal file is not a valid image. Use a PNG or JPG. |
| A white box shows around the seal | The seal is a JPG or a PNG without transparency. Use a PNG with a transparent background. |
| The seal overlaps something or sits too low | Adjust `seal` in the config (`gap_cm`, `offset_cm`, `width_cm`). It always follows the end of the document, so check a sheet in Word. |
| `The template has no text '...'` | A label in `template_fields` is not in the template as typed. Copy it from the template, or correct it in the config. |
| `The template contains '...' in N places` | A label must appear exactly once. Make the text unique in the template. |
| `The template has no table whose header row contains ...` | The marks table headers differ from `marks_table` in the config. |
| `Roll ... appears in rows ... with different details` | The same roll number has different names, mobile numbers or marks. Fix the workbook, or set `duplicates` to `keep-first` or `keep-last`. |
| `... marks cannot be split ... (possible totals: a to b)` | That student's total is outside what the section rules allow. Correct the total, change the answers per section with `--ask-sections`, or set `unreachable_marks` to `skip`. |
| `The section rules for this template are not saved yet and no one is there to answer` | Run once in a normal terminal to answer the section questions, or pass `--answer 5,4`. |
| `No PDF engine is available` | Install Microsoft Word with `pywin32`, or LibreOffice, or use `--keep-docx-only`. Use `--soffice PATH` if LibreOffice is installed but not found. |
| `Could not open ... as a Word (.docx) document` or `... as an Excel workbook` | The file is not a `.docx` / `.xlsx`, for example an old `.doc` / `.xls`, or Office's temporary file that starts with `~$`. |
| `FAILED roll ...: PermissionError` | An output file is open in Word, Excel or a PDF viewer. Close it and run again. |
| `'Mobile Number' has 9 digits instead of 10` | Excel drops a leading zero from numbers. Format the column as Text. This is only a warning. |
| The signature looks like an ordinary font | The five handwriting fonts are not installed on that computer, so Word substituted another. |
| The settings screen is in your way | Use `--no-prompt`. |
| You want to start over | Delete `config.local.json` and `section_rules.json`. |

## Known limitations

- Marks are whole numbers. Half marks are not supported, and `17.5` is reported as an error.
- In every section exactly N questions are answered, each with at least 1 mark. A student's total therefore cannot be below the sum of the N values (9 with answers of 5 and 4), so a total of 0 cannot be split. See `unreachable_marks`.
- A section is a run of consecutive question rows.
- The template needs one marks table, found by its header texts. Question labels must be unique.
- Labels are found in the document body and in table cells, not in headers, footers or text boxes. A label must be plain text, not part of a hyperlink or a content control. The value replaces everything after the label in its paragraph.
- An output file that is open in another program cannot be overwritten.
- The Word PDF export cannot be exercised by the automated tests, which run on Linux and Windows without Word. PDFs made by LibreOffice can look slightly different from Word's.
- The five signature fonts are Windows fonts.
- The seal is stamped at the end of the document, not at a fixed place on a fixed page. If the last page is almost full, the seal can run past its bottom edge, so look at a finished sheet in Word. Only one seal image is used per run, and "already has a seal" is judged from the name or description of pictures in the sheet.
- Pictures with a wrong checksum are recovered; any other damaged part makes the file unusable. A file that is not a Word file at all is skipped.
- Completing submitted sheets reads the roll number from the one field you configure. It does not check that a signature is present, and text boxes and shapes are not read. A student's text is never corrected, only compared.

## Intended use

This tool prepares CA1 / CA2 top sheets from a workbook of marks. Please note what it generates and what it does not:

- **The question-wise marks are generated.** They are spread randomly within each question's allotted marks so that they add up to the total in the workbook. They are not an examiner's own question-by-question assessment.
- **In generate mode the typed signature is generated.** It is the student's name in a handwriting-style font, not a handwritten signature. Set `"signature": {"enabled": false}` in the config to leave it out.
- **In submissions mode the program adds no signature at all.** The student's own signature, typed details and pictures stay exactly as the student made them, which is the better choice for anything official.

- **Sheets made for students who sent nothing carry a generated signature.** It is the student's name in a handwriting-style font, not something the student signed. The run report marks these sheets `GENERATED (typed signature)` so they can always be told apart from sheets the students submitted.
- **The seal is an authentication mark.** The program only stamps an image file that you give it, for that run, and never stores it. The file should be held by the person entitled to apply the seal, and sheets should be stamped only with that person's authority.

Use the tool only where your institution permits sheets to be prepared this way, and do not present generated content as an examiner's or a student's original record. The software is provided as is, without warranty (see [LICENSE](LICENSE)).

## Version

This is version 2.3.0. See [CHANGELOG.md](CHANGELOG.md) for what changed, or run `py .\generate_top_sheets.py --version`.

## Notes

- PDF export is intentionally sequential, using one Word instance for the whole batch. Running multiple Microsoft Word automation instances in parallel can cause file locks and orphaned Word processes.
- The five signature fonts are common Windows fonts: Brush Script MT, Freestyle Script, Lucida Handwriting, Mistral, and Viner Hand ITC. Word will substitute a font if one is missing. A PDF made by LibreOffice can look slightly different from one made by Word.
- Input workbooks, templates, and generated files are excluded by `.gitignore` so student data is not committed accidentally.
