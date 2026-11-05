# Changelog

All notable changes to this project are listed here, newest first.
The version number is also printed by `py .\generate_top_sheets.py --version`.

## 2.0.0 - 2026-11-05

### Added
- `config.json`: the exam date, full marks, which workbook column feeds which template label, fixed values, fonts, signature settings and the duplicate, existing-file and unreachable-marks rules are now settings instead of code.
- A settings screen at the start of every interactive run for the values that change per exam (exam date, full marks, fixed fields such as paper code and subject teacher, font mode). Saved answers go to the git-ignored `config.local.json`. `--no-prompt` skips it.
- The marks table is read from the template: any number of question rows, any column order, with the sections ("answer any N of ...") asked for once per template and saved in `section_rules.json`.
- Roll number as the primary key: identical duplicate rows are merged and conflicting ones stop the run with every row listed (`duplicates` setting).
- `font_mode`: `template` copies the template's own font; `custom` uses a font you name.
- `--check` (validate and report, write nothing), `run_report.csv`, LibreOffice as a PDF engine (`--pdf-engine`), `--exam-date`, `--full-marks`, `--existing`, `--duplicates`.
- Warnings for mobile numbers of the wrong length and for a template whose "Full Marks" text disagrees with the configured full marks.
- `examples/` with a sample workbook and template, `run.bat` for double-click use on Windows, and automated tests (`python -m unittest`).

### Changed
- Question-wise marks: in every section exactly N questions are answered, each with at least 1 mark and at most its allotted marks, and the marks add up exactly to the student's total.
- The exam date is no longer built into the code; the shipped `config.json` leaves it empty so the first run asks.
- PDF export with Word reuses one Word instance for the whole batch instead of starting one per file.
- Problems in the workbook, config or template are all reported together, before anything is generated.

### Fixed
- Values filled into the template lost the label's bold and font; they now keep it.
- The signature line is found even when it sits inside a table.
- One failing student no longer aborts the whole batch.
- A wrong file type (an old `.doc`, `.xls` or a `~$` temporary file) gives a readable message instead of a traceback.

## 1.0.0 - Initial release
- Generated a DOCX and PDF top sheet for every student in a workbook from a Word template, with the labels, question rows and mark rules fixed in the code.
