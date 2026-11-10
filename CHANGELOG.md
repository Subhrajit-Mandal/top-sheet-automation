# Changelog

All notable changes to this project are listed here, newest first.
The version number is also printed by `py .\generate_top_sheets.py --version`.

## 2.3.0 - 2026-11-09

### Added
- Pictures saved with a wrong checksum (for example a checksum of zero, which some apps write) no longer make a student's file unusable. The picture is read exactly as it is, the sheet is completed, and the sheet is listed for a look. JPEGs and PNGs are checked, and the note says whether the picture "looks complete" or "may be damaged". Other damaged parts are still refused.
- Students who sent no file at all can get a sheet made from the standard template, with workbook data, the question-wise marks and the typed signature (and the seal, if chosen): give `--template` together with `--submissions`, or answer the new question in an interactive run. The report marks them `GENERATED (typed signature)` with the note "no sheet was submitted".
- A tip when roll numbers on submitted sheets are not in the workbook: add them and run again with `--existing skip`.

### Changed
- A student whose file was received but could not be used is reported as `FILE RECEIVED, NOT USED` (with the file name and the reason) instead of `NO SUBMISSION`, and is never replaced by a generated sheet. The roll number taken from the file name is used for this label only.
- `--template` and `--submissions` can now be used together.

## 2.2.0 - 2026-11-07

### Added
- Optional official seal, kept as a separate image file and stamped on every sheet the program creates, in both modes. An interactive run asks "Stamp the official seal on all N sheet(s)?" just before the sheets are made and opens a file picker for the image; `--seal PATH` does the same without asking.
- The seal is placed at the very end of the document, towards the bottom right, a small distance below the last line, without changing the page layout. `seal` settings in `config.json` set its width, gap, alignment and offset.
- A sheet that already has a picture named like a seal is not stamped a second time and is listed. The summary reports how many sheets were stamped, and `run_report.csv` has a new `Seal` column.
- `seal/` and `seals/` are ignored by git.

### Changed
- The README's intended-use note now covers the seal.

## 2.1.0 - 2026-11-07

### Added
- Complete sheets that students filled in themselves: `--submissions FOLDER`. Each sheet is matched to the workbook by the roll number written in it, the marks and question-wise marks are added, and the student's own text, signature and pictures are left untouched. Differences from the workbook are listed, blank fields are filled, and students with no sheet are listed.
- A sheet that cannot be used (not a readable Word file, a picture missing inside it, a different layout, an unknown or repeated roll number) is skipped and listed by name; it never stops the run.
- `where` setting for each field in `template_fields`: `"after-label"` (default) or `"next-cell"`, for sheets that keep the value in the cell to the right of the label. Works when generating sheets as well.
- `presets/ca2.json` for the CA2 sheet, and a question at the start of an interactive run: generate new sheets, or complete submitted ones.
- `submissions/` and `*.local.json` are ignored by git.

### Changed
- A new Marks Awarded entry takes the alignment of the Marks Allotted column when the cell has none of its own.
- The README's intended-use note now distinguishes generated signatures from sheets the students signed themselves.

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
