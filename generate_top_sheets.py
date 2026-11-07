"""Create one completed CA1 top sheet (DOCX and PDF) for every student row.

Nothing about the sheet's layout is hard-coded any more. What to fill, which
workbook column feeds which label, the exam date, the full marks, the fonts
and the duplicate / collision rules all live in ``config.json``.

The marks table is read from the template itself: every question row, its
"Marks Allotted" cap and the sections the questions fall into (for example
"answer any 5 of 1.a to 1.g").  For each student the random question-wise
marks are then spread so that they add up exactly to the student's total.

Nothing is ever guessed: a label or column that is declared but cannot be
found stops the run with a message naming it.  Anything that is not declared
is left exactly as it appears in the template.
"""

from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import json
import os
import random
import re
import shutil
import subprocess
import sys
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook


__version__ = "2.2.0"

HERE = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = HERE / "config.json"
RULES_FILENAME = "section_rules.json"
DOCX_FOLDER = "docx"
PDF_FOLDER = "pdf"
REPORT_FILENAME = "run_report.csv"

DEFAULT_CONFIG: dict[str, Any] = {
    # Printed in the "Date of Examination" field and under the signature.
    # Empty by default: the settings screen asks for it at the start of a run.
    "exam_date": "",
    # Maximum marks a student can obtain (checked against the workbook).
    "full_marks": 25,
    # "template": copy the font of the template's own label text.
    # "custom":   use custom_font for everything the program fills in.
    "font_mode": "template",
    "custom_font": {"name": "Calibri", "size_pt": None},
    # Workbook headers that identify a student. Roll number is the primary key.
    "columns": {
        "roll_number": "UNI Roll number",
        "name": "Name",
        "marks_obtained": "Marks Obtained",
    },
    # Values in the marks column that mean "absent" (no marks are awarded).
    "absent_markers": ["Ab"],
    # Every label the program fills. Use "column" (exact workbook header) or
    # "value" (fixed text; {exam_date} and {full_marks} are expanded).
    # A fixed value of "" leaves that label exactly as the template has it.
    # Optional "where": "after-label" (default: the value follows the label in the same
    # paragraph) or "next-cell" (the value is in the table cell to the right of the label).
    # (The settings screen saves its answers for fixed values in "fixed_values".)
    "template_fields": [
        {"label": "Name of the Student:", "column": "Name"},
        {"label": "Roll Number:", "column": "UNI Roll number"},
        {"label": "Mobile Number:", "column": "Mobile Number"},
        {"label": "Marks obtained:", "column": "Marks Obtained"},
        {"label": "Date of Examination:", "value": "{exam_date}"},
        {"label": "Paper (course) Code:", "value": ""},
        {"label": "Subject Teacher:", "value": ""},
        {"label": "UPID:", "value": ""},
    ],
    # Overrides for fixed-value labels, as {"label": "text"}. Normally written
    # by the settings screen into config.local.json, not edited by hand.
    "fixed_values": {},
    # Header texts of the marks table inside the template.
    # Set to null to switch question-wise marks off.
    "marks_table": {
        "question_header": "Q. No.",
        "allotted_header": "Marks Allotted",
        "awarded_header": "Marks Awarded",
    },
    # Rows of the marks table whose question label is one of these are ignored.
    "ignore_question_labels": ["total", "grand total"],
    "signature": {
        "enabled": True,
        "label": "Signature of the student with date",
        "fonts": [
            "Brush Script MT",
            "Freestyle Script",
            "Lucida Handwriting",
            "Mistral",
            "Viner Hand ITC",
        ],
        "name_size_pt": 16,
        "date_size_pt": 9,
    },
    # Same roll number twice with different details:
    # "stop", "keep-first" or "keep-last". Identical rows are always merged.
    "duplicates": "stop",
    # Files already in the output folder: "overwrite" or "skip".
    "existing_files": "overwrite",
    # A total that cannot be split under the section rules: "stop" or "skip".
    "unreachable_marks": "stop",
    # Where the optional official seal goes (the image itself is never stored in a config:
    # it is chosen at the start of each run or given with --seal). It is placed just below the
    # last line of the document. "align": "right", "left" or "center"; "offset_cm" moves it
    # inward from the chosen side.
    "seal": {"width_cm": 3.3, "gap_cm": 0.3, "align": "right", "offset_cm": 0.0},
    # Optional sanity checks that only ever warn. null switches them off.
    "full_marks_label": "Full Marks:",
    "mobile_check": {"column": "Mobile Number", "digits": 10},
}


class UserError(Exception):
    """A problem the user can fix: bad input, config or template."""


def info(message: str = "") -> None:
    print(message, flush=True)


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

def merge_config(base: dict, override: dict, where: str = "config.json", prefix: str = "") -> dict:
    result = copy.deepcopy(base)
    for key, value in override.items():
        if key.startswith("_"):  # free-form notes are allowed
            continue
        if key not in base:
            raise UserError(
                f"Unknown setting '{prefix}{key}' in {where}. "
                f"Valid settings here: {', '.join(base)}."
            )
        if key == "fixed_values" and isinstance(value, dict):
            result[key] = {**result[key], **value}
        elif isinstance(base[key], dict) and isinstance(value, dict):
            result[key] = merge_config(base[key], value, where, f"{prefix}{key}.")
        else:
            result[key] = value
    return result


def validate_config(config: dict) -> None:
    def one_of(key: str, allowed: tuple[str, ...]) -> None:
        if config[key] not in allowed:
            raise UserError(
                f"config.{key} must be one of {', '.join(allowed)} "
                f"(found {config[key]!r})."
            )

    one_of("font_mode", ("template", "custom"))
    one_of("duplicates", ("stop", "keep-first", "keep-last"))
    one_of("existing_files", ("overwrite", "skip"))
    one_of("unreachable_marks", ("stop", "skip"))

    if not isinstance(config["exam_date"], str):
        raise UserError('config.exam_date must be text, for example "02-09-2026".')
    if not isinstance(config["full_marks"], int) or config["full_marks"] < 1:
        raise UserError("config.full_marks must be a whole number of at least 1.")
    if config["font_mode"] == "custom" and not str(config["custom_font"]["name"] or "").strip():
        raise UserError("config.font_mode is 'custom' but custom_font.name is empty.")
    size = config["custom_font"].get("size_pt")
    if size is not None and (not isinstance(size, (int, float)) or size <= 0):
        raise UserError("config.custom_font.size_pt must be a positive number or null.")
    if not config["signature"]["fonts"]:
        raise UserError("config.signature.fonts must list at least one font.")
    for key, value in config["columns"].items():
        if not str(value or "").strip():
            raise UserError(f"config.columns.{key} must not be empty.")

    seen_labels: set[str] = set()
    for index, entry in enumerate(config["template_fields"], start=1):
        where = f"config.template_fields[{index}]"
        if not isinstance(entry, dict) or not str(entry.get("label", "")).strip():
            raise UserError(f"{where} needs a non-empty 'label'.")
        extra = set(entry) - {"label", "column", "value", "where"}
        if extra:
            raise UserError(f"{where} has unknown key(s): {', '.join(sorted(extra))}.")
        if ("column" in entry) == ("value" in entry):
            raise UserError(f"{where} ('{entry['label']}') needs exactly one of 'column' or 'value'.")
        if entry.get("where", "after-label") not in ("after-label", "next-cell"):
            raise UserError(f"{where} ('{entry['label']}'): 'where' must be 'after-label' or 'next-cell'.")
        if entry["label"] in seen_labels:
            raise UserError(f"{where}: label '{entry['label']}' is declared twice.")
        seen_labels.add(entry["label"])

    seal = config["seal"]
    for key in ("width_cm", "gap_cm", "offset_cm"):
        if not isinstance(seal[key], (int, float)) or isinstance(seal[key], bool) or seal[key] < 0:
            raise UserError(f"config.seal.{key} must be a number of at least 0.")
    if seal["width_cm"] <= 0 or seal["width_cm"] > 15:
        raise UserError("config.seal.width_cm must be more than 0 and at most 15.")
    if seal["align"] not in ("left", "center", "right"):
        raise UserError("config.seal.align must be 'left', 'center' or 'right'.")

    fixed_labels = [e["label"] for e in config["template_fields"] if "value" in e]
    for label, value in config["fixed_values"].items():
        if label not in fixed_labels:
            raise UserError(
                f"config.fixed_values has '{label}', which is not a fixed-value entry in "
                f"template_fields. Fixed-value labels: {', '.join(fixed_labels) or '(none)'}."
            )
        if not isinstance(value, str):
            raise UserError(f"config.fixed_values['{label}'] must be text.")

    table = config["marks_table"]
    if table is not None:
        for key in ("question_header", "allotted_header", "awarded_header"):
            if not str(table.get(key, "")).strip():
                raise UserError(f"config.marks_table.{key} must not be empty.")


def local_config_path(path: Path | None) -> Path:
    """config.json -> config.local.json beside it (git-ignored, one per user)."""
    base = path or DEFAULT_CONFIG_PATH
    return base.with_name(f"{base.stem}.local{base.suffix}")


def read_json_object(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise UserError(f"{path.name} is not valid JSON: {error}") from error
    if not isinstance(data, dict):
        raise UserError(f"{path.name} must contain a JSON object.")
    return data


def read_local_config(path: Path | None) -> dict:
    local = local_config_path(path)
    return read_json_object(local) if local.exists() else {}


def apply_fixed_values(config: dict) -> None:
    for entry in config["template_fields"]:
        if "value" in entry and entry["label"] in config["fixed_values"]:
            entry["value"] = config["fixed_values"][entry["label"]]


def load_config(path: Path | None = None) -> dict:
    """config.json (shipped defaults) with config.local.json (your saved answers) on top."""
    config_path = path or DEFAULT_CONFIG_PATH
    if config_path.exists():
        config = merge_config(DEFAULT_CONFIG, read_json_object(config_path), config_path.name)
    elif path is not None:
        raise UserError(f"Config file not found: {config_path}")
    else:
        config = copy.deepcopy(DEFAULT_CONFIG)
    local = read_local_config(path)
    if local:
        config = merge_config(config, local, local_config_path(path).name)
    validate_config(config)
    apply_fixed_values(config)
    return config


def save_local_config(path: Path | None, changes: dict) -> Path:
    """Merge `changes` into config.local.json, leaving everything else in it alone."""
    target = local_config_path(path)
    data = read_json_object(target) if target.exists() else {}

    def merge(into: dict, new: dict) -> None:
        for key, value in new.items():
            if isinstance(value, dict) and isinstance(into.get(key), dict):
                merge(into[key], value)
            else:
                into[key] = copy.deepcopy(value)

    merge(data, changes)
    target.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return target


def expand(text: str, config: dict) -> str:
    return text.replace("{exam_date}", str(config["exam_date"])).replace(
        "{full_marks}", str(config["full_marks"])
    )


def font_settings(config: dict) -> dict:
    custom = config["custom_font"]
    return {
        "mode": config["font_mode"],
        "name": custom["name"],
        "size_pt": custom.get("size_pt"),
    }


# --------------------------------------------------------------------------
# Small text helpers
# --------------------------------------------------------------------------

def as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalise_roll(value: str) -> str:
    value = value.strip()
    if re.fullmatch(r"\d+\.0+", value):
        value = value.split(".")[0]
    return value


def clean_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", " ", value).strip(". ")
    return value or "student"


def squash(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip().casefold()


# --------------------------------------------------------------------------
# Workbook
# --------------------------------------------------------------------------

@dataclass
class Student:
    row: int
    roll: str
    name: str
    values: dict[str, str]
    total: int | None = None  # None means absent


def needed_headers(config: dict) -> list[str]:
    headers = list(config["columns"].values())
    headers += [f["column"] for f in config["template_fields"] if "column" in f]
    mobile = config.get("mobile_check")
    if mobile:
        headers.append(mobile["column"])
    return list(dict.fromkeys(headers))


def read_students(workbook_path: Path, sheet_name: str | None, config: dict) -> tuple[list[Student], list[str]]:
    try:
        workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    except Exception as error:
        raise UserError(
            f"Could not open '{workbook_path.name}' as an Excel workbook ({type(error).__name__}). "
            "Check that it is an .xlsx file (not an old .xls or Excel's temporary '~$' file)."
        ) from error
    notes: list[str] = []
    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise UserError(
                    f"Sheet '{sheet_name}' not found. Sheets in the workbook: "
                    f"{', '.join(workbook.sheetnames)}."
                )
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook.active
        try:
            worksheet.reset_dimensions()
        except Exception:
            pass

        rows = worksheet.iter_rows(values_only=True)
        try:
            header_row = next(rows)
        except StopIteration:
            raise UserError("The worksheet is empty.") from None

        positions: dict[str, int] = {}
        repeated: set[str] = set()
        for index, cell in enumerate(header_row):
            header = as_text(cell)
            if not header:
                continue
            if header in positions:
                repeated.add(header)
            else:
                positions[header] = index

        needed = needed_headers(config)
        missing = [h for h in needed if h not in positions]
        if missing:
            raise UserError(
                "The workbook has no column with the header(s) declared in the config: "
                + ", ".join(f"'{h}'" for h in missing)
                + ".\nHeaders found in the workbook: "
                + ", ".join(f"'{h}'" for h in positions)
                + ".\nFix the header in the workbook or the matching entry in config.json "
                "(headers must match exactly; nothing is matched by guesswork)."
            )
        ambiguous = [h for h in needed if h in repeated]
        if ambiguous:
            raise UserError(
                "These headers appear more than once in the workbook, so the right column "
                "cannot be known: " + ", ".join(f"'{h}'" for h in ambiguous) + "."
            )

        roll_header = config["columns"]["roll_number"]
        name_header = config["columns"]["name"]
        students: list[Student] = []
        for row_number, row in enumerate(rows, start=2):
            values = {
                h: as_text(row[positions[h]]) if positions[h] < len(row) else ""
                for h in needed
            }
            if not any(values.values()):
                continue
            values[roll_header] = normalise_roll(values[roll_header])
            roll, name = values[roll_header], values[name_header]
            if not roll or not name:
                notes.append(f"Row {row_number}: skipped, the name or roll number is blank.")
                continue
            students.append(Student(row=row_number, roll=roll, name=name, values=values))
        return students, notes
    finally:
        workbook.close()


def resolve_duplicates(students: list[Student], policy: str) -> tuple[list[Student], list[str], list[str]]:
    """Roll number is the primary key. Returns (students, warnings, conflicts)."""
    groups: dict[str, list[Student]] = {}
    for student in students:
        groups.setdefault(student.roll.casefold(), []).append(student)

    kept: list[Student] = []
    warnings: list[str] = []
    conflicts: list[str] = []
    for group in groups.values():
        if len(group) == 1:
            kept.append(group[0])
            continue
        rows = ", ".join(str(s.row) for s in group)
        signatures = {tuple((h, squash(v)) for h, v in sorted(s.values.items())) for s in group}
        if len(signatures) == 1:
            warnings.append(
                f"Roll {group[0].roll}: identical data in rows {rows}; using row {group[0].row} once."
            )
            kept.append(group[0])
            continue
        if policy == "stop":
            lines = [f"Roll {group[0].roll} appears in rows {rows} with different details:"]
            for s in group:
                details = ", ".join(f"{h}={v!r}" for h, v in s.values.items())
                lines.append(f"    row {s.row}: {details}")
            conflicts.append("\n".join(lines))
            continue
        chosen = group[0] if policy == "keep-first" else group[-1]
        warnings.append(
            f"Roll {chosen.roll}: conflicting rows {rows}; kept row {chosen.row} ({policy})."
        )
        kept.append(chosen)
    kept.sort(key=lambda s: s.row)
    return kept, warnings, conflicts


def parse_total_mark(value: str, config: dict) -> int | None:
    text = value.strip()
    if not text:
        raise ValueError(f"'{config['columns']['marks_obtained']}' is blank.")
    if text.casefold() in {m.casefold() for m in config["absent_markers"]}:
        return None
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    if not re.fullmatch(r"\d+", text):
        markers = ", ".join(config["absent_markers"])
        raise ValueError(
            f"marks must be a whole number or one of [{markers}] (found {value!r})."
        )
    total = int(text)
    if total > config["full_marks"]:
        raise ValueError(f"{total} is above the full marks ({config['full_marks']}).")
    return total


def validate_students(students: list[Student], config: dict) -> tuple[list[str], list[str]]:
    """Parse totals and run the data checks. Returns (errors, warnings)."""
    errors: list[str] = []
    warnings: list[str] = []
    marks_header = config["columns"]["marks_obtained"]
    for student in students:
        try:
            student.total = parse_total_mark(student.values[marks_header], config)
        except ValueError as error:
            errors.append(f"Row {student.row} (roll {student.roll}, {student.name}): {error}")

    mobile = config.get("mobile_check")
    if mobile:
        by_length: dict[int, list[str]] = {}
        for student in students:
            digits = re.sub(r"\D", "", student.values[mobile["column"]])
            if digits and len(digits) != mobile["digits"]:
                by_length.setdefault(len(digits), []).append(str(student.row))
        for length, rows in by_length.items():
            warnings.append(
                f"'{mobile['column']}' has {length} digits instead of {mobile['digits']} in "
                f"{len(rows)} row(s): {', '.join(rows)}."
            )

    skip = set(config["columns"].values())
    for entry in config["template_fields"]:
        column = entry.get("column")
        if not column or column in skip:
            continue
        blank = [str(s.row) for s in students if not s.values[column]]
        if blank:
            warnings.append(
                f"Column '{column}' is blank in row(s) {', '.join(blank)}; "
                f"'{entry['label']}' is left as the template has it on those sheets."
            )
    return errors, warnings


# --------------------------------------------------------------------------
# Word helpers
# --------------------------------------------------------------------------

def iter_tables(container) -> list[Table]:
    found: list[Table] = []
    for table in container.tables:
        found.append(table)
        for row in table.rows:
            for cell in row.cells:
                found.extend(iter_tables(cell))
    return found


def iter_template_paragraphs(document) -> list[Paragraph]:
    paragraphs = list(document.paragraphs)
    seen_cells: set[int] = set()

    def walk(table: Table) -> None:
        for row in table.rows:
            for cell in row.cells:
                if id(cell._tc) in seen_cells:
                    continue
                seen_cells.add(id(cell._tc))
                paragraphs.extend(cell.paragraphs)
                for nested in cell.tables:
                    walk(nested)

    for table in document.tables:
        walk(table)
    return paragraphs


def paragraph_text(paragraph: Paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def find_label_paragraph(document, label: str, required: bool = True) -> Paragraph | None:
    matches = [p for p in iter_template_paragraphs(document) if label in paragraph_text(p)]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        if required:
            raise UserError(f"The template has no text '{label}'.")
        return None
    raise UserError(
        f"The template contains '{label}' in {len(matches)} places; it must appear exactly once."
    )


def label_cell(paragraph: Paragraph) -> _Cell | None:
    parent = paragraph._parent
    return parent if isinstance(parent, _Cell) else None


def next_cell(cell: _Cell) -> _Cell | None:
    """The table cell to the right of `cell` in the same row (merged cells count once)."""
    tc = cell._tc.getnext()
    while tc is not None and tc.tag != qn("w:tc"):
        tc = tc.getnext()
    return _Cell(tc, cell._parent) if tc is not None else None


def value_cell(paragraph: Paragraph, label: str) -> _Cell:
    cell = label_cell(paragraph)
    if cell is None:
        raise UserError(f"'{label}' is not inside a table cell, so its value cannot be in the next cell.")
    target = next_cell(cell)
    if target is None:
        raise UserError(f"There is no cell to the right of '{label}' to hold its value.")
    return target


def read_field(paragraph: Paragraph, label: str, where: str = "after-label") -> str:
    """What is written for a label: the text after it, or the text of the cell to its right."""
    if where == "next-cell":
        return re.sub(r"\s+", " ", value_cell(paragraph, label).text).strip()
    return paragraph_text(paragraph).split(label, 1)[1].strip()


def set_run_font(run, name: str, size_pt: float | None = None) -> None:
    run.font.name = name
    fonts = run._r.rPr.rFonts
    for attribute in ("eastAsia", "cs"):
        fonts.set(qn(f"w:{attribute}"), name)
    for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{attribute}"), None)
    if size_pt:
        run.font.size = Pt(size_pt)


def copy_run_format(source_run, target_run, font: dict | None) -> None:
    """Give target_run the formatting of source_run (and the custom font if chosen)."""
    source_properties = source_run._r.rPr if source_run is not None else None
    if source_properties is not None:
        existing = target_run._r.rPr
        if existing is not None:
            target_run._r.remove(existing)
        target_run._r.insert(0, copy.deepcopy(source_properties))
    if font and font["mode"] == "custom":
        set_run_font(target_run, font["name"], font.get("size_pt"))


def fill_field(paragraph: Paragraph, label: str, value: str, font: dict) -> None:
    """Replace everything after `label` in the paragraph with `value`.

    The label and anything before it are untouched; the value takes the
    formatting of the label's last run, so the sheet keeps its own look.
    """
    runs = paragraph.runs
    texts = [run.text for run in runs]
    start = "".join(texts).find(label)
    end = start + len(label)
    position = 0
    label_index = offset = None
    for index, text in enumerate(texts):
        if position < end <= position + len(text):
            label_index, offset = index, end - position
            break
        position += len(text)
    if label_index is None:
        raise UserError(f"Could not locate the end of '{label}' inside its paragraph.")

    runs[label_index].text = texts[label_index][:offset]
    for later in runs[label_index + 1:]:
        later.text = ""
    separator = "" if label[-1:].isspace() else " "
    new_run = paragraph.add_run(separator + value)
    copy_run_format(runs[label_index], new_run, font)


def fill_cell(cell: _Cell, value: str, label_paragraph: Paragraph, font: dict) -> None:
    """Put `value` in a value cell, looking like what Word would type there."""
    paragraph = cell.paragraphs[0]
    runs = paragraph.runs
    if runs:  # keep the formatting the cell already has
        runs[0].text = value
        for extra in runs[1:]:
            extra.text = ""
        run = runs[0]
    else:
        run = paragraph.add_run(value)
        properties = paragraph._p.pPr
        mark = properties.find(qn("w:rPr")) if properties is not None else None
        if mark is not None:  # the formatting Word would use for typing in this cell
            formatting = copy.deepcopy(mark)
            for tracked in ("w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange"):
                for element in formatting.findall(qn(tracked)):
                    formatting.remove(element)
            run._r.insert(0, formatting)
        else:  # otherwise the label's own look, without its bold
            label_runs = label_paragraph.runs
            copy_run_format(label_runs[-1] if label_runs else None, run, None)
            run.bold = False
    if font["mode"] == "custom":
        set_run_font(run, font["name"], font.get("size_pt"))


def write_field(paragraph: Paragraph, entry: dict, value: str, font: dict) -> None:
    if entry.get("where", "after-label") == "next-cell":
        fill_cell(value_cell(paragraph, entry["label"]), value, paragraph, font)
    else:
        fill_field(paragraph, entry["label"], value, font)


# --------------------------------------------------------------------------
# Marks table
# --------------------------------------------------------------------------

@dataclass
class QuestionRow:
    label: str
    allotted: int | None  # None: the allotted cell is blank
    table_row: int


@dataclass
class MarksTable:
    questions: list[QuestionRow]
    question_col: int
    allotted_col: int
    awarded_col: int
    fingerprint: str


def find_marks_table(document, config: dict) -> tuple[Table, MarksTable]:
    settings = config["marks_table"]
    wanted = {k: settings[k].strip().casefold() for k in
              ("question_header", "allotted_header", "awarded_header")}
    candidates = []
    for table in iter_tables(document):
        for row_index, row in enumerate(table.rows[:5]):
            try:
                texts = [cell.text.strip().casefold() for cell in row.cells]
            except Exception:
                continue
            if all(text in texts for text in wanted.values()):
                candidates.append((table, row_index, texts))
                break
    if not candidates:
        raise UserError(
            "The template has no table whose header row contains "
            f"'{settings['question_header']}', '{settings['allotted_header']}' and "
            f"'{settings['awarded_header']}'. Check config.marks_table, or set it to null "
            "to switch question-wise marks off."
        )
    if len(candidates) > 1:
        raise UserError("More than one table in the template looks like the marks table.")

    table, header_index, texts = candidates[0]
    question_col = texts.index(wanted["question_header"])
    allotted_col = texts.index(wanted["allotted_header"])
    awarded_col = texts.index(wanted["awarded_header"])
    ignored = {text.casefold() for text in config["ignore_question_labels"]}

    questions: list[QuestionRow] = []
    seen: set[str] = set()
    for row_index in range(header_index + 1, len(table.rows)):
        cells = table.rows[row_index].cells
        if len(cells) <= max(question_col, allotted_col, awarded_col):
            continue
        label = cells[question_col].text.strip()
        if not label or label.casefold() in ignored:
            continue
        if label in seen:
            raise UserError(f"Question '{label}' appears twice in the marks table.")
        seen.add(label)
        allotted_text = cells[allotted_col].text.strip()
        if not allotted_text:
            allotted = None
        elif re.fullmatch(r"\d+", allotted_text):
            allotted = int(allotted_text)
        else:
            raise UserError(
                f"'{settings['allotted_header']}' for question '{label}' is "
                f"{allotted_text!r}, which is not a whole number."
            )
        questions.append(QuestionRow(label, allotted, row_index))

    if not questions:
        raise UserError("The marks table has a header row but no question rows.")
    text = "\n".join(f"{q.label}|{q.allotted if q.allotted is not None else ''}" for q in questions)
    fingerprint = hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]
    return table, MarksTable(questions, question_col, allotted_col, awarded_col, fingerprint)


# --------------------------------------------------------------------------
# Sections ("answer any N of these rows")
# --------------------------------------------------------------------------

@dataclass
class Section:
    labels: list[str]
    answer: int


_SUB_PART = re.compile(r"^(?:Q\.?\s*)?(\d+)\s*[.(\-]?\s*([A-Za-z]{1,4})\s*\)?$")
_PLAIN_NUMBER = re.compile(r"^(?:Q\.?\s*)?(\d+)$")


def detect_sections(questions: list[QuestionRow]) -> list[list[int]]:
    """Group question rows: 1.a, 1.b ... together; consecutive plain numbers together."""
    keys: list[tuple] = []
    for index, question in enumerate(questions):
        label = question.label.strip()
        sub = _SUB_PART.match(label)
        if sub:
            keys.append(("sub", sub.group(1)))
        elif _PLAIN_NUMBER.match(label):
            keys.append(("plain",))
        else:
            keys.append(("other", index))
    groups: list[list[int]] = []
    for index, key in enumerate(keys):
        if groups and keys[groups[-1][0]] == key:
            groups[-1].append(index)
        else:
            groups.append([index])
    return groups


def awardable(questions: list[QuestionRow], indexes: list[int]) -> list[int]:
    return [questions[i].allotted for i in indexes if questions[i].allotted]


def describe_group(questions: list[QuestionRow], indexes: list[int]) -> str:
    labels = [questions[i].label for i in indexes]
    span = labels[0] if len(labels) == 1 else f"{labels[0]} to {labels[-1]}"
    caps = awardable(questions, indexes)
    if not caps:
        carry = "none carry marks"
    elif len(set(caps)) == 1:
        carry = f"{len(caps)} carry marks: {caps[0]} each"
    else:
        carry = f"{len(caps)} carry marks: {', '.join(str(c) for c in sorted(set(caps)))}"
    return f"{span}  ({len(indexes)} rows; {carry})"


def parse_ranges(text: str, total: int) -> list[list[int]]:
    groups: list[list[int]] = []
    next_row = 1
    for part in re.split(r"[,\s]+", text.strip()):
        if not part:
            continue
        match = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not match:
            raise ValueError(f"'{part}' is not a row or a range like 1-7.")
        first = int(match.group(1))
        last = int(match.group(2) or first)
        if first != next_row or last < first or last > total:
            raise ValueError(
                f"Ranges must run in order with no gaps or overlaps, covering rows 1-{total} "
                f"(expected the next range to start at {next_row})."
            )
        groups.append(list(range(first - 1, last)))
        next_row = last + 1
    if next_row != total + 1:
        raise ValueError(f"The ranges stop at row {next_row - 1}; they must cover rows 1-{total}.")
    return groups


def section_limits(questions: list[QuestionRow], sections: list[Section]) -> tuple[int, int]:
    """Lowest and highest total a student can have under the section rules."""
    by_label = {q.label: q for q in questions}
    low = high = 0
    for section in sections:
        caps = sorted((by_label[l].allotted or 0 for l in section.labels), reverse=True)
        caps = [c for c in caps if c > 0]
        low += section.answer
        high += sum(caps[: section.answer])
    return low, high


def load_rules(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as error:
        raise UserError(f"{path.name} is not valid JSON ({error}). Delete it to be asked again.") from error


def save_rules(path: Path, marks: MarksTable, template_name: str, sections: list[Section]) -> None:
    rules = load_rules(path)
    rules[marks.fingerprint] = {
        "template": template_name,
        "saved_on": date.today().isoformat(),
        "sections": [{"labels": s.labels, "answer": s.answer} for s in sections],
    }
    path.write_text(json.dumps(rules, indent=2) + "\n", encoding="utf-8")


def show_sections(marks: MarksTable, groups: list[list[int]], answers: list[int] | None = None) -> None:
    for number, indexes in enumerate(groups, start=1):
        suffix = f"   -> answer any {answers[number - 1]}" if answers else ""
        info(f"  Section {number}: {describe_group(marks.questions, indexes)}{suffix}")


def configure_sections(
    marks: MarksTable,
    config: dict,
    rules_path: Path,
    template_name: str,
    *,
    answers: list[int] | None = None,
    force_ask: bool = False,
    save: bool = True,
    input_fn: Callable[[str], str] = input,
) -> list[Section]:
    questions = marks.questions
    by_label = {q.label: q for q in questions}
    saved = load_rules(rules_path).get(marks.fingerprint)

    if saved and not force_ask and answers is None:
        sections = [Section(list(item["labels"]), int(item["answer"])) for item in saved["sections"]]
        if all(label in by_label for s in sections for label in s.labels):
            info("Using the section rules saved for this template "
                 f"({rules_path.name}; pass --ask-sections to change them):")
            groups = [[questions.index(by_label[l]) for l in s.labels] for s in sections]
            show_sections(marks, groups, [s.answer for s in sections])
            return sections

    info(f"Question rows found in the marks table ({len(questions)}):")
    for number, question in enumerate(questions, start=1):
        cap = question.allotted if question.allotted else "none"
        info(f"  {number:>2}. {question.label:<8} marks allotted: {cap}")
    groups = detect_sections(questions)
    info("\nThe questions fall into these sections:")
    show_sections(marks, groups)

    try:
        if answers is None:
            while True:
                reply = input_fn(
                    "\nPress Enter to accept this grouping, or type the row ranges "
                    f"for the sections (for example 1-7,8-{len(questions)}): "
                ).strip()
                if not reply:
                    break
                try:
                    groups = parse_ranges(reply, len(questions))
                except ValueError as error:
                    info(f"  {error}")
                    continue
                info("\nSections now:")
                show_sections(marks, groups)
            chosen: list[int] = []
            info()
            for number, indexes in enumerate(groups, start=1):
                available = len(awardable(questions, indexes))
                if available == 0:
                    info(f"Section {number} has no row that carries marks; it stays blank.")
                    chosen.append(0)
                    continue
                while True:
                    reply = input_fn(
                        f"Answer how many questions from Section {number}? "
                        f"[1-{available}, Enter = {available}]: "
                    ).strip()
                    if not reply:
                        count = available
                    elif reply.isdigit():
                        count = int(reply)
                    else:
                        info("  Please type a whole number.")
                        continue
                    if 1 <= count <= available:
                        chosen.append(count)
                        break
                    info(f"  It must be between 1 and {available}.")
        else:
            if len(answers) != len(groups):
                raise UserError(
                    f"--answer gave {len(answers)} number(s) but {len(groups)} section(s) were "
                    "found. Run without --answer to see and adjust the sections."
                )
            chosen = answers
            for number, (indexes, count) in enumerate(zip(groups, answers), start=1):
                available = len(awardable(questions, indexes))
                if count > available or (count < 1 and available):
                    raise UserError(
                        f"--answer: Section {number} has {available} row(s) that carry marks, "
                        f"so {count} is not possible."
                    )
    except EOFError:
        raise UserError(
            "The section rules for this template are not saved yet and no one is there to "
            "answer. Run once in a normal terminal, or pass --answer (for example --answer 5,4)."
        ) from None

    sections = [
        Section([questions[i].label for i in indexes], count)
        for indexes, count in zip(groups, chosen)
    ]
    low, high = section_limits(questions, sections)
    info(f"\nA student's total can range from {low} to {high} under these rules "
         f"(full marks in the config: {config['full_marks']}).")
    if high < config["full_marks"]:
        info(f"  WARNING: the highest total reachable ({high}) is below the full marks.")
    elif high > config["full_marks"]:
        info(f"  Note: the rules allow up to {high}, above the full marks.")
    else:
        info("  OK: this matches the full marks.")

    if save:
        keep = True
        if answers is None:
            try:
                reply = input_fn(f"Save these answers in {rules_path.name} for next time? [Y/n]: ")
                keep = reply.strip().casefold() not in {"n", "no"}
            except EOFError:
                keep = True
        if keep:
            save_rules(rules_path, marks, template_name, sections)
            info(f"Saved to {rules_path}.")
    return sections


# --------------------------------------------------------------------------
# Random question-wise marks
# --------------------------------------------------------------------------

def allocate(total: int, marks: MarksTable, sections: list[Section], student_key: str) -> dict[str, int]:
    """Split `total` over the questions: exactly N answered per section, 1..cap each."""
    caps = {q.label: q.allotted for q in marks.questions}
    pools = [
        [(label, caps[label]) for label in section.labels if caps[label]]
        for section in sections
    ]
    counts = [section.answer for section in sections]
    needed = sum(counts)
    low, high = section_limits(marks.questions, sections)
    if not low <= total <= high:
        raise ValueError(
            f"{total} marks cannot be split with exactly the required number of answered "
            f"questions (possible totals: {low} to {high})."
        )

    rng = random.Random(hashlib.sha256(student_key.encode("utf-8")).digest())
    chosen = None
    for _ in range(200):
        attempt = [rng.sample(pool, count) for pool, count in zip(pools, counts) if count]
        if sum(cap for picked in attempt for _, cap in picked) >= total:
            chosen = attempt
            break
    if chosen is None:  # always reachable because total <= high
        chosen = [sorted(pool, key=lambda item: -item[1])[:count]
                  for pool, count in zip(pools, counts) if count]

    allocation = {label: 1 for picked in chosen for label, _ in picked}
    limits = {label: cap for picked in chosen for label, cap in picked}
    for _ in range(total - needed):
        open_rows = [label for label, mark in allocation.items() if mark < limits[label]]
        allocation[rng.choice(open_rows)] += 1

    assert sum(allocation.values()) == total
    return allocation


def populate_marks_awarded(document, config: dict, allocation: dict[str, int], font: dict) -> None:
    table, marks = find_marks_table(document, config)
    by_label = {q.label: q for q in marks.questions}
    for label, mark in allocation.items():
        question = by_label[label]
        cells = table.rows[question.table_row].cells
        awarded, allotted = cells[marks.awarded_col], cells[marks.allotted_col]
        paragraph = awarded.paragraphs[0]
        if paragraph.alignment is None:  # line up with the Marks Allotted column
            paragraph.alignment = next((p.alignment for p in allotted.paragraphs
                                        if p.alignment is not None), None)
        for run in paragraph.runs:
            run.text = ""
        reference = next((r for p in allotted.paragraphs for r in p.runs if r._r.rPr is not None), None)
        run = paragraph.add_run(str(mark))
        copy_run_format(reference, run, font)


# --------------------------------------------------------------------------
# Signature (unchanged algorithm)
# --------------------------------------------------------------------------

def add_student_signature(document, student: Student, config: dict) -> None:
    settings = config["signature"]
    label_paragraph = find_label_paragraph(document, settings["label"])

    seed = hashlib.sha256(f"{student.roll}-signature".encode("utf-8")).digest()
    font_name = random.Random(seed).choice(settings["fonts"])

    signature = label_paragraph.insert_paragraph_before()
    signature.alignment = label_paragraph.alignment or WD_ALIGN_PARAGRAPH.LEFT
    signature.paragraph_format.left_indent = label_paragraph.paragraph_format.left_indent
    signature.paragraph_format.space_after = Pt(0)

    name_run = signature.add_run(student.name)
    name_run.font.name = font_name
    name_run.font.size = Pt(settings["name_size_pt"])
    name_run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

    exam_date = str(config["exam_date"]).strip()
    if exam_date:
        date_run = signature.add_run("\n" + exam_date)
        date_run.font.size = Pt(settings["date_size_pt"])


# --------------------------------------------------------------------------
# Template pre-flight
# --------------------------------------------------------------------------

def field_value(entry: dict, student: Student, config: dict) -> str:
    if "column" in entry:
        return student.values[entry["column"]]
    return expand(entry["value"], config).strip()


def open_template(path: Path):
    try:
        return Document(path)
    except Exception as error:
        raise UserError(
            f"Could not open '{path.name}' as a Word (.docx) document ({type(error).__name__}). "
            "Check that it is a .docx file (not an old .doc or Word's temporary '~$' file)."
        ) from error


def read_template_full_marks(document, label: str | None) -> int | None:
    """The number after the template's "Full Marks:" text, or None if it is not there."""
    if not label:
        return None
    try:
        paragraph = find_label_paragraph(document, label, required=False)
    except UserError:
        return None
    if paragraph is None:
        return None
    found = re.search(r"\d+", paragraph_text(paragraph).split(label, 1)[1])
    return int(found.group()) if found else None


def analyse_template(template_path: Path, config: dict) -> tuple[MarksTable | None, list[str]]:
    """Check that every declared label exists; return the marks table and warnings."""
    document = open_template(template_path)
    warnings: list[str] = []
    for entry in config["template_fields"]:
        required = "column" in entry or bool(entry["value"].strip())
        paragraph = find_label_paragraph(document, entry["label"], required=required)
        if paragraph is None:
            warnings.append(f"The template has no '{entry['label']}' (its fixed value is empty, so nothing is lost).")
    if config["signature"]["enabled"]:
        find_label_paragraph(document, config["signature"]["label"])

    label = config.get("full_marks_label")
    stated = read_template_full_marks(document, label)
    if stated is not None and stated != config["full_marks"]:
        warnings.append(
            f"The template says '{label} {stated}' but full marks is set to {config['full_marks']}."
        )

    marks = None
    if config["marks_table"] is not None:
        _, marks = find_marks_table(document, config)
    return marks, warnings


# --------------------------------------------------------------------------
# Official seal (optional): stamped just below the last line of every sheet
# --------------------------------------------------------------------------

def check_seal_image(path: Path) -> list[str]:
    """Raise UserError if the seal cannot be used; otherwise return any warnings."""
    if not path.is_file():
        raise UserError(f"The seal image was not found: {path}")
    try:
        from docx.image.image import Image
        image = Image.from_file(str(path))
    except Exception as error:
        raise UserError(
            f"'{path.name}' could not be read as a picture ({type(error).__name__}). Use a PNG or JPG file."
        ) from error
    warnings: list[str] = []
    if image.content_type != "image/png":
        warnings.append(
            f"The seal image is not a PNG ({image.content_type}). It cannot have a transparent "
            "background, so a white box may show around the seal. A PNG with a transparent background is best."
        )
    size = path.stat().st_size
    if size > 1_500_000:
        warnings.append(
            f"The seal image is {size / 1_000_000:.1f} MB and a copy goes into every sheet; "
            "a smaller file keeps the sheets small."
        )
    return warnings


def has_seal_picture(document) -> bool:
    """True if the document already holds a picture or shape whose name or description says 'seal'."""
    for doc_pr in document.element.body.iter(qn("wp:docPr")):
        if "seal" in f"{doc_pr.get('name', '')} {doc_pr.get('descr', '')}".casefold():
            return True
    return False


def stamp_seal(document, image_path: Path, config: dict) -> str:
    """Put the seal just below the last line of the document, towards the bottom right.

    A one-point-high paragraph is added after the last line and the seal floats from it, so the
    seal follows wherever the document really ends and the page layout does not change.
    Returns "stamped", or "already had a seal" if the sheet already contains one.
    """
    if has_seal_picture(document):
        return "already had a seal"
    settings = config["seal"]
    section = document.sections[-1]
    text_width = section.page_width - section.left_margin - section.right_margin
    width = Cm(settings["width_cm"])
    offset = Cm(settings["offset_cm"])
    if settings["align"] == "right":
        x = text_width - width - offset
    elif settings["align"] == "left":
        x = offset
    else:
        x = (text_width - width) // 2
    x = max(int(x), 0)
    y = int(Cm(settings["gap_cm"]))

    paragraph = document.add_paragraph()
    spacing = paragraph.paragraph_format
    spacing.space_before = Pt(0)
    spacing.space_after = Pt(0)
    spacing.line_spacing = Pt(1)
    spacing.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    run = paragraph.add_run()
    run.add_picture(str(image_path), width=width)

    inline = run._r.xpath(".//wp:inline")[0]
    extent = inline.find(qn("wp:extent"))
    shape_id = inline.find(qn("wp:docPr")).get("id")
    graphic = inline.find(qn("a:graphic"))
    anchor = parse_xml(
        f'<wp:anchor {nsdecls("wp", "a")} distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{extent.get("cx")}" cy="{extent.get("cy")}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{shape_id}" name="Seal" descr="Official seal"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '</wp:anchor>'
    )
    anchor.append(graphic)
    inline.getparent().replace(inline, anchor)
    return "stamped"


def choose_seal(args: argparse.Namespace, count: int,
                input_fn: Callable[[str], str]) -> tuple[Path | None, list[str]]:
    """Ask whether to stamp the seal (unless --seal says so) and let the user pick the image."""
    if count == 0:
        return None, []
    path = args.seal
    if path is None:
        if args.no_prompt or args.check:
            return None, []
        try:
            reply = input_fn(f"\nStamp the official seal on all {count} sheet(s)? [y/N]: ")
        except EOFError:
            return None, []
        if reply.strip().casefold() not in {"y", "yes"}:
            return None, []
        try:
            path = choose_input_file("seal image", (".png", ".jpg", ".jpeg"))
        except UserError:
            info("No seal image was chosen, so the sheets will be created without a seal.")
            return None, []
        except Exception as error:
            raise UserError(
                f"The file picker could not be opened ({type(error).__name__}). "
                "Give the seal image with --seal PATH instead."
            ) from error
    return path, check_seal_image(path)


# --------------------------------------------------------------------------
# DOCX job (runs in worker processes)
# --------------------------------------------------------------------------

def create_docx(template_path: Path, student: Student, allocation: dict[str, int] | None,
                output_path: Path, config: dict, seal_path: Path | None = None) -> str:
    """Write one sheet. Returns the seal outcome: "", "stamped" or "already had a seal"."""
    document = Document(template_path)
    font = font_settings(config)
    for entry in config["template_fields"]:
        value = field_value(entry, student, config)
        if value:
            write_field(find_label_paragraph(document, entry["label"]), entry, value, font)
    if allocation:
        populate_marks_awarded(document, config, allocation, font)
    if config["signature"]["enabled"]:
        add_student_signature(document, student, config)
    seal = stamp_seal(document, seal_path, config) if seal_path else ""
    document.save(output_path)
    return seal


def docx_job(job: tuple) -> tuple[str, str | None, str]:
    template_path, student, allocation, output_path, config, seal_path = job
    try:
        seal = create_docx(Path(template_path), student, allocation, Path(output_path), config,
                           Path(seal_path) if seal_path else None)
        return student.roll, None, seal
    except Exception as error:  # reported per student, the batch continues
        return student.roll, f"{type(error).__name__}: {error}", ""


# --------------------------------------------------------------------------
# PDF export
# --------------------------------------------------------------------------

def word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def find_soffice(explicit: str | None) -> str | None:
    if explicit:
        return explicit
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(candidate).exists():
            return candidate
    return None


def export_with_word(pairs: list[tuple[Path, Path]]) -> list[str]:
    import win32com.client

    failures: list[str] = []
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for docx_path, pdf_path in pairs:
            try:
                document = word.Documents.Open(str(docx_path.resolve()))
                try:
                    document.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
                finally:
                    document.Close(False)
            except Exception as error:
                failures.append(f"{pdf_path.name}: {error}")
    finally:
        word.Quit()
    return failures


def export_with_libreoffice(pairs: list[tuple[Path, Path]], soffice: str) -> list[str]:
    failures: list[str] = []
    for start in range(0, len(pairs), 20):
        chunk = pairs[start:start + 20]
        out_dir = chunk[0][1].parent
        command = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir)]
        command += [str(docx.resolve()) for docx, _ in chunk]
        try:
            subprocess.run(command, check=True, capture_output=True, timeout=900)
        except (subprocess.SubprocessError, OSError) as error:
            failures.append(f"LibreOffice call failed: {error}")
        for _, pdf_path in chunk:
            if not pdf_path.exists():
                failures.append(f"{pdf_path.name}: LibreOffice did not produce it.")
    return failures


def export_pdfs(pairs: list[tuple[Path, Path]], engine: str, soffice: str | None) -> list[str]:
    if not pairs:
        return []
    if engine == "auto":
        engine = "word" if word_available() else "libreoffice"
    if engine == "word":
        if not word_available():
            raise UserError("--pdf-engine word needs Windows with Microsoft Word and pywin32 installed.")
        info(f"Exporting {len(pairs)} PDF file(s) with Microsoft Word...")
        return export_with_word(pairs)
    path = find_soffice(soffice)
    if not path:
        raise UserError(
            "No PDF engine is available. Install Microsoft Word with pywin32 (Windows) or "
            "LibreOffice, or run with --keep-docx-only. A LibreOffice location can be given "
            "with --soffice PATH."
        )
    info(f"Exporting {len(pairs)} PDF file(s) with LibreOffice...")
    return export_with_libreoffice(pairs, path)


# --------------------------------------------------------------------------
# Run planning
# --------------------------------------------------------------------------

@dataclass
class Job:
    student: Student
    stem: str
    docx_path: Path
    pdf_path: Path
    allocation: dict[str, int] | None
    write_docx: bool = True
    write_pdf: bool = True
    seal: str = ""
    notes: list[str] = field(default_factory=list)


def plan_jobs(students: list[Student], marks: MarksTable | None, sections: list[Section] | None,
              output_dir: Path, config: dict, make_pdf: bool) -> tuple[list[Job], list[str], list[str]]:
    jobs: list[Job] = []
    errors: list[str] = []
    skipped: list[str] = []
    stems: dict[str, Student] = {}
    overwrite = config["existing_files"] == "overwrite"

    if marks and sections:
        low, high = section_limits(marks.questions, sections)

    for student in students:
        allocation = None
        if student.total is not None and marks and sections:
            if not low <= student.total <= high:
                message = (
                    f"Row {student.row} (roll {student.roll}, {student.name}): {student.total} marks "
                    f"cannot be split with exactly the required number of answered questions "
                    f"(possible totals: {low} to {high})."
                )
                if config["unreachable_marks"] == "stop":
                    errors.append(message)
                else:
                    skipped.append(message + " Skipped.")
                continue
            allocation = allocate(student.total, marks, sections, student.roll)

        stem = clean_filename(f"{student.roll}_{student.name}")
        clash = stems.get(stem.casefold())
        if clash:
            errors.append(
                f"Rows {clash.row} and {student.row} would both be saved as '{stem}'. "
                "Roll numbers are unique, so their names must clean up to the same text."
            )
            continue
        stems[stem.casefold()] = student

        job = Job(student, stem, output_dir / DOCX_FOLDER / f"{stem}.docx",
                  output_dir / PDF_FOLDER / f"{stem}.pdf", allocation)
        if job.docx_path.exists() and not overwrite:
            job.write_docx = False
            job.notes.append("docx already existed, kept")
        if make_pdf and job.pdf_path.exists() and not overwrite:
            job.write_pdf = False
            job.notes.append("pdf already existed, kept")
        jobs.append(job)
    return jobs, errors, skipped


def write_report(path: Path, jobs: list[Job], marks: MarksTable | None, failures: dict[str, str],
                 notes: list[str], marks_header: str) -> None:
    order = [q.label for q in marks.questions] if marks else []
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Roll", "Name", "Workbook row", "Marks obtained", "Question-wise marks",
                         "Status", "DOCX", "PDF", "Seal", "Notes"])
        for job in jobs:
            student = job.student
            split = ""
            if job.allocation:
                split = "; ".join(f"{l}={job.allocation[l]}" for l in order if l in job.allocation)
            status = "FAILED" if student.roll in failures else ("kept existing" if not job.write_docx else "written")
            writer.writerow([student.roll, student.name, student.row,
                             student.values.get(marks_header, ""),
                             split, status, job.docx_path.name, job.pdf_path.name, job.seal,
                             "; ".join(job.notes + ([failures[student.roll]] if student.roll in failures else []))])
        for note in notes:
            writer.writerow(["", "", "", "", "", "NOTE", "", "", "", note])


# --------------------------------------------------------------------------
# Settings screen: per-exam values asked for at the start of a run
# --------------------------------------------------------------------------

def parse_exam_date(text: str) -> str | None:
    """DD-MM-YYYY text normalised (2-9-2026 -> 02-09-2026), or None if it is not a date."""
    try:
        return datetime.strptime(text.strip(), "%d-%m-%Y").strftime("%d-%m-%Y")
    except ValueError:
        return None


def fixed_value_entries(config: dict) -> list[dict]:
    """Fixed-value labels that are typed by hand (not derived, like {exam_date})."""
    return [e for e in config["template_fields"] if "value" in e and "{" not in e["value"]]


def set_fixed_value(config: dict, label: str, value: str) -> None:
    config["fixed_values"][label] = value
    apply_fixed_values(config)


def settings_rows(config: dict) -> list[tuple[str, str, str]]:
    """(key, title, shown value) for each line of the settings screen."""
    rows = [
        ("exam_date", "Exam date", config["exam_date"] or "(blank)"),
        ("full_marks", "Full marks", str(config["full_marks"])),
    ]
    for entry in fixed_value_entries(config):
        rows.append((f"fixed:{entry['label']}", entry["label"].rstrip(": "), entry["value"] or "(blank)"))
    mode = config["font_mode"]
    rows.append(("font_mode", "Font mode",
                 mode if mode == "template" else f"custom ({config['custom_font']['name']})"))
    return rows


def review_settings(config: dict, changes: dict, *, template_full_marks: int | None = None,
                    full_marks_from_template: bool = False,
                    input_fn: Callable[[str], str] = input) -> None:
    """Show the per-exam settings and let the user change them.

    `config` is updated as the user answers; every change is also recorded in
    `changes` (in config.local.json's shape) so the caller can offer to save it.
    """
    def ask(prompt: str) -> str:
        return input_fn(prompt).strip()

    while True:
        rows = settings_rows(config)
        info("\nSettings for this run")
        for number, (_, title, shown) in enumerate(rows, start=1):
            info(f"  {number}. {title:<26}{shown}")
        if template_full_marks is not None:
            if full_marks_from_template:
                info(f"     (Full marks {config['full_marks']} was read from the template.)")
            elif template_full_marks != config["full_marks"]:
                info(f"     (Note: the template itself says Full Marks: {template_full_marks}.)")

        reply = ask("\nPress Enter to continue, or type a number to change it: ")
        if not reply:
            return
        if not (reply.isdigit() and 1 <= int(reply) <= len(rows)):
            info("  Please type one of the numbers shown, or just press Enter.")
            continue
        key, title, _ = rows[int(reply) - 1]

        if key == "exam_date":
            current = config["exam_date"] or "blank"
            answer = ask(f"Exam date as DD-MM-YYYY [{current}] (Enter keeps, - clears): ")
            if not answer:
                continue
            if answer == "-":
                new = ""
            else:
                new = parse_exam_date(answer)
                if new is None:
                    use = ask(f"  '{answer}' does not look like DD-MM-YYYY. Use it anyway? [y/N]: ")
                    if use.casefold() not in {"y", "yes"}:
                        continue
                    new = answer
            config["exam_date"] = new
            changes["exam_date"] = new

        elif key == "full_marks":
            answer = ask(f"Full marks [{config['full_marks']}]: ")
            if not answer:
                continue
            if not answer.isdigit() or int(answer) < 1:
                info("  Full marks must be a whole number of at least 1.")
                continue
            config["full_marks"] = int(answer)
            changes["full_marks"] = int(answer)
            full_marks_from_template = False

        elif key.startswith("fixed:"):
            label = key.split(":", 1)[1]
            entry = next(e for e in config["template_fields"] if e["label"] == label)
            answer = ask(f"{title} [{entry['value'] or 'blank'}] (Enter keeps, - clears): ")
            if not answer:
                continue
            new = "" if answer == "-" else answer
            set_fixed_value(config, label, new)
            changes.setdefault("fixed_values", {})[label] = new

        elif key == "font_mode":
            answer = ask("Font: 1 = the template's own font, 2 = a custom font. Choose 1 or 2: ")
            if answer == "1":
                config["font_mode"] = "template"
                changes["font_mode"] = "template"
            elif answer == "2":
                current = config["custom_font"]["name"]
                name = ask(f"Font name [{current}]: ") or current
                if not name:
                    info("  A font name is needed for a custom font.")
                    continue
                config["font_mode"] = "custom"
                config["custom_font"]["name"] = name
                changes["font_mode"] = "custom"
                changes.setdefault("custom_font", {})["name"] = name
            elif answer:
                info("  Please type 1 or 2.")


# --------------------------------------------------------------------------
# Completing sheets that students have filled in
# --------------------------------------------------------------------------

EMPTY_ROLL = "field is empty"


@dataclass
class Submission:
    path: Path
    roll: str = ""
    student: Student | None = None
    marks: MarksTable | None = None
    problem: str = ""
    notes: list[str] = field(default_factory=list)
    stem: str = ""
    allocation: dict[str, int] | None = None
    status: str = ""
    seal: str = ""


def roll_entry(config: dict) -> dict:
    for entry in config["template_fields"]:
        if entry.get("column") == config["columns"]["roll_number"]:
            return entry
    raise UserError(
        "Completing submitted sheets needs a template_fields entry fed by the roll number column "
        f"('{config['columns']['roll_number']}'), so each sheet can be matched to the workbook."
    )


def roll_key(text: str) -> str:
    """Roll numbers are compared without case or spaces: 'LME 2026410' equals 'lme2026410'."""
    return re.sub(r"\s+", "", normalise_roll(text)).casefold()


def list_submission_files(folder: Path) -> tuple[list[Path], list[str]]:
    files: list[Path] = []
    refused: list[str] = []
    for path in sorted(folder.iterdir()):
        if not path.is_file() or path.name.startswith("~$"):  # "~$" is Word's temporary lock file
            continue
        suffix = path.suffix.casefold()
        if suffix == ".docx":
            files.append(path)
        elif suffix in {".docm", ".doc", ".dotx", ".dotm", ".odt", ".pdf"}:
            refused.append(f"{path.name}: not read (only .docx files are accepted)")
    if not files:
        raise UserError(f"There are no .docx files in {folder}.")
    return files, refused


def scan_submission(path: Path, config: dict, entry: dict) -> Submission:
    """Open one submitted sheet and read the roll number from it. Never raises."""
    submission = Submission(path)
    try:
        document = Document(path)
    except Exception as error:
        submission.problem = f"could not be opened as a Word document ({type(error).__name__}: {error})"
        return submission
    try:
        for declared in config["template_fields"]:
            required = "column" in declared or bool(declared["value"].strip())
            find_label_paragraph(document, declared["label"], required=required)
        paragraph = find_label_paragraph(document, entry["label"])
        submission.roll = read_field(paragraph, entry["label"], entry.get("where", "after-label"))
        if config["marks_table"] is not None:
            submission.marks = find_marks_table(document, config)[1]
    except UserError as error:
        submission.problem = f"does not match the expected sheet: {error}"
        return submission
    if not submission.roll:
        submission.problem = f"the '{entry['label']}' {EMPTY_ROLL}"
    return submission


def same_value(a: str, b: str, column: str, config: dict) -> bool:
    mobile = config.get("mobile_check")
    if mobile and column == mobile["column"]:
        return re.sub(r"\D", "", a) == re.sub(r"\D", "", b)
    return squash(a) == squash(b)


def complete_submission(submission: Submission, config: dict, output_path: Path,
                        seal_path: Path | None = None) -> None:
    """Add what the student's sheet is missing. What the student wrote is never overwritten."""
    document = Document(submission.path)
    font = font_settings(config)
    student = submission.student
    marks_header = config["columns"]["marks_obtained"]
    roll_label = roll_entry(config)["label"]
    for entry in config["template_fields"]:
        label = entry["label"]
        where = entry.get("where", "after-label")
        paragraph = find_label_paragraph(document, label, required=False)
        if paragraph is None:
            continue
        current = read_field(paragraph, label, where)
        if entry.get("column") == marks_header:  # the program's own field
            value = student.values[marks_header]
            if value:
                write_field(paragraph, entry, value, font)
                if current and squash(current) != squash(value):
                    submission.notes.append(f"'{label}' held '{current}', replaced with '{value}'")
        elif "column" in entry:  # the student's own field
            if label == roll_label:
                continue
            expected = student.values[entry["column"]]
            if not current:
                if expected:
                    write_field(paragraph, entry, expected, font)
                    submission.notes.append(f"'{label}' was blank, filled from the workbook")
            elif expected and not same_value(current, expected, entry["column"], config):
                submission.notes.append(
                    f"'{label}' in the sheet is '{current}' but the workbook has '{expected}' (kept the sheet's)"
                )
        else:  # a fixed value: only fills a blank
            value = expand(entry["value"], config).strip()
            if value and not current:
                write_field(paragraph, entry, value, font)
    if submission.allocation:
        populate_marks_awarded(document, config, submission.allocation, font)
    if seal_path:
        submission.seal = stamp_seal(document, seal_path, config)
    document.save(output_path)


def write_submission_report(path: Path, submissions: list[Submission], missing: list[Student],
                            marks_header: str) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Roll", "Name", "Submitted file", "Workbook row", "Marks obtained",
                         "Question-wise marks", "Status", "Seal", "Notes"])
        for item in submissions:
            student = item.student
            split = ""
            if item.allocation and item.marks:
                order = [q.label for q in item.marks.questions]
                split = "; ".join(f"{l}={item.allocation[l]}" for l in order if l in item.allocation)
            writer.writerow([
                student.roll if student else item.roll, student.name if student else "", item.path.name,
                student.row if student else "", student.values.get(marks_header, "") if student else "",
                split, item.status or ("SKIPPED" if item.problem else ""), item.seal,
                "; ".join(([item.problem] if item.problem else []) + item.notes)])
        for student in missing:
            writer.writerow([student.roll, student.name, "", student.row,
                             student.values.get(marks_header, ""), "", "NO SUBMISSION", "", ""])


def run_submissions(args: argparse.Namespace, config: dict, folder: Path, students: list[Student],
                    warnings: list[str], input_fn: Callable[[str], str]) -> int:
    entry = roll_entry(config)
    files, refused = list_submission_files(folder)
    info(f"Submissions: {len(files)} .docx file(s) in {folder}")
    scanned = [scan_submission(path, config, entry) for path in files]

    if all(item.problem.endswith(EMPTY_ROLL) for item in scanned):
        raise UserError(
            f"No roll number could be read from any of the {len(scanned)} sheet(s). If the students "
            f"typed it in the cell to the right of '{entry['label']}', set \"where\": \"next-cell\" "
            "on that entry in the config (see presets/ca2.json). Otherwise the sheets may use a "
            "different label than the config expects."
        )

    by_key: dict[str, Student] = {}
    for student in students:
        by_key.setdefault(roll_key(student.roll), student)
    groups: dict[str, list[Submission]] = {}
    for item in scanned:
        if item.problem:
            continue
        student = by_key.get(roll_key(item.roll))
        if student is None:
            item.problem = f"roll number '{item.roll}' is not in the workbook"
            continue
        item.student = student
        groups.setdefault(roll_key(student.roll), []).append(item)
    for group in groups.values():
        if len(group) > 1:
            names = ", ".join(item.path.name for item in group)
            for item in group:
                item.problem = (f"roll {group[0].student.roll} was submitted more than once ({names}); "
                                "none of these files was used")

    rules_path = (args.config.resolve().parent if args.config else HERE) / RULES_FILENAME
    answers = [int(x) for x in args.answer.split(",")] if args.answer else None
    layouts: dict[str, list[Submission]] = {}
    for item in scanned:
        if not item.problem and item.marks:
            layouts.setdefault(item.marks.fingerprint, []).append(item)
    sections_for: dict[str, list[Section] | None] = {}
    for fingerprint, items in layouts.items():
        marks = items[0].marks
        if len(layouts) > 1:
            info(f"\nAnother sheet layout was found in {len(items)} file(s), for example {items[0].path.name}.")
        if args.check and answers is None and not load_rules(rules_path).get(fingerprint):
            info("\nSection rules are not set for this layout yet; the next normal run will ask. "
                 "Detected sections:")
            show_sections(marks, detect_sections(marks.questions))
            sections_for[fingerprint] = None
            continue
        try:
            sections_for[fingerprint] = configure_sections(
                marks, config, rules_path, items[0].path.name, answers=answers,
                force_ask=args.ask_sections, save=not args.check, input_fn=input_fn)
        except UserError as error:
            for item in items:
                item.problem = str(error)

    for item in scanned:
        if item.problem:
            continue
        sections = sections_for.get(item.marks.fingerprint) if item.marks else None
        total = item.student.total
        if total is not None and sections:
            low, high = section_limits(item.marks.questions, sections)
            if not low <= total <= high:
                item.problem = (f"{total} marks cannot be split with exactly the required number of "
                                f"answered questions (possible totals: {low} to {high})")
                continue
            item.allocation = allocate(total, item.marks, sections, item.student.roll)
        item.stem = clean_filename(f"{item.student.roll}_{item.student.name}")

    output_dir = args.output
    make_pdf = not args.docx_only
    ready = [item for item in scanned if not item.problem]
    kept_existing = 0
    for item in ready:
        docx_path = output_dir / DOCX_FOLDER / f"{item.stem}.docx"
        if docx_path.exists() and config["existing_files"] == "skip":
            item.status = "kept existing"
            kept_existing += 1
    if kept_existing:
        warnings.append(f"{kept_existing} file(s) already exist in the output folder and will be kept.")
    warnings += refused
    if warnings:
        report_problems("Warnings", warnings)

    done_roll = {roll_key(item.student.roll) for item in ready}
    missing = [s for s in students if roll_key(s.roll) not in done_roll
               and roll_key(s.roll) not in groups]
    if args.check:
        if args.seal:
            for message in check_seal_image(args.seal):
                info(f"  Seal warning: {message}")
            info(f"The seal image {args.seal.name} can be used.")
        info(f"\nCheck finished: {len(ready)} sheet(s) could be completed, "
             f"{sum(1 for i in scanned if i.problem)} would be skipped, "
             f"{len(missing)} student(s) have no submission. No files were written.")
        for item in scanned:
            if item.problem:
                info(f"  skipped {item.path.name}: {item.problem}")
        return 0

    seal_path, seal_warnings = choose_seal(
        args, sum(1 for item in ready if item.status != "kept existing"), input_fn)
    if seal_warnings:
        report_problems("Warnings about the seal image", seal_warnings)
    (output_dir / DOCX_FOLDER).mkdir(parents=True, exist_ok=True)
    if make_pdf:
        (output_dir / PDF_FOLDER).mkdir(parents=True, exist_ok=True)
    info(f"\nCompleting {len(ready)} sheet(s)...")
    written: list[Submission] = []
    for item in ready:
        if item.status == "kept existing":
            continue
        try:
            complete_submission(item, config, output_dir / DOCX_FOLDER / f"{item.stem}.docx", seal_path)
            item.status = "completed"
            written.append(item)
        except Exception as error:
            item.problem = f"could not be completed ({type(error).__name__}: {error})"

    pdf_failures: list[str] = []
    if make_pdf:
        pairs = [(output_dir / DOCX_FOLDER / f"{i.stem}.docx", output_dir / PDF_FOLDER / f"{i.stem}.pdf")
                 for i in written]
        pdf_failures = export_pdfs(pairs, args.pdf_engine, args.soffice)

    skipped = [item for item in scanned if item.problem]
    differences = [(i, n) for i in scanned if not i.problem for n in i.notes if "workbook has" in n]
    filled = [(i, n) for i in scanned if not i.problem for n in i.notes if "was blank" in n]
    replaced = [(i, n) for i in scanned if not i.problem for n in i.notes if "replaced with" in n]
    write_submission_report(output_dir / REPORT_FILENAME, scanned, missing, config["columns"]["marks_obtained"])

    info(f"\nCompleted: {len(written)} sheet(s)" + (f", kept existing: {kept_existing}" if kept_existing else ""))
    if seal_path:
        info(f"Seal: stamped on {sum(1 for i in written if i.seal == 'stamped')} of {len(written)} sheet(s) ({seal_path.name}).")
        had = [i for i in written if i.seal == "already had a seal"]
        if had:
            report_problems("Sheets that already had a seal (not stamped again)", [i.path.name for i in had])
    if skipped:
        report_problems("Files that were not used", [f"{i.path.name}: {i.problem}" for i in skipped])
    if differences:
        report_problems("Details in the sheet that differ from the workbook (the sheet's text was kept)",
                        [f"{i.path.name}: {n}" for i, n in differences])
    if filled:
        report_problems("Blank fields filled from the workbook", [f"{i.path.name}: {n}" for i, n in filled])
    if replaced:
        report_problems("Marks fields that already held text (replaced)", [f"{i.path.name}: {n}" for i, n in replaced])
    if missing:
        listed = [f"{s.roll} {s.name}" for s in missing[:30]]
        if len(missing) > 30:
            listed.append(f"... and {len(missing) - 30} more (see {REPORT_FILENAME})")
        report_problems("Students in the workbook with no submitted sheet", listed)
    for message in pdf_failures:
        info(f"  PDF problem: {message}")
    info(f"\nFiles are in: {output_dir}\nRun report: {output_dir / REPORT_FILENAME}")
    return 1 if skipped or pdf_failures else 0


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

def choose_input_file(label: str, suffixes: tuple[str, ...]) -> Path:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title=f"Select the {label}",
        filetypes=[(label, " ".join(f"*{s}" for s in suffixes)), ("All files", "*.*")],
    )
    root.destroy()
    if not selected:
        raise UserError(f"No {label} was selected.")
    return Path(selected)


def choose_folder(label: str) -> Path:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askdirectory(title=f"Select the {label}")
    root.destroy()
    if not selected:
        raise UserError(f"No {label} was selected.")
    return Path(selected)


def choose_sources(args: argparse.Namespace, input_fn: Callable[[str], str]) -> tuple[Path | None, Path | None]:
    """Returns (submissions folder, template). Exactly one of them is set."""
    if args.submissions and args.template:
        raise UserError("Use either --template (generate new sheets) or --submissions "
                        "(complete sheets filled in by students), not both.")
    if args.submissions:
        return args.submissions, None
    if args.template:
        return None, args.template
    info("\nWhat would you like to do?")
    info("  1. Generate new sheets from a template")
    info("  2. Complete sheets that students have filled in")
    try:
        reply = input_fn("Choose 1 or 2 [1]: ").strip()
    except EOFError:
        reply = ""
    if reply == "2":
        return choose_folder("folder of the students' filled-in sheets"), None
    return None, choose_input_file("top sheet template", (".docx",))


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create one completed CA1 top sheet per student.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--workbook", type=Path, help="Excel workbook with the student data")
    parser.add_argument("--template", type=Path, help="Word top-sheet template (to generate new sheets)")
    parser.add_argument("--submissions", type=Path,
                        help="Folder of sheets filled in by students; they are completed instead of "
                             "generating new sheets (use instead of --template)")
    parser.add_argument("--sheet", help="Worksheet name (default: the active sheet)")
    parser.add_argument("--output", type=Path, default=Path("output"), help="Output folder")
    parser.add_argument("--config", type=Path, help="Config file (default: config.json beside the script)")
    parser.add_argument("--workers", type=int, default=max(1, min(os.cpu_count() or 1, 4)),
                        help="Parallel DOCX workers")
    parser.add_argument("--keep-docx-only", "--no-pdf", dest="docx_only", action="store_true",
                        help="Create DOCX files only, no PDFs")
    parser.add_argument("--check", action="store_true",
                        help="Validate everything and report; write no files")
    parser.add_argument("--seal", type=Path,
                        help="Image of the official seal to stamp on every sheet (PNG with a transparent "
                             "background is best). Without it, an interactive run asks whether to stamp one")
    parser.add_argument("--no-prompt", action="store_true",
                        help="Do not show the settings screen; use config.json and config.local.json as they are")
    parser.add_argument("--ask-sections", action="store_true",
                        help="Ask about the sections again even if rules are saved")
    parser.add_argument("--answer", help="Questions to answer per section, e.g. 5,4 "
                                         "(skips the questions; uses the detected sections)")
    parser.add_argument("--pdf-engine", choices=("auto", "word", "libreoffice"), default="auto")
    parser.add_argument("--soffice", help="Path to LibreOffice's soffice program")
    parser.add_argument("--exam-date", help="Overrides config exam_date for this run")
    parser.add_argument("--full-marks", type=int, help="Overrides config full_marks for this run")
    parser.add_argument("--font-mode", choices=("template", "custom"), help="Overrides config font_mode")
    parser.add_argument("--custom-font", help="Font name for font mode 'custom'")
    parser.add_argument("--duplicates", choices=("stop", "keep-first", "keep-last"))
    parser.add_argument("--existing", choices=("overwrite", "skip"), dest="existing_files")
    return parser.parse_args(argv)


def apply_overrides(config: dict, args: argparse.Namespace) -> set[str]:
    """Apply command-line overrides; return the setting names they set."""
    given: set[str] = set()
    if args.exam_date is not None:
        config["exam_date"] = args.exam_date
        given.add("exam_date")
    if args.full_marks is not None:
        config["full_marks"] = args.full_marks
        given.add("full_marks")
    if args.font_mode:
        config["font_mode"] = args.font_mode
        given.add("font_mode")
    if args.custom_font:
        config["custom_font"]["name"] = args.custom_font
        config["font_mode"] = args.font_mode or "custom"
        given.update({"font_mode", "custom_font"})
    if args.duplicates:
        config["duplicates"] = args.duplicates
        given.add("duplicates")
    if args.existing_files:
        config["existing_files"] = args.existing_files
        given.add("existing_files")
    validate_config(config)
    return given


def report_problems(title: str, items: list[str]) -> None:
    info(f"\n{title} ({len(items)}):")
    for item in items:
        info("  - " + item.replace("\n", "\n    "))


def run(args: argparse.Namespace, input_fn: Callable[[str], str] = input) -> int:
    config = load_config(args.config)
    explicit = set(read_local_config(args.config)) | apply_overrides(config, args)

    workbook_path = args.workbook or choose_input_file("student data workbook", (".xlsx", ".xlsm"))
    submissions_dir, template_path = choose_sources(args, input_fn)
    if not workbook_path.is_file():
        raise UserError(f"The workbook file was not found: {workbook_path}")
    if template_path is not None and not template_path.is_file():
        raise UserError(f"The template file was not found: {template_path}")
    if submissions_dir is not None and not submissions_dir.is_dir():
        raise UserError(f"The submissions folder was not found: {submissions_dir}")

    # Full marks: offer the number the template itself states, unless the user has set one.
    stated = (read_template_full_marks(open_template(template_path), config.get("full_marks_label"))
              if template_path is not None else None)
    from_template = False
    if stated is not None and "full_marks" not in explicit and stated != config["full_marks"]:
        config["full_marks"] = stated
        from_template = True

    if not args.no_prompt and not args.check:
        changes: dict = {}
        try:
            review_settings(config, changes, template_full_marks=stated,
                            full_marks_from_template=from_template, input_fn=input_fn)
            if changes:
                local_name = local_config_path(args.config).name
                reply = input_fn(f"\nSave these changes as your defaults in {local_name}? [Y/n]: ")
                if reply.strip().casefold() not in {"n", "no"}:
                    info(f"Saved to {save_local_config(args.config, changes)}.")
        except EOFError:
            info("\n(No one is answering the questions; continuing with the settings shown.)")
        validate_config(config)

    students, notes = read_students(workbook_path, args.sheet, config)
    info(f"Workbook: {workbook_path.name} - {len(students)} student row(s) read.")
    students, duplicate_notes, conflicts = resolve_duplicates(students, config["duplicates"])
    warnings = notes + duplicate_notes
    if submissions_dir is None and not config["exam_date"].strip():
        warnings.append(
            "The exam date is empty, so 'Date of Examination' and the date under the signature "
            "are left blank. Set it on the settings screen or with --exam-date."
        )
    errors = list(conflicts)
    data_errors, data_warnings = validate_students(students, config)
    errors += data_errors
    warnings += data_warnings

    if submissions_dir is not None:
        return run_submissions(args, config, submissions_dir, students, warnings, input_fn)

    marks, template_warnings = analyse_template(template_path, config)
    warnings += template_warnings
    info(f"Template: {template_path.name}")
    if marks:
        info(f"Marks table: {len(marks.questions)} question row(s) found.")

    if errors:
        report_problems("Problems that stop the run", errors)
        if warnings:
            report_problems("Warnings", warnings)
        info("\nNothing was generated. Fix the items above and run again.")
        return 1

    sections = None
    if marks:
        rules_path = (args.config.resolve().parent if args.config else HERE) / RULES_FILENAME
        answers = [int(x) for x in args.answer.split(",")] if args.answer else None
        if args.check and answers is None and not load_rules(rules_path).get(marks.fingerprint):
            info("\nSection rules are not set for this template yet; the next normal run will ask. "
                 "Detected sections:")
            show_sections(marks, detect_sections(marks.questions))
        else:
            sections = configure_sections(
                marks, config, rules_path, template_path.name, answers=answers,
                force_ask=args.ask_sections, save=not args.check, input_fn=input_fn)

    make_pdf = not args.docx_only
    output_dir = args.output
    jobs, plan_errors, skipped = plan_jobs(students, marks if sections else None, sections,
                                           output_dir, config, make_pdf)
    warnings += skipped
    if plan_errors:
        report_problems("Problems that stop the run", plan_errors)
        if warnings:
            report_problems("Warnings", warnings)
        info("\nNothing was generated. Fix the items above and run again.")
        return 1

    existing = [j for j in jobs if (j.docx_path.exists() or (make_pdf and j.pdf_path.exists()))]
    if existing:
        action = "will be overwritten" if config["existing_files"] == "overwrite" else "will be kept"
        warnings.append(f"{len(existing)} file(s) already exist in the output folder and {action}.")
    if warnings:
        report_problems("Warnings", warnings)

    absent = sum(1 for j in jobs if j.student.total is None)
    info(f"\nReady: {len(jobs)} sheet(s), {absent} absent. Font mode: {config['font_mode']}. "
         f"Exam date: {config['exam_date'] or '(blank)'}.")
    if args.check:
        if args.seal:
            for message in check_seal_image(args.seal):
                info(f"Seal warning: {message}")
            info(f"The seal image {args.seal.name} can be used.")
        info("Check finished; no files were written.")
        return 0

    seal_path, seal_warnings = choose_seal(args, sum(1 for j in jobs if j.write_docx), input_fn)
    if seal_warnings:
        report_problems("Warnings about the seal image", seal_warnings)

    (output_dir / DOCX_FOLDER).mkdir(parents=True, exist_ok=True)
    if make_pdf:
        (output_dir / PDF_FOLDER).mkdir(parents=True, exist_ok=True)

    to_write = [j for j in jobs if j.write_docx]
    failures: dict[str, str] = {}
    if to_write:
        info(f"Creating {len(to_write)} DOCX file(s) with {args.workers} worker(s)...")
        work = [(str(template_path), j.student, j.allocation, str(j.docx_path), config,
                 str(seal_path) if seal_path else None) for j in to_write]
        if args.workers > 1 and len(work) > 1:
            with ProcessPoolExecutor(max_workers=args.workers) as pool:
                results = list(pool.map(docx_job, work))
        else:
            results = [docx_job(item) for item in work]
        failures = {roll: message for roll, message, _ in results if message}
        seal_outcome = {roll: outcome for roll, _, outcome in results}
        for job in to_write:
            job.seal = seal_outcome.get(job.student.roll, "")
        for roll, message in failures.items():
            info(f"  FAILED roll {roll}: {message}")

    pdf_failures: list[str] = []
    if make_pdf:
        pairs = [(j.docx_path, j.pdf_path) for j in jobs
                 if j.write_pdf and j.student.roll not in failures and j.docx_path.exists()]
        pdf_failures = export_pdfs(pairs, args.pdf_engine, args.soffice)
        for message in pdf_failures:
            info(f"  PDF problem: {message}")

    report_path = output_dir / REPORT_FILENAME
    write_report(report_path, jobs, marks if sections else None, failures,
                 warnings + [f"PDF problem: {m}" for m in pdf_failures],
                 config["columns"]["marks_obtained"])
    info(f"\nDone. Files are in: {output_dir}")
    if seal_path:
        info(f"Seal: stamped on {sum(1 for j in to_write if j.seal == 'stamped')} of {len(to_write)} sheet(s) "
             f"({seal_path.name}).")
        had = [j for j in to_write if j.seal == "already had a seal"]
        if had:
            report_problems("Sheets that already had a seal (not stamped again)", [j.docx_path.name for j in had])
    info(f"Run report: {report_path}")
    return 1 if failures or pdf_failures else 0


def main(argv: list[str] | None = None) -> int:
    try:
        return run(parse_args(argv))
    except UserError as error:
        info(f"\nERROR: {error}")
        return 1
    except KeyboardInterrupt:
        info("\nStopped.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
