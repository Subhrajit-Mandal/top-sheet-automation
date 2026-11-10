"""Tests for generate_top_sheets.py. Run with:  python -m unittest -v

The tests build their own small Word templates and workbooks, so they need
neither your real template nor any student data.
"""

from __future__ import annotations

import base64
import copy
import hashlib
import io
import random
import struct
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from openpyxl import Workbook

import generate_top_sheets as g


# --------------------------------------------------------------------------
# Fixtures
# --------------------------------------------------------------------------

def make_template(path: Path, questions: list[tuple[str, str]], columns: tuple[str, ...] = (
        "Q. No.", "Marks Allotted", "Marks Awarded", "Remarks"), full_marks: str = "25") -> Path:
    """A small top sheet: info table, marks table, signature line."""
    document = Document()
    info = document.add_table(rows=6, cols=2)
    labels = [
        ("Name of the Student:", 0, 0), ("Roll Number:", 1, 0), ("Mobile Number:", 2, 0),
        ("Date of Examination:", 3, 0), ("Paper (course) Code:", 4, 0), ("Marks obtained:", 5, 0),
    ]
    for text, row, col in labels:
        paragraph = info.cell(row, col).paragraphs[0]
        run = paragraph.add_run(text + " ")
        run.bold = True
        run.font.name = "Twentieth Century"
        run.font.size = Pt(10)
    info.cell(0, 1).paragraphs[0].add_run(f"Full Marks: {full_marks}")

    table = document.add_table(rows=1 + len(questions), cols=len(columns))
    for index, header in enumerate(columns):
        table.cell(0, index).text = header
    allotted_col = next(i for i, c in enumerate(columns) if "llot" in c)
    question_col = next(i for i, c in enumerate(columns) if c in ("Q. No.", "No"))
    for row, (label, allotted) in enumerate(questions, start=1):
        table.cell(row, question_col).text = label
        run = table.cell(row, allotted_col).paragraphs[0].add_run(allotted)
        run.font.name = "Twentieth Century"
        run.font.size = Pt(10)
    document.add_paragraph("Signature of the student with date")
    document.save(path)
    return path


def make_workbook(path: Path, rows: list[list], header=("Name", "UNI Roll number", "Mobile Number", "Marks Obtained")) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(list(header))
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


YOURS = [("1.a)", "1"), ("1.b)", "1"), ("1.c)", "1"), ("1.d)", "1"), ("1.e)", "1"), ("1.f)", "1"),
         ("1.g)", ""), ("2", "5"), ("3", "5"), ("4", "5"), ("5", "5"), ("6", "5"), ("7", "")]


def config() -> dict:
    cfg = copy.deepcopy(g.DEFAULT_CONFIG)
    cfg["template_fields"] = [
        f for f in cfg["template_fields"] if f["label"] in
        {"Name of the Student:", "Roll Number:", "Mobile Number:", "Marks obtained:",
         "Date of Examination:", "Paper (course) Code:"}
    ]
    return cfg


def read_marks(template: Path, cfg: dict) -> g.MarksTable:
    return g.find_marks_table(Document(template), cfg)[1]


def feeder(*replies):
    """A fake input(): returns the replies in order, then behaves like a closed stdin."""
    remaining = iter(replies)

    def fake_input(prompt=""):
        try:
            return next(remaining)
        except StopIteration:
            raise EOFError from None
    return fake_input


def quiet(function, *args, **kwargs):
    with redirect_stdout(io.StringIO()) as buffer:
        result = function(*args, **kwargs)
    return result, buffer.getvalue()


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.dir = Path(self._dir.name)

    def tearDown(self) -> None:
        self._dir.cleanup()


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

class TextHelpers(unittest.TestCase):
    def test_roll_normalisation(self):
        self.assertEqual(g.normalise_roll(" 34900725001 "), "34900725001")
        self.assertEqual(g.normalise_roll("34900725001.0"), "34900725001")
        self.assertEqual(g.normalise_roll("ME/2026/01"), "ME/2026/01")

    def test_filename_format_is_unchanged(self):
        self.assertEqual(g.clean_filename("34900730090_A"), "34900730090_A")
        self.assertEqual(g.clean_filename('349_A/B:C'), "349_A_B_C")

    def test_total_marks_parsing(self):
        cfg = config()
        self.assertEqual(g.parse_total_mark("17", cfg), 17)
        self.assertEqual(g.parse_total_mark("17.0", cfg), 17)
        self.assertIsNone(g.parse_total_mark("Ab", cfg))
        self.assertIsNone(g.parse_total_mark("AB", cfg))
        for bad in ("", "17.5", "x", "26", "-3"):
            with self.assertRaises(ValueError, msg=bad):
                g.parse_total_mark(bad, cfg)

    def test_full_marks_comes_from_config(self):
        cfg = config()
        cfg["full_marks"] = 30
        self.assertEqual(g.parse_total_mark("28", cfg), 28)


# --------------------------------------------------------------------------
# Roll number is the primary key
# --------------------------------------------------------------------------

def student(row, roll, name="A", mobile="9876543210", marks="10") -> g.Student:
    return g.Student(row, roll, name, {"Name": name, "UNI Roll number": roll,
                                       "Mobile Number": mobile, "Marks Obtained": marks})


class Duplicates(unittest.TestCase):
    def test_identical_rows_collapse(self):
        kept, warnings, conflicts = g.resolve_duplicates(
            [student(2, "1"), student(3, "1"), student(4, "2")], "stop")
        self.assertEqual([s.row for s in kept], [2, 4])
        self.assertEqual(conflicts, [])
        self.assertEqual(len(warnings), 1)

    def test_whitespace_and_case_do_not_make_a_conflict(self):
        kept, _, conflicts = g.resolve_duplicates(
            [student(2, "1", name="Ana  Roy"), student(3, "1", name="ana roy")], "stop")
        self.assertEqual(len(kept), 1)
        self.assertEqual(conflicts, [])

    def test_conflict_stops_and_lists_every_row(self):
        kept, _, conflicts = g.resolve_duplicates(
            [student(2, "1", marks="10"), student(3, "1", marks="12"), student(4, "2")], "stop")
        self.assertEqual([s.roll for s in kept], ["2"])
        self.assertEqual(len(conflicts), 1)
        self.assertIn("row 2", conflicts[0])
        self.assertIn("row 3", conflicts[0])

    def test_keep_first_and_keep_last(self):
        rows = [student(2, "1", marks="10"), student(3, "1", marks="12")]
        self.assertEqual(g.resolve_duplicates(rows, "keep-first")[0][0].row, 2)
        self.assertEqual(g.resolve_duplicates(rows, "keep-last")[0][0].row, 3)


# --------------------------------------------------------------------------
# The marks table is read from the template, not hard-coded
# --------------------------------------------------------------------------

class AdaptiveMarksTable(TempCase):
    def test_your_current_layout(self):
        template = make_template(self.dir / "t.docx", YOURS)
        marks = read_marks(template, config())
        self.assertEqual(len(marks.questions), 13)
        groups = g.detect_sections(marks.questions)
        self.assertEqual([len(x) for x in groups], [7, 6])
        sections = [g.Section([marks.questions[i].label for i in groups[0]], 5),
                    g.Section([marks.questions[i].label for i in groups[1]], 4)]
        self.assertEqual(g.section_limits(marks.questions, sections), (9, 25))

    def test_every_total_splits_correctly_on_your_layout(self):
        template = make_template(self.dir / "t.docx", YOURS)
        marks = read_marks(template, config())
        groups = g.detect_sections(marks.questions)
        sections = [g.Section([marks.questions[i].label for i in groups[0]], 5),
                    g.Section([marks.questions[i].label for i in groups[1]], 4)]
        caps = {q.label: q.allotted for q in marks.questions}
        for total in range(9, 26):
            for roll in ("1", "2", "34900730090", "abc"):
                split = g.allocate(total, marks, sections, roll)
                self.assertEqual(sum(split.values()), total)
                self.assertEqual(sum(1 for l in split if l.startswith("1.")), 5)
                self.assertEqual(sum(1 for l in split if not l.startswith("1.")), 4)
                for label, mark in split.items():
                    self.assertTrue(1 <= mark <= caps[label], (total, roll, label, mark))
                self.assertNotIn("1.g)", split)
                self.assertNotIn("7", split)

    def test_totals_outside_the_reachable_range_are_refused(self):
        template = make_template(self.dir / "t.docx", YOURS)
        marks = read_marks(template, config())
        groups = g.detect_sections(marks.questions)
        sections = [g.Section([marks.questions[i].label for i in groups[0]], 5),
                    g.Section([marks.questions[i].label for i in groups[1]], 4)]
        for total in (0, 8, 26):
            with self.assertRaises(ValueError):
                g.allocate(total, marks, sections, "1")

    def test_same_student_same_split(self):
        template = make_template(self.dir / "t.docx", YOURS)
        marks = read_marks(template, config())
        groups = g.detect_sections(marks.questions)
        sections = [g.Section([marks.questions[i].label for i in groups[0]], 5),
                    g.Section([marks.questions[i].label for i in groups[1]], 4)]
        self.assertEqual(g.allocate(17, marks, sections, "77"), g.allocate(17, marks, sections, "77"))

    def test_a_different_sheet_more_rows_fewer_columns(self):
        """Different row count, different caps, no Remarks/Bloom columns, columns reordered."""
        rows = ([(f"1.{c})", "2") for c in "abcdefgh"]            # 8 parts, 2 marks each, answer any 3
                + [(f"2.{c})", "3") for c in "abc"]               # 3 parts, 3 marks each, answer any 2
                + [("3", "4"), ("4", "4"), ("5", "4"), ("6", "4")]  # 4 long, answer any 2
                + [("Total", "")])                                # ignored row
        template = make_template(self.dir / "t.docx", rows,
                                 columns=("Marks Awarded", "Q. No.", "Marks Allotted"))
        cfg = config()
        marks = read_marks(template, cfg)
        self.assertEqual(len(marks.questions), 15)  # "Total" ignored
        groups = g.detect_sections(marks.questions)
        self.assertEqual([len(x) for x in groups], [8, 3, 4])
        sections = [g.Section([marks.questions[i].label for i in x], n)
                    for x, n in zip(groups, (3, 2, 2))]
        low, high = g.section_limits(marks.questions, sections)
        self.assertEqual((low, high), (7, 6 + 6 + 8))
        for total in range(low, high + 1):
            split = g.allocate(total, marks, sections, f"r{total}")
            self.assertEqual(sum(split.values()), total)

        # and the numbers really land in the right cells of the reordered table
        document = Document(template)
        split = g.allocate(15, marks, sections, "x")
        g.populate_marks_awarded(document, cfg, split, g.font_settings(cfg))
        table, found = g.find_marks_table(document, cfg)
        for question in found.questions:
            cell = table.rows[question.table_row].cells[found.awarded_col].text
            self.assertEqual(cell, str(split.get(question.label, "")))

    def test_blank_allotted_rows_never_receive_marks(self):
        template = make_template(self.dir / "t.docx", [("1", "5"), ("2", ""), ("3", "5")])
        marks = read_marks(template, config())
        sections = [g.Section(["1", "2", "3"], 2)]
        split = g.allocate(7, marks, sections, "q")
        self.assertEqual(set(split), {"1", "3"})

    def test_bad_allotted_cell_is_an_error(self):
        template = make_template(self.dir / "t.docx", [("1", "five")])
        with self.assertRaises(g.UserError):
            read_marks(template, config())

    def test_missing_marks_table_names_what_it_looked_for(self):
        template = make_template(self.dir / "t.docx", YOURS, columns=("No", "Allotted", "Awarded", "Remarks"))
        # headers differ from the config, so nothing is guessed
        with self.assertRaises(g.UserError) as caught:
            read_marks(template, config())
        self.assertIn("Marks Allotted", str(caught.exception))


# --------------------------------------------------------------------------
# Filling fields: formatting, fonts, no guessing
# --------------------------------------------------------------------------

class Fields(TempCase):
    def test_value_copies_the_labels_formatting(self):
        template = make_template(self.dir / "t.docx", YOURS)
        document = Document(template)
        paragraph = g.find_label_paragraph(document, "Name of the Student:")
        g.fill_field(paragraph, "Name of the Student:", "Ana Roy", g.font_settings(config()))
        runs = paragraph.runs
        self.assertEqual(runs[-1].text, " Ana Roy")
        self.assertTrue(runs[-1].bold)
        self.assertEqual(runs[-1].font.name, "Twentieth Century")
        self.assertEqual(runs[-1].font.size.pt, 10)
        self.assertEqual(paragraph.text.strip(), "Name of the Student: Ana Roy")

    def test_custom_font_changes_only_the_value(self):
        template = make_template(self.dir / "t.docx", YOURS)
        document = Document(template)
        cfg = config()
        cfg["font_mode"], cfg["custom_font"] = "custom", {"name": "Calibri", "size_pt": 11}
        paragraph = g.find_label_paragraph(document, "Roll Number:")
        g.fill_field(paragraph, "Roll Number:", "123", g.font_settings(cfg))
        label_run, value_run = paragraph.runs[0], paragraph.runs[-1]
        self.assertEqual(label_run.font.name, "Twentieth Century")
        self.assertEqual(value_run.font.name, "Calibri")
        self.assertEqual(value_run.font.size.pt, 11)
        self.assertTrue(value_run.bold)

    def test_refilling_replaces_existing_text_after_the_label(self):
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Year: ").bold = True
        paragraph.add_run("3rd")
        g.fill_field(paragraph, "Year:", "4th", g.font_settings(config()))
        self.assertEqual(paragraph.text.strip(), "Year: 4th")

    def test_missing_label_stops_with_its_name(self):
        template = make_template(self.dir / "t.docx", YOURS)
        cfg = config()
        cfg["template_fields"].append({"label": "Guardian Name:", "column": "Name"})
        with self.assertRaises(g.UserError) as caught:
            g.analyse_template(template, cfg)
        self.assertIn("Guardian Name:", str(caught.exception))

    def test_undeclared_labels_stay_exactly_as_in_the_template(self):
        template = make_template(self.dir / "t.docx", YOURS)
        document = Document(template)
        g.create_docx(template, student(2, "1", name="Ana"), None, self.dir / "o.docx", config())
        out = Document(self.dir / "o.docx")
        for paragraph in g.iter_template_paragraphs(out):
            if "Paper (course) Code:" in paragraph.text:
                self.assertEqual(paragraph.text.strip(), "Paper (course) Code:")  # fixed value "" -> untouched

    def test_label_found_inside_a_table_for_signature(self):
        document = Document()
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Signature of the student with date"
        g.add_student_signature(document, student(2, "1", name="Ana"), config())
        self.assertEqual(table.cell(0, 0).paragraphs[0].text.splitlines()[0], "Ana")

    def test_signature_font_algorithm_is_the_original_one(self):
        cfg = config()
        original_fonts = ("Brush Script MT", "Freestyle Script", "Lucida Handwriting", "Mistral", "Viner Hand ITC")
        self.assertEqual(tuple(cfg["signature"]["fonts"]), original_fonts)
        for roll in ("34900730090", "34900730105", "1"):
            seed = hashlib.sha256((roll + "-signature").encode("utf-8")).digest()
            expected = random.Random(seed).choice(original_fonts)
            document = Document()
            document.add_paragraph("Signature of the student with date")
            g.add_student_signature(document, student(2, roll, name="Ana"), cfg)
            self.assertEqual(document.paragraphs[0].runs[0].font.name, expected)


# --------------------------------------------------------------------------
# Workbook reading
# --------------------------------------------------------------------------

class Workbook_(TempCase):
    def test_header_mismatch_lists_what_was_found(self):
        path = make_workbook(self.dir / "w.xlsx", [["A", 1, "9876543210", 10]],
                             header=("Name", "Roll", "Mobile Number", "Marks Obtained"))
        with self.assertRaises(g.UserError) as caught:
            g.read_students(path, None, config())
        message = str(caught.exception)
        self.assertIn("UNI Roll number", message)
        self.assertIn("'Roll'", message)  # what it did find

    def test_messy_rows(self):
        path = make_workbook(self.dir / "w.xlsx", [
            ["A", 34900730090.0, "9876543210", 10],   # float roll
            ["B", " 34900730091 ", "9876543210", "Ab"],
            ["", 34900730092, "9876543210", 12],      # blank name -> skipped
            [None, None, None, None],                 # empty row
            ["C", 34900730093, None, 13],             # blank mobile
        ])
        students, notes = g.read_students(path, None, config())
        self.assertEqual([s.roll for s in students], ["34900730090", "34900730091", "34900730093"])
        self.assertTrue(any("Row 4" in n for n in notes))
        errors, warnings = g.validate_students(students, config())
        self.assertEqual(errors, [])
        self.assertIsNone(students[1].total)
        self.assertTrue(any("Mobile Number" in w and "blank" in w for w in warnings))

    def test_all_bad_marks_reported_together(self):
        path = make_workbook(self.dir / "w.xlsx", [
            ["A", 1, "9876543210", 17.5], ["B", 2, "9876543210", None], ["C", 3, "9876543210", 40]])
        students, _ = g.read_students(path, None, config())
        errors, _ = g.validate_students(students, config())
        self.assertEqual(len(errors), 3)

    def test_mobile_length_warning_is_grouped(self):
        path = make_workbook(self.dir / "w.xlsx", [["A", 1, "985412578", 10], ["B", 2, "985412579", 11]])
        students, _ = g.read_students(path, None, config())
        _, warnings = g.validate_students(students, config())
        self.assertEqual(len([w for w in warnings if "digits" in w]), 1)


# --------------------------------------------------------------------------
# Collisions and the whole run
# --------------------------------------------------------------------------

class EndToEnd(TempCase):
    def setUp(self):
        super().setUp()
        self.template = make_template(self.dir / "t.docx", YOURS)
        self.rows = [["A", 1, "9876543210", 10], ["B", 2, "9876543210", "Ab"], ["C", 3, "9876543210", 25]]
        self.workbook = make_workbook(self.dir / "w.xlsx", self.rows)
        self.out = self.dir / "out"
        self.cfg = self.dir / "config.json"
        self.cfg.write_text('{"template_fields": [{"label": "Name of the Student:", "column": "Name"},'
                            '{"label": "Roll Number:", "column": "UNI Roll number"},'
                            '{"label": "Mobile Number:", "column": "Mobile Number"},'
                            '{"label": "Marks obtained:", "column": "Marks Obtained"},'
                            '{"label": "Date of Examination:", "value": "{exam_date}"},'
                            '{"label": "Paper (course) Code:", "value": ""}]}')

    def go(self, *extra, answers=("", "5", "4", ""), screen=False):
        argv = ["--workbook", str(self.workbook), "--template", str(self.template),
                "--output", str(self.out), "--config", str(self.cfg), "--keep-docx-only",
                "--workers", "1", *extra]
        if not screen:
            argv.append("--no-prompt")
        return quiet(g.run, g.parse_args(argv), feeder(*answers))

    def sheet_text(self, name="1_A.docx") -> str:
        document = Document(self.out / "docx" / name)
        return "\n".join(p.text for p in g.iter_template_paragraphs(document))

    def test_full_run_and_saved_rules(self):
        (code, text) = self.go()
        self.assertEqual(code, 0, text)
        files = sorted(p.name for p in (self.out / "docx").glob("*.docx"))
        self.assertEqual(files, ["1_A.docx", "2_B.docx", "3_C.docx"])
        self.assertTrue((self.dir / g.RULES_FILENAME).exists())
        self.assertTrue((self.out / g.REPORT_FILENAME).exists())

        # absent student: no question-wise marks at all, marks field says Ab
        absent = Document(self.out / "docx" / "2_B.docx")
        text_ = "\n".join(p.text for p in g.iter_template_paragraphs(absent))
        self.assertIn("Marks obtained: Ab", text_)

        # second run: saved rules are used, nobody is asked anything
        (code, text) = self.go(answers=())
        self.assertEqual(code, 0, text)
        self.assertIn("Using the section rules saved", text)

    def test_conflicting_duplicates_stop_everything(self):
        make_workbook(self.workbook, self.rows + [["C2", 3, "9876543210", 20]])
        (code, text) = self.go()
        self.assertEqual(code, 1)
        self.assertIn("Roll 3 appears in rows", text)
        self.assertFalse((self.out / "docx").exists())

    def test_identical_duplicate_is_merged_silently_into_one_sheet(self):
        make_workbook(self.workbook, self.rows + [["C", 3, "9876543210", 25]])
        (code, _) = self.go()
        self.assertEqual(code, 0)
        self.assertEqual(len(list((self.out / "docx").glob("*.docx"))), 3)

    def test_unreachable_total_stops_and_names_the_student(self):
        make_workbook(self.workbook, self.rows + [["D", 4, "9876543210", 5]])
        (code, text) = self.go()
        self.assertEqual(code, 1)
        self.assertIn("roll 4", text)
        self.assertIn("9 to 25", text)

    def test_unreachable_total_can_be_skipped_by_config(self):
        make_workbook(self.workbook, self.rows + [["D", 4, "9876543210", 5]])
        (code, text) = self.go("--answer", "5,4", answers=())
        self.assertEqual(code, 1)  # still stops: policy is 'stop'
        cfg = g.json.loads(self.cfg.read_text())
        cfg["unreachable_marks"] = "skip"
        self.cfg.write_text(g.json.dumps(cfg))
        (code, text) = self.go(answers=())
        self.assertEqual(code, 0, text)
        self.assertFalse((self.out / "docx" / "4_D.docx").exists())

    def test_existing_files_overwrite_or_skip(self):
        self.go()
        target = self.out / "docx" / "1_A.docx"
        target.write_bytes(b"old")
        self.go("--existing", "skip", answers=())
        self.assertEqual(target.read_bytes(), b"old")
        self.go(answers=())
        self.assertNotEqual(target.read_bytes(), b"old")

    def test_check_mode_writes_nothing(self):
        (code, _) = self.go("--check")
        self.assertEqual(code, 0)
        self.assertFalse(self.out.exists())
        self.assertFalse((self.dir / g.RULES_FILENAME).exists())

    def test_exam_date_and_font_overrides(self):
        self.go("--exam-date", "01-01-2027", "--font-mode", "custom", "--custom-font", "Calibri")
        document = Document(self.out / "docx" / "1_A.docx")
        text = "\n".join(p.text for p in g.iter_template_paragraphs(document))
        self.assertIn("Date of Examination: 01-01-2027", text)
        self.assertIn("01-01-2027", document.paragraphs[0].text + "".join(p.text for p in document.paragraphs))

    def test_unknown_config_key_is_rejected(self):
        self.cfg.write_text('{"exam_dat": "x"}')
        with self.assertRaises(g.UserError) as caught:
            g.load_config(self.cfg)
        self.assertIn("exam_dat", str(caught.exception))

    # ---- settings screen, end to end -------------------------------------
    # Answer order for a first run: screen replies..., [save?], then the section
    # questions: accept grouping "", 5, 4, save section rules "".

    def test_screen_answers_reach_the_sheets_and_are_saved(self):
        code, text = self.go(screen=True, answers=("1", "5-9-2026", "3", "ME-301", "", "",
                                                   "", "5", "4", ""))
        self.assertEqual(code, 0, text)
        sheet = self.sheet_text()
        self.assertIn("Date of Examination: 05-09-2026", sheet)
        self.assertIn("Paper (course) Code: ME-301", sheet)
        self.assertEqual(sheet.count("05-09-2026"), 2)  # the field and under the signature

        local = g.json.loads((self.dir / "config.local.json").read_text())
        self.assertEqual(local["exam_date"], "05-09-2026")
        self.assertEqual(local["fixed_values"], {"Paper (course) Code:": "ME-301"})
        # the shipped file is never written to
        self.assertNotIn("exam_date", g.json.loads(self.cfg.read_text()))

        # next run: saved values are used, no screen needed
        self.go(answers=())
        self.assertIn("Date of Examination: 05-09-2026", self.sheet_text())

    def test_declining_to_save_applies_to_this_run_only(self):
        code, _ = self.go(screen=True, answers=("1", "05-09-2026", "", "n", "", "5", "4", ""))
        self.assertEqual(code, 0)
        self.assertIn("Date of Examination: 05-09-2026", self.sheet_text())
        self.assertFalse((self.dir / "config.local.json").exists())

    def test_no_changes_means_nothing_to_save(self):
        code, text = self.go(screen=True, answers=("", "", "5", "4", ""))
        self.assertEqual(code, 0, text)
        self.assertNotIn("Save these changes", text)
        self.assertFalse((self.dir / "config.local.json").exists())

    def test_no_one_answering_continues_with_the_settings_shown(self):
        code, text = self.go("--answer", "5,4", screen=True, answers=())
        self.assertEqual(code, 0, text)
        self.assertIn("No one is answering", text)

    def test_no_prompt_skips_the_screen(self):
        _, text = self.go()
        self.assertNotIn("Settings for this run", text)

    def test_screen_is_skipped_in_check_mode(self):
        _, text = self.go("--check", screen=True, answers=())
        self.assertNotIn("Settings for this run", text)

    def test_blank_exam_date_is_a_warning_not_a_guess(self):
        code, text = self.go()
        self.assertEqual(code, 0, text)
        self.assertIn("exam date is empty", text)
        self.assertEqual(
            [l for l in self.sheet_text().splitlines() if "Date of Examination" in l][0].strip(),
            "Date of Examination:")

    def test_template_full_marks_is_offered_as_the_default(self):
        self.template = make_template(self.dir / "t30.docx", YOURS, full_marks="30")
        _, text = self.go(screen=True, answers=("", "", "5", "4", ""))
        self.assertIn("Full marks 30 was read from the template", text)

    def test_an_explicit_full_marks_wins_over_the_template(self):
        self.template = make_template(self.dir / "t30.docx", YOURS, full_marks="30")
        (self.dir / "config.local.json").write_text('{"full_marks": 25}')
        _, text = self.go(screen=True, answers=("", "", "5", "4", ""))
        self.assertNotIn("was read from the template", text)
        self.assertIn("the template itself says Full Marks: 30", text)

    # ---- official seal, generate mode ------------------------------------

    def seal_file(self) -> Path:
        path = self.dir / "seal.png"
        path.write_bytes(tiny_png())
        return path

    def test_the_seal_is_stamped_on_every_sheet(self):
        code, text = self.go("--seal", str(self.seal_file()))
        self.assertEqual(code, 0, text)
        sheets = sorted((self.out / "docx").glob("*.docx"))
        self.assertEqual(len(sheets), 3)
        for sheet in sheets:
            self.assertEqual(len(seal_anchors(Document(sheet))), 1, sheet.name)
        self.assertIn("Seal: stamped on 3 of 3 sheet(s) (seal.png)", text)
        self.assertEqual({r["Seal"] for r in self.report() if r["Status"] == "written"}, {"stamped"})

    def test_without_a_seal_nothing_is_stamped_and_nothing_is_said(self):
        code, text = self.go()
        for sheet in (self.out / "docx").glob("*.docx"):
            self.assertEqual(seal_anchors(Document(sheet)), [])
        self.assertNotIn("Seal:", text)
        self.assertEqual({r["Seal"] for r in self.report()}, {""})

    def test_answering_yes_then_choosing_a_file_stamps_every_sheet(self):
        with mock.patch.object(g, "choose_input_file", return_value=self.seal_file()):
            code, text = self.go(screen=True, answers=("", "", "5", "4", "", "y"))
        self.assertEqual(code, 0, text)
        for sheet in (self.out / "docx").glob("*.docx"):
            self.assertEqual(len(seal_anchors(Document(sheet))), 1)

    def test_answering_no_stamps_nothing(self):
        with mock.patch.object(g, "choose_input_file") as picker:
            code, _ = self.go(screen=True, answers=("", "", "5", "4", "", "n"))
        picker.assert_not_called()
        for sheet in (self.out / "docx").glob("*.docx"):
            self.assertEqual(seal_anchors(Document(sheet)), [])

    def test_a_bad_seal_file_stops_before_anything_is_written(self):
        bad = self.dir / "seal.png"
        bad.write_text("not a picture")
        with self.assertRaises(g.UserError) as caught:
            self.go("--seal", str(bad))
        self.assertIn("could not be read as a picture", str(caught.exception))
        self.assertFalse((self.out / "docx").exists())

    def test_check_mode_looks_at_the_seal_but_writes_nothing(self):
        code, text = self.go("--check", "--seal", str(self.seal_file()))
        self.assertEqual(code, 0, text)
        self.assertIn("seal.png can be used", text)
        self.assertFalse(self.out.exists())

    def test_a_template_that_already_has_a_seal_is_not_stamped_twice(self):
        template = Document(self.template)
        template.add_picture(io.BytesIO(tiny_png()))
        next(template.element.body.iter(qn("wp:docPr"))).set("name", "College Seal")
        template.save(self.template)
        code, text = self.go("--seal", str(self.seal_file()))
        self.assertEqual(code, 0, text)
        self.assertIn("already had a seal", text)
        for sheet in (self.out / "docx").glob("*.docx"):
            self.assertEqual(seal_anchors(Document(sheet)), [])

    def report(self):
        with (self.out / g.REPORT_FILENAME).open(encoding="utf-8-sig", newline="") as handle:
            return list(g.csv.DictReader(handle))

    def test_shipped_config_matches_the_code_defaults(self):
        shipped = g.json.loads((g.HERE / "config.json").read_text(encoding="utf-8-sig"))
        self.assertEqual(shipped, g.DEFAULT_CONFIG)


class SettingsScreen(unittest.TestCase):
    def run_screen(self, *replies, cfg=None, **kwargs):
        cfg = cfg or config()
        changes: dict = {}
        quiet(g.review_settings, cfg, changes, input_fn=feeder(*replies), **kwargs)
        return cfg, changes

    def test_lists_only_what_a_person_would_type(self):
        titles = [title for _, title, _ in g.settings_rows(config())]
        self.assertEqual(titles, ["Exam date", "Full marks", "Paper (course) Code", "Font mode"])
        # derived or column-fed labels are not offered

    def test_exam_date_is_normalised(self):
        cfg, changes = self.run_screen("1", "2-9-2026", "")
        self.assertEqual(cfg["exam_date"], "02-09-2026")
        self.assertEqual(changes, {"exam_date": "02-09-2026"})

    def test_an_odd_date_needs_confirmation(self):
        cfg, _ = self.run_screen("1", "Sept 2nd", "n", "")
        self.assertEqual(cfg["exam_date"], "")
        cfg, changes = self.run_screen("1", "Sept 2nd", "y", "")
        self.assertEqual(cfg["exam_date"], "Sept 2nd")

    def test_enter_keeps_and_dash_clears(self):
        cfg = config()
        cfg["exam_date"] = "01-01-2027"
        cfg, changes = self.run_screen("1", "", "", cfg=cfg)
        self.assertEqual(cfg["exam_date"], "01-01-2027")
        self.assertEqual(changes, {})
        cfg, changes = self.run_screen("1", "-", "", cfg=cfg)
        self.assertEqual(cfg["exam_date"], "")

    def test_full_marks_must_be_a_whole_number(self):
        cfg, changes = self.run_screen("2", "abc", "2", "0", "2", "30", "")
        self.assertEqual(cfg["full_marks"], 30)
        self.assertEqual(changes, {"full_marks": 30})

    def test_fixed_field_updates_the_template_entry_and_the_overlay(self):
        cfg, changes = self.run_screen("3", "ME-301", "")
        entry = next(e for e in cfg["template_fields"] if e["label"] == "Paper (course) Code:")
        self.assertEqual(entry["value"], "ME-301")
        self.assertEqual(changes, {"fixed_values": {"Paper (course) Code:": "ME-301"}})

    def test_font_mode_custom_asks_for_the_name(self):
        cfg, changes = self.run_screen("4", "2", "Arial", "")
        self.assertEqual((cfg["font_mode"], cfg["custom_font"]["name"]), ("custom", "Arial"))
        self.assertEqual(changes, {"font_mode": "custom", "custom_font": {"name": "Arial"}})
        cfg, changes = self.run_screen("4", "1", "", cfg=cfg)
        self.assertEqual(cfg["font_mode"], "template")

    def test_a_wrong_menu_choice_is_not_fatal(self):
        cfg, changes = self.run_screen("9", "x", "")
        self.assertEqual(changes, {})

    def test_closing_input_mid_way_raises_for_the_caller_to_handle(self):
        with self.assertRaises(EOFError):
            quiet(g.review_settings, config(), {}, input_fn=feeder())


# --------------------------------------------------------------------------
# Sheets that students have filled in (value in the cell next to the label)
# --------------------------------------------------------------------------

CA2_QUESTIONS = [("1.i)", "1"), ("1.ii)", "1"), ("1.iii)", "1"), ("1.iv)", "1"), ("1v)", "1"),
                 ("2", "5"), ("3", "5"), ("4", "5"), ("5", "5")]


def tiny_png() -> bytes:
    import struct
    import zlib

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b""))


def make_ca2(path: Path, roll: str, name: str, mobile: str = "9876543210", *, teacher: str = "Dr X",
             marks: str = "", picture: bool = True, questions=CA2_QUESTIONS) -> Path:
    """A student-filled sheet: every label has its value in the cell to its right."""
    document = Document()
    rows = [("Name of the Student:", name, "Roll Number:", roll),
            ("UPID:", "", "Date of Examination:", ""),
            ("Paper (course) Code:", "", "Mobile Number:", mobile),
            ("Subject Teacher:", teacher, "Full Marks:", "25"),
            ("Marks Obtained:", marks, "", "")]
    table = document.add_table(rows=len(rows), cols=4)
    for r, values in enumerate(rows):
        for c, text in enumerate(values):
            if text:
                run = table.cell(r, c).paragraphs[0].add_run(text)
                run.bold = c % 2 == 0
                run.font.name = "Trebuchet MS"
    marks_table = document.add_table(rows=1 + len(questions), cols=4)
    for c, header in enumerate(("Q. No.", "Marks Allotted", "Marks Awarded", "Course Outcome")):
        marks_table.cell(0, c).text = header
    for r, (label, allotted) in enumerate(questions, start=1):
        marks_table.cell(r, 0).text = label
        marks_table.cell(r, 1).text = allotted
        marks_table.cell(r, 1).paragraphs[0].alignment = 1  # centred, like the real sheet
    if picture:
        document.add_picture(io.BytesIO(tiny_png()))
    document.add_paragraph("Signature of the student with date")
    document.save(path)
    return path


def zero_checksum(path: Path, member: str, *, store: bool = False) -> Path:
    """Rewrite a docx so that one member carries a checksum of 0, as some apps save pictures."""
    source = zipfile.ZipFile(path)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            kind = zipfile.ZIP_STORED if (store and info.filename == member) else zipfile.ZIP_DEFLATED
            target.writestr(info.filename, source.read(info.filename), compress_type=kind)
    raw = bytearray(buffer.getvalue())
    info = zipfile.ZipFile(io.BytesIO(bytes(raw))).getinfo(member)
    raw[info.header_offset + 14:info.header_offset + 18] = b"\0\0\0\0"           # local header
    position = 0
    while True:                                                                    # central directory
        position = bytes(raw).find(b"PK\x01\x02", position)
        if position < 0:
            break
        name_length = struct.unpack("<H", raw[position + 28:position + 30])[0]
        if bytes(raw[position + 46:position + 46 + name_length]).decode() == member:
            raw[position + 16:position + 20] = b"\0\0\0\0"
        position += 4
    path.write_bytes(bytes(raw))
    return path


def ca2_value(path: Path, label: str) -> str:
    document = Document(path)
    return g.read_field(g.find_label_paragraph(document, label), label, "next-cell")


def awarded_marks(path: Path) -> dict[str, str]:
    table, marks = g.find_marks_table(Document(path), config())
    return {q.label: table.rows[q.table_row].cells[marks.awarded_col].text
            for q in marks.questions if table.rows[q.table_row].cells[marks.awarded_col].text}


def ca2_config_text() -> str:
    return (g.HERE / "presets" / "ca2.json").read_text(encoding="utf-8-sig")


STUDENTS = [["Tania Dandapat", 34900725002, "9434310339", 23],
            ["Sarat Roy", "LME2026410", "9434310338", 17],
            ["Mita Das", 34900725003, "9434310340", "Ab"],
            ["Rohan Sen", 34900725004, "9434310341", 25]]


class Submissions(TempCase):
    def setUp(self):
        super().setUp()
        if not (g.HERE / "presets" / "ca2.json").exists():
            self.skipTest("no presets/ca2.json beside the script")
        self.subs = self.dir / "subs"
        self.subs.mkdir()
        self.workbook = make_workbook(self.dir / "w.xlsx", STUDENTS)
        self.config = self.dir / "ca2.json"
        self.config.write_text(ca2_config_text())
        self.out = self.dir / "out"

    def sheet(self, name, roll, student, **kwargs):
        return make_ca2(self.subs / name, roll, student, **kwargs)

    def go(self, *extra, answers=(), answer="5,4"):
        argv = ["--workbook", str(self.workbook), "--submissions", str(self.subs), "--config",
                str(self.config), "--output", str(self.out), "--keep-docx-only", "--no-prompt", *extra]
        if answer:
            argv += ["--answer", answer]

        def like_main(args, input_fn):  # what main() does with a problem the user can fix
            try:
                return g.run(args, input_fn)
            except g.UserError as error:
                print(f"\nERROR: {error}")
                return 1
        return quiet(like_main, g.parse_args(argv), feeder(*answers))

    def report(self):
        with (self.out / g.REPORT_FILENAME).open(encoding="utf-8-sig", newline="") as handle:
            return list(g.csv.DictReader(handle))

    # ---- official seal, submissions mode ---------------------------------

    def seal_file(self) -> Path:
        path = self.dir / "seal.png"
        path.write_bytes(tiny_png() + b"\0")   # not byte-identical to the students' own picture
        return path

    def test_the_seal_goes_on_completed_sheets_only(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        self.sheet("stranger.docx", "99999999999", "STRANGER")
        code, text = self.go("--seal", str(self.seal_file()))
        self.assertEqual(code, 1)    # the stranger's sheet is skipped
        done = sorted((self.out / "docx").glob("*.docx"))
        self.assertEqual(len(done), 1)
        self.assertEqual(len(seal_anchors(Document(done[0]))), 1)
        self.assertIn("Seal: stamped on 1 of 1 sheet(s) (seal.png)", text)
        seals = {r["Submitted file"]: r["Seal"] for r in self.report()}
        self.assertEqual(seals["t.docx"], "stamped")
        self.assertEqual(seals["stranger.docx"], "")

    def test_the_students_own_pictures_survive_next_to_the_seal(self):
        source = self.sheet("t.docx", "34900725002", "TANIA")
        self.go("--seal", str(self.seal_file()))
        done = Document(self.out / "docx" / "34900725002_Tania Dandapat.docx")
        blobs = [part.blob for part in done.part.package.iter_parts() if part.partname.startswith("/word/media/")]
        self.assertEqual(len(blobs), 2)                       # the student's picture and the seal
        import zipfile
        original = [zipfile.ZipFile(source).read(n) for n in zipfile.ZipFile(source).namelist()
                    if n.startswith("word/media/")]
        self.assertTrue(all(blob in blobs for blob in original))

    def test_a_submitted_sheet_that_already_has_a_seal_is_not_stamped_again(self):
        path = self.sheet("t.docx", "34900725002", "TANIA")
        document = Document(path)
        next(document.element.body.iter(qn("wp:docPr"))).set("name", "College Seal")
        document.save(path)
        code, text = self.go("--seal", str(self.seal_file()))
        self.assertEqual(code, 0, text)
        self.assertIn("already had a seal", text)
        done = Document(self.out / "docx" / "34900725002_Tania Dandapat.docx")
        self.assertEqual(seal_anchors(done), [])

    def test_answering_yes_stamps_the_completed_sheets(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        argv = ["--workbook", str(self.workbook), "--submissions", str(self.subs), "--config",
                str(self.config), "--output", str(self.out), "--keep-docx-only", "--answer", "5,4"]
        with mock.patch.object(g, "choose_input_file", return_value=self.seal_file()):
            code, text = quiet(g.run, g.parse_args(argv + ["--exam-date", "01-01-2027"]),
                               feeder("", "y"))      # Enter on the settings screen, then yes to the seal
        self.assertEqual(code, 0, text)
        done = Document(self.out / "docx" / "34900725002_Tania Dandapat.docx")
        self.assertEqual(len(seal_anchors(done)), 1)

    def test_check_mode_looks_at_the_seal(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        code, text = self.go("--check", "--seal", str(self.seal_file()))
        self.assertEqual(code, 0, text)
        self.assertIn("seal.png can be used", text)
        self.assertFalse(self.out.exists())

    def test_check_mode_stops_on_an_unusable_seal(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        bad = self.dir / "bad.png"
        bad.write_text("x")
        code, text = self.go("--check", "--seal", str(bad))
        self.assertEqual(code, 1)
        self.assertIn("could not be read as a picture", text)

    # ---- damaged pictures are recovered ----------------------------------

    def damaged(self, name, roll, student, picture, member, store=False):
        path = make_ca2(self.subs / name, roll, student, picture=False)
        document = Document(path)
        document.add_picture(io.BytesIO(picture))
        document.save(path)
        return zero_checksum(path, member, store=store)

    def test_a_sheet_with_a_zero_checksum_picture_is_completed(self):
        path = self.damaged("ananya.docx", "34900725002", "TANIA", TINY_JPEG, "word/media/image1.jpg", store=True)
        with self.assertRaises(zipfile.BadZipFile):      # this is the error python-docx gives on its own
            Document(path)
        code, text = self.go()
        self.assertEqual(code, 0, text)
        done = self.out / "docx" / "34900725002_Tania Dandapat.docx"
        self.assertEqual(ca2_value(done, "Marks Obtained:"), "23")
        again = Document(done)                           # the output opens normally
        pictures = [part.blob for part in again.part.package.iter_parts() if part.partname.startswith("/word/media/")]
        self.assertEqual(pictures, [TINY_JPEG])          # the picture is exactly the student's
        self.assertEqual(zipfile.ZipFile(done).testzip(), None)
        self.assertIn("recovering a picture with a wrong checksum", text)
        self.assertTrue(any("wrong checksum" in r["Notes"] for r in self.report()))

    def test_a_damaged_text_part_is_still_skipped(self):
        path = make_ca2(self.subs / "x.docx", "34900725002", "TANIA", picture=False)
        zero_checksum(path, "word/document.xml")
        code, text = self.go()
        self.assertEqual(code, 1)
        self.assertIn("is damaged and is not a picture", text)
        self.assertFalse((self.out / "docx").glob("*.docx").__next__ if False else False)

    # ---- the report tells apart "sent nothing" from "sent something unusable" ----

    def test_a_file_that_cannot_be_used_is_not_reported_as_no_submission(self):
        (self.subs / "34900725004_ANANYA_PAL.docx").write_bytes(b"not a word file")
        self.sheet("t.docx", "34900725002", "TANIA")
        code, text = self.go()
        self.assertEqual(code, 1)
        statuses = {r["Roll"]: r for r in self.report() if r["Roll"]}
        self.assertEqual(statuses["34900725004"]["Status"], "FILE RECEIVED, NOT USED")
        self.assertIn("34900725004_ANANYA_PAL.docx", statuses["34900725004"]["Notes"])
        self.assertEqual(statuses["34900725003"]["Status"], "NO SUBMISSION")
        self.assertIn("Students who sent a file that could not be used", text)
        listed = text.split("Students in the workbook with no sheet")[1]
        self.assertNotIn("34900725004", listed)
        self.assertIn("34900725003", listed)

    def test_the_file_name_is_only_a_label_never_used_for_matching(self):
        self.sheet("34900725003_WRONG.docx", "34900725002", "TANIA")   # name says 003, the sheet says 002
        self.go()
        statuses = {r["Roll"]: r["Status"] for r in self.report() if r["Roll"]}
        self.assertEqual(statuses["34900725002"], "completed")
        self.assertEqual(statuses["34900725003"], "NO SUBMISSION")

    def test_unknown_rolls_come_with_a_tip(self):
        self.sheet("x.docx", "99999999999", "STRANGER")
        _, text = self.go()
        self.assertIn("run again with --existing skip", text)

    # ---- sheets for students who sent nothing -----------------------------

    def template(self):
        return make_ca2(self.dir / "template.docx", "", "", mobile="", teacher="", picture=False)

    def test_students_who_sent_nothing_get_a_sheet_from_the_template(self):
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        code, text = self.go("--template", str(self.template()))
        self.assertEqual(code, 0, text)
        names = sorted(p.name for p in (self.out / "docx").glob("*.docx"))
        self.assertEqual(names, ["34900725002_Tania Dandapat.docx", "34900725003_Mita Das.docx",
                                 "34900725004_Rohan Sen.docx", "LME2026410_Sarat Roy.docx"][:0] or names)
        self.assertEqual(len(names), 4)
        made = self.out / "docx" / "LME2026410_Sarat Roy.docx"
        self.assertEqual(ca2_value(made, "Name of the Student:"), "Sarat Roy")
        self.assertEqual(ca2_value(made, "Roll Number:"), "LME2026410")
        self.assertEqual(ca2_value(made, "Mobile Number:"), "9434310338")
        self.assertEqual(ca2_value(made, "Marks Obtained:"), "17")
        self.assertEqual(sum(int(v) for v in awarded_marks(made).values()), 17)
        # the typed (generated) signature, in one of the configured handwriting fonts
        document = Document(made)
        texts = [p.text for p in document.paragraphs]
        index = texts.index("Signature of the student with date")
        signature = document.paragraphs[index - 1]
        self.assertEqual(signature.runs[0].text, "Sarat Roy")
        self.assertIn(signature.runs[0].font.name, g.DEFAULT_CONFIG["signature"]["fonts"])

    def test_the_run_report_records_which_sheets_were_generated(self):
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        _, text = self.go("--template", str(self.template()))
        rows = {r["Roll"]: r for r in self.report() if r["Roll"]}
        self.assertEqual(rows["34900725002"]["Status"], "completed")
        for roll in ("LME2026410", "34900725003", "34900725004"):
            self.assertEqual(rows[roll]["Status"], "GENERATED (typed signature)", roll)
            self.assertIn("no sheet was submitted", rows[roll]["Notes"])
            self.assertEqual(rows[roll]["Submitted file"], "")
        self.assertEqual(rows["LME2026410"]["Question-wise marks"].count("="), 9)
        self.assertEqual(rows["34900725003"]["Question-wise marks"], "")          # absent
        self.assertIn("Made from the template for students who sent nothing: 3 (typed signature)", text)
        self.assertIn("Sheets made from the template (no sheet was submitted)", text)

    def test_without_a_template_nothing_is_generated(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        _, text = self.go()
        self.assertEqual(len(list((self.out / "docx").glob("*.docx"))), 1)
        self.assertNotIn("Made from the template", text)

    def test_students_with_an_unusable_file_are_not_replaced_by_a_typed_signature(self):
        (self.subs / "LME2026410_SARAT.docx").write_bytes(b"broken")
        self.sheet("t.docx", "34900725002", "TANIA")
        self.go("--template", str(self.template()))
        self.assertFalse((self.out / "docx" / "LME2026410_Sarat Roy.docx").exists())
        statuses = {r["Roll"]: r["Status"] for r in self.report() if r["Roll"]}
        self.assertEqual(statuses["LME2026410"], "FILE RECEIVED, NOT USED")
        self.assertEqual(statuses["34900725004"], "GENERATED (typed signature)")

    def test_a_total_too_low_to_split_is_listed_not_generated(self):
        make_workbook(self.workbook, STUDENTS + [["Low Scorer", 34900725005, "9434310342", 5]])
        self.sheet("t.docx", "34900725002", "TANIA")
        code, text = self.go("--template", str(self.template()))
        self.assertFalse((self.out / "docx" / "34900725005_Low Scorer.docx").exists())
        row = {r["Roll"]: r for r in self.report() if r["Roll"]}["34900725005"]
        self.assertEqual(row["Status"], "NO SUBMISSION")
        self.assertIn("not made from the template", row["Notes"])
        self.assertIn("34900725005", text.split("Students in the workbook with no sheet")[1])

    def test_generated_sheets_carry_the_seal_too(self):
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        seal = self.dir / "seal.png"
        seal.write_bytes(tiny_png() + b"\0")
        code, text = self.go("--template", str(self.template()), "--seal", str(seal))
        self.assertEqual(code, 0, text)
        for sheet in (self.out / "docx").glob("*.docx"):
            self.assertEqual(len(seal_anchors(Document(sheet))), 1, sheet.name)
        self.assertIn("Seal: stamped on 4 of 4 sheet(s)", text)

    def test_with_the_signature_switched_off_the_report_says_so(self):
        cfg = g.json.loads(self.config.read_text())
        cfg["signature"] = {"enabled": False}
        self.config.write_text(g.json.dumps(cfg))
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        self.go("--template", str(self.template()))
        row = {r["Roll"]: r for r in self.report() if r["Roll"]}["LME2026410"]
        self.assertEqual(row["Status"], "GENERATED (no signature)")
        self.assertIn("without a signature", row["Notes"])

    def test_kept_existing_generated_sheets_are_not_made_twice(self):
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        template = str(self.template())
        self.go("--template", template)
        target = self.out / "docx" / "LME2026410_Sarat Roy.docx"
        target.write_bytes(b"earlier")
        self.go("--template", template, "--existing", "skip")
        self.assertEqual(target.read_bytes(), b"earlier")
        row = {r["Roll"]: r for r in self.report() if r["Roll"]}["LME2026410"]
        self.assertEqual(row["Status"], "kept existing")

    def test_a_template_that_does_not_fit_only_costs_a_warning(self):
        blank = Document()
        blank.add_paragraph("nothing the config knows about")
        path = self.dir / "other.docx"
        blank.save(path)
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        code, text = self.go("--template", str(path))
        self.assertEqual(code, 0, text)
        self.assertIn("Sheets could not be made from the template", text)
        self.assertTrue((self.out / "docx" / "34900725002_Tania Dandapat.docx").exists())
        self.assertEqual(len(list((self.out / "docx").glob("*.docx"))), 1)

    def test_check_mode_counts_what_would_be_generated(self):
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT")
        code, text = self.go("--check", "--template", str(self.template()))
        self.assertEqual(code, 0, text)
        self.assertIn("3 would be made from the template", text)
        self.assertFalse(self.out.exists())

    def test_the_preset_is_a_valid_config(self):
        cfg = g.load_config(self.config)
        self.assertTrue(all(e.get("where") == "next-cell" for e in cfg["template_fields"]))

    def test_a_sheet_is_completed_and_nothing_else_changes(self):
        source = self.sheet("tania.docx", "34900725002", "TANIA DANDAPAT")
        code, text = self.go()
        self.assertEqual(code, 0, text)
        done = self.out / "docx" / "34900725002_Tania Dandapat.docx"  # file name from the workbook
        self.assertEqual(ca2_value(done, "Marks Obtained:"), "23")
        self.assertEqual(ca2_value(done, "Name of the Student:"), "TANIA DANDAPAT")
        split = awarded_marks(done)
        self.assertEqual(sum(int(v) for v in split.values()), 23)
        self.assertEqual(len(split), 9)
        self.assertTrue(all(1 <= int(v) <= (1 if k.startswith("1") else 5) for k, v in split.items()))
        # the student's picture is carried over byte for byte
        import zipfile
        before, after = zipfile.ZipFile(source), zipfile.ZipFile(done)
        images = [n for n in before.namelist() if n.startswith("word/media/")]
        self.assertTrue(images)
        for name in images:
            self.assertEqual(before.read(name), after.read(name))

    def test_awarded_marks_follow_the_allotted_columns_alignment(self):
        self.sheet("t.docx", "34900725002", "T")
        self.go()
        done = self.out / "docx" / "34900725002_Tania Dandapat.docx"
        table, marks = g.find_marks_table(Document(done), config())
        row = table.rows[marks.questions[0].table_row]
        self.assertEqual(row.cells[marks.awarded_col].paragraphs[0].alignment, 1)

    def test_what_the_student_wrote_is_kept_and_differences_are_listed(self):
        self.sheet("sarat.docx", " lme 2026410 ", "SARAT ROY", mobile="9434310339")  # spaces, case
        code, text = self.go()
        self.assertEqual(code, 0, text)
        done = self.out / "docx" / "LME2026410_Sarat Roy.docx"
        self.assertEqual(ca2_value(done, "Mobile Number:"), "9434310339")      # the sheet's, not the workbook's
        self.assertIn("9434310338", text)                                       # but the difference is shown
        self.assertNotIn("Name of the Student:' in the sheet", text)            # name differs only by case
        self.assertTrue(any("9434310338" in r["Notes"] for r in self.report()))

    def test_blank_student_fields_are_filled_from_the_workbook(self):
        self.sheet("t.docx", "34900725002", "TANIA DANDAPAT", mobile="")
        code, text = self.go()
        done = self.out / "docx" / "34900725002_Tania Dandapat.docx"
        self.assertEqual(ca2_value(done, "Mobile Number:"), "9434310339")
        self.assertIn("was blank", text)

    def test_fixed_values_only_fill_blanks(self):
        self.sheet("t.docx", "34900725002", "T")
        self.go("--exam-date", "05-11-2026")
        done = self.out / "docx" / "34900725002_Tania Dandapat.docx"
        self.assertEqual(ca2_value(done, "Date of Examination:"), "05-11-2026")   # was blank
        self.assertEqual(ca2_value(done, "Subject Teacher:"), "Dr X")              # student's own, untouched

    def test_absent_student(self):
        self.sheet("m.docx", "34900725003", "MITA DAS")
        code, _ = self.go()
        done = self.out / "docx" / "34900725003_Mita Das.docx"
        self.assertEqual(ca2_value(done, "Marks Obtained:"), "Ab")
        self.assertEqual(awarded_marks(done), {})

    def test_the_same_roll_twice_skips_both_and_the_run_continues(self):
        self.sheet("a.docx", "34900725002", "TANIA")
        self.sheet("b.docx", "34900725002", "TANIA")
        self.sheet("sarat.docx", "LME2026410", "SARAT ROY")
        code, text = self.go()
        self.assertEqual(code, 1)
        self.assertIn("submitted more than once", text)
        self.assertEqual(sorted(p.name for p in (self.out / "docx").glob("*.docx")), ["LME2026410_Sarat Roy.docx"])

    def test_a_roll_that_is_not_in_the_workbook(self):
        self.sheet("x.docx", "99999999999", "STRANGER")
        self.sheet("t.docx", "34900725002", "TANIA")
        code, text = self.go()
        self.assertEqual(code, 1)
        self.assertIn("'99999999999' is not in the workbook", text)
        self.assertEqual(len(list((self.out / "docx").glob("*.docx"))), 1)

    def test_students_without_a_sheet_are_listed(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        _, text = self.go()
        self.assertIn("Students in the workbook with no sheet", text)
        self.assertIn("Rohan Sen", text)
        statuses = {r["Roll"]: r["Status"] for r in self.report()}
        self.assertEqual(statuses["34900725004"], "NO SUBMISSION")
        self.assertEqual(statuses["34900725002"], "completed")

    def test_unreadable_files_are_reported_by_name_and_do_not_stop_the_run(self):
        import zipfile
        good = self.sheet("good.docx", "34900725002", "TANIA")
        (self.subs / "truncated.docx").write_bytes(good.read_bytes()[:300])
        with zipfile.ZipFile(good) as src, zipfile.ZipFile(self.subs / "no_image.docx", "w") as out:
            for name in src.namelist():          # the picture part is gone but still referenced
                if not name.startswith("word/media/"):
                    out.writestr(name, src.read(name))
        (self.subs / "notes.docx").write_text("this is not a Word file")
        code, text = self.go()
        self.assertEqual(code, 1)
        for name in ("truncated.docx", "no_image.docx", "notes.docx"):
            self.assertIn(name, text)
        self.assertIn("could not be opened", text)
        self.assertTrue((self.out / "docx" / "34900725002_Tania Dandapat.docx").exists())

    def test_other_file_types_are_refused_and_lock_files_ignored(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        (self.subs / "macro.docm").write_bytes(b"x")
        (self.subs / "~$t.docx").write_bytes(b"x")
        _, text = self.go()
        self.assertIn("macro.docm: not read", text)
        self.assertNotIn("~$t.docx", text)

    def test_wrong_layout_setting_gives_one_clear_message(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        cfg = g.json.loads(self.config.read_text())
        for entry in cfg["template_fields"]:
            entry.pop("where")                    # the 'after-label' default does not fit these sheets
        self.config.write_text(g.json.dumps(cfg))
        code, text = self.go()
        self.assertEqual(code, 1)
        self.assertIn("next-cell", text)
        self.assertFalse(self.out.exists())

    def test_a_total_that_cannot_be_split_skips_only_that_student(self):
        make_workbook(self.workbook, STUDENTS + [["Low Scorer", 34900725005, "9434310342", 5]])
        self.sheet("low.docx", "34900725005", "LOW SCORER")
        self.sheet("t.docx", "34900725002", "TANIA")
        code, text = self.go()
        self.assertEqual(code, 1)
        self.assertIn("5 marks cannot be split", text)
        self.assertEqual([p.name for p in (self.out / "docx").glob("*.docx")], ["34900725002_Tania Dandapat.docx"])

    def test_a_second_layout_with_other_sections_is_handled_per_layout(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        self.sheet("odd.docx", "LME2026410", "SARAT", questions=[("1", "5"), ("2", "5"), ("3", "5")])
        code, text = self.go()      # --answer 5,4 fits the first layout (2 sections) but not the 3-row one
        self.assertEqual(code, 1)
        self.assertIn("odd.docx", text)
        self.assertTrue((self.out / "docx" / "34900725002_Tania Dandapat.docx").exists())

    def test_existing_output_can_be_kept(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        self.go()
        target = self.out / "docx" / "34900725002_Tania Dandapat.docx"
        target.write_bytes(b"old")
        self.go("--existing", "skip")
        self.assertEqual(target.read_bytes(), b"old")

    def test_check_mode_writes_nothing(self):
        self.sheet("t.docx", "34900725002", "TANIA")
        code, text = self.go("--check")
        self.assertEqual(code, 0, text)
        self.assertIn("Check finished: 1 submitted sheet(s)", text)
        self.assertFalse(self.out.exists())

    def test_a_folder_without_docx_files(self):
        code, text = self.go()
        self.assertEqual(code, 1)
        self.assertIn("no .docx files", text)

    def test_mode_question_when_nothing_is_chosen(self):
        argv = ["--workbook", str(self.workbook), "--config", str(self.config), "--output", str(self.out),
                "--no-prompt", "--keep-docx-only"]
        # no --template/--submissions: the program asks; reply 2 then a folder picker would open,
        # so only the question itself is checked here
        args = g.parse_args(argv)
        text = io.StringIO()
        with redirect_stdout(text):
            try:
                g.choose_sources(args, feeder("2"))
            except Exception:
                pass                                # the folder picker needs a screen
        self.assertIn("Complete sheets that students have filled in", text.getvalue())


class NextCellLayout(TempCase):
    def test_a_merged_cell_counts_once(self):
        document = Document()
        table = document.add_table(rows=1, cols=4)
        table.cell(0, 0).merge(table.cell(0, 1))
        table.cell(0, 0).text = "Label:"
        table.cell(0, 2).text = "value"
        cell = g.label_cell(table.cell(0, 0).paragraphs[0])
        self.assertEqual(g.next_cell(cell).text, "value")

    def test_the_last_cell_has_no_value_cell(self):
        document = Document()
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 1).text = "Label:"
        with self.assertRaises(g.UserError):
            g.value_cell(table.cell(0, 1).paragraphs[0], "Label:")

    def test_a_label_outside_a_table_cannot_use_next_cell(self):
        document = Document()
        paragraph = document.add_paragraph("Label:")
        with self.assertRaises(g.UserError):
            g.value_cell(paragraph, "Label:")

    def test_empty_cell_takes_the_labels_look_without_its_bold(self):
        path = make_ca2(self.dir / "s.docx", "1", "A", picture=False)
        document = Document(path)
        paragraph = g.find_label_paragraph(document, "Marks Obtained:")
        entry = {"label": "Marks Obtained:", "where": "next-cell"}
        g.write_field(paragraph, entry, "17", g.font_settings(config()))
        run = g.value_cell(paragraph, "Marks Obtained:").paragraphs[0].runs[0]
        self.assertEqual((run.text, run.bold, run.font.name), ("17", False, "Trebuchet MS"))

    def test_empty_cell_uses_the_formatting_word_would_type_with(self):
        from docx.oxml import OxmlElement
        document = Document()
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Label:"
        mark_paragraph = table.cell(0, 1).paragraphs[0]
        properties = mark_paragraph._p.get_or_add_pPr()
        run_properties = OxmlElement("w:rPr")
        run_properties.append(OxmlElement("w:i"))
        properties.append(run_properties)
        g.write_field(table.cell(0, 0).paragraphs[0], {"label": "Label:", "where": "next-cell"}, "x",
                      g.font_settings(config()))
        self.assertTrue(table.cell(0, 1).paragraphs[0].runs[0].italic)

    def test_a_filled_cell_keeps_its_own_run_formatting(self):
        document = Document()
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Label:"
        run = table.cell(0, 1).paragraphs[0].add_run("old")
        run.font.name = "Arial"
        g.write_field(table.cell(0, 0).paragraphs[0], {"label": "Label:", "where": "next-cell"}, "new",
                      g.font_settings(config()))
        cell_run = table.cell(0, 1).paragraphs[0].runs[0]
        self.assertEqual((cell_run.text, cell_run.font.name), ("new", "Arial"))

    def test_custom_font_applies_in_a_value_cell(self):
        path = make_ca2(self.dir / "s.docx", "1", "A", picture=False)
        document = Document(path)
        cfg = config()
        cfg["font_mode"], cfg["custom_font"] = "custom", {"name": "Calibri", "size_pt": 11}
        paragraph = g.find_label_paragraph(document, "Marks Obtained:")
        g.write_field(paragraph, {"label": "Marks Obtained:", "where": "next-cell"}, "9", g.font_settings(cfg))
        run = g.value_cell(paragraph, "Marks Obtained:").paragraphs[0].runs[0]
        self.assertEqual((run.font.name, run.font.size.pt), ("Calibri", 11))

    def test_generating_a_new_sheet_also_works_with_next_cell(self):
        template = make_ca2(self.dir / "t.docx", "", "", mobile="", teacher="", picture=False)
        cfg = g.load_config(None)
        cfg = g.merge_config(cfg, g.json.loads(ca2_config_text()) if (g.HERE / "presets" / "ca2.json").exists() else {})
        cfg["signature"]["enabled"] = False
        out = self.dir / "o.docx"
        g.create_docx(template, student(2, "123", name="Ana", mobile="9999999999", marks="12"), None, out, cfg)
        self.assertEqual(ca2_value(out, "Name of the Student:"), "Ana")
        self.assertEqual(ca2_value(out, "Roll Number:"), "123")
        self.assertEqual(ca2_value(out, "Marks Obtained:"), "12")

    def test_where_must_be_a_known_value(self):
        cfg = config()
        cfg["template_fields"][0]["where"] = "below"
        with self.assertRaises(g.UserError):
            g.validate_config(cfg)


# --------------------------------------------------------------------------
# Official seal (optional)
# --------------------------------------------------------------------------

TINY_JPEG = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDABALDA4MChAODQ4SERATGCgaGBYWGDEjJR0oOjM9PDkzODdASFxOQERXRTc4UG1RV19iZ2hnPk1xeXBkeFxlZ2P/2wBDARESEhgVGC8aGi9jQjhCY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2NjY2P/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDFoooryz7w/9k=")


def seal_anchors(document) -> list:
    return [a for a in document.element.body.iter(qn("wp:anchor"))
            if a.find(qn("wp:docPr")).get("name") == "Seal"]


def make_page(path: Path, picture_name: str | None = None) -> Path:
    """A small A4 page shaped like the real sheet (narrow side margins), ending with a text line."""
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin, section.right_margin = Cm(0.8), Cm(0.9)
    document.add_paragraph("Signature of the student with date")
    if picture_name:
        document.add_picture(io.BytesIO(tiny_png()))
        document.element.body.iter(qn("wp:docPr")).__next__().set("name", picture_name)
    document.save(path)
    return path


class SealStamping(TempCase):
    def setUp(self):
        super().setUp()
        self.png = self.dir / "seal.png"
        self.png.write_bytes(tiny_png())

    def stamp(self, **settings):
        cfg = config()
        cfg["seal"].update(settings)
        document = Document(make_page(self.dir / "page.docx"))
        outcome = g.stamp_seal(document, self.png, cfg)
        out = self.dir / "stamped.docx"
        document.save(out)
        return outcome, Document(out), cfg

    def offset(self, anchor, axis):
        return int(anchor.find(qn(f"wp:position{axis}")).find(qn("wp:posOffset")).text)

    def test_stamped_just_below_the_last_line_at_the_right(self):
        outcome, document, _ = self.stamp()
        self.assertEqual(outcome, "stamped")
        (anchor,) = seal_anchors(document)
        text_width = Cm(21) - Cm(0.8) - Cm(0.9)
        self.assertAlmostEqual(self.offset(anchor, "H"), text_width - Cm(3.3), delta=1000)   # flush right
        self.assertEqual(self.offset(anchor, "V"), int(Cm(0.3)))                          # small gap
        self.assertEqual(anchor.find(qn("wp:positionH")).get("relativeFrom"), "margin")
        self.assertEqual(anchor.find(qn("wp:positionV")).get("relativeFrom"), "paragraph")
        self.assertEqual(int(anchor.find(qn("wp:extent")).get("cx")), int(Cm(3.3)))
        self.assertIsNotNone(anchor.find(qn("wp:wrapNone")))                               # does not move text

    def test_the_seal_hangs_from_a_one_point_paragraph_at_the_very_end(self):
        _, document, _ = self.stamp()
        last = document.paragraphs[-1]
        self.assertEqual(document.paragraphs[-2].text, "Signature of the student with date")
        self.assertEqual(last.text, "")
        self.assertEqual(last.paragraph_format.line_spacing, Pt(1))
        self.assertEqual(last.paragraph_format.space_before, Pt(0))
        self.assertEqual(last.paragraph_format.space_after, Pt(0))
        self.assertEqual(len(seal_anchors(document)), 1)
        self.assertIn(last._p, [a.getparent().getparent().getparent() for a in seal_anchors(document)])

    def test_the_picture_goes_in_unchanged(self):
        _, document, _ = self.stamp()
        pictures = [part.blob for part in document.part.package.iter_parts()
                    if part.partname.startswith("/word/media/")]
        self.assertEqual(pictures, [self.png.read_bytes()])

    def test_alignment_width_gap_and_offset_come_from_the_config(self):
        text_width = Cm(21) - Cm(0.8) - Cm(0.9)
        _, document, _ = self.stamp(align="left", offset_cm=1.0, width_cm=2.0, gap_cm=0.5)
        (anchor,) = seal_anchors(document)
        self.assertEqual(self.offset(anchor, "H"), int(Cm(1.0)))
        self.assertEqual(self.offset(anchor, "V"), int(Cm(0.5)))
        self.assertEqual(int(anchor.find(qn("wp:extent")).get("cx")), int(Cm(2.0)))
        _, document, _ = self.stamp(align="center", width_cm=4.0)
        (anchor,) = seal_anchors(document)
        self.assertAlmostEqual(self.offset(anchor, "H"), (text_width - Cm(4.0)) // 2, delta=1000)
        _, document, _ = self.stamp(align="right", offset_cm=0.5)
        (anchor,) = seal_anchors(document)
        self.assertAlmostEqual(self.offset(anchor, "H"), text_width - Cm(3.3) - Cm(0.5), delta=1000)

    def test_a_sheet_that_already_has_a_seal_is_left_alone(self):
        for name in ("College Seal", "Image 6 seal"):
            page = make_page(self.dir / "p.docx", picture_name=name)
            document = Document(page)
            before = len(document.paragraphs)
            self.assertEqual(g.stamp_seal(document, self.png, config()), "already had a seal")
            self.assertEqual(len(document.paragraphs), before)
            self.assertEqual(seal_anchors(document), [])

    def test_other_pictures_do_not_count_as_a_seal(self):
        document = Document(make_page(self.dir / "p.docx", picture_name="Signature of student"))
        self.assertEqual(g.stamp_seal(document, self.png, config()), "stamped")

    def test_the_word_seal_in_other_text_is_not_confused_with_a_picture(self):
        document = Document()
        document.add_paragraph("The seal of the college will be added later")
        self.assertEqual(g.stamp_seal(document, self.png, config()), "stamped")

    def test_seal_settings_are_validated(self):
        for bad in ({"align": "middle"}, {"width_cm": 0}, {"width_cm": 20}, {"gap_cm": -1}, {"offset_cm": "x"}):
            cfg = config()
            cfg["seal"].update(bad)
            with self.assertRaises(g.UserError, msg=str(bad)):
                g.validate_config(cfg)


class SealImage(TempCase):
    def test_a_missing_file(self):
        with self.assertRaises(g.UserError):
            g.check_seal_image(self.dir / "nope.png")

    def test_a_file_that_is_not_a_picture(self):
        bad = self.dir / "seal.png"
        bad.write_text("this is not an image")
        with self.assertRaises(g.UserError) as caught:
            g.check_seal_image(bad)
        self.assertIn("could not be read as a picture", str(caught.exception))

    def test_a_png_is_fine(self):
        good = self.dir / "seal.png"
        good.write_bytes(tiny_png())
        self.assertEqual(g.check_seal_image(good), [])

    def test_a_jpeg_works_with_a_transparency_warning(self):
        jpg = self.dir / "seal.jpg"
        jpg.write_bytes(TINY_JPEG)
        warnings = g.check_seal_image(jpg)
        self.assertEqual(len(warnings), 1)
        self.assertIn("transparent", warnings[0])
        document = Document(make_page(self.dir / "p.docx"))
        self.assertEqual(g.stamp_seal(document, jpg, config()), "stamped")

    def test_a_big_file_is_warned_about(self):
        big = self.dir / "big.png"
        big.write_bytes(tiny_png() + b"\0" * 1_600_000)   # still a valid picture header
        self.assertTrue(any("MB" in w for w in g.check_seal_image(big)))


class SealQuestion(TempCase):
    """The question asked just before the sheets are made, and the picker behind it."""

    def args(self, *extra):
        return g.parse_args(["--workbook", "w.xlsx", *extra])

    def test_an_explicit_seal_is_used_without_asking(self):
        png = self.dir / "s.png"
        png.write_bytes(tiny_png())
        asked = []
        path, warnings = g.choose_seal(self.args("--seal", str(png), "--no-prompt"), 5, asked.append)
        self.assertEqual((path, warnings, asked), (png, [], []))

    def test_no_prompt_and_check_never_ask(self):
        def refuse(prompt):
            raise AssertionError("must not ask")
        self.assertEqual(g.choose_seal(self.args("--no-prompt"), 5, refuse), (None, []))
        self.assertEqual(g.choose_seal(self.args("--check"), 5, refuse), (None, []))

    def test_nothing_to_stamp_means_no_question(self):
        def refuse(prompt):
            raise AssertionError("must not ask")
        self.assertEqual(g.choose_seal(self.args(), 0, refuse), (None, []))

    def test_the_question_names_the_number_of_sheets(self):
        shown = []

        def ask(prompt):
            shown.append(prompt)
            return "n"
        self.assertEqual(g.choose_seal(self.args(), 58, ask), (None, []))
        self.assertIn("58 sheet(s)", shown[0])

    def test_yes_opens_the_picker(self):
        png = self.dir / "s.png"
        png.write_bytes(tiny_png())
        with mock.patch.object(g, "choose_input_file", return_value=png) as picker:
            for reply in ("y", "Y", "yes"):
                self.assertEqual(g.choose_seal(self.args(), 3, lambda _: reply), (png, []))
        self.assertEqual(picker.call_count, 3)

    def test_anything_else_means_no_seal(self):
        with mock.patch.object(g, "choose_input_file") as picker:
            for reply in ("", "n", "no", "maybe"):
                self.assertEqual(g.choose_seal(self.args(), 3, lambda _: reply), (None, []))
        picker.assert_not_called()

    def test_closing_the_input_means_no_seal(self):
        self.assertEqual(g.choose_seal(self.args(), 3, feeder()), (None, []))

    def test_cancelling_the_picker_carries_on_without_a_seal(self):
        with mock.patch.object(g, "choose_input_file", side_effect=g.UserError("No seal image was selected.")):
            with redirect_stdout(io.StringIO()) as shown:
                self.assertEqual(g.choose_seal(self.args(), 3, lambda _: "y"), (None, []))
        self.assertIn("without a seal", shown.getvalue())

    def test_a_picker_that_cannot_open_says_how_to_continue(self):
        with mock.patch.object(g, "choose_input_file", side_effect=RuntimeError("no display")):
            with self.assertRaises(g.UserError) as caught:
                g.choose_seal(self.args(), 3, lambda _: "y")
        self.assertIn("--seal PATH", str(caught.exception))

    def test_a_chosen_file_that_is_not_a_picture_is_refused(self):
        bad = self.dir / "s.png"
        bad.write_text("nope")
        with mock.patch.object(g, "choose_input_file", return_value=bad):
            with self.assertRaises(g.UserError):
                g.choose_seal(self.args(), 3, lambda _: "y")


class RecoveryHelpers(TempCase):
    def sheet(self, picture, member, store=False, damage=None):
        """A sheet with a zero-checksum picture; `damage` swaps in damaged picture bytes afterwards."""
        path = make_ca2(self.dir / "s.docx", "1", "A", picture=False)
        document = Document(path)
        document.add_picture(io.BytesIO(picture))
        document.save(path)
        if damage is not None:
            source = zipfile.ZipFile(path)
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
                for info in source.infolist():
                    data = damage if info.filename == member else source.read(info.filename)
                    target.writestr(info.filename, data)
            path.write_bytes(buffer.getvalue())
        return zero_checksum(path, member, store=store)

    def test_python_docx_alone_cannot_open_such_a_file(self):
        path = self.sheet(tiny_png(), "word/media/image1.png")
        with self.assertRaises(zipfile.BadZipFile):
            Document(path)

    def test_a_png_is_recovered_byte_for_byte(self):
        path = self.sheet(tiny_png(), "word/media/image1.png")
        document, notes, data = g.open_document(path)
        self.assertEqual(len(notes), 1)
        self.assertIn("wrong checksum", notes[0])
        self.assertIn("looks complete", notes[0])
        rebuilt = zipfile.ZipFile(io.BytesIO(data))
        self.assertIsNone(rebuilt.testzip())
        self.assertEqual(rebuilt.read("word/media/image1.png"), tiny_png())
        self.assertGreater(len(document.tables), 0)

    def test_a_stored_jpeg_like_the_real_files(self):
        path = self.sheet(TINY_JPEG, "word/media/image1.jpg", store=True)
        zf = zipfile.ZipFile(path)
        self.assertEqual(zf.getinfo("word/media/image1.jpg").compress_type, zipfile.ZIP_STORED)
        self.assertEqual(zf.getinfo("word/media/image1.jpg").CRC, 0)
        _, notes, data = g.open_document(path)
        self.assertIn("looks complete", notes[0])
        self.assertEqual(zipfile.ZipFile(io.BytesIO(data)).read("word/media/image1.jpg"), TINY_JPEG)

    def test_a_cut_off_jpeg_is_flagged(self):
        _, notes, _ = g.open_document(self.sheet(TINY_JPEG, "word/media/image1.jpg", damage=TINY_JPEG[:-30]))
        self.assertIn("may be damaged", notes[0])

    def test_a_png_with_a_damaged_chunk_is_flagged(self):
        broken = bytearray(tiny_png())
        broken[41] ^= 0xFF                                  # inside the image data
        _, notes, _ = g.open_document(self.sheet(tiny_png(), "word/media/image1.png", damage=bytes(broken)))
        self.assertIn("may be damaged", notes[0])

    def test_png_checking(self):
        good = tiny_png()
        self.assertTrue(g.png_is_intact(good))
        self.assertFalse(g.png_is_intact(good[:-6]))
        self.assertFalse(g.png_is_intact(b"not a png"))
        flipped = bytearray(good)
        flipped[20] ^= 0x01
        self.assertFalse(g.png_is_intact(bytes(flipped)))

    def test_a_damaged_text_part_is_not_recovered(self):
        path = zero_checksum(make_ca2(self.dir / "s.docx", "1", "A", picture=False), "word/document.xml")
        with self.assertRaises(zipfile.BadZipFile) as caught:
            g.open_document(path)
        self.assertIn("not a picture", str(caught.exception))

    def test_a_file_that_is_not_a_zip_is_not_touched(self):
        bad = self.dir / "s.docx"
        bad.write_bytes(b"plain text")
        with self.assertRaises(zipfile.BadZipFile):
            g.open_document(bad)

    def test_a_healthy_file_opens_without_notes(self):
        path = make_ca2(self.dir / "s.docx", "1", "A")
        _, notes, data = g.open_document(path)
        self.assertEqual((notes, data), ([], None))

    def test_scanning_a_recoverable_sheet_reads_its_roll_number(self):
        if not (g.HERE / "presets" / "ca2.json").exists():
            self.skipTest("no presets/ca2.json beside the script")
        (self.dir / "ca2.json").write_text(ca2_config_text())
        cfg = g.load_config(self.dir / "ca2.json")
        path = self.sheet(TINY_JPEG, "word/media/image1.jpg", store=True)
        sub = g.scan_submission(path, cfg, g.roll_entry(cfg))
        self.assertEqual((sub.problem, sub.roll), ("", "1"))
        self.assertTrue(any("wrong checksum" in n for n in sub.notes))
        self.assertIsNotNone(sub.repaired)


class ChooseSources(TempCase):
    def args(self, *extra):
        return g.parse_args(["--workbook", "w.xlsx", *extra])

    def test_a_template_may_accompany_the_submissions(self):
        folder, template = g.choose_sources(self.args("--submissions", "subs", "--template", "t.docx"), feeder())
        self.assertEqual((folder, template), (Path("subs"), Path("t.docx")))

    def test_each_alone_is_unchanged(self):
        self.assertEqual(g.choose_sources(self.args("--submissions", "subs"), feeder()), (Path("subs"), None))
        self.assertEqual(g.choose_sources(self.args("--template", "t.docx"), feeder()), (None, Path("t.docx")))

    def test_the_questions_when_nothing_is_given(self):
        with mock.patch.object(g, "choose_folder", return_value=Path("subs")), \
                mock.patch.object(g, "choose_input_file", return_value=Path("t.docx")):
            self.assertEqual(g.choose_sources(self.args(), feeder("2", "y")), (Path("subs"), Path("t.docx")))
            self.assertEqual(g.choose_sources(self.args(), feeder("2", "n")), (Path("subs"), None))
            self.assertEqual(g.choose_sources(self.args(), feeder("2")), (Path("subs"), None))   # input closed
            self.assertEqual(g.choose_sources(self.args(), feeder("1")), (None, Path("t.docx")))
            self.assertEqual(g.choose_sources(self.args(), feeder()), (None, Path("t.docx")))

    def test_cancelling_the_template_picker_keeps_the_submissions(self):
        with mock.patch.object(g, "choose_folder", return_value=Path("subs")), \
                mock.patch.object(g, "choose_input_file", side_effect=g.UserError("none")):
            with redirect_stdout(io.StringIO()) as shown:
                self.assertEqual(g.choose_sources(self.args(), feeder("2", "y")), (Path("subs"), None))
        self.assertIn("will only be listed", shown.getvalue())


class Release(TempCase):
    def test_version_flag(self):
        with redirect_stdout(io.StringIO()) as buffer:
            with self.assertRaises(SystemExit) as caught:
                g.parse_args(["--version"])
        self.assertEqual(caught.exception.code, 0)
        self.assertIn(g.__version__, buffer.getvalue())

    def test_changelog_leads_with_the_current_version(self):
        changelog = g.HERE / "CHANGELOG.md"
        if not changelog.exists():
            self.skipTest("no CHANGELOG.md beside the script")
        headings = [l for l in changelog.read_text(encoding="utf-8").splitlines() if l.startswith("## ")]
        self.assertTrue(headings[0].startswith(f"## {g.__version__}"), headings[0])

    def test_gitignore_keeps_data_out_but_lets_the_examples_in(self):
        ignore = g.HERE / ".gitignore"
        if not ignore.exists():
            self.skipTest("no .gitignore beside the script")
        text = ignore.read_text()
        for pattern in ("*.xlsx", "*.docx", "output/", "*.local.json", "section_rules.json", "submissions/", "seal/",
                        "!examples/*.docx", "!examples/*.xlsx"):
            self.assertIn(pattern, text)
        # the exceptions must come after the broad patterns, or they do nothing
        self.assertGreater(text.index("!examples/*.docx"), text.index("*.docx"))

    def test_the_shipped_examples_work(self):
        examples = g.HERE / "examples"
        workbook = examples / "Template_student_data.xlsx"
        template = examples / "ME_CA1_2026.docx"
        if not (workbook.exists() and template.exists()):
            self.skipTest("no examples folder beside the script")
        (self.dir / "config.json").write_text((g.HERE / "config.json").read_text(encoding="utf-8-sig"))
        argv = ["--workbook", str(workbook), "--template", str(template), "--config",
                str(self.dir / "config.json"), "--output", str(self.dir / "out"), "--answer", "5,4",
                "--exam-date", "01-01-2027", "--no-prompt", "--keep-docx-only", "--workers", "1"]
        code, text = quiet(g.run, g.parse_args(argv))
        self.assertEqual(code, 0, text)
        self.assertEqual(len(list((self.dir / "out" / "docx").glob("*.docx"))), 16)


class WrongFiles(TempCase):
    def test_a_wrong_template_file_gives_a_readable_error(self):
        bad = self.dir / "t.docx"
        bad.write_bytes(b"not a word file")
        with self.assertRaises(g.UserError) as caught:
            g.open_template(bad)
        self.assertIn("t.docx", str(caught.exception))

    def test_a_wrong_workbook_file_gives_a_readable_error(self):
        bad = self.dir / "w.xlsx"
        bad.write_bytes(b"not an excel file")
        with self.assertRaises(g.UserError) as caught:
            g.read_students(bad, None, config())
        self.assertIn("w.xlsx", str(caught.exception))


class LocalConfig(TempCase):
    def test_local_file_overrides_the_shipped_one(self):
        shipped = self.dir / "config.json"
        shipped.write_text('{"exam_date": "01-01-2026", "full_marks": 30}')
        (self.dir / "config.local.json").write_text('{"exam_date": "02-02-2027"}')
        cfg = g.load_config(shipped)
        self.assertEqual((cfg["exam_date"], cfg["full_marks"]), ("02-02-2027", 30))

    def test_a_typo_in_the_local_file_names_the_file(self):
        shipped = self.dir / "config.json"
        shipped.write_text("{}")
        (self.dir / "config.local.json").write_text('{"exam_dat": "x"}')
        with self.assertRaises(g.UserError) as caught:
            g.load_config(shipped)
        self.assertIn("config.local.json", str(caught.exception))
        self.assertIn("exam_dat", str(caught.exception))

    def test_fixed_values_must_name_a_fixed_field(self):
        shipped = self.dir / "config.json"
        shipped.write_text('{"fixed_values": {"Nope:": "x"}}')
        with self.assertRaises(g.UserError):
            g.load_config(shipped)
        shipped.write_text('{"fixed_values": {"Name of the Student:": "x"}}')  # fed by a column
        with self.assertRaises(g.UserError):
            g.load_config(shipped)

    def test_saving_merges_and_keeps_other_settings(self):
        shipped = self.dir / "config.json"
        shipped.write_text("{}")
        local = self.dir / "config.local.json"
        local.write_text('{"exam_date": "x", "fixed_values": {"UPID:": "1"}, "font_mode": "custom"}')
        g.save_local_config(shipped, {"fixed_values": {"Paper (course) Code:": "ME"}, "full_marks": 30})
        saved = g.json.loads(local.read_text())
        self.assertEqual(saved["exam_date"], "x")
        self.assertEqual(saved["font_mode"], "custom")
        self.assertEqual(saved["fixed_values"], {"UPID:": "1", "Paper (course) Code:": "ME"})
        self.assertEqual(saved["full_marks"], 30)

    def test_local_and_rules_files_are_git_ignored(self):
        ignore = (g.HERE / ".gitignore")
        if ignore.exists():  # the packaged copy ships it; a bare checkout of the script may not
            text = ignore.read_text()
            self.assertIn("*.local.json", text)
            self.assertIn("section_rules.json", text)


if __name__ == "__main__":
    unittest.main()
