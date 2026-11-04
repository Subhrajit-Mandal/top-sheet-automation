"""Tests for generate_top_sheets.py. Run with:  python -m unittest -v

The tests build their own small Word templates and workbooks, so they need
neither your real template nor any student data.
"""

from __future__ import annotations

import copy
import hashlib
import io
import random
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from docx import Document
from docx.shared import Pt
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
                            '{"label": "Date of Examination:", "value": "{exam_date}"}]}')

    def go(self, *extra, answers=("", "5", "4", "")):
        feed = iter(answers)

        def fake_input(prompt):
            try:
                return next(feed)
            except StopIteration:
                raise EOFError

        argv = ["--workbook", str(self.workbook), "--template", str(self.template),
                "--output", str(self.out), "--config", str(self.cfg), "--keep-docx-only",
                "--workers", "1", *extra]
        return quiet(g.run, g.parse_args(argv), fake_input)

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

    def test_shipped_config_matches_the_code_defaults(self):
        shipped = g.json.loads((g.HERE / "config.json").read_text(encoding="utf-8-sig"))
        self.assertEqual(shipped, g.DEFAULT_CONFIG)


if __name__ == "__main__":
    unittest.main()
