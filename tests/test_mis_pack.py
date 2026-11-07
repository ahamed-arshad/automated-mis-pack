"""Tests for the MIS pack.

The main test compares the pipeline against the generator's truth data and answer key,
so it proves both that seeded errors are caught and that clean rows are left correct.
Run `python generate_branch_files.py` first.
"""

import csv
import random
import shutil
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import config  # noqa: E402
import generate_branch_files as generator  # noqa: E402
import mis_pack  # noqa: E402

WEEK = date(2026, 11, 2)
SAMPLE_INPUT = config.INPUT_DIR / WEEK.isoformat()
ANSWER_KEY = generator.ANSWER_KEY

ISSUE_BY_KIND = {
    "branch_text": {"Branch name standardized", "Branch name typo corrected", "Unknown branch"},
    "service_text": {"Service name standardized", "Unknown service"},
    "number_as_text": {"Number format standardized"},
    "date_as_text": {"Date format standardized"},
    "date_as_serial": {"Date format standardized"},
    "blank": {"Missing value"},
    "set": {"Negative value", "Not a number"},
    "over_booked": {"Delivered + returned exceed booked"},
    "on_time_over": {"On-time exceeds delivered"},
    "date": {"Date outside reporting week", "Invalid date"},
    "revenue_x100": {"Revenue per parcel out of range"},
    "exact_duplicate": {"Exact duplicate row"},
    "conflicting_duplicate": {"Conflicting duplicate rows"},
    "missing_day": {"Missing day"},
    "missing_file": {"No data received"},
}


@pytest.fixture
def workspace(tmp_path):
    """Give each test its own history and output, so tests never touch the sample run."""
    history = tmp_path / "history.csv"
    with open(history, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["date", "branch_code", "service", *config.NUMERIC_COLUMNS])
    return tmp_path, history


@pytest.fixture
def sample_run(workspace):
    tmp_path, history = workspace
    return mis_pack.run_pack(WEEK, SAMPLE_INPUT, tmp_path / "out", history_path=history)


def answer_key():
    with open(ANSWER_KEY, newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def test_every_seeded_error_is_caught_with_the_right_outcome(sample_run):
    for seeded in answer_key():
        if seeded["error_kind"] == "missing_file":
            source = ""
        elif seeded["error_kind"] == "missing_day":
            source = ""
        else:
            source = generator.FILE_NAMES[seeded["branch_code"]]
        matches = [
            issue for issue in sample_run.issues
            if issue.action == seeded["expected_outcome"]
            and issue.issue in ISSUE_BY_KIND[seeded["error_kind"]]
            and issue.source_file == source
            and (seeded["error_kind"] in ("conflicting_duplicate", "exact_duplicate")
                 or not seeded["field"] or issue.field == seeded["field"])
        ]
        assert matches, f"Not caught: {seeded}"


def test_no_false_alarms(sample_run):
    """Every exception traces to a seeded error, or to the known currency-format files."""
    expected = set()
    for seeded in answer_key():
        code = seeded["branch_code"]
        source = "" if "missing" in seeded["error_kind"] else generator.FILE_NAMES[code]
        for issue_name in ISSUE_BY_KIND[seeded["error_kind"]]:
            expected.add((source, seeded["expected_outcome"], issue_name))
    for code in generator.LAYOUTS["csv_semicolon"]:
        expected.add((generator.FILE_NAMES[code], "Auto-fixed", "Number format standardized"))

    unexpected = [issue for issue in sample_run.issues
                  if (issue.source_file, issue.action, issue.issue) not in expected]
    assert not unexpected


def test_clean_rows_exactly_match_the_original_truth(sample_run):
    """Fixes must restore the true value, and day-first dates must not be misread."""
    truth = generator.build_week(random.Random(generator.SEED), WEEK)
    excluded = {(item[0], item[1], item[2]) for item in generator.SEEDED_ERRORS
                if apply_outcome(item) == generator.EXCLUDE}
    excluded |= set(generator.CONFLICTING_DUPLICATES)

    expected_keys = {
        key for key in truth
        if key[0] != generator.MISSING_FILE_BRANCH
        and (key[0], key[1]) != generator.MISSING_DAY
        and key not in excluded
    }
    actual = {
        (row.branch_code, row.date.weekday(), row.service): row
        for row in sample_run.clean.itertuples()
    }
    assert set(actual) == expected_keys
    assert len(actual) == len(sample_run.clean)

    for key, row in actual.items():
        true_row = truth[key]
        assert row.date.date() == true_row["date"]
        for column in config.NUMERIC_COLUMNS:
            assert getattr(row, column) == pytest.approx(true_row[column]), (key, column)


def apply_outcome(item):
    _, outcome = generator.apply_error(
        {"booked": 1, "delivered": 1, "on_time": 1, "returned": 0, "revenue": 1.0,
         "fuel_cost": 1.0, "date": WEEK, "branch": "x", "service": "x"},
        item[3], item[4])
    return outcome


def test_row_counts_reconcile(sample_run):
    stats = sample_run.stats
    assert stats["rows_read"] == stats["rows_used"] + stats["rows_excluded"] + stats["rows_removed"]
    assert stats["rows_used"] == len(sample_run.clean)


def test_rerun_replaces_history_instead_of_doubling_it(workspace):
    tmp_path, history = workspace
    first = mis_pack.run_pack(WEEK, SAMPLE_INPUT, tmp_path / "out", history_path=history)
    mis_pack.run_pack(WEEK, SAMPLE_INPUT, tmp_path / "out", history_path=history)
    assert len(pd.read_csv(history)) == len(first.clean)


def test_day_first_file_is_not_read_as_month_first():
    values = ["02/11/2026", "03/11/2026", "08/11/2026"]
    assert mis_pack.choose_date_format(values, WEEK, date(2026, 11, 8)) == "%d/%m/%Y"


def test_month_first_file_is_not_read_as_day_first():
    values = ["11/02/2026", "11/03/2026", "11/08/2026"]
    assert mis_pack.choose_date_format(values, WEEK, date(2026, 11, 8)) == "%m/%d/%Y"


@pytest.mark.parametrize("value, from_excel, expected", [
    (12, True, (12.0, "ok")),
    ("12", False, (12.0, "ok")),
    ("12", True, (12.0, "fixed")),
    ("₹1,234.50", False, (1234.5, "fixed")),
    ("Rs. 2,84,469", False, (284469.0, "fixed")),
    (" 1,204 ", False, (1204.0, "fixed")),
    ("(45)", False, (-45.0, "fixed")),
    ("1,2,3", False, (None, "invalid")),
    ("TBC", False, (None, "invalid")),
    ("", False, (None, "missing")),
    (None, True, (None, "missing")),
])
def test_parse_number(value, from_excel, expected):
    assert mis_pack.parse_number(value, from_excel) == expected


def run_single_file(workspace, build):
    tmp_path, history = workspace
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    build(input_dir)
    return mis_pack.run_pack(WEEK, input_dir, tmp_path / "out", history_path=history)


def test_corrupt_file_is_reported_and_does_not_stop_the_run(workspace):
    def build(folder):
        (folder / "broken.xlsx").write_bytes(b"this is not a workbook")
        shutil.copy(SAMPLE_INPUT / "BR01_weekly_ops.xlsx", folder)

    result = run_single_file(workspace, build)
    assert any(i.issue == "File could not be used" and i.source_file == "broken.xlsx"
               for i in result.issues)
    assert result.stats["files_read"] == 1


def test_file_missing_a_required_column_is_reported(workspace):
    def build(folder):
        wb = Workbook()
        wb.active.append(["Date", "Branch", "Service", "Booked", "Delivered", "On Time",
                          "Returned", "Revenue"])
        wb.active.append([datetime(2026, 11, 2), "BR01", "Express", 10, 9, 8, 1, 150])
        wb.save(folder / "no_fuel.xlsx")
        shutil.copy(SAMPLE_INPUT / "BR01_weekly_ops.xlsx", folder)

    result = run_single_file(workspace, build)
    problem = next(i for i in result.issues if i.source_file == "no_fuel.xlsx")
    assert "fuel_cost" in problem.detail


def test_text_from_branch_files_cannot_run_as_a_formula(workspace):
    def build(folder):
        (folder / "formula.csv").write_text(
            "Date,Branch,Service,Booked,Delivered,On Time,Returned,Revenue,Fuel Cost\n"
            "2026-11-02,=1+1,Express,10,9,8,1,150,9\n",
            encoding="utf-8",
        )
        shutil.copy(SAMPLE_INPUT / "BR01_weekly_ops.xlsx", folder)

    result = run_single_file(workspace, build)
    sheet = load_workbook(result.report_path)["Exceptions"]
    cells = [cell for row in sheet.iter_rows() for cell in row if cell.value == "=1+1"]
    assert cells
    assert all(cell.data_type == "s" for cell in cells)


def test_excel_formulas_match_pandas():
    """Needs a report already opened and saved by Excel, which stores calculated values."""
    report = config.OUTPUT_DIR / f"MIS_Pack_{WEEK.isoformat()}.xlsx"
    if not report.exists():
        pytest.skip("Sample report has not been generated")
    summary = load_workbook(report, data_only=True)["Summary"]
    values = {row[0].value: row[1].value for row in summary.iter_rows(min_row=7, max_row=40)}
    if values.get("Parcels booked") is None:
        pytest.skip("Report has no calculated values until it is saved from Excel")

    clean = pd.read_excel(report, sheet_name="Clean Data")
    totals = clean[["Booked", "Delivered", "On Time", "Returned", "Revenue", "Fuel Cost"]].sum()
    assert values["Parcels booked"] == totals["Booked"]
    assert values["Parcels delivered"] == totals["Delivered"]
    assert values["Delivery rate"] == pytest.approx(totals["Delivered"] / totals["Booked"])
    assert values["On-time delivery"] == pytest.approx(totals["On Time"] / totals["Delivered"])
    assert values["Return rate"] == pytest.approx(totals["Returned"] / totals["Booked"])
    assert values["Revenue"] == pytest.approx(totals["Revenue"])
    assert values["Fuel cost"] == pytest.approx(totals["Fuel Cost"])
    assert values["Check: rows reconcile"] == "OK"
    assert values["Check: KPI totals match Clean Data"] == "OK"
