"""Automated Weekly Operations MIS Pack.

Reads every branch file for a reporting week, standardizes the columns, validates each
row, and writes a formatted Excel report with live KPI formulas and an exceptions sheet.

Usage:
    python mis_pack.py --week-start 2026-11-02

With no --week-start it reports on the last complete Monday-to-Sunday week, which is
what a scheduled Monday-morning run needs. Files are read from data/input/<week-start>/.
"""

import argparse
import csv
import difflib
import io
import logging
import math
import re
import sys
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import config
from report_writer import write_report

log = logging.getLogger("mis_pack")

FIXED = "Auto-fixed"
REMOVED = "Removed duplicate"
EXCLUDED = "Excluded"
FOLLOW_UP = "Follow up"

EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
TEXT_SUFFIXES = {".csv", ".tsv", ".txt"}
DELIMITERS = [",", ";", "\t", "|"]
TOTAL_LABELS = {"total", "totals", "grand total", "subtotal", "sum"}
HEADER_SEARCH_ROWS = 25

# Day-first and month-first dates can look identical, so the file as a whole decides.
AMBIGUOUS_DATE_FORMATS = ["%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y"]
CLEAR_DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y", "%d %B %Y"]
EXCEL_SERIAL_RANGE = (40000, 60000)
EXCEL_EPOCH = date(1899, 12, 30)

PLAIN_NUMBER = re.compile(r"^-?\d+(\.\d+)?$")
GROUPED_NUMBER = re.compile(r"^-?\d{1,3}(,\d{2,3})*,\d{3}(\.\d+)?$")
CURRENCY_PREFIX = re.compile(r"^(₹|rs\.?|inr)", re.IGNORECASE)


class FileProblem(Exception):
    """A branch file that cannot be used at all."""


@dataclass
class Issue:
    action: str
    issue: str
    branch: str = ""
    date: object = ""
    service: str = ""
    field: str = ""
    original: object = ""
    corrected: object = ""
    detail: str = ""
    source_file: str = ""
    source_row: object = ""


@dataclass
class Record:
    source_file: str
    source_row: int
    raw: dict
    from_excel: bool
    parsed_date: date = None
    date_kind: str = None
    code: str = None
    service: str = None
    numbers: dict = field(default_factory=dict)
    is_valid: bool = True


@dataclass
class RunResult:
    clean: pd.DataFrame
    issues: list
    stats: dict
    report_path: Path


def is_blank(value):
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return isinstance(value, str) and not value.strip()


def heading_key(text):
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


ALIASES = {
    heading_key(alias): column
    for column, aliases in config.COLUMN_ALIASES.items()
    for alias in aliases
}


def map_heading(heading):
    """Return (standard column, service) for a heading, or None if it is not recognized."""
    if is_blank(heading):
        return None
    key = heading_key(heading)
    if key in ALIASES:
        return ALIASES[key], None
    words = re.split(r"[\s_\-]+", str(heading).strip().lower(), maxsplit=1)
    if len(words) == 2 and words[0] in config.WIDE_SERVICE_PREFIXES:
        measure = ALIASES.get(heading_key(words[1]))
        if measure:
            return measure, config.WIDE_SERVICE_PREFIXES[words[0]]
    return None


def find_header(grid):
    """Find the row that looks most like a header, skipping any title rows above it."""
    best_index, best_columns, best_score = None, {}, 0
    for index, row in enumerate(grid[:HEADER_SEARCH_ROWS]):
        columns = {}
        taken = set()
        for position, heading in enumerate(row):
            mapped = map_heading(heading)
            if mapped and mapped not in taken:
                columns[position] = mapped
                taken.add(mapped)
        score = len({column for column, _ in columns.values()})
        if score > best_score:
            best_index, best_columns, best_score = index, columns, score
    return best_index, best_columns, best_score


def read_excel_grids(path):
    workbook = load_workbook(path, data_only=True)
    grids = [[list(row) for row in sheet.iter_rows(values_only=True)]
             for sheet in workbook.worksheets]
    workbook.close()
    return grids


def read_text_grids(path):
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252")
    return [list(csv.reader(io.StringIO(text), delimiter=delimiter)) for delimiter in DELIMITERS]


def extract_records(path):
    """Read one branch file into row records, whatever its layout."""
    from_excel = path.suffix.lower() in EXCEL_SUFFIXES
    try:
        grids = read_excel_grids(path) if from_excel else read_text_grids(path)
    except Exception as error:
        raise FileProblem(f"File could not be opened: {error}") from error

    best = max(((*find_header(grid), grid) for grid in grids), key=lambda item: item[2],
               default=(None, {}, 0, []))
    header_index, columns, score, grid = best
    if score < config.MIN_HEADER_MATCHES:
        raise FileProblem("No recognizable header row. Add the new headings to COLUMN_ALIASES.")

    found = {column for column, _ in columns.values()}
    services = sorted({service for _, service in columns.values() if service})
    if services:
        found.add("service")
    missing = sorted(set(config.REQUIRED_COLUMNS) - found)
    if missing:
        raise FileProblem(f"Required columns not found: {', '.join(missing)}")

    records = []
    for offset, row in enumerate(grid[header_index + 1:], start=header_index + 2):
        cells = {position: row[position] if position < len(row) else None for position in columns}
        if all(is_blank(value) for value in cells.values()):
            continue
        id_cells = [cells[p] for p, (column, _) in columns.items() if column in config.ID_COLUMNS]
        if any(str(value).strip().lower() in TOTAL_LABELS for value in id_cells):
            continue

        for service in services or [None]:
            raw = {"service": service} if service else {}
            for position, (column, column_service) in columns.items():
                if column_service in (None, service):
                    raw[column] = cells[position]
            records.append(Record(path.name, offset, raw, from_excel))
    return records


def parse_text_date(text, formats):
    for pattern in formats:
        try:
            return datetime.strptime(text, pattern).date(), pattern
        except ValueError:
            continue
    return None, None


def choose_date_format(values, week_start, week_end):
    """Pick the day-first or month-first reading that puts the file's dates in the week."""
    texts = [value.strip() for value in values if isinstance(value, str) and value.strip()]
    hits = Counter()
    for pattern in AMBIGUOUS_DATE_FORMATS:
        for text in texts:
            parsed, _ = parse_text_date(text, [pattern])
            if parsed and week_start <= parsed <= week_end:
                hits[pattern] += 1
    return hits.most_common(1)[0][0] if hits else None


def read_date(value, formats):
    """Return (date, kind) where kind records how the date was stored."""
    if isinstance(value, datetime):
        return value.date(), "real date"
    if isinstance(value, date):
        return value, "real date"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if EXCEL_SERIAL_RANGE[0] <= value <= EXCEL_SERIAL_RANGE[1]:
            return EXCEL_EPOCH + timedelta(days=int(value)), "Excel serial number"
        return None, None
    if isinstance(value, str):
        return parse_text_date(value.strip(), formats)
    return None, None


def parse_number(value, from_excel):
    """Return (number, status), where status is ok, fixed, missing or invalid."""
    if is_blank(value):
        return None, "missing"
    if isinstance(value, bool):
        return None, "invalid"
    if isinstance(value, (int, float)):
        return float(value), "ok"
    if not isinstance(value, str):
        return None, "invalid"

    text = value.strip()
    negative = text.startswith("(") and text.endswith(")")
    cleaned = CURRENCY_PREFIX.sub("", text.strip("()")).replace(" ", "")
    if GROUPED_NUMBER.match(cleaned):
        cleaned = cleaned.replace(",", "")
    if not PLAIN_NUMBER.match(cleaned):
        return None, "invalid"

    number = float(cleaned)
    is_plain_text = value == text and PLAIN_NUMBER.match(text)
    status = "ok" if is_plain_text and not from_excel else "fixed"
    return (-number if negative else number), status


class Validator:
    """Standardizes rows and records every change or rejection as an Issue."""

    def __init__(self, master, week_start):
        self.master = master
        self.week_start = week_start
        self.week_end = week_start + timedelta(days=config.DAYS_IN_WEEK - 1)
        self.issues = []
        self.codes = {heading_key(code): code for code in master["branch_code"]}
        self.names = {heading_key(row.branch_name): row.branch_code
                      for row in master.itertuples()}

    def add(self, action, issue, record=None, **details):
        if record is not None:
            details.setdefault("branch", record.code or record.raw.get("branch") or "")
            details.setdefault("date", record.parsed_date or "")
            details.setdefault("service", record.service or "")
            details.setdefault("source_file", record.source_file)
            details.setdefault("source_row", record.source_row)
        self.issues.append(Issue(action=action, issue=issue, **details))

    def exclude(self, record, issue, **details):
        record.is_valid = False
        self.add(EXCLUDED, issue, record, **details)

    def parse_dates(self, records):
        """Read one file's dates, letting the whole file settle day-first or month-first."""
        preferred = choose_date_format([r.raw.get("date") for r in records],
                                       self.week_start, self.week_end)
        leading = [preferred] if preferred else []
        formats = leading + AMBIGUOUS_DATE_FORMATS + CLEAR_DATE_FORMATS
        for record in records:
            record.parsed_date, record.date_kind = read_date(record.raw.get("date"), formats)

    def check_dates(self, records):
        """Log date problems, treating the file's usual storage format as normal."""
        kinds = Counter(r.date_kind for r in records if r.date_kind)
        usual = kinds.most_common(1)[0][0] if kinds else None
        for record in records:
            value = record.raw.get("date")
            if is_blank(value):
                self.exclude(record, "Missing value", field="date",
                             detail="Date is blank.")
            elif record.parsed_date is None:
                self.exclude(record, "Invalid date", field="date", original=value,
                             detail="Value could not be read as a date.")
            elif record.date_kind != usual:
                self.add(FIXED, "Date format standardized", record, field="date", original=value,
                         corrected=record.parsed_date.isoformat(),
                         detail=f"Stored as {self.describe_kind(record.date_kind)}, "
                                f"unlike the rest of the file.")
            if record.parsed_date and not self.week_start <= record.parsed_date <= self.week_end:
                self.exclude(record, "Date outside reporting week", field="date",
                             original=record.parsed_date.isoformat(),
                             detail=f"Reporting week is {self.week_start} to {self.week_end}.")

    @staticmethod
    def describe_kind(kind):
        return "text" if kind and kind.startswith("%") else kind

    def resolve_branch(self, record):
        value = record.raw.get("branch")
        if is_blank(value):
            self.exclude(record, "Missing value", field="branch", detail="Branch is blank.")
            return
        key = heading_key(value)
        code = self.codes.get(key) or self.names.get(key)
        issue = "Branch name standardized"
        if code is None:
            close = difflib.get_close_matches(key, self.names, n=2,
                                              cutoff=config.BRANCH_FUZZY_CUTOFF)
            if len(close) == 1:
                code = self.names[close[0]]
                issue = "Branch name typo corrected"
        if code is None:
            self.exclude(record, "Unknown branch", field="branch", original=value,
                         detail="Not in the branch master list.")
            return

        record.code = code
        name = self.master.set_index("branch_code").at[code, "branch_name"]
        if value not in (code, name):
            self.add(FIXED, issue, record, field="branch", original=value, corrected=code,
                     detail=f"Matched to {name}.")

    def resolve_service(self, record):
        value = record.raw.get("service")
        if is_blank(value):
            self.exclude(record, "Missing value", field="service", detail="Service is blank.")
            return
        service = config.SERVICE_ALIASES.get(re.sub(r"[^a-z]", "", str(value).lower()))
        if service is None:
            self.exclude(record, "Unknown service", field="service", original=value,
                         detail=f"Expected one of: {', '.join(config.SERVICES)}.")
            return
        record.service = service
        if value != service:
            self.add(FIXED, "Service name standardized", record, field="service", original=value,
                     corrected=service)

    def resolve_numbers(self, record):
        for column in config.NUMERIC_COLUMNS:
            value = record.raw.get(column)
            number, status = parse_number(value, record.from_excel)
            if status == "missing":
                self.exclude(record, "Missing value", field=column,
                             detail="A required figure is blank.")
            elif status == "invalid":
                self.exclude(record, "Not a number", field=column, original=value,
                             detail="Value could not be read as a number.")
            else:
                record.numbers[column] = number
                if status == "fixed":
                    self.add(FIXED, "Number format standardized", record, field=column,
                             original=value, corrected=number)

    def check_business_rules(self, record):
        numbers = record.numbers
        for column, number in numbers.items():
            if number < 0:
                self.exclude(record, "Negative value", field=column, original=number,
                             detail="Figures cannot be negative.")
        for column in config.COUNT_COLUMNS:
            if column in numbers and not float(numbers[column]).is_integer():
                self.exclude(record, "Not a whole number", field=column,
                             original=numbers[column], detail="Parcel counts must be whole.")

        if {"booked", "delivered", "returned"} <= numbers.keys():
            resolved = numbers["delivered"] + numbers["returned"]
            if resolved > numbers["booked"]:
                self.exclude(record, "Delivered + returned exceed booked", field="delivered",
                             original=resolved, detail=f"Booked is {numbers['booked']:.0f}.")
        if {"delivered", "on_time"} <= numbers.keys() and numbers["on_time"] > numbers["delivered"]:
            self.exclude(record, "On-time exceeds delivered", field="on_time",
                         original=numbers["on_time"],
                         detail=f"Delivered is {numbers['delivered']:.0f}.")
        if {"booked", "revenue"} <= numbers.keys() and numbers["booked"] > 0:
            per_parcel = numbers["revenue"] / numbers["booked"]
            low, high = config.REVENUE_PER_PARCEL_RANGE
            if not low <= per_parcel <= high:
                self.exclude(record, "Revenue per parcel out of range", field="revenue",
                             original=numbers["revenue"],
                             detail=f"Works out at {per_parcel:,.2f} per parcel, "
                                    f"allowed range is {low:g} to {high:g}.")

    def check_duplicates(self, records):
        """Keep one copy of exact repeats and exclude rows that contradict each other."""
        groups = {}
        for record in records:
            if record.is_valid:
                key = (record.parsed_date, record.code, record.service)
                groups.setdefault(key, []).append(record)

        for group in groups.values():
            if len(group) == 1:
                continue
            versions = {tuple(sorted(record.numbers.items())) for record in group}
            if len(versions) == 1:
                for record in group[1:]:
                    record.is_valid = False
                    self.add(REMOVED, "Exact duplicate row", record,
                             detail=f"Same as row {group[0].source_row}, first copy kept.")
            else:
                rows = ", ".join(str(record.source_row) for record in group)
                for record in group:
                    self.exclude(record, "Conflicting duplicate rows",
                                 detail=f"Rows {rows} report different figures for the "
                                        f"same day and service.")

    def check_completeness(self, records):
        """Flag branches that sent nothing, or left out a whole day."""
        days_by_branch = {}
        for record in records:
            if record.code and record.parsed_date:
                days_by_branch.setdefault(record.code, set()).add(record.parsed_date)

        week = [self.week_start + timedelta(days=offset)
                for offset in range(config.DAYS_IN_WEEK)]
        for row in self.master.itertuples():
            days = days_by_branch.get(row.branch_code)
            if days is None:
                self.add(FOLLOW_UP, "No data received", branch=row.branch_code,
                         detail=f"No usable rows for {row.branch_name} this week.")
                continue
            for day in week:
                if day not in days:
                    self.add(FOLLOW_UP, "Missing day", branch=row.branch_code, date=day,
                             detail=f"{row.branch_name} sent no rows for {day:%A}.")

    def collapse_repeated_fixes(self):
        """Replace a run of identical fixes in one file with a single summary line."""
        groups = {}
        for issue in self.issues:
            if issue.action == FIXED:
                groups.setdefault((issue.source_file, issue.issue, issue.field), []).append(issue)

        replaced = {id(issue) for group in groups.values()
                    if len(group) > config.FIX_COLLAPSE_LIMIT for issue in group}
        summaries = [
            Issue(FIXED, issue_name, field=field_name, source_file=source, branch=group[0].branch,
                  detail=f"{len(group)} cells in this file, for example "
                         f"{group[0].original!r} became {group[0].corrected!r}.")
            for (source, issue_name, field_name), group in groups.items()
            if len(group) > config.FIX_COLLAPSE_LIMIT
        ]
        self.issues = [i for i in self.issues if id(i) not in replaced] + summaries


def build_clean_table(records, master):
    rows = [
        {
            "date": record.parsed_date,
            "branch_code": record.code,
            "service": record.service,
            **{column: record.numbers[column] for column in config.NUMERIC_COLUMNS},
            "source_file": record.source_file,
            "source_row": record.source_row,
        }
        for record in records if record.is_valid
    ]
    columns = ["date", "branch_code", "service", *config.NUMERIC_COLUMNS,
               "source_file", "source_row"]
    clean = pd.DataFrame(rows, columns=columns)
    clean = clean.merge(master, on="branch_code", how="left")
    clean["date"] = pd.to_datetime(clean["date"])
    clean[config.COUNT_COLUMNS] = clean[config.COUNT_COLUMNS].astype("int64")
    return clean.sort_values(["branch_code", "date", "service"]).reset_index(drop=True)


def attach_prior_week(clean, history_path):
    """Add last week's figures for the same branch, weekday and service to each row."""
    prior_columns = {column: f"lw_{column}" for column in config.NUMERIC_COLUMNS}
    if history_path.exists():
        history = pd.read_csv(history_path, parse_dates=["date"])
    else:
        history = pd.DataFrame(columns=["date", "branch_code", "service",
                                        *config.NUMERIC_COLUMNS])
    prior = history.rename(columns=prior_columns)
    prior["date"] = pd.to_datetime(prior["date"]) + pd.Timedelta(days=7)
    merged = clean.merge(prior, on=["date", "branch_code", "service"], how="left")
    merged[list(prior_columns.values())] = merged[list(prior_columns.values())].fillna(0)
    return merged


def update_history(clean, history_path):
    """Save this week's clean rows. Re-running a week replaces it rather than doubling it."""
    columns = ["date", "branch_code", "service", *config.NUMERIC_COLUMNS]
    latest = clean[columns].copy()
    if history_path.exists():
        history = pd.read_csv(history_path, parse_dates=["date"])
        history = history[~history["date"].isin(latest["date"].unique())]
        latest = pd.concat([history, latest], ignore_index=True)
    latest = latest.sort_values(["date", "branch_code", "service"])
    latest.to_csv(history_path, index=False, date_format="%Y-%m-%d")


def network_kpis(clean):
    """Calculate headline KPIs in pandas, as an independent check on the Excel formulas."""
    if clean.empty:
        return {}
    totals = clean[config.NUMERIC_COLUMNS].sum()
    return {
        "booked": int(totals["booked"]),
        "delivered": int(totals["delivered"]),
        "delivery_rate": totals["delivered"] / totals["booked"],
        "on_time_rate": totals["on_time"] / totals["delivered"],
        "return_rate": totals["returned"] / totals["booked"],
        "revenue": float(totals["revenue"]),
        "fuel_cost": float(totals["fuel_cost"]),
        "fuel_per_delivered": totals["fuel_cost"] / totals["delivered"],
    }


def run_pack(week_start, input_dir=None, output_dir=None, master_path=None, history_path=None):
    """Run the whole pipeline for one week and return the result."""
    started = time.perf_counter()
    input_dir = Path(input_dir or config.INPUT_DIR / week_start.isoformat())
    output_dir = Path(output_dir or config.OUTPUT_DIR)
    master_path = Path(master_path or config.BRANCH_MASTER)
    history_path = Path(history_path or config.HISTORY_FILE)

    if not input_dir.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")
    master = pd.read_csv(master_path, dtype=str)
    validator = Validator(master, week_start)

    paths = sorted(
        path for path in input_dir.iterdir()
        if path.suffix.lower() in EXCEL_SUFFIXES | TEXT_SUFFIXES
        and not path.name.startswith("~$")
    )
    if not paths:
        raise FileNotFoundError(f"No branch files found in {input_dir}")

    all_records, files_read = [], 0
    for path in paths:
        try:
            records = extract_records(path)
        except FileProblem as problem:
            log.warning("Skipped %s: %s", path.name, problem)
            validator.add(FOLLOW_UP, "File could not be used", detail=str(problem),
                          source_file=path.name)
            continue
        files_read += 1
        validator.parse_dates(records)
        for record in records:
            validator.resolve_branch(record)
            validator.resolve_service(record)
        validator.check_dates(records)
        for record in records:
            validator.resolve_numbers(record)
            validator.check_business_rules(record)
        all_records.extend(records)
        log.info("Read %-34s %3d rows", path.name, len(records))

    validator.check_duplicates(all_records)
    validator.check_completeness(all_records)
    validator.collapse_repeated_fixes()

    clean = attach_prior_week(build_clean_table(all_records, master), history_path)
    removed = sum(issue.action == REMOVED for issue in validator.issues)
    stats = {
        "files_found": len(paths),
        "files_read": files_read,
        "branches_expected": len(master),
        "rows_read": len(all_records),
        "rows_removed": removed,
        "rows_excluded": len(all_records) - len(clean) - removed,
        "rows_used": len(clean),
        "rows_expected": len(master) * config.DAYS_IN_WEEK * len(config.SERVICES),
        "generated": datetime.now(),
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / f"MIS_Pack_{week_start.isoformat()}.xlsx"
    write_report(report_path, clean, validator.issues, master, stats, week_start)
    update_history(clean, history_path)

    stats["seconds"] = time.perf_counter() - started
    return RunResult(clean, validator.issues, stats, report_path)


def last_complete_week(today=None):
    """Return the Monday of the most recent week that has fully finished."""
    today = today or date.today()
    return today - timedelta(days=today.weekday() + 7)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Build the weekly operations MIS pack.")
    parser.add_argument("--week-start", help="Monday of the reporting week, YYYY-MM-DD")
    parser.add_argument("--input-dir", help="Folder of branch files for the week")
    parser.add_argument("--output-dir", help="Folder for the report")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s", stream=sys.stdout)
    try:
        week_start = (datetime.strptime(args.week_start, "%Y-%m-%d").date()
                      if args.week_start else last_complete_week())
    except ValueError:
        log.error("--week-start must look like 2026-11-02")
        return 2
    if week_start.weekday() != 0:
        log.error("--week-start must be a Monday, %s is a %s", week_start, f"{week_start:%A}")
        return 2

    try:
        result = run_pack(week_start, args.input_dir, args.output_dir)
    except FileNotFoundError as error:
        log.error("%s", error)
        return 1
    except PermissionError as error:
        log.error("Could not save the report. Close it in Excel and run again. (%s)", error)
        return 1

    stats = result.stats
    counts = Counter(issue.action for issue in result.issues)
    kpis = network_kpis(result.clean)
    log.info("")
    log.info("Report:     %s", result.report_path)
    log.info("Files:      %d read of %d found, %d branches expected",
             stats["files_read"], stats["files_found"], stats["branches_expected"])
    log.info("Rows:       %d read, %d used, %d excluded, %d duplicates removed",
             stats["rows_read"], stats["rows_used"], stats["rows_excluded"],
             stats["rows_removed"])
    log.info("Exceptions: %d auto-fixed, %d excluded, %d follow up",
             counts[FIXED], counts[EXCLUDED], counts[FOLLOW_UP])
    if kpis:
        log.info("On-time:    %.1f%%", kpis["on_time_rate"] * 100)
    log.info("Runtime:    %.1f seconds", stats["seconds"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
