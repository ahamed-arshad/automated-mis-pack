"""Create the synthetic branch files used to demonstrate the MIS pack.

Writes 15 branch files in 6 different layouts, seeds known data errors into them, and
saves an answer key so the pipeline can be tested against what was planted. Surat
(BR16) deliberately submits no file. All data is invented.

Usage:
    python generate_branch_files.py [--week-start 2026-11-02]
"""

import argparse
import csv
import random
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

from config import BRANCH_MASTER, HISTORY_FILE, INPUT_DIR, PROJECT_ROOT, SERVICES

SEED = 45
ANSWER_KEY = PROJECT_ROOT / "data" / "reference" / "seeded_errors_answer_key.csv"

BRANCHES = [
    ("BR01", "Ahmedabad", "West"),
    ("BR02", "Bengaluru", "South"),
    ("BR03", "Chandigarh", "North"),
    ("BR04", "Chennai", "South"),
    ("BR05", "Delhi", "North"),
    ("BR06", "Hyderabad", "South"),
    ("BR07", "Indore", "West"),
    ("BR08", "Jaipur", "North"),
    ("BR09", "Kochi", "South"),
    ("BR10", "Kolkata", "East"),
    ("BR11", "Lucknow", "North"),
    ("BR12", "Mumbai", "West"),
    ("BR13", "Nagpur", "West"),
    ("BR14", "Patna", "East"),
    ("BR15", "Pune", "West"),
    ("BR16", "Surat", "West"),
]
NAME_BY_CODE = {code: name for code, name, _ in BRANCHES}
MISSING_FILE_BRANCH = "BR16"

# Each branch has its own volume and service quality, so the KPIs differ meaningfully.
ON_TIME_RATE = {
    "BR01": 0.95, "BR02": 0.93, "BR03": 0.89, "BR04": 0.86, "BR05": 0.94, "BR06": 0.96,
    "BR07": 0.90, "BR08": 0.84, "BR09": 0.92, "BR10": 0.97, "BR11": 0.88, "BR12": 0.91,
    "BR13": 0.94, "BR14": 0.95, "BR15": 0.93, "BR16": 0.92,
}
WEEKDAY_VOLUME = [1.0, 1.05, 1.1, 1.08, 1.15, 0.7, 0.4]

# Errors seeded into the data as (branch, day of week, service, kind, value).
# Day 0 is Monday. Every entry appears in the answer key with its expected outcome.
FIX = "Auto-fixed"
EXCLUDE = "Excluded"
SEEDED_ERRORS = [
    # Safe to fix automatically.
    ("BR05", 1, "Standard", "branch_text", "  DELHI "),
    ("BR08", 2, "Express", "branch_text", "br-08"),
    ("BR04", 3, "Standard", "branch_text", "Chenai"),
    ("BR13", 2, "Express", "service_text", "EXP."),
    ("BR02", 4, "Standard", "service_text", "standard "),
    ("BR09", 0, "Standard", "service_text", "Std"),
    ("BR01", 2, "Standard", "number_as_text", "booked"),
    ("BR12", 3, "Express", "number_as_text", "revenue"),
    ("BR03", 1, "Express", "date_as_text", None),
    ("BR02", 2, "Express", "date_as_serial", None),
    # Rows that need a person to decide.
    ("BR03", 3, "Standard", "blank", "booked"),
    ("BR12", 1, "Standard", "blank", "revenue"),
    ("BR14", 2, "Standard", "blank", "fuel_cost"),
    ("BR02", 1, "Standard", "set", ("returned", -4)),
    ("BR09", 3, "Express", "set", ("fuel_cost", -1205.5)),
    ("BR07", 0, "Standard", "over_booked", None),
    ("BR15", 2, "Express", "over_booked", None),
    ("BR10", 2, "Standard", "on_time_over", None),
    ("BR06", 3, "Express", "on_time_over", None),
    ("BR01", 1, "Express", "date", date(2025, 11, 3)),
    ("BR08", 4, "Standard", "date", date(2026, 10, 30)),
    ("BR11", 4, "Express", "date", "TBC"),
    ("BR05", 3, "Express", "date", "31-Nov-2026"),
    ("BR12", 0, "Express", "branch_text", "Head Office"),
    ("BR09", 2, "Standard", "branch_text", "BR99"),
    ("BR04", 1, "Express", "set", ("delivered", "TBC")),
    ("BR13", 3, "Standard", "set", ("revenue", "n/a")),
    ("BR02", 3, "Express", "service_text", "Overnight"),
    ("BR05", 4, "Standard", "revenue_x100", None),
    ("BR15", 1, "Standard", "revenue_x100", None),
]
EXACT_DUPLICATES = [("BR06", 1, "Standard"), ("BR11", 2, "Standard"), ("BR13", 0, "Express")]
CONFLICTING_DUPLICATES = [("BR03", 4, "Express"), ("BR10", 0, "Express")]
MISSING_DAY = ("BR13", 6)  # Nagpur sends nothing for Sunday.

# Sample layouts, each mimicking how a different branch keeps its own report.
LAYOUTS = {
    "standard": ["BR01", "BR02", "BR03"],
    "banner": ["BR04", "BR05", "BR06"],
    "csv_comma": ["BR07", "BR08", "BR09"],
    "csv_semicolon": ["BR10", "BR11"],
    "cover_sheet": ["BR12", "BR13"],
    "wide": ["BR14", "BR15"],
}
FILE_NAMES = {
    "BR01": "BR01_weekly_ops.xlsx",
    "BR02": "BR02_weekly_ops.xlsx",
    "BR03": "BR03 Chandigarh weekly.xlsx",
    "BR04": "Chennai Ops Report FINAL.xlsx",
    "BR05": "delhi-report (1).xlsx",
    "BR06": "Hyderabad_week45_v2.xlsx",
    "BR07": "indore_wk45.csv",
    "BR08": "BR08_export.csv",
    "BR09": "kochi weekly export.csv",
    "BR10": "branch10_report.csv",
    "BR11": "LKO_station_report.csv",
    "BR12": "Mumbai MIS.xlsx",
    "BR13": "Nagpur MIS new.xlsx",
    "BR14": "patna_side_by_side.xlsx",
    "BR15": "Pune_Weekly.xlsx",
}


def make_row(rng, code, day, service, week_start, scale=1.0):
    """Build one internally consistent row of branch activity."""
    base = 120 + (int(code[2:]) * 37) % 260
    service_mix = 0.35 if service == "Express" else 1.0
    booked = max(5, round(base * service_mix * WEEKDAY_VOLUME[day] * scale * rng.uniform(0.9, 1.1)))
    returned = round(booked * rng.uniform(0.01, 0.055))
    delivered = min(booked - returned, round(booked * rng.uniform(0.9, 0.97)))
    on_time = min(delivered, round(delivered * ON_TIME_RATE[code] * rng.uniform(0.97, 1.03)))
    price = rng.uniform(95, 130) if service == "Express" else rng.uniform(48, 70)
    return {
        "date": week_start + timedelta(days=day),
        "code": code,
        "branch": code,
        "service": service,
        "booked": booked,
        "delivered": delivered,
        "on_time": on_time,
        "returned": returned,
        "revenue": round(booked * price, 2),
        "fuel_cost": round(delivered * rng.uniform(7, 13), 2),
    }


def apply_error(row, kind, value):
    """Corrupt a row in place and return the answer-key field name and outcome."""
    if kind == "branch_text":
        row["branch"] = value
        return "branch", EXCLUDE if value in ("Head Office", "BR99") else FIX
    if kind == "service_text":
        row["service"] = value
        return "service", EXCLUDE if value == "Overnight" else FIX
    if kind == "number_as_text":
        row[value] = f" {row[value]:,} "
        return value, FIX
    if kind == "date_as_text":
        row["date"] = row["date"].strftime("%d %b %Y")
        return "date", FIX
    if kind == "date_as_serial":
        row["date"] = (row["date"] - date(1899, 12, 30)).days
        return "date", FIX
    if kind == "blank":
        row[value] = None
        return value, EXCLUDE
    if kind == "set":
        field, new_value = value
        row[field] = new_value
        return field, EXCLUDE
    if kind == "over_booked":
        row["delivered"] = row["booked"] + 12
        return "delivered", EXCLUDE
    if kind == "on_time_over":
        row["on_time"] = row["delivered"] + 9
        return "on_time", EXCLUDE
    if kind == "date":
        row["date"] = value
        return "date", EXCLUDE
    if kind == "revenue_x100":
        row["revenue"] = round(row["revenue"] * 100, 2)
        return "revenue", EXCLUDE
    raise ValueError(f"Unknown error kind: {kind}")


def build_week(rng, week_start):
    """Create the clean rows for every branch and return them keyed for error seeding."""
    rows = {}
    for code, _, _ in BRANCHES:
        for day in range(7):
            for service in SERVICES:
                rows[(code, day, service)] = make_row(rng, code, day, service, week_start)
    return rows


def seed_errors(rows):
    """Plant every seeded error and return the rows per branch plus the answer key."""
    key = []
    for code, day, service, kind, value in SEEDED_ERRORS:
        field, outcome = apply_error(rows[(code, day, service)], kind, value)
        key.append([code, day, service, kind, field, outcome])

    by_branch = {code: [] for code, _, _ in BRANCHES}
    for (code, day, service), row in rows.items():
        if (code, day) != MISSING_DAY:
            by_branch[code].append(row)

    for code, day, service in EXACT_DUPLICATES:
        by_branch[code].append(dict(rows[(code, day, service)]))
        key.append([code, day, service, "exact_duplicate", "", "Removed duplicate"])

    for code, day, service in CONFLICTING_DUPLICATES:
        clone = dict(rows[(code, day, service)])
        clone["booked"] += 25
        by_branch[code].append(clone)
        key.append([code, day, service, "conflicting_duplicate", "booked", EXCLUDE])

    key.append([MISSING_DAY[0], MISSING_DAY[1], "", "missing_day", "", "Follow up"])
    key.append([MISSING_FILE_BRANCH, "", "", "missing_file", "", "Follow up"])
    return by_branch, key


def text_date(value, pattern):
    return value.strftime(pattern) if isinstance(value, date) else value


def money_text(value):
    return f"₹{value:,.2f}" if isinstance(value, (int, float)) else value


def export_value(value):
    return "" if value is None else value


def label(row, use_name):
    """Show a clean branch as its name when the file uses names, keep seeded text as-is."""
    if row["branch"] == row["code"]:
        return NAME_BY_CODE[row["code"]] if use_name else row["code"]
    return row["branch"]


def body_row(row, date_value, branch_value, money=lambda v: v):
    return [
        date_value,
        branch_value,
        row["service"],
        row["booked"],
        row["delivered"],
        row["on_time"],
        row["returned"],
        money(row["revenue"]),
        money(row["fuel_cost"]),
    ]


def write_standard(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Date", "Branch Code", "Service", "Booked", "Delivered", "On Time",
               "Returned", "Revenue", "Fuel Cost"])
    for row in rows:
        ws.append(body_row(row, row["date"], label(row, use_name=False)))
    for cell in ws["A"][1:]:
        if isinstance(cell.value, date):
            cell.number_format = "yyyy-mm-dd"
    wb.save(path)


def write_banner(path, rows):
    """Title rows above the table and a totals row beneath it."""
    code = rows[0]["code"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly"
    ws.append([f"{NAME_BY_CODE[code]} Branch - Weekly Operations Report"])
    ws.append(["Prepared by branch admin team"])
    ws.append([])
    ws.append(["Report Date", "Branch", "Service Type", "Parcels Booked", "Parcels Delivered",
               "Delivered On Time", "Parcels Returned", "Revenue (Rs)", "Fuel (Rs)"])
    for row in rows:
        ws.append(body_row(row, text_date(row["date"], "%d-%b-%Y"), label(row, use_name=True)))
    numeric = [row for row in rows if isinstance(row["booked"], int)]
    ws.append(["Total", "", "", sum(row["booked"] for row in numeric)])
    wb.save(path)


def write_csv_comma(path, rows):
    """Comma-separated export with US month-first dates."""
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["date", "branch_id", "product", "booked_qty", "delivered_qty",
                         "ontime_qty", "returned_qty", "sales", "fuel_spend"])
        for row in rows:
            cells = body_row(row, text_date(row["date"], "%m/%d/%Y"), label(row, use_name=False))
            writer.writerow([export_value(cell) for cell in cells])


def write_csv_semicolon(path, rows):
    """Semicolon-separated export with day-first dates and rupee-formatted money."""
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(["Dt", "Stn", "Svc", "Bkd", "Dlvd", "OT", "Rtn", "Rev", "Fuel"])
        for row in rows:
            cells = body_row(row, text_date(row["date"], "%d/%m/%Y"), label(row, use_name=False),
                             money=money_text)
            writer.writerow([export_value(cell) for cell in cells])


def write_cover_sheet(path, rows):
    """The data sits on a second sheet, behind a cover page."""
    code = rows[0]["code"]
    wb = Workbook()
    cover = wb.active
    cover.title = "Cover"
    cover.append([f"{NAME_BY_CODE[code]} Weekly MIS"])
    cover.append(["See the Daily Ops sheet for figures."])
    ws = wb.create_sheet("Daily Ops")
    ws.append(["Day", "Location", "Shipment Class", "Consignments Booked",
               "Consignments Delivered", "On-Time Deliveries", "Returns", "Gross Revenue",
               "Fuel Expense"])
    for row in rows:
        ws.append(body_row(row, text_date(row["date"], "%Y-%m-%d"), label(row, use_name=True)))
    wb.save(path)


def write_wide(path, rows):
    """One row per day, with Express and Standard columns side by side."""
    measures = [("Booked", "booked"), ("Delivered", "delivered"), ("On Time", "on_time"),
                ("Returned", "returned"), ("Revenue", "revenue"), ("Fuel", "fuel_cost")]
    wb = Workbook()
    ws = wb.active
    ws.title = "Side by side"
    ws.append(["Date", "Branch"]
              + [f"Exp {title}" for title, _ in measures]
              + [f"Std {title}" for title, _ in measures])
    days = {}
    for row in rows:
        days.setdefault(row["date"], {})[row["service"]] = row
    for day, pair in sorted(days.items()):
        express, standard = pair["Express"], pair["Standard"]
        ws.append([day, NAME_BY_CODE[express["code"]]]
                  + [express[field] for _, field in measures]
                  + [standard[field] for _, field in measures])
    for cell in ws["A"][1:]:
        cell.number_format = "dd/mm/yyyy"
    wb.save(path)


WRITERS = {
    "standard": write_standard,
    "banner": write_banner,
    "csv_comma": write_csv_comma,
    "csv_semicolon": write_csv_semicolon,
    "cover_sheet": write_cover_sheet,
    "wide": write_wide,
}


def write_reference_files():
    BRANCH_MASTER.parent.mkdir(parents=True, exist_ok=True)
    with open(BRANCH_MASTER, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["branch_code", "branch_name", "region"])
        writer.writerows(BRANCHES)


def write_prior_week_history(rng, week_start):
    """Seed last week's clean rows so the first report can show week-on-week change."""
    fields = ["date", "branch_code", "service", "booked", "delivered", "on_time", "returned",
              "revenue", "fuel_cost"]
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(HISTORY_FILE, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(fields)
        for code, _, _ in BRANCHES:
            for day in range(7):
                for service in SERVICES:
                    row = make_row(rng, code, day, service, week_start,
                                   scale=rng.uniform(0.93, 1.03))
                    writer.writerow([row["date"].isoformat(), code] + [
                        row[field] for field in fields[2:]
                    ])


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--week-start", default="2026-11-02",
                        help="Monday of the reporting week (YYYY-MM-DD)")
    args = parser.parse_args()
    week_start = datetime.strptime(args.week_start, "%Y-%m-%d").date()

    rng = random.Random(SEED)
    target_dir = Path(INPUT_DIR) / week_start.isoformat()
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.mkdir(parents=True)

    write_reference_files()
    rows_by_branch, answer_key = seed_errors(build_week(rng, week_start))

    for layout, codes in LAYOUTS.items():
        for code in codes:
            WRITERS[layout](target_dir / FILE_NAMES[code], rows_by_branch[code])

    with open(ANSWER_KEY, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["branch_code", "day", "service", "error_kind", "field",
                         "expected_outcome"])
        writer.writerows(answer_key)

    write_prior_week_history(rng, week_start - timedelta(days=7))
    print(f"Wrote {sum(len(c) for c in LAYOUTS.values())} branch files to {target_dir}")
    print(f"Seeded {len(answer_key)} known issues, answer key at {ANSWER_KEY}")


if __name__ == "__main__":
    main()
