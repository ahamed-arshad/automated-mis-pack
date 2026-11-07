"""Write the formatted MIS pack workbook with openpyxl.

KPI cells are live Excel formulas over the Clean Data sheet, so any figure can be traced
by clicking it. Named ranges (cd_booked, th_ontime and so on) keep the formulas readable.
"""

from collections import Counter
from datetime import timedelta

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

import config

FONT = "Arial"
NAVY = "1F3A5F"
PALE = "EAF0F7"
GRID = "D0D7DE"
PRIOR_HEADER = "5B7083"
CHART_BLUE = "2A78D6"
CHART_ORANGE = "EB6834"
CHART_RED = "E34948"
CHART_GRID = "E5E7EB"

COUNT_FORMAT = "#,##0"
MONEY_FORMAT = '"₹"#,##0'
UNIT_MONEY_FORMAT = '"₹"#,##0.00'
RATE_FORMAT = "0.0%"
CHANGE_FORMAT = "+0.0%;-0.0%;0.0%"
POINTS_FORMAT = '+0.0" pts";-0.0" pts";0.0" pts"'
DATE_FORMAT = "ddd dd mmm yyyy"

ACTION_ORDER = {"Follow up": 0, "Excluded": 1, "Removed duplicate": 2, "Auto-fixed": 3}
ACTION_FILLS = {
    "Follow up": ("FFC7CE", "9C0006"),
    "Excluded": ("FFEB9C", "9C5700"),
    "Removed duplicate": ("E7E6E6", "404040"),
    "Auto-fixed": ("C6EFCE", "006100"),
}
OPEN_ACTIONS = {"Follow up", "Excluded"}
THIN = Side(style="thin", color=GRID)

CLEAN_COLUMNS = [
    ("Date", "date", 18, DATE_FORMAT, "cd_date"),
    ("Branch Code", "branch_code", 12, None, "cd_branch"),
    ("Branch Name", "branch_name", 15, None, None),
    ("Region", "region", 12, None, "cd_region"),
    ("Service", "service", 11, None, "cd_service"),
    ("Booked", "booked", 10, COUNT_FORMAT, "cd_booked"),
    ("Delivered", "delivered", 10, COUNT_FORMAT, "cd_delivered"),
    ("On Time", "on_time", 10, COUNT_FORMAT, "cd_ontime"),
    ("Returned", "returned", 10, COUNT_FORMAT, "cd_returned"),
    ("Revenue", "revenue", 12, UNIT_MONEY_FORMAT, "cd_revenue"),
    ("Fuel Cost", "fuel_cost", 12, UNIT_MONEY_FORMAT, "cd_fuel"),
    ("LW Booked", "lw_booked", 11, COUNT_FORMAT, "cd_lw_booked"),
    ("LW Delivered", "lw_delivered", 12, COUNT_FORMAT, "cd_lw_delivered"),
    ("LW On Time", "lw_on_time", 11, COUNT_FORMAT, "cd_lw_ontime"),
    ("LW Returned", "lw_returned", 12, COUNT_FORMAT, "cd_lw_returned"),
    ("LW Revenue", "lw_revenue", 12, UNIT_MONEY_FORMAT, "cd_lw_revenue"),
    ("LW Fuel Cost", "lw_fuel_cost", 12, UNIT_MONEY_FORMAT, "cd_lw_fuel"),
    ("Source File", "source_file", 32, None, None),
    ("Source Row", "source_row", 11, None, None),
]

EXCEPTION_COLUMNS = [
    ("Action", 18), ("Issue", 32), ("Branch", 13), ("Date", 13), ("Service", 11), ("Field", 11),
    ("Original Value", 18), ("Corrected Value", 16), ("Detail", 58), ("Source File", 32),
    ("Source Row", 11), ("Follow-Up Status", 16),
]

KPI_DEFINITIONS = [
    ("Delivery Rate", "Delivered / Booked", "Share of booked parcels delivered."),
    ("On-Time %", "On Time / Delivered", "Measured against delivered parcels, not booked."),
    ("Return Rate", "Returned / Booked", "Lower is better."),
    ("Fuel / Delivered", "Fuel Cost / Delivered", "Fuel spend for each parcel delivered."),
    ("Revenue / Booked", "Revenue / Booked", "Average revenue for each parcel booked."),
    ("Coverage", "Rows Used / Rows Expected",
     "Shows how complete the week's data is. Volume totals are understated below 100%."),
    ("Week-on-week", "This Week vs same rows last week",
     "Compares only the branch, weekday and service rows present this week, so a missing "
     "row does not look like a fall in volume."),
]

VALIDATION_RULES = [
    ("Auto-fixed", "Date stored differently from the rest of the file; branch or service spelling, "
                   "case or spacing; numbers stored as text or with ₹ and commas."),
    ("Removed duplicate", "Exact repeat of a row. The first copy is kept."),
    ("Excluded", "Missing or non-numeric figures; invalid dates or dates outside the week; "
                 "unknown branch or service; negative values; part-parcel counts; delivered + "
                 "returned above booked; on-time above delivered; revenue per parcel outside "
                 "the allowed range; duplicate rows with conflicting figures."),
    ("Follow up", "Branch with no usable data, a whole day missing, or an unreadable file."),
]

ASSUMPTIONS = [
    "All branch data in this sample is synthetic.",
    "Each row is one branch, day and service. Figures describe parcels booked that day, so "
    "delivered + returned cannot exceed booked.",
    "Blue cells above are inputs. Changing one updates the KPI statuses in this workbook only. "
    "Update config.py to carry a change into future weeks.",
    "Excluded rows stay out of every KPI until the branch corrects its file and the pack is re-run.",
]


def style_title(ws, title, subtitle=None):
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = Font(name=FONT, size=10, color="52514E")


def style_section(ws, row, text):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name=FONT, size=12, bold=True, color=NAVY)


def style_header(ws, row, headers, fill=NAVY, start_column=1):
    for offset, heading in enumerate(headers):
        cell = ws.cell(row=row, column=start_column + offset, value=heading)
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_cell(ws, row, column, value, number_format=None, bold=False, fill=None, align=None):
    cell = ws.cell(row=row, column=column, value=value)
    cell.font = Font(name=FONT, size=10, bold=bold)
    cell.border = Border(bottom=THIN)
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if align:
        cell.alignment = Alignment(horizontal=align, vertical="center")
    return cell


def write_text(ws, row, column, value):
    """Write text from a branch file so it can never run as a formula in the report."""
    text = ILLEGAL_CHARACTERS_RE.sub("", str(value))
    cell = write_cell(ws, row, column, text)
    cell.data_type = "s"
    return cell


def set_widths(ws, widths):
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def set_print_layout(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def add_name(wb, name, reference):
    wb.defined_names[name] = DefinedName(name, attr_text=reference)


def add_status_colors(ws, cell_range):
    rules = [
        ('"On target"', "C6EFCE", "006100"),
        ('"OK"', "C6EFCE", "006100"),
        ('"Watch"', "FFEB9C", "9C5700"),
        ('"Low coverage"', "FFEB9C", "9C5700"),
        ('"Below target"', "FFC7CE", "9C0006"),
        ('"Above limit"', "FFC7CE", "9C0006"),
        ('"CHECK"', "FFC7CE", "9C0006"),
        ('"No data"', "E7E6E6", "404040"),
    ]
    for value, fill, text in rules:
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="equal", formula=[value],
                       fill=PatternFill("solid", bgColor=fill),
                       font=Font(color=text, bold=True)),
        )


def write_clean_data(wb, clean):
    ws = wb["Clean Data"]
    style_header(ws, 1, [column[0] for column in CLEAN_COLUMNS])
    for index, (heading, _, _, _, _) in enumerate(CLEAN_COLUMNS, start=1):
        if heading.startswith("LW "):
            ws.cell(row=1, column=index).fill = PatternFill("solid", fgColor=PRIOR_HEADER)

    for row_number, row in enumerate(clean.to_dict("records"), start=2):
        for index, (_, key, _, number_format, _) in enumerate(CLEAN_COLUMNS, start=1):
            value = row[key]
            if key == "date":
                value = value.to_pydatetime()
            elif key.startswith("lw_") and key[3:] in config.COUNT_COLUMNS:
                value = int(value)
            if isinstance(value, str):
                write_text(ws, row_number, index, value)
            else:
                write_cell(ws, row_number, index, value, number_format)

    last_row = max(len(clean) + 1, 2)
    for index, (_, _, width, _, name) in enumerate(CLEAN_COLUMNS, start=1):
        letter = get_column_letter(index)
        ws.column_dimensions[letter].width = width
        if name:
            add_name(wb, name, f"'Clean Data'!${letter}$2:${letter}${last_row}")
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CLEAN_COLUMNS))}{last_row}"
    ws.freeze_panes = "C2"
    ws.row_dimensions[1].height = 30
    set_print_layout(ws)


def write_definitions(wb):
    ws = wb["Definitions"]
    style_title(ws, "Definitions and Thresholds")
    set_widths(ws, [26, 44, 70])

    style_section(ws, 3, "Thresholds")
    style_header(ws, 4, ["Setting", "Value", "Used For"])
    settings = [
        ("On-time target", config.THRESHOLDS["on_time_target"], RATE_FORMAT, "th_ontime",
         "At or above this is On target."),
        ("Watch band below target", config.THRESHOLDS["on_time_tolerance"], RATE_FORMAT,
         "th_tolerance", "Within this distance of target is Watch, further below is Below target."),
        ("Return rate limit", config.THRESHOLDS["return_rate_limit"], RATE_FORMAT, "th_return",
         "Network return rate above this is Above limit."),
        ("Minimum coverage", config.THRESHOLDS["min_coverage"], RATE_FORMAT, "th_coverage",
         "A branch below this shows Low coverage."),
        ("Rows expected per branch", config.DAYS_IN_WEEK * len(config.SERVICES), "0", "exp_rows",
         "Days in the week multiplied by the number of services."),
    ]
    for offset, (label, value, number_format, name, note) in enumerate(settings):
        row = 5 + offset
        write_cell(ws, row, 1, label)
        cell = write_cell(ws, row, 2, value, number_format, align="center")
        cell.font = Font(name=FONT, size=10, color="0000FF")
        write_cell(ws, row, 3, note)
        add_name(wb, name, f"Definitions!$B${row}")

    row = 11
    style_section(ws, row, "KPI Definitions")
    style_header(ws, row + 1, ["KPI", "Calculation", "Notes"])
    row += 2
    for label, calculation, note in KPI_DEFINITIONS:
        write_cell(ws, row, 1, label).alignment = Alignment(vertical="top")
        write_cell(ws, row, 2, calculation).alignment = Alignment(vertical="top")
        write_cell(ws, row, 3, note).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    row += 1
    style_section(ws, row, "How Rows Are Treated")
    style_header(ws, row + 1, ["Action", "Applies To"])
    ws.merge_cells(start_row=row + 1, start_column=2, end_row=row + 1, end_column=3)
    row += 2
    for action, rule in VALIDATION_RULES:
        write_cell(ws, row, 1, action).alignment = Alignment(vertical="top")
        write_cell(ws, row, 2, rule).alignment = Alignment(wrap_text=True, vertical="top")
        write_cell(ws, row, 3, None)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.row_dimensions[row].height = 44 if len(rule) > 150 else 18
        row += 1

    row += 1
    style_section(ws, row, "Assumptions")
    row += 1
    for note in ASSUMPTIONS:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell = ws.cell(row=row, column=1, value=note)
        cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1
    set_print_layout(ws)


def write_exceptions(wb, issues):
    ws = wb["Exceptions"]
    style_title(ws, "Exceptions",
                "Every change or rejection made to branch data. Filter by Action, and use "
                "Follow-Up Status to track what has been chased.")
    header_row = 4
    style_header(ws, header_row, [heading for heading, _ in EXCEPTION_COLUMNS])
    set_widths(ws, [width for _, width in EXCEPTION_COLUMNS])

    ordered = sorted(issues, key=lambda i: (ACTION_ORDER[i.action], str(i.branch),
                                            str(i.date), i.source_file, str(i.source_row)))
    for row, issue in enumerate(ordered, start=header_row + 1):
        status = "Open" if issue.action in OPEN_ACTIONS else "No action"
        fill, text = ACTION_FILLS[issue.action]
        cell = write_text(ws, row, 1, issue.action)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.font = Font(name=FONT, size=10, bold=True, color=text)
        write_text(ws, row, 2, issue.issue)
        write_text(ws, row, 3, issue.branch)
        write_cell(ws, row, 4, issue.date or None, "dd mmm yyyy", align="center")
        write_text(ws, row, 5, issue.service)
        write_text(ws, row, 6, issue.field)
        for column, value in ((7, issue.original), (8, issue.corrected)):
            if isinstance(value, (int, float)):
                is_whole = float(value).is_integer()
                write_cell(ws, row, column, value, COUNT_FORMAT if is_whole else "#,##0.00")
            else:
                write_text(ws, row, column, value)
        write_text(ws, row, 9, issue.detail).alignment = Alignment(wrap_text=True, vertical="top")
        write_text(ws, row, 10, issue.source_file)
        write_cell(ws, row, 11, issue.source_row if issue.source_row != "" else None,
                   align="center")
        write_text(ws, row, 12, status)

    last_row = max(header_row + len(ordered), header_row + 1)
    tracker = DataValidation(type="list", formula1='"Open,Chased,Resolved,No action"',
                             allow_blank=True, showErrorMessage=True,
                             errorTitle="Follow-Up Status", error="Pick a status from the list.")
    ws.add_data_validation(tracker)
    tracker.add(f"L{header_row + 1}:L{last_row}")

    ws.auto_filter.ref = f"A{header_row}:L{last_row}"
    ws.freeze_panes = ws.cell(row=header_row + 1, column=3)
    add_name(wb, "ex_action", f"Exceptions!$A${header_row + 1}:$A${last_row}")
    add_name(wb, "ex_issue", f"Exceptions!$B${header_row + 1}:$B${last_row}")
    set_print_layout(ws)


def write_branch_kpis(wb, master):
    ws = wb["Branch KPIs"]
    style_title(ws, "Branch KPIs", "Every figure is a live formula over the Clean Data sheet. "
                                   "Check Coverage before comparing volumes.")
    headers = [
        "Code", "Branch", "Region", "Rows Used", "Coverage", "Data Check", "Booked", "Delivered",
        "On Time", "Returned", "Revenue", "Fuel Cost", "Delivery Rate", "On-Time %",
        "Gap to Target", "Return Rate", "Fuel / Delivered", "Revenue / Booked", "Booked WoW",
        "On-Time WoW", "Status", "Rank Key",
    ]
    header_row = 4
    style_header(ws, header_row, headers)
    ws.row_dimensions[header_row].height = 32
    set_widths(ws, [8, 14, 11, 7, 10, 12, 9, 10, 9, 9, 11, 10, 10, 9, 11, 8, 10, 10, 9, 11, 13, 10])

    first = header_row + 1
    branches = master.sort_values("branch_code").to_dict("records")
    last = first + len(branches) - 1
    total = last + 1

    for row, branch in enumerate(branches, start=first):
        sums = {
            "G": "cd_booked", "H": "cd_delivered", "I": "cd_ontime", "J": "cd_returned",
            "K": "cd_revenue", "L": "cd_fuel",
        }
        write_cell(ws, row, 1, branch["branch_code"])
        write_cell(ws, row, 2, branch["branch_name"])
        write_cell(ws, row, 3, branch["region"])
        write_cell(ws, row, 4, f"=COUNTIFS(cd_branch,$A{row})", COUNT_FORMAT, align="center")
        write_cell(ws, row, 5, f"=D{row}/exp_rows", RATE_FORMAT, align="center")
        write_cell(ws, row, 6, f'=IF(E{row}<th_coverage,"Low coverage","OK")', align="center")
        for letter, name in sums.items():
            column = ord(letter) - 64
            number_format = MONEY_FORMAT if letter in "KL" else COUNT_FORMAT
            write_cell(ws, row, column, f"=SUMIFS({name},cd_branch,$A{row})", number_format)
        write_ratio_cells(ws, row)

        lw_booked = f"SUMIFS(cd_lw_booked,cd_branch,$A{row})"
        lw_delivered = f"SUMIFS(cd_lw_delivered,cd_branch,$A{row})"
        lw_on_time = f"SUMIFS(cd_lw_ontime,cd_branch,$A{row})"
        write_cell(ws, row, 19,
                   f'=IF({lw_booked}=0,"n/a",'
                   f'SUMIFS(cd_booked,cd_branch,$A{row},cd_lw_booked,">0")/{lw_booked}-1)',
                   CHANGE_FORMAT, align="right")
        write_cell(ws, row, 20,
                   f'=IF(OR(NOT(ISNUMBER(N{row})),{lw_delivered}=0),"n/a",'
                   f"(N{row}-{lw_on_time}/{lw_delivered})*100)",
                   POINTS_FORMAT, align="right")
        write_cell(ws, row, 21,
                   f'=IF(NOT(ISNUMBER(N{row})),"No data",IF(N{row}>=th_ontime,"On target",'
                   f'IF(N{row}>=th_ontime-th_tolerance,"Watch","Below target")))',
                   align="center")
        write_cell(ws, row, 22, f'=IF(ISNUMBER(N{row}),N{row}+ROW()/1000000000,"")', "0.000000")

    write_cell(ws, total, 1, "Network", bold=True, fill=PALE)
    for column in (2, 3):
        write_cell(ws, total, column, None, fill=PALE)
    write_cell(ws, total, 4, f"=SUM(D{first}:D{last})", COUNT_FORMAT, bold=True, fill=PALE,
               align="center")
    write_cell(ws, total, 5, f"=D{total}/(exp_rows*COUNTA(A{first}:A{last}))", RATE_FORMAT,
               bold=True, fill=PALE, align="center")
    write_cell(ws, total, 6, f'=IF(E{total}<th_coverage,"Low coverage","OK")', bold=True,
               fill=PALE, align="center")
    for column in range(7, 13):
        letter = get_column_letter(column)
        number_format = MONEY_FORMAT if letter in "KL" else COUNT_FORMAT
        write_cell(ws, total, column, f"=SUM({letter}{first}:{letter}{last})", number_format,
                   bold=True, fill=PALE)
    write_ratio_cells(ws, total, bold=True, fill=PALE)
    write_cell(ws, total, 19,
               '=IF(SUM(cd_lw_booked)=0,"n/a",'
               'SUMIFS(cd_booked,cd_lw_booked,">0")/SUM(cd_lw_booked)-1)',
               CHANGE_FORMAT, bold=True, fill=PALE, align="right")
    write_cell(ws, total, 20,
               f'=IF(OR(NOT(ISNUMBER(N{total})),SUM(cd_lw_delivered)=0),"n/a",'
               f"(N{total}-SUM(cd_lw_ontime)/SUM(cd_lw_delivered))*100)",
               POINTS_FORMAT, bold=True, fill=PALE, align="right")
    write_cell(ws, total, 21,
               f'=IF(NOT(ISNUMBER(N{total})),"No data",IF(N{total}>=th_ontime,"On target",'
               f'IF(N{total}>=th_ontime-th_tolerance,"Watch","Below target")))',
               bold=True, fill=PALE, align="center")

    ws.column_dimensions["V"].hidden = True
    add_status_colors(ws, f"U{first}:U{total}")
    add_status_colors(ws, f"F{first}:F{total}")
    ws.freeze_panes = ws.cell(row=first, column=3)
    add_name(wb, "kp_branch", f"'Branch KPIs'!$B${first}:$B${last}")
    add_name(wb, "kp_ontime", f"'Branch KPIs'!$N${first}:$N${last}")
    add_name(wb, "kp_status", f"'Branch KPIs'!$U${first}:$U${last}")
    add_name(wb, "kp_rank", f"'Branch KPIs'!$V${first}:$V${last}")
    set_print_layout(ws)
    return first, last, total


def write_ratio_cells(ws, row, bold=False, fill=None):
    """Write the rate columns, which are built from totals rather than averaged."""
    formulas = [
        (13, f'=IF(G{row}=0,"n/a",H{row}/G{row})', RATE_FORMAT),
        (14, f'=IF(H{row}=0,"n/a",I{row}/H{row})', RATE_FORMAT),
        (15, f'=IF(ISNUMBER(N{row}),(N{row}-th_ontime)*100,"n/a")', POINTS_FORMAT),
        (16, f'=IF(G{row}=0,"n/a",J{row}/G{row})', RATE_FORMAT),
        (17, f'=IF(H{row}=0,"n/a",L{row}/H{row})', UNIT_MONEY_FORMAT),
        (18, f'=IF(G{row}=0,"n/a",K{row}/G{row})', UNIT_MONEY_FORMAT),
    ]
    for column, formula, number_format in formulas:
        write_cell(ws, row, column, formula, number_format, bold=bold, fill=fill, align="right")


def write_metric_table(ws, top, title, label, items, criteria, label_format=None, total=True):
    """Write one breakdown table and return its first and last data rows."""
    style_section(ws, top, title)
    style_header(ws, top + 1, [label, "Booked", "Delivered", "On Time", "Returned", "Revenue",
                               "Fuel Cost", "Delivery Rate", "On-Time %", "Return Rate"])
    first = top + 2
    measures = ["cd_booked", "cd_delivered", "cd_ontime", "cd_returned", "cd_revenue", "cd_fuel"]
    for row, item in enumerate(items, start=first):
        write_cell(ws, row, 1, item, label_format, align="left")
        for column, measure in enumerate(measures, start=2):
            number_format = MONEY_FORMAT if column >= 6 else COUNT_FORMAT
            write_cell(ws, row, column, f"=SUMIFS({measure},{criteria},$A{row})", number_format)
        write_rate_cells(ws, row)
    last = first + len(items) - 1

    if total:
        row = last + 1
        write_cell(ws, row, 1, "Total", bold=True, fill=PALE)
        for column in range(2, 8):
            letter = get_column_letter(column)
            number_format = MONEY_FORMAT if column >= 6 else COUNT_FORMAT
            write_cell(ws, row, column, f"=SUM({letter}{first}:{letter}{last})", number_format,
                       bold=True, fill=PALE)
        write_rate_cells(ws, row, bold=True, fill=PALE)
    return first, last


def write_rate_cells(ws, row, bold=False, fill=None):
    write_cell(ws, row, 8, f'=IF(B{row}=0,"n/a",C{row}/B{row})', RATE_FORMAT, bold, fill, "right")
    write_cell(ws, row, 9, f'=IF(C{row}=0,"n/a",D{row}/C{row})', RATE_FORMAT, bold, fill, "right")
    write_cell(ws, row, 10, f'=IF(B{row}=0,"n/a",E{row}/B{row})', RATE_FORMAT, bold, fill, "right")


def write_breakdowns(wb, master, issues, week_start, ranked_count):
    ws = wb["Breakdowns"]
    style_title(ws, "Breakdowns", "Region, service and day views, plus the branch ranking and "
                                  "exception counts.")
    set_widths(ws, [20, 13, 11, 11, 11, 11, 11, 10, 10, 10])

    regions = sorted(master["region"].unique())
    first, last = write_metric_table(ws, 4, "By Region", "Region", regions, "cd_region")

    top = last + 4
    first, last = write_metric_table(ws, top, "By Service", "Service", config.SERVICES,
                                     "cd_service")

    top = last + 4
    days = [week_start + timedelta(days=offset) for offset in range(config.DAYS_IN_WEEK)]
    day_first, day_last = write_metric_table(ws, top, "By Day", "Date", days, "cd_date",
                                             label_format="ddd dd mmm")
    day_header = day_first - 1

    top = day_last + 4
    style_section(ws, top, "On-Time Ranking")
    header = top + 1
    style_header(ws, header, ["Rank", "Branch", "On-Time %", "Gap to Target", "Above Target",
                              "Below Target"])
    for offset in range(ranked_count):
        row = header + 1 + offset
        position = f"MATCH(LARGE(kp_rank,$A{row}),kp_rank,0)"
        write_cell(ws, row, 1, offset + 1, align="left")
        write_cell(ws, row, 2, f"=INDEX(kp_branch,{position})")
        write_cell(ws, row, 3, f"=INDEX(kp_ontime,{position})", RATE_FORMAT)
        write_cell(ws, row, 4, f"=(C{row}-th_ontime)*100", POINTS_FORMAT)
        write_cell(ws, row, 5, f"=MAX(D{row},0)", '0.0;-0.0;""')
        write_cell(ws, row, 6, f"=MIN(D{row},0)", '0.0;-0.0;""')
    rank_last = header + ranked_count

    top = rank_last + 3 if ranked_count else header + 2
    style_section(ws, top, "Exceptions by Issue")
    style_header(ws, top + 1, ["Action", "Issue", None, None, "Count"])
    ws.merge_cells(start_row=top + 1, start_column=2, end_row=top + 1, end_column=4)
    counts = Counter((issue.action, issue.issue) for issue in issues)
    ordered = sorted(counts, key=lambda key: (ACTION_ORDER[key[0]], -counts[key], key[1]))
    for row, (action, issue_name) in enumerate(ordered, start=top + 2):
        write_cell(ws, row, 1, action)
        for column in (2, 3, 4):
            write_cell(ws, row, column, issue_name if column == 2 else None)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        write_cell(ws, row, 5, f"=COUNTIFS(ex_action,$A{row},ex_issue,$B{row})", COUNT_FORMAT)
    set_print_layout(ws)
    return {"day_header": day_header, "day_last": day_last, "rank_header": header,
            "rank_last": rank_last}


def write_summary(wb, stats, week_start, kpi_rows, ranked_count):
    ws = wb["Summary"]
    ws.sheet_view.showGridLines = False
    week_end = week_start + timedelta(days=config.DAYS_IN_WEEK - 1)
    style_title(ws, "Weekly Operations MIS Pack",
                f"Week of {week_start:%d %b} to {week_end:%d %b %Y}   |   "
                f"Generated {stats['generated']:%d %b %Y %H:%M}")
    set_widths(ws, [34, 13, 13, 12, 11, 14, 3])

    first, last, total = kpi_rows
    kpi = "'Branch KPIs'!"
    style_section(ws, 5, "Network KPIs")
    style_header(ws, 6, ["KPI", "This Week", "Last Week", "Change", "Target", "Status"])

    def like_for_like(measure, lw_measure):
        return f'=IF(C{{r}}=0,"n/a",SUMIFS({measure},{lw_measure},">0")/C{{r}}-1)'

    rows = [
        ("Parcels booked", f"={kpi}G{total}", "=SUM(cd_lw_booked)",
         like_for_like("cd_booked", "cd_lw_booked"), COUNT_FORMAT, CHANGE_FORMAT, None, None),
        ("Parcels delivered", f"={kpi}H{total}", "=SUM(cd_lw_delivered)",
         like_for_like("cd_delivered", "cd_lw_delivered"), COUNT_FORMAT, CHANGE_FORMAT, None, None),
        ("Delivery rate", f"={kpi}M{total}",
         '=IF(SUM(cd_lw_booked)=0,"n/a",SUM(cd_lw_delivered)/SUM(cd_lw_booked))',
         "points", RATE_FORMAT, POINTS_FORMAT, None, None),
        ("On-time delivery", f"={kpi}N{total}",
         '=IF(SUM(cd_lw_delivered)=0,"n/a",SUM(cd_lw_ontime)/SUM(cd_lw_delivered))',
         "points", RATE_FORMAT, POINTS_FORMAT, "=th_ontime",
         '=IF(B{r}>=E{r},"On target",IF(B{r}>=E{r}-th_tolerance,"Watch","Below target"))'),
        ("Return rate", f"={kpi}P{total}",
         '=IF(SUM(cd_lw_booked)=0,"n/a",SUM(cd_lw_returned)/SUM(cd_lw_booked))',
         "points", RATE_FORMAT, POINTS_FORMAT, "=th_return",
         '=IF(B{r}<=E{r},"On target","Above limit")'),
        ("Revenue", f"={kpi}K{total}", "=SUM(cd_lw_revenue)",
         like_for_like("cd_revenue", "cd_lw_revenue"), MONEY_FORMAT, CHANGE_FORMAT, None, None),
        ("Fuel cost", f"={kpi}L{total}", "=SUM(cd_lw_fuel)",
         like_for_like("cd_fuel", "cd_lw_fuel"), MONEY_FORMAT, CHANGE_FORMAT, None, None),
        ("Fuel cost per delivered parcel", f"={kpi}Q{total}",
         '=IF(SUM(cd_lw_delivered)=0,"n/a",SUM(cd_lw_fuel)/SUM(cd_lw_delivered))',
         "ratio", UNIT_MONEY_FORMAT, CHANGE_FORMAT, None, None),
        ("Revenue per booked parcel", f"={kpi}R{total}",
         '=IF(SUM(cd_lw_booked)=0,"n/a",SUM(cd_lw_revenue)/SUM(cd_lw_booked))',
         "ratio", UNIT_MONEY_FORMAT, CHANGE_FORMAT, None, None),
    ]
    for row, (label, current, prior, change, value_format, change_format, target, status) in (
            enumerate(rows, start=7)):
        write_cell(ws, row, 1, label)
        write_cell(ws, row, 2, current, value_format, align="right")
        write_cell(ws, row, 3, prior, value_format, align="right")
        if change == "points":
            change = f'=IF(OR(NOT(ISNUMBER(B{row})),NOT(ISNUMBER(C{row}))),"n/a",(B{row}-C{row})*100)'
        elif change == "ratio":
            change = f'=IF(OR(NOT(ISNUMBER(B{row})),NOT(ISNUMBER(C{row}))),"n/a",B{row}/C{row}-1)'
        else:
            change = change.format(r=row)
        write_cell(ws, row, 4, change, change_format, align="right")
        write_cell(ws, row, 5, target, RATE_FORMAT, align="right")
        write_cell(ws, row, 6, status.format(r=row) if status else None, align="center")
    kpi_last = 6 + len(rows)
    add_status_colors(ws, f"F7:F{kpi_last}")

    note = ws.cell(row=kpi_last + 1, column=1,
                   value="Last Week covers the same branch, day and service rows as This Week, "
                         "so missing data does not distort Change.")
    note.font = Font(name=FONT, size=9, italic=True, color="52514E")
    ws.merge_cells(start_row=kpi_last + 1, start_column=1, end_row=kpi_last + 1, end_column=6)
    note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[kpi_last + 1].height = 26

    top = kpi_last + 3
    style_section(ws, top, "Data Quality")
    style_header(ws, top + 1, ["Measure", "Value"])
    rows_read = top + 4
    excluded = rows_read + 2
    used = rows_read + 3
    expected = rows_read + 4
    coverage = rows_read + 5
    quality = [
        ("Branches reporting", f"=COUNTIF({kpi}D{first}:D{last},\">0\")", COUNT_FORMAT),
        ("Branches expected", stats["branches_expected"], COUNT_FORMAT),
        ("Rows read from files", stats["rows_read"], COUNT_FORMAT),
        ("Exact duplicates removed", '=COUNTIF(ex_action,"Removed duplicate")', COUNT_FORMAT),
        ("Rows excluded", stats["rows_excluded"], COUNT_FORMAT),
        ("Rows used in KPIs", "=COUNTA(cd_branch)", COUNT_FORMAT),
        ("Rows expected", stats["rows_expected"], COUNT_FORMAT),
        ("Coverage of expected rows", f"=B{used}/B{expected}", RATE_FORMAT),
        ("Auto-fixed lines on Exceptions", '=COUNTIF(ex_action,"Auto-fixed")', COUNT_FORMAT),
        ("Excluded lines on Exceptions", '=COUNTIF(ex_action,"Excluded")', COUNT_FORMAT),
        ("Follow-ups for branches", '=COUNTIF(ex_action,"Follow up")', COUNT_FORMAT),
        ("Check: rows reconcile",
         f'=IF(B{rows_read}-B{rows_read + 1}-B{excluded}=B{used},"OK","CHECK")', None),
        ("Check: KPI totals match Clean Data",
         f'=IF(AND({kpi}G{total}=SUM(cd_booked),ABS({kpi}K{total}-SUM(cd_revenue))<0.01),'
         f'"OK","CHECK")', None),
    ]
    for row, (label, value, number_format) in enumerate(quality, start=top + 2):
        write_cell(ws, row, 1, label)
        write_cell(ws, row, 2, value, number_format, align="right")
    quality_last = top + 1 + len(quality)
    add_status_colors(ws, f"B{quality_last - 1}:B{quality_last}")

    ws.merge_cells("A3:F3")
    banner = ws["A3"]
    banner.value = (
        f'=IF(B{coverage}>=1,"All expected rows were received and passed validation.",'
        f'"Volume and revenue totals are understated: "&TEXT(1-B{coverage},"0.0%")&'
        f'" of expected rows are missing or excluded. See the Exceptions sheet.")'
    )
    banner.font = Font(name=FONT, size=10, bold=True, color="9C5700")
    banner.fill = PatternFill("solid", fgColor="FFF4CE")
    banner.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 32

    top = quality_last + 2
    style_section(ws, top, "Lowest On-Time Branches")
    style_header(ws, top + 1, ["Branch", "On-Time %", "Gap to Target", "Status"])
    for offset in range(min(5, ranked_count)):
        row = top + 2 + offset
        position = f"MATCH(SMALL(kp_rank,{offset + 1}),kp_rank,0)"
        write_cell(ws, row, 1, f"=INDEX(kp_branch,{position})")
        write_cell(ws, row, 2, f"=INDEX(kp_ontime,{position})", RATE_FORMAT, align="right")
        write_cell(ws, row, 3, f"=(B{row}-th_ontime)*100", POINTS_FORMAT, align="right")
        write_cell(ws, row, 4, f"=INDEX(kp_status,{position})", align="center")
    add_status_colors(ws, f"D{top + 2}:D{top + 6}")
    set_print_layout(ws)


def style_axes(chart, y_title):
    """Keep axes visible and recessive. Excel can hide them unless told otherwise."""
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.title = y_title
    chart.y_axis.title.overlay = False
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=CHART_GRID))
    chart.title.overlay = False
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.roundedCorners = False


def add_charts(wb, positions, ranked_count):
    summary = wb["Summary"]
    breakdowns = wb["Breakdowns"]

    if ranked_count:
        header, last = positions["rank_header"], positions["rank_last"]
        ranking = BarChart()
        ranking.type = "col"
        ranking.overlap = 100
        ranking.gapWidth = 40
        ranking.title = "On-time gap to target by branch"
        ranking.add_data(Reference(breakdowns, min_col=5, max_col=6, min_row=header, max_row=last),
                         titles_from_data=True)
        ranking.set_categories(Reference(breakdowns, min_col=2, min_row=header + 1, max_row=last))
        for series, color in zip(ranking.series, (CHART_BLUE, CHART_RED)):
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = color
            series.invertIfNegative = False
        style_axes(ranking, "Points vs target")
        ranking.y_axis.number_format = "0"
        ranking.width, ranking.height = 22, 9.5
        summary.add_chart(ranking, "H5")

    header, last = positions["day_header"], positions["day_last"]
    daily = LineChart()
    daily.title = "Parcels booked and delivered by day (rows used)"
    daily.add_data(Reference(breakdowns, min_col=2, max_col=3, min_row=header, max_row=last),
                   titles_from_data=True)
    daily.set_categories(Reference(breakdowns, min_col=1, min_row=header + 1, max_row=last))
    for series, color in zip(daily.series, (CHART_BLUE, CHART_ORANGE)):
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 19050
        series.marker.symbol = "circle"
        series.marker.size = 7
        series.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        series.marker.graphicalProperties.line.solidFill = color
        series.smooth = False
    style_axes(daily, "Parcels")
    daily.x_axis.number_format = "ddd"
    daily.y_axis.number_format = "#,##0"
    daily.y_axis.scaling.min = 0
    daily.width, daily.height = 22, 9.5
    summary.add_chart(daily, "H25")


def write_report(path, clean, issues, master, stats, week_start):
    """Build and save the whole workbook."""
    wb = Workbook()
    wb.active.title = "Summary"
    for name in ("Branch KPIs", "Breakdowns", "Exceptions", "Clean Data", "Definitions"):
        wb.create_sheet(name)
    wb.calculation.fullCalcOnLoad = True

    has_delivered = clean.groupby("branch_code")["delivered"].sum() > 0
    ranked_count = int(has_delivered.sum())

    write_clean_data(wb, clean)
    write_definitions(wb)
    write_exceptions(wb, issues)
    kpi_rows = write_branch_kpis(wb, master)
    positions = write_breakdowns(wb, master, issues, week_start, ranked_count)
    write_summary(wb, stats, week_start, kpi_rows, ranked_count)
    add_charts(wb, positions, ranked_count)

    for name, color in (("Summary", NAVY), ("Exceptions", "E34948")):
        wb[name].sheet_properties.tabColor = color
    wb.save(path)
