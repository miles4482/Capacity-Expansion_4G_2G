#!/usr/bin/env python3
"""Rule store for the GSM capacity workbooks.

Rule versions are v1, v2, v3, and so on. Each one sits beside the capacity
file version that used it. GSM_Rules_Store_v1.xlsx stays as the first store.
This writer builds GSM_Rules_Store_v4.xlsx and can add the same sheets to a
capacity workbook. Each version sheet is the first step, G900 reduction,
then the second step, G1800 addition.
"""

import xlsxwriter

STORE_VERSION = "v4"
STORE_FILE = f"GSM_Rules_Store_{STORE_VERSION}.xlsx"
ACTIVE_RULE = "v9"
ACTIVE_CAPACITY = "v12"
SAVED = "9 Nov 2026"
G900_STEP = "1. G900 TRX reduction, first step"
G1800_STEP = "2. G1800 TRX addition, second step"

NAVY = "#1F4E79"
ORANGE = "#C65911"
GREEN = "#548235"
PALE_GREEN = "#E2EFDA"
PALE_GRAY = "#F2F2F2"
WHITE = "#FFFFFF"


def capacity_file(version):
    return f"GSM_Capacity_Enhancement_{version}.xlsx"


OLD_GATES = [
    ("All of these are required.", [
        ("Cell", "G900"),
        ("Configured TRX", "above 2"),
        ("SDCCH congestion", "below 1%"),
        ("TCH congestion", "below 1%"),
    ]),
    ("Any one of these is also required.", [
        ("Busy-hour voice load", "at or below 60%"),
        ("Daily CDR_N", "at least 10"),
        ("DL and UL bad quality", "both above 2%"),
    ]),
]
OLD_AFTER = "The cell then keeps 2 TRX. Voice load after the cut = current voice load x configured TRX / 2. No further G900 TRX is removed after that."
OLD_G1800 = [
    "Load. Busy-hour voice load at least 70% proposes one G1800 cell. This line does not choose how many G1800 TRX.",
    "HOSR. HOSR_Out at or below 90% proposes one G1800 cell.",
    "Quality. Both DL and UL bad quality above 2% proposes one G1800 cell.",
    "CDR. Daily CDR_N average at least 10 proposes one G1800 cell.",
    "CDR and ICM. CDR above 0.2% and ICM band 4-5 above 50%, together, proposes one G1800 cell.",
]
OLD_CLOSE = "Any one of these proposes the cell. A cell that already has 1800 TCH is left off the list."
NEW_GATES = [
    ("All five must be true.", [
        ("Configured TRX", "above 2"),
        ("TCH congestion", "below 1%"),
        ("SDCCH congestion", "below 0.5%"),
        ("Busy-hour voice load", "at or below 40%"),
        ("Busy-hour RR307", "above 99%"),
    ]),
]
NEW_AFTER = "The cell then keeps 2 TRX. TCH after the cut = available TCH x 2 / configured TRX, and voice load is recalculated from that."
V8_MORE = "If the cell was not reduced and at least 3 TRX remain, CDR removes 1 more G900 TRX, and bad quality can remove a second while at least 3 still remain. A cell already at 2 loses nothing more."
V9_MORE = "If 3 or more TRX are still there after this step, and the G1800 step is adding a TRX for CDR, bad quality, or high ICM, remove exactly 1 more G900 TRX. A cell already at 2 loses nothing more."
NEW_G1800_INTRO = "Voice load for these lines is the load after the reduction. A cell that was not reduced uses its current load."
NEW_G1800 = [
    "Load. If that voice load is at least 80%, add the smallest number of G1800 TRX that brings pooled voice load under 60%. Cap is 12 TRX. Pooled capacity is the remaining G900 Erlang plus the new G1800 Erlang. A new G1800 cell has TCH = 8 x TRX - 1 BCCH - one SDCCH per two TRX, then x 1.60, then the 5% table.",
    "CDR. If daily CDR_N average is at least 10, add 1 G1800 TRX.",
    "Quality. If DL bad quality is above 2% and UL bad quality is above 2%, add 1 G1800 TRX.",
    "High ICM. If ICM band 4-5 is above 50%, and the reduction, the 80% load line, CDR, and bad quality did not trigger, add 1 G1800 TRX.",
]
NEW_CLOSE = "The four G1800 counts are added. A cell that needs 3 TRX for load and also has CDR_N at least 10 receives 4 G1800 TRX."
NEW_CAPACITY = "Capacity, Erl is Erlang B at 5% blocking. Equivalent channels = available TCH x (1 + 0.60). Voice load % = busy-hour K3014 / Capacity x 100. Busy-hour K3014 is already one hour."

def rule_page(rule, capacity, note, gates, after, more, intro, lines, close, capacity_note):
    return {
        "rule": rule,
        "capacity": capacity,
        "note": note,
        "gates": gates,
        "after": after,
        "more": more,
        "intro": intro,
        "lines": lines,
        "close": close,
        "capacity_note": capacity_note,
    }


OLD_MEASURED = "Capacity, Erl = available TCH x (1 + measured half-rate ratio). "
RULES = [
    rule_page("v1", "v1", "Busy-hour K3014 is multiplied by 24.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 multiplied by 24."),
    rule_page("v2", "v2", "Same steps as v1. Each KPI formula sits in the column after the result. Busy-hour K3014 is still multiplied by 24.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 multiplied by 24."),
    rule_page("v3", "v3", "Same steps as v1. Busy-hour K3014 is the one-hour value. Daily traffic x24 is reference only.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 with no x24."),
    rule_page("v4", "v4", "Same steps as v3. The formula beside each KPI is written as text.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 with no x24."),
    rule_page("v5", "v5", "Same steps as v3. Formula-text columns are removed. Each KPI keeps numerator, denominator, and result.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 with no x24."),
    rule_page("v6", "v6", "Same steps as v3. Average headers use the counter with _avg.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 with no x24."),
    rule_page("v7", "v7", "Same steps as v3. Headers are description, counter, then Avg. This file reduced 3,180 G900 cells and proposed 2,463 G1800 cells.", OLD_GATES, OLD_AFTER, "", "A G1800 cell is proposed on a G900 cell. TRX count is not a filter.", OLD_G1800, OLD_CLOSE, OLD_MEASURED + "Voice load uses busy-hour K3014 with no x24."),
    rule_page("v8", "v8", "Erlang B at 5% and a fixed 60% half rate. The extra G900 removal can take a second TRX. This file removed 3,618 G900 TRX and added 2,999 G1800 TRX.", NEW_GATES, NEW_AFTER, V8_MORE, NEW_G1800_INTRO, NEW_G1800, NEW_CLOSE, NEW_CAPACITY),
    rule_page("v9", "v9", "Same G1800 step as v8. The extra G900 removal is one TRX. Capacity v9 through v12 use this page. Removed G900 TRX is 3,608. G1800 TRX is 2,999.", NEW_GATES, NEW_AFTER, V9_MORE, NEW_G1800_INTRO, NEW_G1800, NEW_CLOSE, NEW_CAPACITY),
]

CATALOG = [(item["rule"], item["capacity"], "Kept", item["note"]) for item in RULES]
CATALOG.append(("v9", "v10", "Kept", "Same rule v9. First capacity file that carried the rule store."))
CATALOG.append(("v9", "v11", "Kept", "Same rule v9."))
CATALOG.append(("v9", ACTIVE_CAPACITY, "Active", "Active capacity file. G900 reduction is the first step. G1800 addition is the second step."))


def formats(workbook):
    base = {"font_name": "Calibri", "font_size": 11, "valign": "vcenter"}
    return {
        "title": workbook.add_format({**base, "bold": True, "font_size": 18, "font_color": NAVY}),
        "section": workbook.add_format({**base, "bold": True, "font_size": 13, "font_color": NAVY, "bottom": 1, "bottom_color": NAVY}),
        "header": workbook.add_format({**base, "bold": True, "font_color": WHITE, "bg_color": NAVY, "border": 1, "text_wrap": True, "align": "center"}),
        "text": workbook.add_format({**base, "text_wrap": True, "border": 1}),
        "open": workbook.add_format({**base, "text_wrap": True}),
        "center": workbook.add_format({**base, "border": 1, "align": "center", "text_wrap": True}),
        "name": workbook.add_format({**base, "border": 1, "bold": True, "font_color": NAVY, "align": "center"}),
        "gate": workbook.add_format({**base, "border": 1, "bold": True, "font_color": NAVY}),
        "active": workbook.add_format({**base, "border": 1, "align": "center", "bg_color": PALE_GREEN, "bold": True}),
        "kept": workbook.add_format({**base, "border": 1, "align": "center", "bg_color": PALE_GRAY}),
        "int": workbook.add_format({**base, "border": 1, "align": "center", "num_format": "0"}),
        "num": workbook.add_format({**base, "border": 1, "align": "center", "num_format": "0.00"}),
    }


def add_rule_sheets(workbook):
    """Add Store, v1 through v9, and Examples. Safe inside a capacity workbook."""
    fmt = formats(workbook)
    write_store(workbook, fmt)
    for item in RULES:
        write_rule(workbook, fmt, item)
    write_examples(workbook, fmt)
    return ["Store", *[item["rule"] for item in RULES], "Examples"]


def write_store(workbook, fmt):
    sheet = workbook.add_worksheet("Store")
    sheet.set_tab_color(GREEN)
    for col, width in enumerate([18, 26, 46, 14, 88]):
        sheet.set_column(col, col, width)
    sheet.freeze_panes(6, 0)
    sheet.write("A1", "GSM rule store", fmt["title"])
    sheet.merge_range(
        "A2:E2",
        "Open a version sheet for the rule. G900 TRX reduction is the first step. G1800 TRX addition is the second step. "
        f"Active rule {ACTIVE_RULE}. Active capacity file {capacity_file(ACTIVE_CAPACITY)}. Saved {SAVED}.",
        fmt["open"],
    )
    sheet.set_row(1, 36)
    headers = ["Rule version", "Capacity file version", "Capacity file", "Status", "What changed in the file"]
    for col, header in enumerate(headers):
        sheet.write(4, col, header, fmt["header"])
    for index, (rule, capacity, status, change) in enumerate(CATALOG):
        row = 5 + index
        sheet.write(row, 0, rule, fmt["name"])
        sheet.write(row, 1, capacity, fmt["name"])
        sheet.write(row, 2, capacity_file(capacity), fmt["text"])
        sheet.write(row, 3, status, fmt["active"] if status == "Active" else fmt["kept"])
        sheet.write(row, 4, change, fmt["text"])
        sheet.set_row(row, 32)
    note_row = 7 + len(CATALOG)
    notes = [
        "Earlier store files stay in the repo. This v4 store writes each rule as the two steps.",
        "A changed rule is the next rule version, saved in the next store file, with the next capacity file version beside it.",
    ]
    for offset, line in enumerate(notes):
        sheet.merge_range(note_row + offset, 0, note_row + offset, 4, line, fmt["text"])
        sheet.set_row(note_row + offset, 24)


def write_line(sheet, fmt, row, text, style, height=36):
    sheet.merge_range(row, 0, row, 1, text, style)
    sheet.set_row(row, height)
    return row + 1


def write_rule(workbook, fmt, item):
    rule = item["rule"]
    capacity = item["capacity"]
    sheet = workbook.add_worksheet(rule)
    sheet.set_tab_color(ORANGE if rule == ACTIVE_RULE else GREEN)
    sheet.set_column("A:A", 32)
    sheet.set_column("B:B", 118)
    also = ""
    if rule == ACTIVE_RULE:
        also = f" Also used by capacity v10, v11, and {ACTIVE_CAPACITY}."
    sheet.write("A1", f"Rule {rule}", fmt["title"])
    sheet.merge_range(
        "A2:B2",
        f"Capacity file version {capacity}: {capacity_file(capacity)}.{also}",
        fmt["open"],
    )
    sheet.set_row(1, 24)
    row = write_line(sheet, fmt, 3, item["note"], fmt["open"], 32)
    row = write_line(sheet, fmt, row + 1, G900_STEP, fmt["section"], 22)
    for lead, gates in item["gates"]:
        row = write_line(sheet, fmt, row, lead, fmt["open"], 20)
        sheet.write(row, 0, "Gate", fmt["header"])
        sheet.write(row, 1, "Threshold", fmt["header"])
        row += 1
        for gate, threshold in gates:
            sheet.write(row, 0, gate, fmt["gate"])
            sheet.write(row, 1, threshold, fmt["text"])
            sheet.set_row(row, 22)
            row += 1
        row += 1
    row = write_line(sheet, fmt, row, item["after"], fmt["open"], 36)
    if item["more"]:
        row = write_line(sheet, fmt, row, item["more"], fmt["open"], 36)
    row = write_line(sheet, fmt, row + 1, G1800_STEP, fmt["section"], 22)
    row = write_line(sheet, fmt, row, item["intro"], fmt["open"], 32)
    for number, line in enumerate(item["lines"], start=1):
        row = write_line(sheet, fmt, row, f"{number}. {line}", fmt["open"], 48)
    row = write_line(sheet, fmt, row, item["close"], fmt["open"], 32)
    write_line(sheet, fmt, row + 1, item["capacity_note"], fmt["open"], 36)


def write_examples(workbook, fmt):
    from build_gsm_capacity import build_erlang_table, decide, defaults, g1800_caps

    sheet = workbook.add_worksheet("Examples")
    sheet.set_tab_color(GREEN)
    sheet.write("A1", "Examples for rule v9", fmt["title"])
    sheet.merge_range(
        "A2:V2",
        f"These rows are rule v9, used by {capacity_file('v9')}, {capacity_file('v10')}, and {capacity_file(ACTIVE_CAPACITY)}.",
        fmt["open"],
    )
    headers = [
        "Case", "Band", "TRX", "Voice %", "TCH cong %", "SDCCH cong %", "RR307",
        "CDR_N", "DL %", "UL %", "ICM %", "K3014 Erl", "Available TCH",
        "Reduction", "TRX final", "TRX removed", "G1800 add",
        "For load", "For CDR", "For quality", "For ICM", "What rule v9 does",
    ]
    for col, header in enumerate(headers):
        sheet.set_column(col, col, 16 if col else 32)
    sheet.set_column(21, 21, 64)
    for col, header in enumerate(headers):
        sheet.write(4, col, header, fmt["header"])
    cases = [
        ("Quiet cell, all five gates", "G900", 4, 20, 0.2, 0.1, 99.5, 1, 0.4, 0.4, 10, 2.0, 30, "Keep 2 TRX. No G1800 add."),
        ("Voice already high, only 2 TRX", "G900", 2, 90, 0.1, 0.1, 100, 0, 0.1, 0.1, 1, 16, 14, "No reduction. G1800 comes from the load rule. No extra G900 removal."),
        ("CDR and bad quality, 4 TRX", "G900", 4, 50, 0.2, 0.2, 100, 12, 3, 3, 80, 8, 30, "Add 2 G1800 TRX. Remove exactly 1 G900 TRX."),
        ("High ICM only, 4 TRX", "G900", 4, 20, 0.2, 0.2, 98, 1, 0.2, 0.2, 70, 3, 30, "Add 1 G1800 TRX and remove 1 G900 TRX."),
        ("Reduced, and CDR is high", "G900", 4, 20, 0.2, 0.1, 100, 15, 0.2, 0.2, 10, 2, 30, "Keep 2 TRX. Add 1 G1800 TRX. Remove no further G900 TRX."),
        ("CDR and quality, only 3 TRX", "G900", 3, 50, 0.2, 0.2, 100, 12, 3, 4, 10, 6, 21, "Add 2 G1800 TRX. Remove 1 G900 TRX."),
        ("G1800 cell with every trigger", "G1800", 4, 10, 0, 0, 100, 20, 5, 5, 90, 4, 20, "No action."),
    ]
    table = build_erlang_table()
    cfg = defaults()
    caps = g1800_caps(table, cfg["pHR"])
    for index, case in enumerate(cases):
        label, band, trx, voice, tch, sdcch, rr307, cdr, dl, ul, icm, carried, avail, why = case
        plan = decide(band, trx, voice, tch, sdcch, rr307, cdr, dl, ul, icm, carried, avail, table, caps, cfg)
        values = [
            label, band, trx, voice, tch, sdcch, rr307, cdr, dl, ul, icm, carried, avail,
            "Yes" if plan["reduce"] else "No", plan["trx_final"], plan["trx_remove"], plan["n_add"],
            plan["n_load"], plan["n_cdr"], plan["n_qual"], plan["n_icm"], why,
        ]
        row = 5 + index
        for col, value in enumerate(values):
            style = fmt["name"] if col == 0 else fmt["text"] if col == 21 else fmt["num"] if isinstance(value, float) else fmt["int"] if isinstance(value, int) else fmt["center"]
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                sheet.write_number(row, col, value, style)
            else:
                sheet.write(row, col, value, style)
        sheet.set_row(row, 28)
        if plan["removed_cdr"] + plan["removed_qual"] + plan["removed_icm"] > 1:
            raise SystemExit(f"{label} removed more than one extra G900 TRX")


def build():
    workbook = xlsxwriter.Workbook(STORE_FILE)
    names = add_rule_sheets(workbook)
    workbook.worksheets_objs.sort(key=lambda item: names.index(item.name))
    workbook.close()
    print(f"workbook {STORE_FILE}")
    print(f"active rule {ACTIVE_RULE} beside capacity {ACTIVE_CAPACITY}")


if __name__ == "__main__":
    build()
