#!/usr/bin/env python3
"""Rule store for the GSM capacity workbooks.

Rule versions are v1, v2, v3, and so on. Each one sits beside the capacity
file version that used it. GSM_Rules_Store_v1.xlsx stays as the first store.
This writer builds GSM_Rules_Store_v2.xlsx and can add the same sheets to a
capacity workbook.
"""

import xlsxwriter

STORE_VERSION = "v2"
STORE_FILE = f"GSM_Rules_Store_{STORE_VERSION}.xlsx"
ACTIVE_RULE = "v9"
ACTIVE_CAPACITY = "v10"
SAVED = "9 Nov 2026"

NAVY = "#1F4E79"
ORANGE = "#C65911"
GREEN = "#548235"
PALE_GREEN = "#E2EFDA"
PALE_GRAY = "#F2F2F2"
WHITE = "#FFFFFF"


def capacity_file(version):
    return f"GSM_Capacity_Enhancement_{version}.xlsx"


def or_rule(busy_hour_times_24):
    if busy_hour_times_24:
        voice = (
            "Capacity, Erl = available TCH x (1 + measured half-rate ratio). "
            "Half-rate ratio = busy-hour K3034 / busy-hour K3014. "
            "Voice load uses busy-hour K3014 multiplied by 24."
        )
    else:
        voice = (
            "Capacity, Erl = available TCH x (1 + measured half-rate ratio). "
            "Half-rate ratio = busy-hour K3034 / busy-hour K3014. "
            "Voice load uses busy-hour K3014 as the one-hour value, with no x24. "
            "Daily K3014, K3034, and K3004 are hourly averages, so those reference columns are x24 and are not used in voice load."
        )
    return [
        voice,
        "G900 reduction requires every one of these: the cell is G900, configured TRX is above 2, SDCCH congestion is below 1%, and TCH congestion is below 1%.",
        "It then needs any one of these: busy-hour voice load at or below 60%, daily CDR_N average at least 10, or both DL and UL bad quality above 2%.",
        "The cell keeps 2 TRX. Voice load after the cut = current voice load x configured TRX / 2. The cell still keeps 2 TRX when that projected load goes above 60%.",
        "A G1800 cell is proposed when any one of these is true: busy-hour voice load at least 70%, HOSR_Out at or below 90%, both DL and UL bad quality above 2%, daily CDR_N average at least 10, or CDR above 0.2% and ICM band 4-5 above 50%.",
        "TRX count does not filter the G1800 list. The list names the G900 cell. It does not choose how many G1800 TRX, an ARFCN, or a BCCH.",
        "A cell that already has 1800 TCH is left off the list. L900 from 5 MHz to 10 MHz is outside this rule.",
    ]


V8_LINES = [
    "Capacity, Erl is Erlang B at 5% blocking. Equivalent channels = available TCH x (1 + 0.60). The measured half-rate ratio stays on the sheet and is not used in capacity.",
    "Voice load % = busy-hour K3014 / Capacity x 100. Busy-hour K3014 is already one hour. Daily traffic x24 is reference only.",
    "G900 reduction runs first and needs every gate: configured TRX above 2, TCH congestion below 1%, SDCCH congestion below 0.5%, busy-hour voice load at or below 40%, and busy-hour RR307 above 99%. The cell then keeps 2 TRX.",
    "TCH after the cut = available TCH x 2 / configured TRX. Capacity after the cut is that TCH count times 1.60, read from the 5% table.",
    "G1800 is counted after that step. If voice load after the step is at least 80%, add the smallest number of G1800 TRX that takes pooled voice load under 60%. The load count stops at 12. A new G1800 cell has TCH = 8 x TRX - 1 BCCH - one SDCCH per two TRX, then x 1.60, then the 5% table.",
    "Daily CDR_N average at least 10 adds 1 G1800 TRX. If the cell was not reduced to 2 and at least 3 G900 TRX remain, remove 1 G900 TRX.",
    "DL and UL bad quality both above 2% add 1 G1800 TRX and can remove one more G900 TRX when at least 3 still remain after the CDR removal. Both lines together can remove 2 G900 TRX.",
    "ICM band 4-5 above 50% adds 1 G1800 TRX only when the reduction, the 80% load line, CDR, and bad quality did not trigger. It then removes 1 G900 TRX when at least 3 remain.",
    "On this file the full reduction is 2,513 cells and 3,319 TRX. The extra path removes 299 TRX. Total G900 TRX removed is 3,618. G1800 is 2,334 cells and 2,999 TRX.",
]

V9_LINES = [
    "Capacity, Erl is Erlang B at 5% blocking. Equivalent channels = available TCH x (1 + 0.60). The measured half-rate ratio stays on the sheet and is not used in capacity.",
    "Voice load % = busy-hour K3014 / Capacity x 100. Busy-hour K3014 is already one hour. Daily traffic x24 is reference only.",
    "G900 reduction runs first and needs every gate: configured TRX above 2, TCH congestion below 1%, SDCCH congestion below 0.5%, busy-hour voice load at or below 40%, and busy-hour RR307 above 99%. The cell then keeps 2 TRX.",
    "TCH after the cut = available TCH x 2 / configured TRX. Capacity after the cut is that TCH count times 1.60, read from the 5% table.",
    "G1800 is counted after that step. If voice load after the step is at least 80%, add the smallest number of G1800 TRX that takes pooled voice load under 60%. The load count stops at 12.",
    "Daily CDR_N average at least 10 adds 1 G1800 TRX. DL and UL bad quality both above 2% add 1 more G1800 TRX. The four G1800 counts are added together.",
    "The extra G900 removal is judged after the reduction step. A cell already at 2 TRX loses nothing more. A cell that still has 3 or more TRX loses exactly 1 G900 TRX when CDR, bad quality, or the ICM line is true. CDR and bad quality share that one removal.",
    "ICM band 4-5 above 50% adds 1 G1800 TRX only when the reduction, the 80% load line, CDR, and bad quality did not trigger.",
    "On capacity v9 the full reduction is 2,513 cells and 3,319 TRX. The extra path is 289 cells and 289 TRX. Total G900 TRX removed is 3,608. G1800 is 2,334 cells and 2,999 TRX. Capacity v10 uses this same rule and carries this store.",
]


RULES = [
    ("v1", "v1", "Busy-hour and daily traffic are both multiplied by 24. G900 reduction is the OR gate. Capacity uses the measured half-rate ratio.", or_rule(True)),
    ("v2", "v2", "Same decision as rule v1. Each KPI formula sits in the column immediately after its result. Busy-hour traffic is still multiplied by 24.", or_rule(True)),
    ("v3", "v3", "Busy-hour Erlang stays the one-hour value. Daily traffic stays x24 and is reference only. The OR gate is unchanged.", or_rule(False)),
    ("v4", "v4", "Same decision as rule v3. The formula column beside each KPI is the formula written as text.", or_rule(False)),
    ("v5", "v5", "Same decision as rule v3. The formula-text columns are removed. Each KPI keeps the numerator average, the denominator average, and the result.", or_rule(False)),
    ("v6", "v6", "Same decision as rule v3. Average headers use the counter with _avg, such as K3014_avg.", or_rule(False)),
    ("v7", "v7", "Same decision as rule v3. Average headers are the counter description, the counter, then Avg. This file reduced 3,180 G900 cells to 2 TRX and proposed 2,463 G1800 cells.", or_rule(False)),
    ("v8", "v8", "Replaces the OR gate. Capacity is Erlang B at 5% with a fixed 60% half rate. CDR and bad quality can each remove a G900 TRX.", V8_LINES),
    ("v9", "v9", "Same capacity method and the same five-gate reduction as rule v8. The extra G900 removal is one TRX at most, judged after the reduction step.", V9_LINES),
]

# Capacity v10 uses rule v9 and is the file that carries this store.
CATALOG = [(rule, capacity, "Kept", change) for rule, capacity, change, _lines in RULES]
CATALOG.append((
    "v9",
    "v10",
    "Active",
    "Active capacity file. The decision is rule v9. This version adds the rule store beside the capacity result.",
))


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
        "active": workbook.add_format({**base, "border": 1, "align": "center", "bg_color": PALE_GREEN, "bold": True}),
        "kept": workbook.add_format({**base, "border": 1, "align": "center", "bg_color": PALE_GRAY}),
        "int": workbook.add_format({**base, "border": 1, "align": "center", "num_format": "0"}),
        "num": workbook.add_format({**base, "border": 1, "align": "center", "num_format": "0.00"}),
    }


def add_rule_sheets(workbook):
    """Add Store, v1 through v9, and Examples. Safe inside a capacity workbook."""
    fmt = formats(workbook)
    write_store(workbook, fmt)
    for rule, capacity, change, lines in RULES:
        write_rule(workbook, fmt, rule, capacity, change, lines)
    write_examples(workbook, fmt)
    return ["Store", *[rule for rule, _capacity, _change, _lines in RULES], "Examples"]


def write_store(workbook, fmt):
    sheet = workbook.add_worksheet("Store")
    sheet.set_tab_color(GREEN)
    for col, width in enumerate([18, 26, 46, 14, 88]):
        sheet.set_column(col, col, width)
    sheet.freeze_panes(6, 0)
    sheet.write("A1", "GSM rule store", fmt["title"])
    sheet.merge_range(
        "A2:E2",
        f"Rule versions are v1, v2, v3, and so on. The capacity file version is the next column. "
        f"Active rule: {ACTIVE_RULE}. Active capacity file: {capacity_file(ACTIVE_CAPACITY)}. Saved {SAVED}.",
        fmt["open"],
    )
    sheet.set_row(1, 32)
    sheet.write("A4", "Versions", fmt["section"])
    headers = ["Rule version", "Capacity file version", "Capacity file", "Status", "What this version is"]
    for col, header in enumerate(headers):
        sheet.write(5, col, header, fmt["header"])
    for index, (rule, capacity, status, change) in enumerate(CATALOG):
        row = 6 + index
        sheet.write(row, 0, rule, fmt["name"])
        sheet.write(row, 1, capacity, fmt["name"])
        sheet.write(row, 2, capacity_file(capacity), fmt["text"])
        sheet.write(row, 3, status, fmt["active"] if status == "Active" else fmt["kept"])
        sheet.write(row, 4, change, fmt["text"])
        sheet.set_row(row, 36)
    note_row = 8 + len(CATALOG)
    sheet.write(note_row, 0, "How the next change is saved", fmt["section"])
    notes = [
        "Earlier rule rows stay on this sheet. A changed decision becomes the next rule version.",
        "The capacity workbook that uses the new rule gets the next capacity version, and that pair is a new row.",
        "GSM_Rules_Store_v1.xlsx stays in the repo. It named the current rule Rule1. This v2 store is the version list.",
        "The next store file will be GSM_Rules_Store_v3.xlsx. It keeps these rows and adds the new one.",
        "Ask to switch to a rule version, for example rule v8 or rule v9. The capacity file version beside it is the workbook that used that rule.",
    ]
    for offset, line in enumerate(notes):
        sheet.merge_range(note_row + 1 + offset, 0, note_row + 1 + offset, 4, line, fmt["text"])
        sheet.set_row(note_row + 1 + offset, 22)


def write_rule(workbook, fmt, rule, capacity, change, lines):
    sheet = workbook.add_worksheet(rule)
    sheet.set_tab_color(ORANGE if rule == ACTIVE_RULE else GREEN)
    sheet.set_column("A:A", 22)
    sheet.set_column("B:B", 120)
    sheet.write("A1", f"Rule {rule}", fmt["title"])
    sheet.merge_range(
        "A2:B2",
        f"Capacity file version {capacity}: {capacity_file(capacity)}."
        + (f" Also used by capacity file version {ACTIVE_CAPACITY}: {capacity_file(ACTIVE_CAPACITY)}." if rule == ACTIVE_RULE else ""),
        fmt["open"],
    )
    sheet.set_row(1, 32)
    rows = [("What changed", change)]
    for line in lines:
        rows.append(("Rule", line))
    sheet.write(3, 0, "Part", fmt["header"])
    sheet.write(3, 1, "Text", fmt["header"])
    for index, (part, text) in enumerate(rows):
        row = 4 + index
        sheet.write(row, 0, part, fmt["name"])
        sheet.write(row, 1, text, fmt["text"])
        sheet.set_row(row, 36)


def write_examples(workbook, fmt):
    from build_gsm_capacity import build_erlang_table, decide, defaults, g1800_caps

    sheet = workbook.add_worksheet("Examples")
    sheet.set_tab_color(GREEN)
    sheet.write("A1", "Examples for rule v9", fmt["title"])
    sheet.merge_range(
        "A2:V2",
        f"These rows are rule v9, used by {capacity_file('v9')} and {capacity_file(ACTIVE_CAPACITY)}.",
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
