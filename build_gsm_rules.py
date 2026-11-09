#!/usr/bin/env python3
"""Rule store for the GSM capacity workbooks.

Rule versions are v1, v2, v3, and so on. Each one sits beside the capacity
file version that used it. GSM_Rules_Store_v1.xlsx stays as the first store.
This writer builds GSM_Rules_Store_v3.xlsx and can add the same sheets to a
capacity workbook. Each version has two rules under G900 TRX reductions and
two rules under G1800 TRX additions.
"""

import xlsxwriter

STORE_VERSION = "v3"
STORE_FILE = f"GSM_Rules_Store_{STORE_VERSION}.xlsx"
ACTIVE_RULE = "v9"
ACTIVE_CAPACITY = "v11"
SAVED = "9 Nov 2026"
G900 = "1. G900 TRX reductions"
G1800 = "2. G1800 TRX additions"

NAVY = "#1F4E79"
ORANGE = "#C65911"
GREEN = "#548235"
PALE_GREEN = "#E2EFDA"
PALE_GRAY = "#F2F2F2"
WHITE = "#FFFFFF"


def capacity_file(version):
    return f"GSM_Capacity_Enhancement_{version}.xlsx"


G900_OLD_1 = (
    "The cell is G900, configured TRX is above 2, SDCCH congestion is below 1%, and TCH congestion is below 1%. All four are required."
)
G900_OLD_2 = (
    "Any one of these is also required: busy-hour voice load at or below 60%, daily CDR_N average at least 10, or both DL and UL bad quality above 2%. "
    "The cell then keeps 2 TRX. This version has no further G900 removal after that."
)
G1800_OLD_1 = (
    "Busy-hour voice load at least 70% proposes one G1800 cell on that G900 cell. This rule does not choose how many G1800 TRX."
)
G1800_OLD_2 = (
    "A G1800 cell is also proposed when any one of these is true: HOSR_Out at or below 90%, both DL and UL bad quality above 2%, "
    "daily CDR_N average at least 10, or CDR above 0.2% and ICM band 4-5 above 50%. TRX count is not a filter. A cell that already has 1800 TCH is left off the list."
)
G900_NEW_1 = (
    "All five are required: configured TRX above 2, TCH congestion below 1%, SDCCH congestion below 0.5%, busy-hour voice load at or below 40%, and busy-hour RR307 above 99%. "
    "The cell keeps 2 TRX. TCH after the cut = available TCH x 2 / configured TRX, then Erlang B at 5% with the 60% design half rate."
)
G900_V8_2 = (
    "After Rule 1, a cell already at 2 TRX loses nothing more. If it was not reduced and at least 3 TRX remain, CDR removes 1 G900 TRX, and bad quality can remove a second while at least 3 still remain."
)
G900_V9_2 = (
    "After Rule 1, look at the TRX still there. Already at 2: remove nothing more. Still 3 or more, and CDR, bad quality, or the ICM line is true: remove exactly 1 G900 TRX. CDR and bad quality share that one removal."
)
G1800_NEW_1 = (
    "After the G900 step, if voice load is at least 80%, add the smallest number of G1800 TRX that takes pooled voice load under 60%. The cap is 12. "
    "Pooled capacity is the remaining G900 Erlang plus the new G1800 Erlang. New G1800 TCH = 8 x TRX - 1 BCCH - one SDCCH per two TRX, then x 1.60, then the 5% table."
)
G1800_NEW_2 = (
    "Daily CDR_N average at least 10 adds 1 G1800 TRX. DL and UL bad quality both above 2% add 1 G1800 TRX. "
    "ICM band 4-5 above 50% adds 1 G1800 TRX only when the G900 reduction, the 80% load rule, CDR, and bad quality did not trigger. These counts are added to Rule 1."
)

RULES = [
    ("v1", "v1", "Busy-hour K3014 is multiplied by 24. Capacity uses the measured half-rate ratio.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v2", "v2", "Same two rules as v1. Each KPI formula sits in the column after the result. Busy-hour K3014 is still multiplied by 24.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v3", "v3", "Same two rules in each section. Busy-hour K3014 is the one-hour value. Daily traffic x24 is reference only.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v4", "v4", "Same two rules as v3. The formula beside each KPI is written as text.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v5", "v5", "Same two rules as v3. Formula-text columns are removed. Each KPI keeps numerator, denominator, and result.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v6", "v6", "Same two rules as v3. Average headers use the counter with _avg.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v7", "v7", "Same two rules as v3. Headers are description, counter, then Avg. This file reduced 3,180 G900 cells and proposed 2,463 G1800 cells.", G900_OLD_1, G900_OLD_2, G1800_OLD_1, G1800_OLD_2),
    ("v8", "v8", "Erlang B at 5% and a fixed 60% half rate. G900 Rule 2 can remove a second TRX. This file removed 3,618 G900 TRX and added 2,999 G1800 TRX.", G900_NEW_1, G900_V8_2, G1800_NEW_1, G1800_NEW_2),
    ("v9", "v9", "Same G1800 rules as v8. G900 Rule 2 removes one TRX at most. Capacity v9, v10, and v11 use it. Removed G900 TRX is 3,608. G1800 TRX is 2,999.", G900_NEW_1, G900_V9_2, G1800_NEW_1, G1800_NEW_2),
]

CATALOG = [(rule, capacity, "Kept", note) for rule, capacity, note, *_rest in RULES]
CATALOG.append(("v9", "v10", "Kept", "Same rule v9. First capacity file that carried the rule store."))
CATALOG.append((
    "v9",
    ACTIVE_CAPACITY,
    "Active",
    "Active capacity file. Rule v9, with two rules under G900 TRX reductions and two rules under G1800 TRX additions.",
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
    for item in RULES:
        write_rule(workbook, fmt, item)
    write_examples(workbook, fmt)
    return ["Store", *[item[0] for item in RULES], "Examples"]


def write_section_table(sheet, fmt, start_row, section, pairs):
    sheet.merge_range(start_row, 0, start_row, 4, section, fmt["section"])
    headers = ["Rule", "Capacity file version", "Capacity file", "Status", "Rule text"]
    for col, header in enumerate(headers):
        sheet.write(start_row + 1, col, header, fmt["header"])
    for offset, (rule_no, capacity, status, text) in enumerate(pairs):
        row = start_row + 2 + offset
        sheet.write(row, 0, rule_no, fmt["name"])
        sheet.write(row, 1, capacity, fmt["name"])
        sheet.write(row, 2, capacity_file(capacity), fmt["text"])
        sheet.write(row, 3, status, fmt["active"] if status == "Active" else fmt["kept"])
        sheet.write(row, 4, text, fmt["text"])
        sheet.set_row(row, 48)
    return start_row + 2 + len(pairs)


def write_store(workbook, fmt):
    sheet = workbook.add_worksheet("Store")
    sheet.set_tab_color(GREEN)
    for col, width in enumerate([16, 24, 42, 14, 88]):
        sheet.set_column(col, col, width)
    sheet.freeze_panes(5, 0)
    sheet.write("A1", "GSM rule store", fmt["title"])
    sheet.merge_range(
        "A2:E2",
        "Two sections. Each section has two rules. The capacity file version is beside the rule. "
        f"Active rule {ACTIVE_RULE}. Active capacity file {capacity_file(ACTIVE_CAPACITY)}. Saved {SAVED}.",
        fmt["open"],
    )
    sheet.set_row(1, 32)

    g900_rows = []
    g1800_rows = []
    for rule, capacity, note, g900_1, g900_2, g1800_1, g1800_2 in RULES:
        status = "Kept"
        g900_rows.append(("Rule 1", capacity, status, g900_1))
        g900_rows.append(("Rule 2", capacity, status, g900_2))
        g1800_rows.append(("Rule 1", capacity, status, g1800_1))
        g1800_rows.append(("Rule 2", capacity, status, g1800_2))
    for capacity, status in (("v10", "Kept"), (ACTIVE_CAPACITY, "Active")):
        g900_rows.append(("Rule 1", capacity, status, G900_NEW_1))
        g900_rows.append(("Rule 2", capacity, status, G900_V9_2))
        g1800_rows.append(("Rule 1", capacity, status, G1800_NEW_1))
        g1800_rows.append(("Rule 2", capacity, status, G1800_NEW_2))

    next_row = write_section_table(sheet, fmt, 3, G900, g900_rows)
    next_row = write_section_table(sheet, fmt, next_row + 1, G1800, g1800_rows)
    sheet.write(next_row + 1, 0, "Capacity file beside each rule version", fmt["section"])
    headers = ["Rule version", "Capacity file version", "Capacity file", "Status", "What changed in the file"]
    for col, header in enumerate(headers):
        sheet.write(next_row + 2, col, header, fmt["header"])
    for index, (rule, capacity, status, change) in enumerate(CATALOG):
        row = next_row + 3 + index
        sheet.write(row, 0, rule, fmt["name"])
        sheet.write(row, 1, capacity, fmt["name"])
        sheet.write(row, 2, capacity_file(capacity), fmt["text"])
        sheet.write(row, 3, status, fmt["active"] if status == "Active" else fmt["kept"])
        sheet.write(row, 4, change, fmt["text"])
        sheet.set_row(row, 32)
    note_row = next_row + 4 + len(CATALOG)
    notes = [
        "GSM_Rules_Store_v1.xlsx and GSM_Rules_Store_v2.xlsx stay in the repo. This v3 store keeps every earlier rule and splits each one into the two sections.",
        "A changed rule is the next rule version, saved in GSM_Rules_Store_v4.xlsx, with the next capacity file version beside it.",
        "Ask to switch to a rule version, for example rule v8 or rule v9. Both sections switch together.",
    ]
    for offset, line in enumerate(notes):
        sheet.merge_range(note_row + offset, 0, note_row + offset, 4, line, fmt["text"])
        sheet.set_row(note_row + offset, 28)


def write_rule(workbook, fmt, item):
    rule, capacity, change, g900_1, g900_2, g1800_1, g1800_2 = item
    sheet = workbook.add_worksheet(rule)
    sheet.set_tab_color(ORANGE if rule == ACTIVE_RULE else GREEN)
    sheet.set_column("A:A", 28)
    sheet.set_column("B:B", 18)
    sheet.set_column("C:C", 110)
    also = ""
    if rule == ACTIVE_RULE:
        also = f" Also used by capacity v10 and capacity {ACTIVE_CAPACITY}."
    sheet.write("A1", f"Rule {rule}", fmt["title"])
    sheet.merge_range(
        "A2:C2",
        f"Capacity file version {capacity}: {capacity_file(capacity)}.{also} Two rules under each section.",
        fmt["open"],
    )
    sheet.set_row(1, 32)
    sheet.merge_range("A4:C4", change, fmt["text"])
    sheet.set_row(3, 32)
    rows = [
        (G900, "Rule 1", g900_1),
        (G900, "Rule 2", g900_2),
        (G1800, "Rule 1", g1800_1),
        (G1800, "Rule 2", g1800_2),
    ]
    for col, header in enumerate(["Section", "Rule", "Rule text"]):
        sheet.write(5, col, header, fmt["header"])
    for index, (section, rule_no, text) in enumerate(rows):
        row = 6 + index
        sheet.write(row, 0, section, fmt["name"])
        sheet.write(row, 1, rule_no, fmt["name"])
        sheet.write(row, 2, text, fmt["text"])
        sheet.set_row(row, 56)


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
