#!/usr/bin/env python3
"""Save the GSM decision rules as a versioned store.

Rule1 v1 is the rule set inside GSM_Capacity_Enhancement_v9.xlsx.
This file is the store. Do not overwrite GSM_Rules_Store_v1.xlsx.
A later rule change writes GSM_Rules_Store_v2.xlsx and keeps this file,
so an earlier rule can be switched back on by name.
"""

from build_gsm_capacity import (
    ACTIVE_RULE,
    ACTIVE_RULE_FILE,
    ACTIVE_RULE_VERSION,
    G1800_TABLE_TRX,
    PARAMS,
    VERSION,
    build_erlang_table,
    decide,
    defaults,
    g1800_caps,
    g1800_tch,
    offered_at,
)
import xlsxwriter

STORE_VERSION = "v1"
OUT = ACTIVE_RULE_FILE
SAVED = "9 Nov 2026"
CAPACITY_FILE = f"GSM_Capacity_Enhancement_{VERSION}.xlsx"

NAVY = "#1F4E79"
ORANGE = "#C65911"
BLUE = "#2F5496"
GREEN = "#548235"
PALE_YELLOW = "#FFF2CC"
PALE_BLUE = "#D6EAF8"
PALE_GREEN = "#E2EFDA"
PALE_ORANGE = "#FCE4D6"
PALE_GRAY = "#F2F2F2"
WHITE = "#FFFFFF"


def formats(workbook):
    base = {"font_name": "Calibri", "font_size": 11, "valign": "vcenter"}
    return {
        "title": workbook.add_format({**base, "bold": True, "font_size": 18, "font_color": NAVY}),
        "section": workbook.add_format({**base, "bold": True, "font_size": 13, "font_color": NAVY, "bottom": 1, "bottom_color": NAVY}),
        "header": workbook.add_format({**base, "bold": True, "font_color": WHITE, "bg_color": NAVY, "border": 1, "text_wrap": True, "align": "center"}),
        "header_orange": workbook.add_format({**base, "bold": True, "font_color": WHITE, "bg_color": ORANGE, "border": 1, "text_wrap": True, "align": "center"}),
        "header_blue": workbook.add_format({**base, "bold": True, "font_color": WHITE, "bg_color": BLUE, "border": 1, "text_wrap": True, "align": "center"}),
        "text": workbook.add_format({**base, "text_wrap": True, "border": 1}),
        "text_open": workbook.add_format({**base, "text_wrap": True}),
        "center": workbook.add_format({**base, "border": 1, "align": "center", "text_wrap": True}),
        "int": workbook.add_format({**base, "border": 1, "align": "center", "num_format": "0"}),
        "num": workbook.add_format({**base, "border": 1, "align": "center", "num_format": "0.00"}),
        "yellow": workbook.add_format({**base, "border": 1, "align": "center", "bg_color": PALE_YELLOW, "bold": True, "num_format": "0.##"}),
        "active": workbook.add_format({**base, "border": 1, "align": "center", "bg_color": PALE_GREEN, "bold": True}),
        "name": workbook.add_format({**base, "border": 1, "bold": True, "font_color": NAVY}),
        "step": workbook.add_format({**base, "border": 1, "align": "center", "bold": True, "bg_color": PALE_ORANGE}),
        "note": workbook.add_format({**base, "text_wrap": True, "italic": True, "font_color": "#595959"}),
    }


def write_header(sheet, row, labels, fmt):
    for col, label in enumerate(labels):
        sheet.write(row, col, label, fmt)


def build():
    if ACTIVE_RULE != "Rule1" or ACTIVE_RULE_VERSION != STORE_VERSION or ACTIVE_RULE_FILE != OUT:
        raise SystemExit("Rule store identity does not match the capacity workbook pointer")
    table = build_erlang_table()
    cfg = defaults()
    caps = g1800_caps(table, cfg["pHR"])

    workbook = xlsxwriter.Workbook(OUT)
    fmt = formats(workbook)
    write_store(workbook, fmt)
    write_rule1(workbook, fmt)
    write_parameters(workbook, fmt)
    write_decision(workbook, fmt)
    write_formulas(workbook, fmt)
    write_data(workbook, fmt)
    write_examples(workbook, fmt, table, caps, cfg)
    write_g1800(workbook, fmt, table, caps)
    order = [
        "Store", "Rule1", "Parameters", "Decision", "Formulas", "Data", "Examples", "G1800 capacity",
    ]
    workbook.worksheets_objs.sort(key=lambda item: order.index(item.name))
    workbook.close()
    print(f"workbook {OUT}")
    print(f"active {ACTIVE_RULE} {ACTIVE_RULE_VERSION} matches {CAPACITY_FILE}")


def write_store(workbook, fmt):
    sheet = workbook.add_worksheet("Store")
    sheet.set_tab_color(GREEN)
    sheet.set_column("A:A", 16)
    sheet.set_column("B:B", 16)
    sheet.set_column("C:C", 14)
    sheet.set_column("D:D", 42)
    sheet.set_column("E:E", 16)
    sheet.set_column("F:F", 88)
    sheet.set_row(0, 24)
    sheet.write("A1", "GSM rule store", fmt["title"])
    sheet.merge_range(
        "A2:F2",
        "This workbook is the rule store. Rule1 is the active rule. "
        f"It is the rule set used by {CAPACITY_FILE}. "
        "Ask to switch to Rule1 to use this set again.",
        fmt["text_open"],
    )
    sheet.set_row(1, 32)
    sheet.write("A4", "Saved rules", fmt["section"])
    write_header(
        sheet, 5,
        ["Rule", "Version", "Status", "Capacity workbook", "Saved", "What this rule does"],
        fmt["header"],
    )
    sheet.write(6, 0, "Rule1", fmt["name"])
    sheet.write(6, 1, "v1", fmt["center"])
    sheet.write(6, 2, "Active", fmt["active"])
    sheet.write(6, 3, CAPACITY_FILE, fmt["text"])
    sheet.write(6, 4, SAVED, fmt["center"])
    sheet.write(
        6, 5,
        "G900 cells only. Reduce to 2 TRX when all five gates pass. "
        "Then add G1800 TRX for high load, CDR, bad quality, or high ICM. "
        "After the reduction step, remove at most one more G900 TRX, and only while 3 or more TRX remain.",
        fmt["text"],
    )
    sheet.set_row(6, 48)

    sheet.write("A9", "How a later change is saved", fmt["section"])
    lines = [
        "This file stays as GSM_Rules_Store_v1.xlsx. It is not overwritten when a rule changes.",
        "The next change is saved as GSM_Rules_Store_v2.xlsx. That file keeps Rule1 and adds the new rule.",
        "A revision you still call Rule1 is stored as Rule1 v2. A replacement set is stored as Rule2.",
        "The Status column in the newest store file says which rule is active. Switching back sets Status to the rule you name.",
        "The capacity workbook is also versioned. A switch rebuilds it under the next GSM_Capacity_Enhancement version and leaves the earlier workbook in place.",
        "Yellow cells on Parameters are the saved numbers for Rule1. A changed number belongs in the next store version.",
    ]
    for index, line in enumerate(lines):
        sheet.merge_range(9 + index, 0, 9 + index, 5, line, fmt["text"])
        sheet.set_row(9 + index, 22)

    sheet.write("A17", "Sheets in this file", fmt["section"])
    sheets = [
        ("Store", "The list of saved rules and which one is active."),
        ("Rule1", "The rule in reading order."),
        ("Parameters", "Every threshold, with the name used by the capacity workbook."),
        ("Decision", "The same rule as a step table: when it runs, what TRX it adds, what TRX it removes."),
        ("Formulas", "Capacity, KPI, and decision formulas."),
        ("Data", "Which file supplies each KPI, how a cell is averaged, and how the band is chosen."),
        ("Examples", "Worked cells. A switch back to Rule1 should reproduce these results."),
        ("G1800 capacity", "TCH and Erlang of a new G1800 cell at the saved 60% half rate and 5% blocking."),
    ]
    write_header(sheet, 18, ["Sheet", "What it holds", "", "", "", ""], fmt["header"])
    for index, (name, purpose) in enumerate(sheets):
        row = 19 + index
        sheet.write(row, 0, name, fmt["name"])
        sheet.merge_range(row, 1, row, 5, purpose, fmt["text"])


def write_rule1(workbook, fmt):
    sheet = workbook.add_worksheet("Rule1")
    sheet.set_tab_color(ORANGE)
    sheet.set_column("A:A", 8)
    sheet.set_column("B:B", 24)
    sheet.set_column("C:C", 118)
    sheet.set_row(0, 24)
    sheet.write("A1", "Rule1", fmt["title"])
    sheet.merge_range(
        "A2:C2",
        f"Version v1. Status: Active. Saved {SAVED}. Matches {CAPACITY_FILE}.",
        fmt["text_open"],
    )
    sheet.set_row(1, 20)
    sheet.write("A4", "Read in this order", fmt["section"])
    write_header(sheet, 5, ["Order", "Step", "Rule"], fmt["header"])
    steps = [
        ("1", "Which cells", "Actions apply to G900 cells. A cell that already has 1800 TCH is left unchanged."),
        ("2", "Capacity", "Equivalent channels = available TCH (K3015) x (1 + 0.60). Capacity is the Erlang B offered traffic at 5% blocking for that channel count. A fraction between two rows is read on the straight line between them."),
        ("3", "Voice load", "Voice load % = busy-hour K3014 / Capacity x 100. Busy-hour Erlang is already one hour. Daily K3014, K3034, and K3004 are hourly averages, so those reference columns are multiplied by 24 and are not used in voice load."),
        ("4", "Half rate on the sheet", "Half Rate Traffic Ratio = busy-hour K3034 / K3014. It is a display KPI. Capacity uses the design ratio 0.60."),
        ("5", "G900 reduction", "Run this first. A G900 cell is reduced only when every gate is true: configured TRX above 2, TCH congestion below 1%, SDCCH congestion below 0.5%, busy-hour voice load at or below 40%, and busy-hour RR307 above 99%."),
        ("6", "TRX kept", "A cell that passes step 5 keeps 2 TRX. TCH after the cut = available TCH x 2 / configured TRX. Voice load is recalculated from that TCH on the 5% table."),
        ("7", "G1800, load", "Use the voice load from step 6. If it is at least 80%, add the smallest number of G1800 TRX that takes pooled voice load under 60%. Pooled capacity is the remaining G900 Erlang plus the new G1800 Erlang. The load count stops at 12 TRX."),
        ("8", "G1800, CDR", "If daily CDR_N average is at least 10, add 1 G1800 TRX."),
        ("9", "G1800, quality", "If DL bad quality is above 2% and UL bad quality is above 2%, add 1 G1800 TRX."),
        ("10", "G1800, high ICM", "If ICM band 4-5 is above 50%, and the reduction, the 80% load line, CDR, and bad quality did not trigger, add 1 G1800 TRX."),
        ("11", "G1800 total", "The four G1800 counts are added. A cell that needs 3 TRX for load and also has CDR_N at least 10 receives 4 G1800 TRX."),
        ("12", "One more G900 TRX", "Judge this on the TRX left after step 5. If 3 or more TRX remain, and CDR, bad quality, or the ICM line is true, remove exactly 1 G900 TRX with the G1800 add. CDR and bad quality share that one removal."),
        ("13", "Already at 2", "If step 5 left the cell at 2 TRX, remove no further G900 TRX. The G1800 add from steps 7 to 10 still stands."),
    ]
    for index, (order, name, rule) in enumerate(steps):
        row = 6 + index
        sheet.write(row, 0, order, fmt["step"])
        sheet.write(row, 1, name, fmt["name"])
        sheet.write(row, 2, rule, fmt["text"])
        sheet.set_row(row, 36)

    gate_row = 21
    sheet.write(gate_row, 0, "Five gates", fmt["section"])
    write_header(sheet, gate_row + 1, ["", "Gate", "Saved test"], fmt["header"])
    gates = [
        ("Configured TRX", '="above "&pTRXAbove'),
        ("TCH congestion", '="below "&pTCHCongMax&"%"'),
        ("SDCCH congestion", '="below "&pSDCCHCongMax&"%"'),
        ("Busy-hour voice load", '="at or below "&pVoiceLoadMax&"%"'),
        ("Busy-hour RR307", '="above "&pTCHAvailMin&"%"'),
    ]
    for index, (gate, formula) in enumerate(gates):
        row = gate_row + 2 + index
        sheet.write(row, 0, index + 1, fmt["step"])
        sheet.write(row, 1, gate, fmt["name"])
        sheet.write_formula(row, 2, formula, fmt["text"])

    outside = 30
    sheet.write(outside, 0, "Outside Rule1", fmt["section"])
    outside_lines = [
        "L900 moving from 5 MHz to 10 MHz is a later phase. Rule1 does not plan it.",
        "Rule1 does not choose a G1800 ARFCN or BCCH.",
        "G1800, Both, and NoTCH cells receive no TRX action.",
        "The extra G900 removal stops at one TRX. A second problem line adds its own G1800 TRX and does not take a second G900 TRX.",
    ]
    for index, line in enumerate(outside_lines):
        row = outside + 1 + index
        sheet.merge_range(row, 0, row, 2, line, fmt["text"])
        sheet.set_row(row, 22)


def write_parameters(workbook, fmt):
    sheet = workbook.add_worksheet("Parameters")
    sheet.set_tab_color("#BF8F00")
    sheet.freeze_panes(5, 0)
    widths = [22, 16, 14, 22, 28, 88]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.write("A1", "Rule1 parameters", fmt["title"])
    sheet.merge_range(
        "A2:F2",
        "Yellow cells are the saved Rule1 values. The capacity workbook reads these same names. "
        "A changed value is a new store version.",
        fmt["text_open"],
    )
    sheet.set_row(1, 32)
    write_header(
        sheet, 4,
        ["Name", "Saved value", "Unit", "Test", "Used in", "Meaning"],
        fmt["header"],
    )
    extra = {
        "pTRXAbove": ("above this", "G900 reduction"),
        "pTRXKeep": ("keep this many", "G900 reduction"),
        "pTCHCongMax": ("below this", "G900 reduction"),
        "pSDCCHCongMax": ("below this", "G900 reduction"),
        "pVoiceLoadMax": ("at or below this", "G900 reduction"),
        "pTCHAvailMin": ("above this", "G900 reduction"),
        "pHR": ("design ratio", "Capacity"),
        "pVoiceLoadHigh": ("at or above this", "G1800 load"),
        "pVoiceLoadTarget": ("bring load under this", "G1800 load"),
        "pCDRnMin": ("at or above this", "G1800 CDR"),
        "pBadQualMin": ("above this on both links", "G1800 quality"),
        "pICMMin": ("above this", "G1800 ICM"),
        "pMaxG1800TRX": ("stop at this", "G1800 load"),
        "pDailyDays": ("flag below this", "Sample length"),
        "pTrafficHours": ("multiply daily Erlang by this", "Daily reference"),
    }
    for offset, (name, value, unit, note) in enumerate(PARAMS):
        row = 5 + offset
        excel_row = row + 1
        test, used = extra[name]
        sheet.write(row, 0, name, fmt["name"])
        number_fmt = fmt["yellow"]
        sheet.write_number(row, 1, value, number_fmt)
        sheet.write(row, 2, unit, fmt["center"])
        sheet.write(row, 3, test, fmt["center"])
        sheet.write(row, 4, used, fmt["center"])
        sheet.write(row, 5, note, fmt["text"])
        workbook.define_name(name, f"=Parameters!$B${excel_row}")
        sheet.set_row(row, 32)


def write_decision(workbook, fmt):
    sheet = workbook.add_worksheet("Decision")
    sheet.set_tab_color(ORANGE)
    sheet.freeze_panes(5, 0)
    for col, width in enumerate([10, 24, 62, 28, 42, 36]):
        sheet.set_column(col, col, width)
    sheet.write("A1", "Rule1 decision order", fmt["title"])
    sheet.merge_range(
        "A2:F2",
        "Work down the rows. G1800 counts from rows 2 to 5 are added. Row 6 removes one G900 TRX at most.",
        fmt["text_open"],
    )
    sheet.set_row(1, 20)
    write_header(
        sheet, 4,
        ["Order", "Step", "Runs when", "G1800 TRX added", "G900 TRX removed", "Then"],
        fmt["header"],
    )
    rows = [
        (
            "1",
            "G900 reduction",
            '="G900, and all five gates: TRX > "&pTRXAbove&", TCH congestion < "&pTCHCongMax&"%, SDCCH congestion < "&pSDCCHCongMax&"%, voice load <= "&pVoiceLoadMax&"%, RR307 > "&pTCHAvailMin&"%"',
            "0",
            '="Keep "&pTRXKeep&" TRX. Removed = configured TRX - "&pTRXKeep',
            "Recalculate voice load, then continue.",
        ),
        (
            "2",
            "G1800 for load",
            '="G900, and voice load after row 1 is at least "&pVoiceLoadHigh&"%"',
            '="Smallest count that takes pooled load under "&pVoiceLoadTarget&"%. Cap "&pMaxG1800TRX',
            "0",
            "Continue. CDR and quality can still add.",
        ),
        (
            "3",
            "G1800 for CDR",
            '="G900, and daily CDR_N average is at least "&pCDRnMin',
            "1",
            "0 on this row. Row 6 may remove 1.",
            "Continue.",
        ),
        (
            "4",
            "G1800 for quality",
            '="G900, and DL bad quality > "&pBadQualMin&"% and UL bad quality > "&pBadQualMin&"%"',
            "1",
            "0 on this row. Row 6 may remove 1.",
            "Continue.",
        ),
        (
            "5",
            "G1800 for high ICM",
            '="G900, rows 1 to 4 did not trigger, and ICM band 4-5 > "&pICMMin&"%"',
            "1",
            "0 on this row. Row 6 may remove 1.",
            "Continue to row 6.",
        ),
        (
            "6",
            "One extra G900 TRX",
            '="Row 1 did not reduce the cell, TRX still there is at least 3, and row 3, 4, or 5 is true"',
            "0. The add is already in rows 2 to 5.",
            "Exactly 1. CDR takes it first, then quality, then ICM.",
            "Stop. A cell already at 2 loses nothing more.",
        ),
    ]
    for index, (order, step, when, add, remove, then) in enumerate(rows):
        row = 5 + index
        sheet.write(row, 0, order, fmt["step"])
        sheet.write(row, 1, step, fmt["name"])
        if when.startswith("="):
            sheet.write_formula(row, 2, when, fmt["text"])
        else:
            sheet.write(row, 2, when, fmt["text"])
        sheet.write(row, 3, add, fmt["text"])
        if isinstance(remove, str) and remove.startswith("="):
            sheet.write_formula(row, 4, remove, fmt["text"])
        else:
            sheet.write(row, 4, remove, fmt["text"])
        sheet.write(row, 5, then, fmt["text"])
        sheet.set_row(row, 48)

    sheet.write("A13", "G1800 cell size", fmt["section"])
    sheet.merge_range(
        "A14:F14",
        "TCH = 8 x TRX - 1 BCCH - one SDCCH for every two TRX. Equivalent channels = TCH x (1 + pHR). "
        "Capacity is that channel count on the 5% Erlang B table. See the G1800 capacity sheet.",
        fmt["text"],
    )
    sheet.set_row(13, 32)
    sheet.merge_range(
        "A15:F15",
        "Pooled capacity = remaining G900 Erlang + new G1800 Erlang. "
        "The load count is the first G1800 row whose capacity is strictly above K3014 / (target / 100) - remaining G900 Erlang.",
        fmt["text"],
    )
    sheet.set_row(14, 32)


def write_formulas(workbook, fmt):
    sheet = workbook.add_worksheet("Formulas")
    sheet.set_tab_color(BLUE)
    sheet.freeze_panes(5, 0)
    for col, width in enumerate([34, 22, 100]):
        sheet.set_column(col, col, width)
    sheet.write("A1", "Rule1 formulas", fmt["title"])
    sheet.merge_range(
        "A2:C2",
        "These are the formulas on All Cells in the capacity workbook. Names such as pHR are the Parameters sheet.",
        fmt["text_open"],
    )
    sheet.set_row(1, 20)
    write_header(sheet, 4, ["Result", "Group", "Formula"], fmt["header"])
    formulas = [
        ("Half Rate Traffic Ratio", "Display KPI", "K3034 / K3014. Zero when K3014 is 0. Not used in Capacity."),
        ("Equivalent channels", "Capacity", "Available TCHs_K3015_Avg * (1 + pHR)"),
        ("Capacity, Erl", "Capacity", "Erlang B at 5% for Equivalent channels. Below 1 channel, use the 1-channel row times the fraction. At 200 or above, use the 200-channel row. A whole number uses that row. A fraction is the straight line between the two nearest rows."),
        ("Voice Load, %", "Capacity", "Traffic Volume on TCH_K3014_Avg / Capacity, Erl * 100"),
        ("TCH Cong, %", "Reduction gate", "TCH Congestion_N_Avg / TCH Congestion_D_Avg * 100. Zero when the denominator is 0."),
        ("SDCCH Cong, %", "Reduction gate", "K3001_Avg / K3000_Avg * 100. Zero when K3000 is 0."),
        ("DL Bad Qual(6~7), %", "Quality", "DL Quality(6~7)_N_Avg / DL Quality_D_Avg * 100"),
        ("UL Bad Qual(6~7), %", "Quality", "UL Quality(6~7)_N_Avg / UL Quality_D_Avg * 100"),
        ("CDR, %", "Display KPI", "CDR_N_Avg / CDR_D_Avg * 100. The CDR gate uses CDR_N_Avg, not this percent."),
        ("ICM_B4~5, %", "ICM", "Interference Band_4-5_Samples_Avg / Interference Band Samples_D_Avg * 100"),
        ("HOSR_Out, %", "Display KPI", "HOSR_N_Avg / HOSR_D_Avg * 100"),
        ("Incoming HOSR, %", "Display KPI", "Incoming HOSR_N_Avg / Incoming HOSR_D_Avg * 100"),
        ("CSSR, %", "Display KPI", "CSSR_N_Avg / CSSR_D_Avg * 100"),
        ("Daily avg TCH, Erl", "Reference", "Daily K3014 average * pTrafficHours. Same pattern for K3034 and K3004. Not used in voice load."),
        ("G900 TRX Reduction", "Decision", 'Yes when Band is G900 and TRX > pTRXAbove and TCH cong < pTCHCongMax and SDCCH cong < pSDCCHCongMax and voice load <= pVoiceLoadMax and RR307 > pTCHAvailMin.'),
        ("TRX after reduction", "Decision", "If reduction is Yes, MIN(pTRXKeep, TRX configured). Otherwise TRX configured."),
        ("TCH after reduction", "Decision", "Available TCH * TRX after reduction / TRX configured."),
        ("Voice load after reduction", "Decision", "K3014 / Erlang B capacity of (TCH after reduction * (1 + pHR)) * 100"),
        ("G1800 TRX for load", "Decision", "On a G900 cell whose voice load after reduction is at least pVoiceLoadHigh: deficit = K3014 / (pVoiceLoadTarget / 100) - capacity after reduction. Add the smallest G1800 TRX count whose capacity is above that deficit, capped at pMaxG1800TRX."),
        ("G1800 TRX for CDR", "Decision", "1 when Band is G900 and CDR_N_Avg >= pCDRnMin."),
        ("G1800 TRX for quality", "Decision", "1 when Band is G900 and DL bad quality > pBadQualMin and UL bad quality > pBadQualMin."),
        ("G1800 TRX for ICM", "Decision", "1 when Band is G900, reduction is not Yes, the load flag is off, CDR is off, quality is off, and ICM > pICMMin."),
        ("G900 removed for CDR", "Decision", "1 when Band is G900, reduction is not Yes, TRX after reduction >= 3, and CDR triggers."),
        ("G900 removed for quality", "Decision", "1 when Band is G900, reduction is not Yes, TRX after reduction >= 3, quality triggers, and CDR does not."),
        ("G900 removed for ICM", "Decision", "1 when the ICM add is 1 and TRX after reduction >= 3."),
        ("G900 TRX final", "Decision", "If reduction is Yes, MIN(pTRXKeep, configured). Otherwise configured minus the three removal flags. Those three flags sum to 0 or 1."),
        ("TRX Remove", "Decision", "Configured TRX - G900 TRX final."),
        ("G1800 TRX Add", "Decision", "Load count + CDR count + quality count + ICM count."),
        ("TCH final", "Decision", "Available TCH * G900 TRX final / TRX configured."),
        ("Capacity final, Erl", "Decision", "Erlang B of (TCH final * (1 + pHR)) plus the G1800 capacity of G1800 TRX Add."),
        ("Voice Load final, %", "Decision", "K3014 / Capacity final * 100."),
    ]
    for index, (result, group, formula) in enumerate(formulas):
        row = 5 + index
        sheet.write(row, 0, result, fmt["name"])
        sheet.write(row, 1, group, fmt["center"])
        sheet.write(row, 2, formula, fmt["text"])
        sheet.set_row(row, 32)


def write_data(workbook, fmt):
    sheet = workbook.add_worksheet("Data")
    sheet.set_tab_color(NAVY)
    sheet.freeze_panes(5, 0)
    for col, width in enumerate([34, 18, 62, 40]):
        sheet.set_column(col, col, width)
    sheet.write("A1", "Rule1 data", fmt["title"])
    sheet.merge_range(
        "A2:D2",
        "Each cell is averaged over the dates that cell has. The window is 26 Oct 2026 through 1 Nov 2026. A short sample is flagged against pDailyDays.",
        fmt["text_open"],
    )
    sheet.set_row(1, 32)
    write_header(sheet, 4, ["Item", "File", "Counter", "How Rule1 uses it"], fmt["header"])
    rows = [
        ("Voice traffic", "Busy hour", "K3014 Traffic Volume on TCH (Erl)", "Date average. Already one hour. Numerator of voice load."),
        ("Half-rate traffic", "Busy hour", "K3034 TCHH Traffic Volume (Erl)", "Date average. Display ratio K3034 / K3014."),
        ("Available TCH", "Busy hour", "K3015 Available TCHs", "Date average. Becomes equivalent channels with the 0.60 design ratio."),
        ("Configured TCH", "Busy hour", "K3016 Configured TCHs", "Date average. Display."),
        ("TRX configured", "Busy hour", "S3655 Number of configured TRXs", "Mode of the rounded daily values for that cell."),
        ("TCH congestion", "Busy hour", "TCH Congestion_N / TCH Congestion_D", "Sum of N over sum of D across the cell's dates."),
        ("SDCCH congestion", "Busy hour", "K3001 / K3000", "Sum of N over sum of D across the cell's dates."),
        ("TCH availability", "Busy hour", "RR307 TCH Availability (%)", "Date average of the busy-hour value. This is the availability gate."),
        ("SDCCH traffic", "Busy hour", "K3004 Traffic Volume on SDCCH (Erl)", "Date average. Display."),
        ("Band", "Busy hour", "R3027A + R3028A for 900, R3027B + R3028B for 1800", "G900 when 900 samples are above 0 and 1800 samples are 0. G1800 is the reverse. Both when each side is above 0. NoTCH when both are 0."),
        ("CDR", "Daily", "CDR_N and CDR_D", "Date average of each. The gate uses CDR_N average. CDR % is N / D."),
        ("Bad quality", "Daily", "DL and UL Quality (6~7) N and D", "Date average of each. The gate needs both DL % and UL % above the threshold."),
        ("ICM", "Daily", "Interference Band 4-5 samples and the sample denominator", "Date average of each. ICM % is N / D."),
        ("HOSR, incoming HOSR, CSSR", "Daily", "Matching N and D counters", "Date average of each. Display KPIs. They do not trigger a TRX action."),
        ("Daily traffic", "Daily", "K3014, K3034, K3004", "Date average, then x 24, because the daily file stores an hourly average. Reference only."),
    ]
    for index, values in enumerate(rows):
        row = 5 + index
        for col, value in enumerate(values):
            style = fmt["name"] if col == 0 else fmt["text"]
            sheet.write(row, col, value, style)
        sheet.set_row(row, 36)

    sheet.write("A22", "Files", fmt["section"])
    files = [
        "V1 busy hour and V1 daily: MYMH05, MYMH06, MYMH07, SYLH04.",
        "V2 busy hour and V2 daily: TNGH12.",
        "Join key: source, GBSC, cell CI, cell name.",
    ]
    for index, line in enumerate(files):
        sheet.merge_range(22 + index, 0, 22 + index, 3, line, fmt["text"])


def write_examples(workbook, fmt, table, caps, cfg):
    sheet = workbook.add_worksheet("Examples")
    sheet.set_tab_color(GREEN)
    sheet.freeze_panes(5, 0)
    headers = [
        "Case", "Band", "TRX", "Voice %", "TCH cong %", "SDCCH cong %", "RR307",
        "CDR_N", "DL %", "UL %", "ICM %", "K3014 Erl", "Available TCH",
        "Reduction", "TRX final", "TRX removed", "G1800 add",
        "For load", "For CDR", "For quality", "For ICM", "What Rule1 does",
    ]
    for col, header in enumerate(headers):
        sheet.set_column(col, col, 16 if col else 28)
    sheet.set_column(21, 21, 62)
    sheet.write("A1", "Rule1 examples", fmt["title"])
    sheet.merge_range(
        "A2:V2",
        "A switch back to Rule1 should reproduce these rows at the saved parameter values.",
        fmt["text_open"],
    )
    write_header(sheet, 4, headers, fmt["header"])
    cases = [
        ("Quiet cell, all five gates", "G900", 4, 20, 0.2, 0.1, 99.5, 1, 0.4, 0.4, 10, 2.0, 30,
         "Keep 2 TRX. No G1800 add."),
        ("Voice already high, only 2 TRX", "G900", 2, 90, 0.1, 0.1, 100, 0, 0.1, 0.1, 1, 16, 14,
         "Reduction does not run. G1800 is sized from the load rule. No extra G900 removal."),
        ("CDR and bad quality, 4 TRX", "G900", 4, 50, 0.2, 0.2, 100, 12, 3, 3, 80, 8, 30,
         "No full reduction. Add 1 G1800 for CDR and 1 for quality. Remove exactly 1 G900 TRX."),
        ("High ICM only, 4 TRX", "G900", 4, 20, 0.2, 0.2, 98, 1, 0.2, 0.2, 70, 3, 30,
         "Availability blocks the reduction. ICM adds 1 G1800 and removes 1 G900."),
        ("Reduced, and CDR is high", "G900", 4, 20, 0.2, 0.1, 100, 15, 0.2, 0.2, 10, 2, 30,
         "Keep 2 TRX. Add 1 G1800 for CDR. Remove no further G900 TRX."),
        ("CDR and quality, only 3 TRX", "G900", 3, 50, 0.2, 0.2, 100, 12, 3, 4, 10, 6, 21,
         "Add 2 G1800 TRX. Remove 1 G900 TRX, so the cell ends at 2."),
        ("Quality only, 5 TRX", "G900", 5, 50, 0.2, 0.2, 100, 1, 3, 3, 10, 8, 40,
         "Add 1 G1800 for quality. Remove 1 G900 TRX."),
        ("G1800 cell with every trigger", "G1800", 4, 10, 0, 0, 100, 20, 5, 5, 90, 4, 20,
         "No action. Rule1 changes G900 cells."),
    ]
    for index, case in enumerate(cases):
        label, band, trx, voice, tch, sdcch, rr307, cdr, dl, ul, icm, carried, avail, why = case
        plan = decide(band, trx, voice, tch, sdcch, rr307, cdr, dl, ul, icm, carried, avail, table, caps, cfg)
        values = [
            label, band, trx, voice, tch, sdcch, rr307, cdr, dl, ul, icm, carried, avail,
            "Yes" if plan["reduce"] else "No",
            plan["trx_final"], plan["trx_remove"], plan["n_add"],
            plan["n_load"], plan["n_cdr"], plan["n_qual"], plan["n_icm"], why,
        ]
        row = 5 + index
        for col, value in enumerate(values):
            if col == 0:
                style = fmt["name"]
            elif col == 21:
                style = fmt["text"]
            elif isinstance(value, float):
                style = fmt["num"]
            elif isinstance(value, int):
                style = fmt["int"]
            else:
                style = fmt["center"]
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                sheet.write_number(row, col, value, style)
            else:
                sheet.write(row, col, value, style)
        sheet.set_row(row, 32)
        if plan["removed_cdr"] + plan["removed_qual"] + plan["removed_icm"] > 1:
            raise SystemExit(f"{label} removed more than one extra G900 TRX")


def write_g1800(workbook, fmt, table, caps):
    sheet = workbook.add_worksheet("G1800 capacity")
    sheet.set_tab_color(BLUE)
    sheet.freeze_panes(5, 0)
    for col, width in enumerate([12, 14, 12, 12, 12, 24, 18]):
        sheet.set_column(col, col, width)
    sheet.write("A1", "New G1800 cell under Rule1", fmt["title"])
    sheet.merge_range(
        "A2:G2",
        "TCH = 8 x TRX - 1 BCCH - one SDCCH per two TRX. Equivalent channels = TCH x (1 + pHR). "
        "Capacity is the 5% Erlang B value at the saved half-rate ratio.",
        fmt["text_open"],
    )
    sheet.set_row(1, 32)
    write_header(
        sheet, 4,
        ["TRX", "Timeslots", "BCCH", "SDCCH", "TCH", "Equivalent channels", "Capacity, Erl"],
        fmt["header_blue"],
    )
    for trx in range(0, G1800_TABLE_TRX + 1):
        row = 5 + trx
        excel_row = row + 1
        sheet.write_number(row, 0, trx, fmt["int"])
        sheet.write_formula(row, 1, f"=A{excel_row}*8", fmt["int"], trx * 8)
        sheet.write_formula(row, 2, f"=IF(A{excel_row}=0,0,1)", fmt["int"], 0 if trx == 0 else 1)
        sheet.write_formula(row, 3, f"=IF(A{excel_row}=0,0,ROUNDUP(A{excel_row}/2,0))", fmt["int"], 0 if trx == 0 else (trx + 1) // 2)
        sheet.write_formula(row, 4, f"=IF(A{excel_row}=0,0,B{excel_row}-C{excel_row}-D{excel_row})", fmt["int"], g1800_tch(trx))
        sheet.write_formula(row, 5, f"=E{excel_row}*(1+pHR)", fmt["num"], g1800_tch(trx) * (1 + defaults()["pHR"]))
        sheet.write_number(row, 6, caps[trx], fmt["num"])

    sheet.write("I4", "Erlang B 5% checks", fmt["section"])
    sheet.write("I5", "Channels", fmt["header"])
    sheet.write("J5", "Offered Erl", fmt["header"])
    sheet.write("K5", "Why it is marked", fmt["header"])
    sheet.set_column("I:I", 14)
    sheet.set_column("J:J", 16)
    sheet.set_column("K:K", 42)
    checks = [
        (6, "1 TRX, BCCH + SDCCH. Huawei prints 2.96"),
        (7, "1 TRX, BCCH and SDCCH combined. Huawei prints 3.74"),
        (10, "10 full-rate TCH. 6.22 Erl"),
        (14, "2 TRX. Huawei 5% column is 9.73 Erl"),
        (16, "10 TCH at 60% half rate. 11.54 Erl"),
        (21, "3 TRX. Huawei 5% column is 16.19 Erl"),
        (29, "4 TRX. Huawei 5% column is 23.83 Erl"),
        (36, "5 TRX. Huawei 5% column is 30.66 Erl"),
    ]
    for index, (channels, why) in enumerate(checks):
        row = 5 + index
        sheet.write_number(row, 8, channels, fmt["int"])
        sheet.write_number(row, 9, offered_at(channels, table), fmt["num"])
        sheet.write(row, 10, why, fmt["text"])


if __name__ == "__main__":
    build()
