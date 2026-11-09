#!/usr/bin/env python3
"""Erlang B reference tables at 5% and 2% blocking, for GSM TCH dimensioning.

Offered traffic is the largest A (Erlang) such that the Erlang B blocking
probability on N channels is the sheet's blocking. Half-rate rows use
equivalent channels

    N x (1 + half-rate traffic ratio)

and read that channel count from the matching Erlang B sheet. A fractional
channel count is interpolated between the two nearest integer rows.

Erlang_B_Table_5pct_v1.xlsx stays as the first 5% file. This writer builds
Erlang_B_Table_v2.xlsx with both blocking rates. The capacity workbook is
not changed by this file. Capacity still reads the 5% half-rate column.
"""

from math import ceil
from pathlib import Path

import xlsxwriter

ROOT = Path(__file__).resolve().parent
VERSION = "v2"
OUT = ROOT / f"Erlang_B_Table_{VERSION}.xlsx"
GOS = 0.05
GOS_2 = 0.02
MAX_CHANNELS = 120
MAX_PHYSICAL = 60
# Fractions keep 10 x (1 + 3/5) equal to 16 in Excel. A decimal 0.6 does not.
HR_RATIOS = [
    (0.0, "0", "0% half rate"),
    (0.2, "1/5", "20% half rate"),
    (0.4, "2/5", "40% half rate"),
    (0.5, "1/2", "50% half rate"),
    (0.6, "3/5", "60% half rate"),
    (0.8, "4/5", "80% half rate"),
    (1.0, "1", "100% half rate"),
]

# Channel counts Huawei prints, plus the 10 TCH and 16 TCH cases discussed.
HUAWEI_NOTES = {
    6: "1 TRX: 8 timeslots minus BCCH and one SDCCH",
    7: "1 TRX: BCCH and SDCCH share one timeslot",
    10: "10 physical TCH, full rate",
    13: "2 TRX with BCCH and two SDCCH",
    14: "2 TRX: 16 timeslots minus BCCH and one SDCCH. Huawei table row",
    16: "10 physical TCH at 60% half rate (10 x 1.6)",
    21: "3 TRX: 24 timeslots minus BCCH and two SDCCH. Huawei example",
    22: "3 TRX with BCCH and one SDCCH",
    29: "4 TRX: 32 timeslots minus BCCH and two SDCCH. Huawei table row",
    30: "4 TRX with BCCH and one SDCCH",
    36: "5 TRX: 40 timeslots minus BCCH and three SDCCH. Huawei table row",
}

# Published Huawei course values at 5% blocking, used as a build check.
HUAWEI_PUBLISHED = {
    6: 2.960,
    7: 3.738,
    14: 9.730,
    21: 16.19,
    29: 23.83,
    36: 30.66,
}


def erlang_b(channels, offered):
    """Blocking probability. Calls that find every channel busy are cleared."""
    if channels <= 0:
        return 1.0
    if offered <= 0:
        return 0.0
    blocking = 1.0
    for k in range(1, channels + 1):
        blocking = offered * blocking / (k + offered * blocking)
    return blocking


def offered_traffic(channels, gos=GOS):
    """Offered Erlang at the target blocking, by bisection."""
    low = 1e-12
    high = max(float(channels) * 2.0, 1.0)
    for _ in range(80):
        mid = (low + high) / 2.0
        if erlang_b(channels, mid) < gos:
            low = mid
        else:
            high = mid
    return (low + high) / 2.0


def offered_at(channels, table):
    """Integer lookup, or linear interpolation between the two nearest rows."""
    if channels <= 0:
        return 0.0
    if channels >= MAX_CHANNELS:
        return table[MAX_CHANNELS]
    lower = int(channels)
    if lower < 1:
        return table[1] * channels
    fraction = channels - lower
    if fraction < 1e-12 or lower >= MAX_CHANNELS:
        return table[lower]
    return table[lower] + fraction * (table[lower + 1] - table[lower])


def build_table(gos):
    table = {}
    for channels in range(1, MAX_CHANNELS + 1):
        traffic = offered_traffic(channels, gos)
        blocking = erlang_b(channels, traffic)
        if abs(blocking - gos) > 1e-9:
            raise SystemExit(f"N={channels} blocking {blocking} is not {gos}")
        table[channels] = traffic
    if abs(gos - GOS) < 1e-12:
        for channels, published in HUAWEI_PUBLISHED.items():
            if abs(table[channels] - published) > 0.006:
                raise SystemExit(
                    f"N={channels} is {table[channels]:.4f}, Huawei prints {published}"
                )
    return table


def main():
    table = build_table(GOS)
    table2 = build_table(GOS_2)
    workbook = xlsxwriter.Workbook(str(OUT), {"strings_to_urls": False})
    navy = "#1F4E79"
    header = workbook.add_format({
        "bold": True, "font_name": "Calibri", "font_size": 10, "font_color": "white",
        "bg_color": navy, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1,
    })
    header_hr = workbook.add_format({
        "bold": True, "font_name": "Calibri", "font_size": 10, "font_color": "white",
        "bg_color": "#C65911", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1,
    })
    title = workbook.add_format({"bold": True, "font_name": "Calibri", "font_size": 16, "font_color": navy})
    section = workbook.add_format({"bold": True, "font_name": "Calibri", "font_size": 12, "font_color": navy})
    note = workbook.add_format({"font_name": "Calibri", "font_size": 11, "text_wrap": True, "valign": "top"})
    n_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "num_format": "0",
    })
    n_hi = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "num_format": "0",
        "bg_color": "#FFF2CC", "bold": True,
    })
    erl = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "num_format": "0.00",
    })
    erl_hi = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "num_format": "0.00",
        "bg_color": "#FFF2CC", "bold": True,
    })
    text = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "valign": "vcenter", "text_wrap": True,
    })
    text_hi = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "valign": "vcenter", "text_wrap": True,
        "bg_color": "#FFF2CC",
    })
    label = workbook.add_format({"font_name": "Calibri", "font_size": 11, "bold": True, "font_color": navy})

    write_erlang_sheet(workbook, table, "5%", "Erlang B 5%", header, title, note, n_fmt, n_hi, erl, erl_hi, text, text_hi)
    write_half_rate_sheet(workbook, table, "5%", "Half Rate 5%", "Erlang B 5%", header, header_hr, title, note, n_fmt, n_hi, erl, erl_hi, label)
    write_erlang_sheet(workbook, table2, "2%", "Erlang B 2%", header, title, note, n_fmt, n_hi, erl, erl_hi, text, text_hi)
    write_half_rate_sheet(workbook, table2, "2%", "Half Rate 2%", "Erlang B 2%", header, header_hr, title, note, n_fmt, n_hi, erl, erl_hi, label)
    write_trx_sheet(workbook, table, header, title, note, n_fmt, erl, text, label)
    write_readme(workbook, table, table2, title, section, note)
    workbook.close()
    print(OUT.name)
    print(f"10 TCH full rate at 5% = {table[10]:.2f} Erl")
    print(f"16 channels (10 TCH, 60% HR) at 5% = {table[16]:.2f} Erl")
    print(f"10 TCH full rate at 2% = {table2[10]:.2f} Erl")
    print(f"16 channels (10 TCH, 60% HR) at 2% = {table2[16]:.2f} Erl")
    print(f"14 TCH at 2% = {table2[14]:.2f} Erl")
    print(f"21 TCH at 2% = {table2[21]:.2f} Erl")


def write_erlang_sheet(workbook, table, pct, sheet_name, header, title, note, n_fmt, n_hi, erl, erl_hi, text, text_hi):
    sheet = workbook.add_worksheet(sheet_name)
    sheet.set_tab_color("#1F4E79" if pct == "5%" else "#2F5496")
    sheet.freeze_panes(5, 1)
    sheet.set_column("A:A", 16)
    sheet.set_column("B:B", 28)
    sheet.set_column("C:C", 18)
    sheet.set_column("D:D", 62)
    sheet.write("A1", f"Erlang B table at {pct} blocking  {VERSION}", title)
    if pct == "5%":
        intro = ("Offered traffic is the Erlang a cell can be offered when 5% of calls are blocked. "
                 "Huawei's capacity course prints this as the 5 column. The 2% table is on Erlang B 2% and Half Rate 2%.")
    else:
        intro = ("Offered traffic is the Erlang a cell can be offered when 2% of calls are blocked. "
                 "Huawei's capacity course uses 2% for a dense area. The 5% table stays on Erlang B 5% and Half Rate 5%.")
    sheet.write("A2", intro, note)
    sheet.write("A3", f"B(N, A) = (A^N / N!) / sum(A^k / k! for k = 0 to N) = {pct}. A is offered traffic. N is the number of channels.", note)
    sheet.set_row(1, 32)
    sheet.set_row(2, 32)
    sheet.set_row(4, 30)
    headers = ["Channels N", f"Offered traffic at {pct}, Erl", "Erl per channel", "Where this channel count is used"]
    for col, name in enumerate(headers):
        sheet.write(4, col, name, header)
    sheet.autofilter(4, 0, 4 + MAX_CHANNELS, 3)
    for channels in range(1, MAX_CHANNELS + 1):
        row = 5 + channels - 1
        marked = channels in HUAWEI_NOTES
        sheet.write_number(row, 0, channels, n_hi if marked else n_fmt)
        sheet.write_number(row, 1, table[channels], erl_hi if marked else erl)
        sheet.write_formula(
            row, 2, f"=IF(A{row + 1}=0,\"\",B{row + 1}/A{row + 1})",
            erl_hi if marked else erl, table[channels] / channels,
        )
        sheet.write(row, 3, HUAWEI_NOTES.get(channels, ""), text_hi if marked else text)
    sheet.repeat_rows(4)
    sheet.set_landscape()
    sheet.fit_to_pages(1, 0)
    sheet.set_header(f"&LErlang B {pct}&R" + VERSION)
    sheet.set_footer("&LChannels 1 to 120&RPage &P")


def write_half_rate_sheet(workbook, table, pct, sheet_name, source_sheet, header, header_hr, title, note, n_fmt, n_hi, erl, erl_hi, label):
    sheet = workbook.add_worksheet(sheet_name)
    sheet.set_tab_color("#C65911" if pct == "5%" else "#833C0C")
    sheet.freeze_panes(6, 1)
    sheet.set_column("A:A", 16)
    sheet.write("A1", f"Half-rate capacity read from the {pct} Erlang B table", title)
    sheet.write("A2", "Equivalent channels = physical TCH x (1 + half-rate traffic ratio). "
                "Half-rate traffic ratio = TCHH traffic / TCH traffic = K3034 / K3014. "
                "One physical TCH used as half rate carries two calls, so 60% half rate on 10 TCH is 16 channels.", note)
    sheet.write("A3", f"Offered Erlang is looked up on {source_sheet}. A value between two channel counts is a straight-line interpolation of those two rows.", note)
    sheet.set_row(1, 36)
    sheet.set_row(2, 32)
    sheet.set_row(4, 22)
    sheet.set_row(5, 30)
    sheet.merge_range(4, 0, 5, 0, "Physical TCH", header)
    for index, (_ratio, expr, ratio_label) in enumerate(HR_RATIOS):
        col = 1 + index * 2
        sheet.set_column(col, col, 16)
        sheet.set_column(col + 1, col + 1, 16)
        group = header_hr if expr == "3/5" else header
        sheet.merge_range(4, col, 4, col + 1, ratio_label, group)
        sheet.write(5, col, "Equivalent channels", group)
        sheet.write(5, col + 1, f"Offered Erl at {pct}", group)
    first_data = 6
    for physical in range(1, MAX_PHYSICAL + 1):
        row = first_data + physical - 1
        marked = physical in (6, 7, 10, 13, 14, 21, 29, 36)
        sheet.write_number(row, 0, physical, n_hi if marked else n_fmt)
        for index, (ratio, expr, _ratio_label) in enumerate(HR_RATIOS):
            col = 1 + index * 2
            excel_row = row + 1
            equiv_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, col)
            equivalent = physical * (1.0 + ratio)
            highlight = marked or expr == "3/5"
            sheet.write_formula(
                row, col,
                f"=$A{excel_row}*(1+{expr})",
                erl_hi if highlight else erl,
                equivalent,
            )
            sheet.write_formula(
                row, col + 1,
                lookup_formula(equiv_cell, source_sheet),
                erl_hi if highlight else erl,
                offered_at(equivalent, table),
            )
    # Physical TCH 10 is Excel row 16 (header uses rows 5 and 6).
    example_col = 1 + len(HR_RATIOS) * 2 + 1
    sheet.set_column(example_col, example_col, 42)
    sheet.set_column(example_col + 1, example_col + 1, 14)
    sheet.write(4, example_col, "10 TCH example", label)
    sheet.write(5, example_col, f"Full rate, offered Erl at {pct}", note)
    sheet.write_formula(5, example_col + 1, "=C16", erl, table[10])
    sheet.write(6, example_col, "60% half rate, equivalent channels", note)
    sheet.write_formula(6, example_col + 1, "=J16", erl, 16)
    sheet.write(7, example_col, f"60% half rate, offered Erl at {pct}", note)
    sheet.write_formula(7, example_col + 1, "=K16", erl, table[16])
    sheet.write(9, example_col, f"Row 16 is 10 physical TCH. Full rate is {table[10]:.2f} Erl. "
                f"At 60% half rate those 10 TCH are 16 channels and carry {table[16]:.2f} Erl at {pct} blocking.", note)
    sheet.repeat_rows(4, 5)
    sheet.set_landscape()
    sheet.fit_to_pages(1, 0)
    sheet.set_header(f"&LHalf rate at {pct} blocking&R" + VERSION)


def lookup_formula(equiv_cell, source_sheet):
    """INDEX/MATCH interpolation. Avoids functions that show #NAME? in older spreadsheets."""
    channels = f"'{source_sheet}'!$A$6:$A$125"
    offered = f"'{source_sheet}'!$B$6:$B$125"
    lower = f"INT({equiv_cell})"
    exact = (
        f"INDEX({offered},MATCH({equiv_cell},{channels},0))"
    )
    base = f"INDEX({offered},MATCH({lower},{channels},0))"
    upper = f"INDEX({offered},MATCH({lower}+1,{channels},0))"
    return (
        f'=IF({equiv_cell}=INT({equiv_cell}),{exact},'
        f'{base}+({equiv_cell}-INT({equiv_cell}))*({upper}-{base}))'
    )


def write_trx_sheet(workbook, table, header, title, note, n_fmt, erl, text, label):
    sheet = workbook.add_worksheet("Huawei TRX")
    sheet.set_tab_color("#548235")
    sheet.freeze_panes(5, 0)
    sheet.set_column("A:A", 12)
    sheet.set_column("B:E", 14)
    sheet.set_column("F:J", 22)
    sheet.set_column("K:K", 54)
    sheet.write("A1", "Typical Huawei TRX channel count at 5% blocking", title)
    sheet.write("A2", "Each TRX has 8 timeslots. One timeslot is the BCCH. "
                "Huawei's course uses about one SDCCH for every two TRX. "
                "TCH = 8 x TRX - 1 BCCH - SDCCH. Offered Erlang is read from Erlang B 5%.", note)
    sheet.set_row(1, 32)
    sheet.set_row(4, 32)
    headers = [
        "TRX", "Timeslots", "BCCH", "SDCCH", "TCH",
        "Full rate, Erl at 5%",
        "Channels at 60% HR", "60% HR, Erl at 5%",
        "Channels at 100% HR", "100% HR, Erl at 5%",
        "Channel rule",
    ]
    for col, name in enumerate(headers):
        sheet.write(4, col, name, header)
    for trx in range(1, 9):
        row = 4 + trx
        sdcch = ceil(trx / 2)
        tch = 8 * trx - 1 - sdcch
        excel_row = row + 1
        sheet.write_number(row, 0, trx, n_fmt)
        sheet.write_formula(row, 1, f"=A{excel_row}*8", n_fmt, 8 * trx)
        sheet.write_number(row, 2, 1, n_fmt)
        sheet.write_number(row, 3, sdcch, n_fmt)
        sheet.write_formula(row, 4, f"=B{excel_row}-C{excel_row}-D{excel_row}", n_fmt, tch)
        sheet.write_formula(row, 5, lookup_formula(f"E{excel_row}", "Erlang B 5%"), erl, table[tch])
        sheet.write_formula(row, 6, f"=E{excel_row}*8/5", erl, tch * 1.6)
        sheet.write_formula(row, 7, lookup_formula(f"G{excel_row}", "Erlang B 5%"), erl, offered_at(tch * 1.6, table))
        sheet.write_formula(row, 8, f"=E{excel_row}*2", erl, tch * 2)
        sheet.write_formula(row, 9, lookup_formula(f"I{excel_row}", "Erlang B 5%"), erl, table[tch * 2])
        sheet.write(row, 10, f"{tch} TCH. Full rate {table[tch]:.2f} Erl at 5%.", text)
    sheet.write("A15", "Worked Huawei rows on the 5% column", label)
    sheet.write("A16", "14 TCH (2 TRX) = 9.73 Erl. 21 TCH (3 TRX) = 16.19 Erl. "
                "29 TCH (4 TRX) = 23.83 Erl. 36 TCH (5 TRX) = 30.66 Erl. "
                "These match the 5 column in Huawei's Erlang B slide.", note)
    sheet.set_row(15, 36)
    sheet.set_landscape()
    sheet.fit_to_pages(1, 1)


def write_readme(workbook, table, table2, title, section, note):
    sheet = workbook.add_worksheet("How to read")
    sheet.set_tab_color("#833C0C")
    sheet.set_column("A:A", 28)
    sheet.set_column("B:B", 110)
    sheet.write("A1", f"How to use Erlang_B_Table_{VERSION}.xlsx", title)
    rows = [
        ("Two blocking rates", "Erlang B 5% and Half Rate 5% are the 5% tables. Erlang B 2% and Half Rate 2% are the same layout at 2% blocking."),
        ("What the number means", "Offered traffic in Erlang. It is the design capacity of that many channels, not the carried traffic already on the cell."),
        ("Erlang B sheets", "One row per channel count, from 1 to 120. Yellow rows are the Huawei TRX counts and the 10 TCH / 16 channel examples."),
        ("10 TCH at 5%", f"Full rate {table[10]:.2f} Erl. At 60% half rate, 16 channels carry {table[16]:.2f} Erl."),
        ("10 TCH at 2%", f"Full rate {table2[10]:.2f} Erl. At 60% half rate, 16 channels carry {table2[16]:.2f} Erl."),
        ("Half-rate rule", "Equivalent channels = physical TCH x (1 + K3034 / K3014). The offered Erlang is then read from the matching blocking sheet. Half rate does not multiply the Erlang by 1.6."),
        ("Huawei 5% check", "14 TCH = 9.73 Erl, 21 TCH = 16.19 Erl, 29 TCH = 23.83 Erl, 36 TCH = 30.66 Erl. Same 5% column as the Huawei capacity course."),
        ("Huawei 2% column", f"14 TCH = {table2[14]:.2f} Erl, 21 TCH = {table2[21]:.2f} Erl, 29 TCH = {table2[29]:.2f} Erl, 36 TCH = {table2[36]:.2f} Erl."),
        ("Capacity file", "GSM_Capacity_Enhancement_v14.xlsx still reads Offered Erl at 5% from the 60% column. This 2% table is the reference beside it. Erlang_B_Table_5pct_v1.xlsx stays unchanged."),
    ]
    sheet.write("A3", "Item", section)
    sheet.write("B3", "Rule", section)
    for offset, (item, rule) in enumerate(rows):
        sheet.write(3 + offset, 0, item, section)
        sheet.write(3 + offset, 1, rule, note)
        sheet.set_row(3 + offset, 36)
    sheet.set_row(1, 22)


if __name__ == "__main__":
    main()
