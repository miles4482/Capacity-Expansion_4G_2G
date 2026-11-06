#!/usr/bin/env python3
"""Build the GSM capacity-enhancement workbook from the busy-hour capacity files.

Uses only:
  2G Capacity Exercise_V1_BH.xlsb
  2G Capacity Exercise_V2_BH.xlsb

The daily capacity files are the same KPI family for one full day. The rules
given for this step are busy-hour rules, so the daily files are not inputs.
"""

from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pyxlsb import open_workbook
import xlsxwriter

FILES = [
    ("V1", "2G Capacity Exercise_V1_BH.xlsb"),
    ("V2", "2G Capacity Exercise_V2_BH.xlsb"),
]
OUT = "GSM_Capacity_Enhancement.xlsx"
EXCEL_EPOCH = datetime(1899, 12, 30)

# Default thresholds. The same numbers live on the Parameters sheet and the
# All Cells formulas read them by name, so a threshold edit recalculates.
PARAMS = [
    ("pVoiceLoadMax", 60, "%", "G900 reduction: voice load at or below this is one OR condition"),
    ("pCDRnMin", 10, "drops", "CDR_N busy-hour mean at or above this"),
    ("pBadQualMin", 2, "%", "DL and UL bad quality (6-7) each above this"),
    ("pCongMax", 1, "%", "SDCCH congestion and TCH congestion must each be below this"),
    ("pHOSRMax", 90, "%", "G1800 addition: HOSR_Out at or below this"),
    ("pVoiceLoadHigh", 70, "%", "Labels the G1800 reason as the high-load or low-load line"),
    ("pCDRMin", 0.2, "%", "G1800 interference line: CDR above this"),
    ("pICMMin", 50, "%", "G1800 interference line: ICM band 4-5 above this"),
    ("pTRXAbove", 2, "TRX", "G900 reduction applies only when configured TRX is above this"),
    ("pTRXKeep", 2, "TRX", "Configured TRX kept on a cell that is reduced"),
]


def num(value):
    if value is None or value == "" or value == " ":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def clean_id(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if text.endswith(".0"):
        head = text[:-2]
        if head.lstrip("-").isdigit():
            return int(head)
    return text


def col_letter(index):
    """1-based column index to Excel letters."""
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def load_cells():
    cells = {}
    for source, path in FILES:
        with open_workbook(path) as workbook:
            with workbook.get_sheet("Sheet1") as sheet:
                rows = sheet.rows()
                header = [cell.v for cell in next(rows)]
                index = {name: pos for pos, name in enumerate(header)}

                def col(rowmap, name):
                    return num(rowmap.get(index[name]))

                for row in rows:
                    rowmap = {cell.c: cell.v for cell in row}
                    gbsc = str(rowmap.get(index["GBSC"])).strip()
                    ci = clean_id(rowmap.get(index["Cell CI"]))
                    name = str(rowmap.get(index["Cell Name"])).strip()
                    key = (source, gbsc, ci, name)
                    item = cells.get(key)
                    if item is None:
                        item = {
                            "source": source,
                            "gbsc": gbsc,
                            "ci": ci,
                            "name": name,
                            "cell_index": clean_id(rowmap.get(index["CellIndex"])),
                            "days": 0,
                            "trx": [],
                            "dates": set(),
                            "k3014": 0.0,
                            "k3034": 0.0,
                            "k3015": 0.0,
                            "tchn": 0.0,
                            "tchd": 0.0,
                            "sdn": 0.0,
                            "sdd": 0.0,
                            "dln": 0.0,
                            "dld": 0.0,
                            "uln": 0.0,
                            "uld": 0.0,
                            "cdrn": 0.0,
                            "cdrd": 0.0,
                            "cdrn_max": 0.0,
                            "hosrn": 0.0,
                            "hosrd": 0.0,
                            "icmn": 0.0,
                            "icmd": 0.0,
                            "a900": 0.0,
                            "b1800": 0.0,
                            "h900": 0.0,
                            "h1800": 0.0,
                        }
                        cells[key] = item
                    item["days"] += 1
                    item["trx"].append(col(rowmap, "S3655:Number of configured TRXs in a cell"))
                    raw_date = rowmap.get(index["Date"])
                    if isinstance(raw_date, (int, float)):
                        item["dates"].add(int(raw_date))
                    item["k3014"] += col(rowmap, "K3014:Traffic Volume on TCH(Erl)")
                    item["k3034"] += col(rowmap, "K3034:TCHH Traffic Volume(Erl)")
                    item["k3015"] += col(rowmap, "K3015:Available TCHs")
                    item["tchn"] += col(rowmap, "TCH Congestion_N")
                    item["tchd"] += col(rowmap, "TCH Congestion_D")
                    item["sdn"] += col(rowmap, "K3001:Failed SDCCH Seizures due to Busy SDCCH")
                    item["sdd"] += col(rowmap, "K3000:SDCCH Seizure Requests")
                    item["dln"] += col(rowmap, "DL Quality(6~7)_N")
                    item["dld"] += col(rowmap, "DL Quality_D")
                    item["uln"] += col(rowmap, "UL Quality(6~7)_N")
                    item["uld"] += col(rowmap, "UL Quality_D")
                    cdrn = col(rowmap, "CDR_N")
                    item["cdrn"] += cdrn
                    item["cdrn_max"] = max(item["cdrn_max"], cdrn)
                    item["cdrd"] += col(rowmap, "CDR_D")
                    item["hosrn"] += col(rowmap, "HOSR_N")
                    item["hosrd"] += col(rowmap, "HOSR_D")
                    item["icmn"] += col(rowmap, "Interference Band_4-5_Samples")
                    item["icmd"] += col(rowmap, "Interference Band Samples_D")
                    item["a900"] += col(rowmap, "R3027A:Number of Available Channels (TCHF) (900/850/810 band)")
                    item["b1800"] += col(rowmap, "R3027B:Number of Available Channels (TCHF) (1800/1900 Band)")
                    item["h900"] += col(rowmap, "R3028A:Number of Available Channels (TCHH) (900/850/810 band)")
                    item["h1800"] += col(rowmap, "R3028B:Number of Available Channels (TCHH) (1800/1900 Band)")
    return cells


def kpi(item):
    days = item["days"]
    carried = item["k3014"] / days
    tchh = item["k3034"] / days
    avail = item["k3015"] / days
    half_rate = (tchh / carried) if carried > 0 else 0.0
    capacity = avail * (1.0 + half_rate)
    voice = (carried / capacity * 100.0) if capacity > 0 else None
    tch_cong = (item["tchn"] / item["tchd"] * 100.0) if item["tchd"] > 0 else 0.0
    sdcch_cong = (item["sdn"] / item["sdd"] * 100.0) if item["sdd"] > 0 else 0.0
    dl = (item["dln"] / item["dld"] * 100.0) if item["dld"] > 0 else None
    ul = (item["uln"] / item["uld"] * 100.0) if item["uld"] > 0 else None
    cdr_mean = item["cdrn"] / days
    cdr = (item["cdrn"] / item["cdrd"] * 100.0) if item["cdrd"] > 0 else None
    hosr = (item["hosrn"] / item["hosrd"] * 100.0) if item["hosrd"] > 0 else None
    icm = (item["icmn"] / item["icmd"] * 100.0) if item["icmd"] > 0 else None
    trx_counts = [int(round(value)) for value in item["trx"]]
    trx = Counter(trx_counts).most_common(1)[0][0]
    g900 = item["a900"] + item["h900"]
    g1800 = item["b1800"] + item["h1800"]
    if g900 > 0 and g1800 == 0:
        band = "G900"
    elif g1800 > 0 and g900 == 0:
        band = "G1800"
    elif g900 > 0 and g1800 > 0:
        band = "Both"
    else:
        band = "NoTCH"

    load_le60 = voice is not None and voice <= 60
    cdrn_ge10 = cdr_mean >= 10
    both_bad = dl is not None and ul is not None and dl > 2 and ul > 2
    cong_ok = tch_cong < 1 and sdcch_cong < 1
    hosr_le90 = hosr is not None and hosr <= 90
    icm_pair = cdr is not None and icm is not None and cdr > 0.2 and icm > 50
    reduce = band == "G900" and trx > 2 and cong_ok and (load_le60 or cdrn_ge10 or both_bad)
    add = band == "G900" and (hosr_le90 or both_bad or cdrn_ge10 or icm_pair)
    load_after = (voice * trx / 2.0) if (reduce and voice is not None and trx > 0) else None

    reasons_r = []
    if reduce and load_le60:
        reasons_r.append("Voice load<=60%")
    if reduce and cdrn_ge10:
        reasons_r.append("CDR_N>=10")
    if reduce and both_bad:
        reasons_r.append("DL and UL bad quality>2%")
    reasons_a = []
    if add and voice is not None and voice >= 70:
        reasons_a.append("Voice load>=70%")
    if add and voice is not None and voice < 70:
        reasons_a.append("Voice load<70%")
    if add and hosr_le90:
        reasons_a.append("HOSR_Out<=90%")
    if add and both_bad:
        reasons_a.append("DL and UL bad quality>2%")
    if add and cdrn_ge10:
        reasons_a.append("CDR_N>=10")
    if add and icm_pair:
        reasons_a.append("CDR>0.2% and ICM_B4~5>50%")

    notes = []
    if band == "G1800":
        notes.append("G1800 cell already on air")
    elif band == "Both":
        notes.append("Cell already has 900 and 1800 TCH")
    elif band == "NoTCH":
        notes.append("No 900 or 1800 TCH samples in the busy hour")
    if reduce and load_after is not None and load_after > 60:
        notes.append("Projected BH voice load exceeds 60% after cutting to 2 TRX")
    if days < 7:
        notes.append("BH sample is shorter than 7 days")

    if reduce and add:
        action = "Reduce G900 TRX and add G1800 cell"
    elif reduce:
        action = "Reduce G900 TRX"
    elif add:
        action = "Add G1800 cell"
    else:
        action = "No action"

    item.update({
        "carried": carried,
        "tchh": tchh,
        "avail": avail,
        "half_rate": half_rate,
        "capacity": capacity,
        "voice": voice,
        "tch_cong": tch_cong,
        "sdcch_cong": sdcch_cong,
        "dl": dl,
        "ul": ul,
        "cdr_mean": cdr_mean,
        "cdr": cdr,
        "hosr": hosr,
        "icm": icm,
        "trx": trx,
        "band": band,
        "reduce": reduce,
        "add": add,
        "load_after": load_after,
        "reason_r": ", ".join(reasons_r),
        "reason_a": ", ".join(reasons_a),
        "notes": "; ".join(notes),
        "action": action,
        "load_le60": load_le60,
        "cdrn_ge10": cdrn_ge10,
        "both_bad": both_bad,
        "hosr_le90": hosr_le90,
        "icm_pair": icm_pair,
    })
    return item


def blank(value):
    return "" if value is None else value


def build(records):
    records = sorted(records, key=lambda item: (item["gbsc"], item["name"], str(item["ci"])))
    last = len(records) + 1

    workbook = xlsxwriter.Workbook(OUT, {"strings_to_urls": False, "nan_inf_to_errors": True})
    workbook.set_properties({
        "title": "GSM Capacity Enhancement",
        "subject": "G900 TRX reduction and G1800 cell addition from busy-hour KPIs",
        "category": "Radio capacity",
    })
    workbook.set_calc_mode("auto")

    for name, _value, _unit, _note in PARAMS:
        # Names are defined after the Parameters sheet cells are known.
        pass

    navy = "#1F4E79"
    orange = "#C65911"
    blue = "#2E75B6"
    teal = "#0F6C6C"
    pale_orange = "#FCE4D6"
    pale_blue = "#DDEBF7"
    pale_gray = "#F2F2F2"
    pale_yellow = "#FFF2CC"
    pale_green = "#E2EFDA"
    pale_red = "#F8CBAD"
    white = "#FFFFFF"

    title_fmt = workbook.add_format({"bold": True, "font_name": "Calibri", "font_size": 18, "font_color": navy})
    section_fmt = workbook.add_format({"bold": True, "font_name": "Calibri", "font_size": 13, "font_color": navy})
    text_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11, "valign": "top", "text_wrap": True})
    label_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11})
    bold_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11, "bold": True})
    int_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11, "num_format": "#,##0", "align": "center"})
    big_int = workbook.add_format({"font_name": "Calibri", "font_size": 16, "bold": True, "num_format": "#,##0", "align": "center", "font_color": navy})
    header = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "bold": True, "font_color": white,
        "bg_color": navy, "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1,
    })
    header_orange = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "bold": True, "font_color": white,
        "bg_color": orange, "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1,
    })
    header_blue = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "bold": True, "font_color": white,
        "bg_color": blue, "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1,
    })
    header_gray = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "bold": True, "font_color": white,
        "bg_color": "#5B5B5B", "text_wrap": True, "valign": "vcenter", "align": "center", "border": 1,
    })
    cell_text = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "valign": "vcenter"})
    cell_center = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "valign": "vcenter"})
    cell_int = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0", "align": "center", "valign": "vcenter"})
    cell_num = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0.00", "align": "center", "valign": "vcenter"})
    cell_ratio = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "0.000", "align": "center", "valign": "vcenter"})
    input_num = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0.00", "align": "center", "bg_color": pale_gray, "valign": "vcenter"})
    input_int = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0", "align": "center", "bg_color": pale_gray, "valign": "vcenter"})
    input_text = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "bg_color": pale_gray, "valign": "vcenter"})
    input_center = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "bg_color": pale_gray, "valign": "vcenter"})
    param_value = workbook.add_format({
        "font_name": "Calibri", "font_size": 12, "bold": True, "border": 1,
        "bg_color": pale_yellow, "align": "center", "num_format": "0.0",
    })
    param_int = workbook.add_format({
        "font_name": "Calibri", "font_size": 12, "bold": True, "border": 1,
        "bg_color": pale_yellow, "align": "center", "num_format": "0",
    })
    wrap = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "text_wrap": True, "valign": "vcenter"})
    total_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "border": 1, "bg_color": pale_green, "num_format": "#,##0", "align": "center"})
    total_text = workbook.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "border": 1, "bg_color": pale_green})
    note_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 10, "italic": True, "font_color": "#595959", "text_wrap": True, "valign": "top"})
    kpi_label = workbook.add_format({"font_name": "Calibri", "font_size": 11, "border": 1, "valign": "vcenter"})
    kpi_orange = workbook.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "border": 1, "bg_color": pale_orange, "num_format": "#,##0", "align": "center"})
    kpi_blue = workbook.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "border": 1, "bg_color": pale_blue, "num_format": "#,##0", "align": "center"})
    kpi_label_o = workbook.add_format({"font_name": "Calibri", "font_size": 11, "border": 1, "bg_color": pale_orange, "valign": "vcenter"})
    kpi_label_b = workbook.add_format({"font_name": "Calibri", "font_size": 11, "border": 1, "bg_color": pale_blue, "valign": "vcenter"})

    # Column map for All Cells. Formulas below depend on this order.
    columns = [
        ("Source", "input_text", 10),
        ("GBSC", "input_text", 12),
        ("Cell CI", "input_int", 12),
        ("Cell Name", "input_text", 16),
        ("Cell Index", "input_int", 12),
        ("Band", "input_center", 12),
        ("BH Days", "input_int", 11),
        ("TRX Configured", "input_int", 12),
        ("Carried Traffic, Erl", "input_num", 14),
        ("TCHH Traffic, Erl", "input_num", 14),
        ("Avail TCH Mean", "input_num", 14),
        ("Half Rate Traffic Ratio", "formula_ratio", 14),
        ("Capacity, Erl", "formula_num", 13),
        ("Voice Load, %", "formula_num", 13),
        ("TCH Cong, %", "formula_num", 12),
        ("SDCCH Cong, %", "formula_num", 13),
        ("DL Bad Qual(6~7), %", "formula_num", 14),
        ("UL Bad Qual(6~7), %", "formula_num", 14),
        ("CDR_N (BH mean)", "input_num", 14),
        ("CDR_N (worst BH)", "input_num", 14),
        ("CDR, %", "formula_num", 12),
        ("HOSR_Out, %", "formula_num", 13),
        ("ICM_B4~5, %", "formula_num", 13),
        ("G900 TRX Reduction", "formula_center", 14),
        ("TRX Keep", "formula_int", 11),
        ("TRX Remove", "formula_int", 12),
        ("Voice Load after 2 TRX, %", "formula_num", 16),
        ("Reduction Reason", "formula_wrap", 36),
        ("G1800 Cell Addition", "formula_center", 14),
        ("Addition Reason", "formula_wrap", 42),
        ("Action", "formula_wrap", 32),
        ("Review", "formula_wrap", 42),
        ("TCH Cong_N", "counter", 12),
        ("TCH Cong_D", "counter", 12),
        ("SDCCH Busy Fail_N", "counter", 14),
        ("SDCCH Request_D", "counter", 14),
        ("DL Bad_N", "counter", 12),
        ("DL Qual_D", "counter", 12),
        ("UL Bad_N", "counter", 12),
        ("UL Qual_D", "counter", 12),
        ("CDR_N Sum", "counter", 12),
        ("CDR_D Sum", "counter", 12),
        ("HOSR_N", "counter", 12),
        ("HOSR_D", "counter", 12),
        ("ICM 4-5 Samples", "counter", 14),
        ("ICM Samples_D", "counter", 14),
        ("TCHF 900", "counter", 12),
        ("TCHF 1800", "counter", 12),
        ("TCHH 900", "counter", 12),
        ("TCHH 1800", "counter", 12),
    ]
    letter = {name: col_letter(pos) for pos, (name, _kind, _width) in enumerate(columns, start=1)}

    def L(name):
        return letter[name]

    def ref(name, row):
        return f"{letter[name]}{row}"

    # ---------- Parameters ----------
    parameters = workbook.add_worksheet("Parameters")
    parameters.set_tab_color(teal)
    parameters.set_column("A:A", 24)
    parameters.set_column("B:B", 14)
    parameters.set_column("C:C", 12)
    parameters.set_column("D:D", 88)
    parameters.write("A1", "Thresholds used by the All Cells formulas", title_fmt)
    parameters.write("A2", "Yellow cells are the live inputs. Edit a yellow cell and Excel recalculates All Cells. The G900 TRX Reduction and G1800 Addition sheets are the result at these default values.", note_fmt)
    parameters.set_row(1, 32)
    for col, text in enumerate(["Name", "Value", "Unit", "Role"]):
        parameters.write(3, col, text, header)
    for row_offset, (name, value, unit, note) in enumerate(PARAMS):
        excel_row = 5 + row_offset
        parameters.write(excel_row - 1, 0, name, cell_text)
        value_fmt = param_int if float(value).is_integer() else param_value
        parameters.write_number(excel_row - 1, 1, value, value_fmt)
        parameters.write(excel_row - 1, 2, unit, cell_center)
        parameters.write(excel_row - 1, 3, note, cell_text)
        workbook.define_name(name, f"=Parameters!$B${excel_row}")
    parameters.write("A16", "These names are workbook names. The decision formulas on All Cells call them directly, so the threshold is not hard-typed into eleven thousand rows.", text_fmt)
    parameters.set_row(15, 32)

    # ---------- All Cells ----------
    all_cells = workbook.add_worksheet("All Cells")
    all_cells.set_tab_color("#7F7F7F")
    all_cells.freeze_panes(1, 4)
    all_cells.autofilter(0, 0, len(records), len(columns) - 1)
    all_cells.set_row(0, 36)
    kind_header = {
        "input_text": header_gray,
        "input_int": header_gray,
        "input_center": header_gray,
        "input_num": header_gray,
        "counter": header_gray,
        "formula_ratio": header,
        "formula_num": header,
        "formula_center": header,
        "formula_int": header,
        "formula_wrap": header,
    }
    # Decision columns get a stronger header.
    decision_names = {
        "G900 TRX Reduction", "TRX Keep", "TRX Remove", "Voice Load after 2 TRX, %",
        "Reduction Reason", "G1800 Cell Addition", "Addition Reason", "Action", "Review",
    }
    for pos, (name, kind, width) in enumerate(columns):
        fmt = header_orange if name in {"G900 TRX Reduction", "TRX Keep", "TRX Remove", "Voice Load after 2 TRX, %", "Reduction Reason"} else kind_header[kind]
        if name in {"G1800 Cell Addition", "Addition Reason"}:
            fmt = header_blue
        if name in {"Action", "Review"}:
            fmt = header
        all_cells.set_column(pos, pos, width)
        all_cells.write(0, pos, name, fmt)

    value_formats = {
        "input_text": input_text,
        "input_int": input_int,
        "input_center": input_center,
        "input_num": input_num,
        "counter": input_num,
        "formula_ratio": cell_ratio,
        "formula_num": cell_num,
        "formula_center": cell_center,
        "formula_int": cell_int,
        "formula_wrap": wrap,
    }

    for offset, item in enumerate(records):
        row = offset + 2
        r = row

        def F(name):
            return ref(name, r)

        formulas = {
            "Half Rate Traffic Ratio": f'=IF({F("Carried Traffic, Erl")}=0,0,{F("TCHH Traffic, Erl")}/{F("Carried Traffic, Erl")})',
            "Capacity, Erl": f'={F("Avail TCH Mean")}*(1+{F("Half Rate Traffic Ratio")})',
            "Voice Load, %": f'=IF({F("Capacity, Erl")}=0,"",{F("Carried Traffic, Erl")}/{F("Capacity, Erl")}*100)',
            "TCH Cong, %": f'=IF({F("TCH Cong_D")}=0,0,{F("TCH Cong_N")}/{F("TCH Cong_D")}*100)',
            "SDCCH Cong, %": f'=IF({F("SDCCH Request_D")}=0,0,{F("SDCCH Busy Fail_N")}/{F("SDCCH Request_D")}*100)',
            "DL Bad Qual(6~7), %": f'=IF({F("DL Qual_D")}=0,"",{F("DL Bad_N")}/{F("DL Qual_D")}*100)',
            "UL Bad Qual(6~7), %": f'=IF({F("UL Qual_D")}=0,"",{F("UL Bad_N")}/{F("UL Qual_D")}*100)',
            "CDR, %": f'=IF({F("CDR_D Sum")}=0,"",{F("CDR_N Sum")}/{F("CDR_D Sum")}*100)',
            "HOSR_Out, %": f'=IF({F("HOSR_D")}=0,"",{F("HOSR_N")}/{F("HOSR_D")}*100)',
            "ICM_B4~5, %": f'=IF({F("ICM Samples_D")}=0,"",{F("ICM 4-5 Samples")}/{F("ICM Samples_D")}*100)',
            "G900 TRX Reduction": (
                f'=IF(AND({F("Band")}="G900",{F("TRX Configured")}>pTRXAbove,'
                f'{F("TCH Cong, %")}<pCongMax,{F("SDCCH Cong, %")}<pCongMax,'
                f'OR(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}<=pVoiceLoadMax),'
                f'{F("CDR_N (BH mean)")}>=pCDRnMin,'
                f'AND(ISNUMBER({F("DL Bad Qual(6~7), %")}),{F("DL Bad Qual(6~7), %")}>pBadQualMin,'
                f'ISNUMBER({F("UL Bad Qual(6~7), %")}),{F("UL Bad Qual(6~7), %")}>pBadQualMin))),"Yes","No")'
            ),
            "TRX Keep": f'=IF({F("G900 TRX Reduction")}="Yes",MIN(pTRXKeep,{F("TRX Configured")}),{F("TRX Configured")})',
            "TRX Remove": f'=IF({F("G900 TRX Reduction")}="Yes",MAX(0,{F("TRX Configured")}-pTRXKeep),0)',
            "Voice Load after 2 TRX, %": (
                f'=IF(AND({F("G900 TRX Reduction")}="Yes",pTRXKeep>0,ISNUMBER({F("Voice Load, %")})),'
                f'{F("Voice Load, %")}*{F("TRX Configured")}/pTRXKeep,"")'
            ),
            "Reduction Reason": (
                f'=IF({F("G900 TRX Reduction")}<>"Yes","",MID('
                f'IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}<=pVoiceLoadMax),", Voice load<="&pVoiceLoadMax&"%","")'
                f'&IF({F("CDR_N (BH mean)")}>=pCDRnMin,", CDR_N>="&pCDRnMin,"")'
                f'&IF(AND(ISNUMBER({F("DL Bad Qual(6~7), %")}),{F("DL Bad Qual(6~7), %")}>pBadQualMin,'
                f'ISNUMBER({F("UL Bad Qual(6~7), %")}),{F("UL Bad Qual(6~7), %")}>pBadQualMin),", DL and UL bad quality>"&pBadQualMin&"%","")'
                f',3,300))'
            ),
            "G1800 Cell Addition": (
                f'=IF(AND({F("Band")}="G900",OR('
                f'AND(ISNUMBER({F("HOSR_Out, %")}),{F("HOSR_Out, %")}<=pHOSRMax),'
                f'AND(ISNUMBER({F("DL Bad Qual(6~7), %")}),{F("DL Bad Qual(6~7), %")}>pBadQualMin,'
                f'ISNUMBER({F("UL Bad Qual(6~7), %")}),{F("UL Bad Qual(6~7), %")}>pBadQualMin),'
                f'{F("CDR_N (BH mean)")}>=pCDRnMin,'
                f'AND(ISNUMBER({F("CDR, %")}),{F("CDR, %")}>pCDRMin,ISNUMBER({F("ICM_B4~5, %")}),{F("ICM_B4~5, %")}>pICMMin)'
                f')),"Yes","No")'
            ),
            "Addition Reason": (
                f'=IF({F("G1800 Cell Addition")}<>"Yes","",MID('
                f'IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}>=pVoiceLoadHigh),", Voice load>="&pVoiceLoadHigh&"%","")'
                f'&IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}<pVoiceLoadHigh),", Voice load<"&pVoiceLoadHigh&"%","")'
                f'&IF(AND(ISNUMBER({F("HOSR_Out, %")}),{F("HOSR_Out, %")}<=pHOSRMax),", HOSR_Out<="&pHOSRMax&"%","")'
                f'&IF(AND(ISNUMBER({F("DL Bad Qual(6~7), %")}),{F("DL Bad Qual(6~7), %")}>pBadQualMin,'
                f'ISNUMBER({F("UL Bad Qual(6~7), %")}),{F("UL Bad Qual(6~7), %")}>pBadQualMin),", DL and UL bad quality>"&pBadQualMin&"%","")'
                f'&IF({F("CDR_N (BH mean)")}>=pCDRnMin,", CDR_N>="&pCDRnMin,"")'
                f'&IF(AND(ISNUMBER({F("CDR, %")}),{F("CDR, %")}>pCDRMin,ISNUMBER({F("ICM_B4~5, %")}),{F("ICM_B4~5, %")}>pICMMin),'
                f'", CDR>"&pCDRMin&"% and ICM_B4~5>"&pICMMin&"%","")'
                f',3,400))'
            ),
            "Action": (
                f'=IF(AND({F("G900 TRX Reduction")}="Yes",{F("G1800 Cell Addition")}="Yes"),'
                f'"Reduce G900 TRX and add G1800 cell",'
                f'IF({F("G900 TRX Reduction")}="Yes","Reduce G900 TRX",'
                f'IF({F("G1800 Cell Addition")}="Yes","Add G1800 cell","No action")))'
            ),
            "Review": (
                f'=MID('
                f'IF({F("Band")}="G1800","; G1800 cell already on air","")'
                f'&IF({F("Band")}="Both","; Cell already has 900 and 1800 TCH","")'
                f'&IF({F("Band")}="NoTCH","; No 900 or 1800 TCH samples in the busy hour","")'
                f'&IF(AND({F("G900 TRX Reduction")}="Yes",ISNUMBER({F("Voice Load after 2 TRX, %")}),{F("Voice Load after 2 TRX, %")}>pVoiceLoadMax),'
                f'"; Projected BH voice load exceeds "&pVoiceLoadMax&"% after cutting to "&pTRXKeep&" TRX","")'
                f'&IF({F("BH Days")}<7,"; BH sample is shorter than 7 days","")'
                f',3,400)'
            ),
        }

        values = {
            "Source": item["source"],
            "GBSC": item["gbsc"],
            "Cell CI": item["ci"],
            "Cell Name": item["name"],
            "Cell Index": item["cell_index"],
            "Band": item["band"],
            "BH Days": item["days"],
            "TRX Configured": item["trx"],
            "Carried Traffic, Erl": item["carried"],
            "TCHH Traffic, Erl": item["tchh"],
            "Avail TCH Mean": item["avail"],
            "CDR_N (BH mean)": item["cdr_mean"],
            "CDR_N (worst BH)": item["cdrn_max"],
            "TCH Cong_N": item["tchn"],
            "TCH Cong_D": item["tchd"],
            "SDCCH Busy Fail_N": item["sdn"],
            "SDCCH Request_D": item["sdd"],
            "DL Bad_N": item["dln"],
            "DL Qual_D": item["dld"],
            "UL Bad_N": item["uln"],
            "UL Qual_D": item["uld"],
            "CDR_N Sum": item["cdrn"],
            "CDR_D Sum": item["cdrd"],
            "HOSR_N": item["hosrn"],
            "HOSR_D": item["hosrd"],
            "ICM 4-5 Samples": item["icmn"],
            "ICM Samples_D": item["icmd"],
            "TCHF 900": item["a900"],
            "TCHF 1800": item["b1800"],
            "TCHH 900": item["h900"],
            "TCHH 1800": item["h1800"],
        }

        for pos, (name, kind, _width) in enumerate(columns):
            fmt = value_formats[kind]
            if name in formulas:
                all_cells.write_formula(row - 1, pos, formulas[name], fmt)
            else:
                value = values[name]
                if isinstance(value, str):
                    all_cells.write_string(row - 1, pos, value, fmt)
                else:
                    all_cells.write_number(row - 1, pos, value, fmt)

    all_cells.conditional_format(1, 0, len(records), len(columns) - 1, {
        "type": "formula",
        "criteria": f'=${L("G900 TRX Reduction")}2="Yes"',
        "format": workbook.add_format({"bg_color": "#FFF4EC"}),
    })
    all_cells.conditional_format(1, 0, len(records), len(columns) - 1, {
        "type": "formula",
        "criteria": f'=AND(${L("G1800 Cell Addition")}2="Yes",${L("G900 TRX Reduction")}2<>"Yes")',
        "format": workbook.add_format({"bg_color": "#F3F8FD"}),
    })

    # ---------- Action snapshots ----------
    reductions = [item for item in records if item["reduce"]]
    additions = [item for item in records if item["add"]]
    reductions.sort(key=lambda item: (item["gbsc"], -item["trx"], item["name"]))
    additions.sort(key=lambda item: (item["gbsc"], item["hosr"] if item["hosr"] is not None else 999, item["name"]))

    def write_table(worksheet, headers, rows, header_fmt):
        worksheet.freeze_panes(1, 4)
        worksheet.autofilter(0, 0, max(1, len(rows)), len(headers) - 1)
        worksheet.set_row(0, 32)
        for pos, (title, width, _fmt) in enumerate(headers):
            worksheet.set_column(pos, pos, width)
            worksheet.write(0, pos, title, header_fmt)
        for r_index, row_values in enumerate(rows, start=1):
            for c_index, (_title, _width, fmt) in enumerate(headers):
                value = row_values[c_index]
                if value == "" or value is None:
                    worksheet.write_blank(r_index, c_index, None, fmt)
                elif isinstance(value, str):
                    worksheet.write_string(r_index, c_index, value, fmt)
                else:
                    worksheet.write_number(r_index, c_index, value, fmt)

    reduction_headers = [
        ("Source", 10, cell_center),
        ("GBSC", 12, cell_text),
        ("Cell CI", 12, cell_int),
        ("Cell Name", 16, cell_text),
        ("BH Days", 11, cell_int),
        ("TRX now", 11, cell_int),
        ("TRX keep", 11, cell_int),
        ("TRX remove", 12, cell_int),
        ("Carried Traffic, Erl", 16, cell_num),
        ("Half Rate Traffic Ratio", 16, cell_ratio),
        ("Avail TCH Mean", 14, cell_num),
        ("Capacity, Erl", 14, cell_num),
        ("Voice Load, %", 14, cell_num),
        ("Voice Load after 2 TRX, %", 18, cell_num),
        ("TCH Cong, %", 13, cell_num),
        ("SDCCH Cong, %", 14, cell_num),
        ("DL Bad Qual(6~7), %", 16, cell_num),
        ("UL Bad Qual(6~7), %", 16, cell_num),
        ("CDR_N (BH mean)", 14, cell_num),
        ("CDR, %", 12, cell_num),
        ("HOSR_Out, %", 13, cell_num),
        ("Projected load check", 22, cell_center),
        ("Reduction Reason", 40, wrap),
        ("Review", 46, wrap),
    ]
    reduction_rows = []
    for item in reductions:
        reduction_rows.append([
            item["source"], item["gbsc"], item["ci"], item["name"], item["days"],
            item["trx"], 2, item["trx"] - 2, item["carried"], item["half_rate"],
            item["avail"], item["capacity"], blank(item["voice"]), blank(item["load_after"]),
            item["tch_cong"], item["sdcch_cong"], blank(item["dl"]), blank(item["ul"]),
            item["cdr_mean"], blank(item["cdr"]), blank(item["hosr"]),
            (
                "Above 100%" if item["load_after"] is not None and item["load_after"] > 100
                else "Above 60%" if item["load_after"] is not None and item["load_after"] > 60
                else "Within 60%"
            ),
            item["reason_r"], item["notes"],
        ])
    reduction_sheet = workbook.add_worksheet("G900 TRX Reduction")
    reduction_sheet.set_tab_color(orange)
    write_table(reduction_sheet, reduction_headers, reduction_rows, header_orange)
    reduction_sheet.conditional_format(1, 13, len(reduction_rows), 13, {
        "type": "cell", "criteria": ">", "value": 60,
        "format": workbook.add_format({"bg_color": pale_red, "font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0.00", "align": "center"}),
    })

    addition_headers = [
        ("Source", 10, cell_center),
        ("GBSC", 12, cell_text),
        ("Cell CI", 12, cell_int),
        ("Cell Name", 16, cell_text),
        ("BH Days", 11, cell_int),
        ("TRX now", 11, cell_int),
        ("Carried Traffic, Erl", 16, cell_num),
        ("Voice Load, %", 14, cell_num),
        ("HOSR_Out, %", 13, cell_num),
        ("DL Bad Qual(6~7), %", 16, cell_num),
        ("UL Bad Qual(6~7), %", 16, cell_num),
        ("CDR_N (BH mean)", 14, cell_num),
        ("CDR, %", 12, cell_num),
        ("ICM_B4~5, %", 13, cell_num),
        ("TCH Cong, %", 13, cell_num),
        ("SDCCH Cong, %", 14, cell_num),
        ("Also reduce G900 TRX", 18, cell_center),
        ("Addition Reason", 55, wrap),
    ]
    addition_rows = []
    for item in additions:
        addition_rows.append([
            item["source"], item["gbsc"], item["ci"], item["name"], item["days"], item["trx"],
            item["carried"], blank(item["voice"]), blank(item["hosr"]), blank(item["dl"]),
            blank(item["ul"]), item["cdr_mean"], blank(item["cdr"]), blank(item["icm"]),
            item["tch_cong"], item["sdcch_cong"], "Yes" if item["reduce"] else "No",
            item["reason_a"],
        ])
    addition_sheet = workbook.add_worksheet("G1800 Addition")
    addition_sheet.set_tab_color(blue)
    write_table(addition_sheet, addition_headers, addition_rows, header_blue)

    # ---------- Summary ----------
    by_bsc = defaultdict(lambda: Counter())
    for item in records:
        bucket = by_bsc[item["gbsc"]]
        bucket["cells"] += 1
        bucket[item["band"]] += 1
        if item["band"] == "G900":
            bucket["g900_trx"] += item["trx"]
        if item["reduce"]:
            bucket["reduce_cells"] += 1
            bucket["trx_remove"] += item["trx"] - 2
            if item["load_after"] is not None and item["load_after"] > 100:
                bucket["over_100"] += 1
            elif item["load_after"] is not None and item["load_after"] > 60:
                bucket["over_60"] += 1
            else:
                bucket["within_60"] += 1
        if item["add"]:
            bucket["add_cells"] += 1
        if item["reduce"] and item["add"]:
            bucket["both"] += 1
        if item["days"] < 7:
            bucket["short"] += 1

    def total(key):
        return sum(bucket[key] for bucket in by_bsc.values())

    reason_counts = Counter()
    for item in additions:
        if item["hosr_le90"]:
            reason_counts["hosr"] += 1
        if item["both_bad"]:
            reason_counts["qual"] += 1
        if item["cdrn_ge10"]:
            reason_counts["cdrn"] += 1
        if item["icm_pair"]:
            reason_counts["icm"] += 1
        if item["voice"] is not None and item["voice"] >= 70:
            reason_counts["high_load"] += 1
        elif item["voice"] is not None:
            reason_counts["low_load"] += 1

    summary = workbook.add_worksheet("Summary")
    summary.set_tab_color(navy)
    summary.set_column("A:A", 3)
    summary.set_column("B:B", 78)
    summary.set_column("C:C", 18)
    summary.set_column("D:D", 18)
    summary.set_column("E:E", 22)
    summary.set_column("F:L", 16)
    summary.hide_gridlines(2)
    summary.write("B1", "GSM capacity enhancement", title_fmt)
    summary.write("B2", "G900 TRX reduction and G1800 cell addition. Busy hour 26 Oct 2026 to 1 Nov 2026.", section_fmt)
    summary.write("B3", "Source files: 2G Capacity Exercise_V1_BH.xlsb and 2G Capacity Exercise_V2_BH.xlsb. Daily files were not used. The rule is a busy-hour rule.", note_fmt)
    summary.set_row(2, 28)

    summary.write("B5", "G900 TRX reduction", header_orange)
    summary.write("C5", "Cells / TRX", header_orange)
    reduction_lines = [
        ("G900 cells assessed", total("G900")),
        ("G900 cells with more than 2 TRX", sum(1 for item in records if item["band"] == "G900" and item["trx"] > 2)),
        ("Cells reduced to 2 TRX", total("reduce_cells")),
        ("G900 TRX removed", total("trx_remove")),
        ("Reduced cells whose projected load stays within 60%", total("within_60")),
        ("Reduced cells whose projected load is above 60% and at or below 100%", total("over_60")),
        ("Reduced cells whose projected load is above 100%", total("over_100")),
    ]
    for index, (label, value) in enumerate(reduction_lines):
        summary.write(5 + index, 1, label, kpi_label_o)
        summary.write_number(5 + index, 2, value, kpi_orange)

    summary.merge_range("E5:F5", "G1800 cell addition", header_blue)
    summary.write("G5", "Cells", header_blue)
    addition_lines = [
        ("G900 cells proposed for a new G1800 cell", total("add_cells")),
        ("Reason includes HOSR_Out <= 90%", reason_counts["hosr"]),
        ("Reason includes DL and UL bad quality > 2%", reason_counts["qual"]),
        ("Reason includes CDR_N >= 10", reason_counts["cdrn"]),
        ("Reason includes CDR > 0.2% and ICM band 4-5 > 50%", reason_counts["icm"]),
        ("Also on the G900 TRX reduction list", total("both")),
        ("High-load band, voice load >= 70%", reason_counts["high_load"]),
    ]
    for index, (label, value) in enumerate(addition_lines):
        summary.merge_range(5 + index, 4, 5 + index, 5, label, kpi_label_b)
        summary.write_number(5 + index, 6, value, kpi_blue)

    summary.write("B14", "Footprint", header)
    summary.write("C14", "Cells", header)
    footprint = [
        ("All cells in the two busy-hour files", len(records)),
        ("G900", total("G900")),
        ("G1800 already on air", total("G1800")),
        ("Both 900 and 1800 TCH", total("Both")),
        ("No 900 or 1800 TCH samples", total("NoTCH")),
        ("Cells with fewer than 7 busy-hour days", total("short")),
    ]
    for index, (label, value) in enumerate(footprint):
        summary.write(14 + index, 1, label, kpi_label)
        summary.write_number(14 + index, 2, value, int_fmt)

    bsc_header_row = 22
    bsc_headers = [
        "GBSC", "Cells", "G900", "G1800 already", "Both", "G900 TRX now",
        "Cells reduced", "TRX removed", "G900 TRX after", "G1800 cells to add", "Both actions",
    ]
    for col, text in enumerate(bsc_headers):
        summary.write(bsc_header_row, 1 + col, text, header)
    bsc_order = ["MYMH05", "MYMH06", "MYMH07", "SYLH04", "TNGH12"]
    for index, gbsc in enumerate(bsc_order):
        bucket = by_bsc[gbsc]
        row = bsc_header_row + 1 + index
        values = [
            gbsc,
            bucket["cells"],
            bucket["G900"],
            bucket["G1800"],
            bucket["Both"],
            bucket["g900_trx"],
            bucket["reduce_cells"],
            bucket["trx_remove"],
            bucket["g900_trx"] - bucket["trx_remove"],
            bucket["add_cells"],
            bucket["both"],
        ]
        for col, value in enumerate(values):
            fmt = total_text if col == 0 else total_fmt
            if col == 0:
                summary.write(row, 1 + col, value, cell_text)
            else:
                summary.write_number(row, 1 + col, value, cell_int)
    total_row = bsc_header_row + 1 + len(bsc_order)
    summary.write(total_row, 1, "Total", total_text)
    for col in range(1, len(bsc_headers)):
        start = col_letter(2 + col)  # column B is 2
        # Data columns start at B. Header col 0 is B, col 1 is C, so Excel column = 2 + col.
        excel_col = col_letter(2 + col)
        first = bsc_header_row + 2
        last_bsc = bsc_header_row + 1 + len(bsc_order)
        summary.write_formula(total_row, 1 + col, f"=SUM({excel_col}{first}:{excel_col}{last_bsc})", total_fmt)

    note_row = total_row + 2
    summary.write(note_row, 1, "How to use the sheets", section_fmt)
    notes = [
        "G900 TRX Reduction is the implementation list. Each row keeps 2 TRX and shows how many TRX come out. Rows with a red voice-load-after-cut figure go above 60% on the remaining 2 TRX, because available TCH is assumed to scale with TRX count.",
        "G1800 Addition is the list of G900 cells proposed for a new DCS1800 cell. A cell already carrying 1800 TCH is not on this list.",
        "All Cells is one row per cell, including cells with no action. KPI and decision columns are Excel formulas. They read the yellow cells on Parameters. Changing a threshold recalculates All Cells.",
        "The two action sheets are the result at the default thresholds. After a threshold change, filter All Cells on G900 TRX Reduction or G1800 Cell Addition.",
        "Rules records the busy-hour aggregation, the formulas, and the readings required where a line could be applied in more than one way.",
    ]
    for index, text in enumerate(notes):
        summary.merge_range(note_row + 1 + index, 1, note_row + 1 + index, 8, text, text_fmt)
        summary.set_row(note_row + 1 + index, 36)

    chart = workbook.add_chart({"type": "column"})
    chart.set_title({"name": "G900 TRX removed and G1800 cells to add"})
    chart.set_y_axis({"name": "Count"})
    chart.set_size({"width": 720, "height": 320})
    first_data = bsc_header_row + 2
    last_data = bsc_header_row + 1 + len(bsc_order)
    chart.add_series({
        "name": "TRX removed",
        "categories": ["Summary", first_data - 1, 1, last_data - 1, 1],
        "values": ["Summary", first_data - 1, 8, last_data - 1, 8],
        "fill": {"color": orange},
    })
    chart.add_series({
        "name": "G1800 cells to add",
        "categories": ["Summary", first_data - 1, 1, last_data - 1, 1],
        "values": ["Summary", first_data - 1, 10, last_data - 1, 10],
        "fill": {"color": blue},
    })
    chart.set_legend({"position": "bottom"})
    summary.insert_chart("B38", chart)

    # ---------- Rules ----------
    rules = workbook.add_worksheet("Rules")
    rules.set_tab_color("#548235")
    rules.set_column("A:A", 3)
    rules.set_column("B:B", 120)
    rules.write("B1", "Rules applied in this workbook", title_fmt)
    rule_lines = [
        "",
        "Purpose",
        "G900 TRX is reduced on cells that can give spectrum back while SDCCH and TCH congestion stay under 1%. Where the busy-hour KPIs show a handover, quality, drop, or interference problem, a G1800 cell is proposed on that G900 cell.",
        "",
        "Data",
        "Every decision uses the cell busy hour from 26 Oct 2026 through 1 Nov 2026. V1_BH covers MYMH05, MYMH06, MYMH07, and SYLH04. V2_BH covers TNGH12. Each file has one busy-hour row per cell per day. Integrity is 1 on every row. TRX configuration does not change across the week for any cell.",
        "Erlang values and CDR_N are the average of the daily busy hours. A ratio uses the sum of the numerators over the sum of the denominators across the available busy hours, so a busier day weighs more. Configured TRX is the configured value in the busy hour (S3655). 10,985 cells have all 7 days. 120 cells have a shorter sample, and the Review column says so.",
        "",
        "Band",
        "The capacity file has no band column. A cell is G900 when the 900-band TCH full-rate and half-rate counters (R3027A and R3028A) are above zero and the 1800-band counters (R3027B and R3028B) are zero across the week. The reverse is G1800. Both means both bands carried TCH. Actions are applied to G900 cells. A cell that already has 1800 TCH is left out of the G1800 addition list.",
        "",
        "Formulas",
        "Half Rate Traffic Ratio = K3034 TCHH Erlang / K3014 TCH Erlang. The ratio is 0 when carried TCH is 0.",
        "Capacity, Erl = Avail TCH Mean x (1 + Half Rate Traffic Ratio). Avail TCH Mean is K3015.",
        "Voice Load, % = Carried Traffic / Capacity x 100. Carried Traffic is K3014.",
        "TCH Cong, % = TCH Congestion_N / TCH Congestion_D x 100. A zero denominator is counted as 0%.",
        "SDCCH Cong, % = K3001 Failed SDCCH seizures due to busy SDCCH / K3000 SDCCH seizure requests x 100. A zero denominator is counted as 0%. This is the SDCCH congestion definition used for the <1% gate.",
        "DL Bad Qual(6~7), % = DL Quality(6~7)_N / DL Quality_D x 100. The file names the denominator DL Quality_D.",
        "UL Bad Qual(6~7), % = UL Quality(6~7)_N / UL Quality_D x 100.",
        "CDR_N is the average busy-hour drop count. CDR, % = CDR_N / CDR_D x 100, using the summed counters.",
        "HOSR_Out, % = HOSR_N / HOSR_D x 100.",
        "ICM_B4~5, % = Interference Band_4-5_Samples / Interference Band Samples_D x 100.",
        "Where a denominator is zero, voice load, bad quality, CDR, HOSR, and ICM are blank, and a greater-than or less-than test on that KPI is not met. Congestion is the exception: no attempts is 0% congestion.",
        "",
        "G900 TRX reduction",
        "The cell is G900, configured TRX is above 2, SDCCH congestion is below 1%, and TCH congestion is below 1%. In addition, any one of these is true: voice load <= 60%, CDR_N >= 10, or both DL and UL bad quality are above 2%.",
        "Action: keep 2 TRX. TRX removed = configured TRX - 2. Cells at 2 TRX or below are not changed.",
        "Voice load after the cut = current voice load x configured TRX / 2. This assumes available TCH scales with TRX count and the half-rate ratio stays as it is. It is a check beside the action. The action still keeps 2 TRX, including where the projected load goes above 60% or above 100%. Those rows are marked in Review and highlighted on the reduction sheet.",
        "",
        "G1800 cell addition",
        "The cell is G900. TRX count is not a filter: the test applies at every TRX count, which is the all-scenario condition.",
        "A new G1800 cell is proposed when any one of these is true: HOSR_Out <= 90%, both DL and UL bad quality are above 2%, CDR_N >= 10, or CDR is above 0.2% and ICM band 4-5 is above 50%.",
        "The high-load line (voice load >= 70%) and the low-load line (voice load < 70%) carry the same KPI tests, so both load bands are in scope. The load band is written on the reason. Voice load on its own does not add a cell.",
        "DL and UL bad quality above 2% means both directions. The same reading is used in the reduction OR list. SDCCH and TCH congestion below 1% likewise means both.",
        "CDR > 0.2 is applied as 0.2 percent.",
        "CDR_N >= 10 is the average busy-hour drop count. The worst single busy hour is on All Cells as CDR_N (worst BH) and is not part of the decision.",
        "",
        "What this file does not do",
        "It does not plan the L900 bandwidth change, and it does not use the cell database, TRX parameter, co-MPT, or inventory files. It does not choose the G1800 ARFCN, the number of G1800 TRX, or the BCCH. The addition list says which G900 cell gets a G1800 cell. It does not say how many TRX that new cell receives.",
    ]
    for index, line in enumerate(rule_lines):
        if line in {
            "Purpose", "Data", "Band", "Formulas", "G900 TRX reduction",
            "G1800 cell addition", "What this file does not do",
        }:
            rules.write(1 + index, 1, line, section_fmt)
        elif line == "":
            rules.write(1 + index, 1, "", text_fmt)
        else:
            rules.write(1 + index, 1, line, text_fmt)
            rules.set_row(1 + index, 18 + 16 * max(1, (len(line) // 145) + 1))

    order = ["Summary", "G900 TRX Reduction", "G1800 Addition", "All Cells", "Parameters", "Rules"]
    workbook.worksheets_objs.sort(key=lambda sheet: order.index(sheet.name))

    workbook.close()

    dates = set()
    for item in records:
        dates.update(item["dates"])
    if dates:
        print("BH dates", sorted(EXCEL_EPOCH + timedelta(days=day) for day in dates))
    print("cells", len(records))
    print("reduce", len(reductions), "trx_remove", sum(item["trx"] - 2 for item in reductions))
    print("add", len(additions), "both", sum(1 for item in records if item["reduce"] and item["add"]))
    print("over60", total("over_60"), "over100", total("over_100"), "within60", total("within_60"))
    print("reasons", dict(reason_counts))
    print("bands", Counter(item["band"] for item in records))
    for gbsc in bsc_order:
        print(gbsc, dict(by_bsc[gbsc]))


def main():
    cells = load_cells()
    records = [kpi(item) for item in cells.values()]
    build(records)


if __name__ == "__main__":
    main()
