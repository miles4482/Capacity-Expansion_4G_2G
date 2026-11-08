#!/usr/bin/env python3
"""GSM capacity enhancement from the four renamed capacity files.

Busy-hour files supply voice load and the capacity KPIs.
Daily files supply CDR, HOSR, quality, interference, and the other KPIs.
Both files have one row per cell per day. Each cell is averaged over the
dates present for that cell. Traffic Erlang is then multiplied by 24.
Voice load uses the busy-hour traffic after that multiplication.

  2G Capacity File_V1_BH.rar
  2G Capacity File_V2_BH.rar
  2G Capacity File_V1_Daily.rar
  2G Capacity File_V2_Daily.rar
"""

from collections import Counter
from pathlib import Path
import csv
import subprocess
import xlsxwriter

ROOT = Path(__file__).resolve().parent
CSV_DIR = ROOT / "source_csv"
SOURCES = [
    ("V1", "BH", "2G Capacity File_V1_BH.rar", "2G Capacity File_V1_BH.csv"),
    ("V2", "BH", "2G Capacity File_V2_BH.rar", "2G Capacity File_V2_BH.csv"),
    ("V1", "Daily", "2G Capacity File_V1_Daily.rar", "2G Capacity File_V1_Daily.csv"),
    ("V2", "Daily", "2G Capacity File_V2_Daily.rar", "2G Capacity File_V2_Daily.csv"),
]
OUT = "GSM_Capacity_Enhancement.xlsx"
TRAFFIC_HOURS = 24
BSC_ORDER = ["MYMH05", "MYMH06", "MYMH07", "SYLH04", "TNGH12"]
SUM_FIELDS = [
    "K3004:Traffic Volume on SDCCH(Erl)",
    "K3014:Traffic Volume on TCH(Erl)",
    "K3034:TCHH Traffic Volume(Erl)",
    "K3015:Available TCHs",
    "K3016:Configured TCHs",
    "TCH Congestion_N",
    "TCH Congestion_D",
    "K3001:Failed SDCCH Seizures due to Busy SDCCH",
    "K3000:SDCCH Seizure Requests",
    "R3027A:Number of Available Channels (TCHF) (900/850/810 band)",
    "R3027B:Number of Available Channels (TCHF) (1800/1900 Band)",
    "R3028A:Number of Available Channels (TCHH) (900/850/810 band)",
    "R3028B:Number of Available Channels (TCHH) (1800/1900 Band)",
    "CDR_N",
    "CDR_D",
    "HOSR_N",
    "HOSR_D",
    "Incoming HOSR_N",
    "Incoming HOSR_D",
    "CSSR_N",
    "CSSR_D",
    "DL Quality(6~7)_N",
    "DL Quality_D",
    "UL Quality(6~7)_N",
    "UL Quality_D",
    "Interference Band_4-5_Samples",
    "Interference Band Samples_D",
]

# Yellow cells on Parameters. All Cells formulas read these names.
PARAMS = [
    ("pVoiceLoadMax", 60, "%", "G900 reduction OR term: busy-hour voice load at or below this"),
    ("pCDRnMin", 10, "drops/day", "OR term: daily-average CDR_N at or above this"),
    ("pBadQualMin", 2, "%", "OR term: both DL and UL bad quality (6-7) above this"),
    ("pCongMax", 1, "%", "Reduction gate: SDCCH congestion and TCH congestion each below this"),
    ("pVoiceLoadHigh", 70, "%", "G1800 OR term: busy-hour voice load at or above this"),
    ("pHOSRMax", 90, "%", "G1800 OR term: HOSR_Out at or below this"),
    ("pCDRMin", 0.2, "%", "G1800 pair: CDR above this, together with ICM"),
    ("pICMMin", 50, "%", "G1800 pair: ICM band 4-5 above this, together with CDR"),
    ("pTRXAbove", 2, "TRX", "G900 reduction applies only when configured TRX is above this"),
    ("pTRXKeep", 2, "TRX", "TRX kept when a G900 cell is reduced"),
    ("pAddTRXMin", 0, "TRX", "G1800 addition applies at every TRX count. 0 means all scenarios"),
    ("pDailyDays", 7, "days", "Flag a cell when fewer than this many dates are available. Averages use the dates present"),
    ("pTrafficHours", 24, "hours", "Multiply busy-hour and daily traffic Erlang by this after the date average"),
]


def num(value):
    if value is None or value == "" or value == " ":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def clean_ci(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].lstrip("-").isdigit():
        return text[:-2]
    return text


def clean_index(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].lstrip("-").isdigit():
        return int(text[:-2])
    try:
        return int(text)
    except ValueError:
        return text


def col_letter(index):
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def ensure_csvs():
    CSV_DIR.mkdir(exist_ok=True)
    for _source, _kind, rar_name, csv_name in SOURCES:
        dest = CSV_DIR / csv_name
        if dest.exists() and dest.stat().st_size > 0:
            continue
        rar_path = ROOT / rar_name
        if not rar_path.exists():
            raise SystemExit(f"Missing source archive {rar_name}")
        subprocess.check_call(["unar", "-f", "-o", str(CSV_DIR), str(rar_path)])
        if not dest.exists():
            raise SystemExit(f"Extracting {rar_name} did not produce {csv_name}")


def load_side(kind):
    cells = {}
    for source, file_kind, _rar_name, csv_name in SOURCES:
        if file_kind != kind:
            continue
        with (CSV_DIR / csv_name).open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                key = (
                    source,
                    row["GBSC"].strip(),
                    clean_ci(row["Cell CI"]),
                    row["Cell Name"].strip(),
                )
                item = cells.get(key)
                if item is None:
                    item = {
                        "source": source,
                        "gbsc": key[1],
                        "ci": key[2],
                        "name": key[3],
                        "cell_index": clean_index(row.get("CellIndex")),
                        "days": 0,
                        "dates": set(),
                        "trx": [],
                    }
                    for field in SUM_FIELDS:
                        item[field] = 0.0
                    cells[key] = item
                item["days"] += 1
                item["dates"].add(row["Date"].strip())
                item["trx"].append(num(row.get("S3655:Number of configured TRXs in a cell")))
                for field in SUM_FIELDS:
                    item[field] += num(row.get(field))
    return cells


def load_bh():
    return load_side("BH")


def load_daily():
    return load_side("Daily")


def ratio(numer, denom):
    if denom > 0:
        return numer / denom * 100.0
    return None


def prepare(bh_cells, daily_cells):
    records = []
    for key, item in bh_cells.items():
        days = item["days"]
        daily = daily_cells.get(key)
        k3014_avg = item["K3014:Traffic Volume on TCH(Erl)"] / days
        k3034_avg = item["K3034:TCHH Traffic Volume(Erl)"] / days
        k3004_avg = item["K3004:Traffic Volume on SDCCH(Erl)"] / days
        carried = k3014_avg * TRAFFIC_HOURS
        tchh = k3034_avg * TRAFFIC_HOURS
        avail = item["K3015:Available TCHs"] / days
        half = (k3034_avg / k3014_avg) if k3014_avg > 0 else 0.0
        capacity = avail * (1.0 + half)
        voice = (carried / capacity * 100.0) if capacity > 0 else None
        tch_cong = (item["TCH Congestion_N"] / item["TCH Congestion_D"] * 100.0) if item["TCH Congestion_D"] > 0 else 0.0
        sdcch_cong = (
            item["K3001:Failed SDCCH Seizures due to Busy SDCCH"] / item["K3000:SDCCH Seizure Requests"] * 100.0
        ) if item["K3000:SDCCH Seizure Requests"] > 0 else 0.0
        trx_values = [int(round(value)) for value in item["trx"]]
        trx = Counter(trx_values).most_common(1)[0][0]
        g900 = item["R3027A:Number of Available Channels (TCHF) (900/850/810 band)"] + item["R3028A:Number of Available Channels (TCHH) (900/850/810 band)"]
        g1800 = item["R3027B:Number of Available Channels (TCHF) (1800/1900 Band)"] + item["R3028B:Number of Available Channels (TCHH) (1800/1900 Band)"]
        if g900 > 0 and g1800 == 0:
            band = "G900"
        elif g1800 > 0 and g900 == 0:
            band = "G1800"
        elif g900 > 0 and g1800 > 0:
            band = "Both"
        else:
            band = "NoTCH"

        if daily is None:
            cdr_n = dl = ul = cdr = hosr = icm = ihosr = cssr = None
            daily_tch = None
            daily_days = 0
            shown = {}
        else:
            daily_days = daily["days"]

            def day_avg(field):
                return daily[field] / daily_days

            cdr_n = day_avg("CDR_N")
            dl = ratio(daily["DL Quality(6~7)_N"], daily["DL Quality_D"])
            ul = ratio(daily["UL Quality(6~7)_N"], daily["UL Quality_D"])
            cdr = ratio(daily["CDR_N"], daily["CDR_D"])
            hosr = ratio(daily["HOSR_N"], daily["HOSR_D"])
            icm = ratio(daily["Interference Band_4-5_Samples"], daily["Interference Band Samples_D"])
            ihosr = ratio(daily["Incoming HOSR_N"], daily["Incoming HOSR_D"])
            cssr = ratio(daily["CSSR_N"], daily["CSSR_D"])
            shown = {
                "k3014": day_avg("K3014:Traffic Volume on TCH(Erl)"),
                "k3034": day_avg("K3034:TCHH Traffic Volume(Erl)"),
                "k3004": day_avg("K3004:Traffic Volume on SDCCH(Erl)"),
                "cdrn": day_avg("CDR_N"),
                "cdrd": day_avg("CDR_D"),
                "hosrn": day_avg("HOSR_N"),
                "hosrd": day_avg("HOSR_D"),
                "ihosrn": day_avg("Incoming HOSR_N"),
                "ihosrd": day_avg("Incoming HOSR_D"),
                "cssrn": day_avg("CSSR_N"),
                "cssrd": day_avg("CSSR_D"),
                "dln": day_avg("DL Quality(6~7)_N"),
                "dld": day_avg("DL Quality_D"),
                "uln": day_avg("UL Quality(6~7)_N"),
                "uld": day_avg("UL Quality_D"),
                "icmn": day_avg("Interference Band_4-5_Samples"),
                "icmd": day_avg("Interference Band Samples_D"),
            }
            daily_tch = shown["k3014"] * TRAFFIC_HOURS

        load_le60 = voice is not None and voice <= 60
        load_ge70 = voice is not None and voice >= 70
        load_lt70 = voice is not None and voice < 70
        cdrn_ge10 = cdr_n is not None and cdr_n >= 10
        both_bad = dl is not None and ul is not None and dl > 2 and ul > 2
        cong_ok = tch_cong < 1 and sdcch_cong < 1
        hosr_le90 = hosr is not None and hosr <= 90
        icm_pair = cdr is not None and icm is not None and cdr > 0.2 and icm > 50
        reduce = band == "G900" and trx > 2 and cong_ok and (load_le60 or cdrn_ge10 or both_bad)
        add = band == "G900" and trx >= 0 and (
            load_ge70 or hosr_le90 or both_bad or cdrn_ge10 or icm_pair
        )
        load_after = (voice * trx / 2.0) if reduce and voice is not None and trx > 0 else None

        reasons_r = []
        if reduce and load_le60:
            reasons_r.append("Voice load<=60%")
        if reduce and cdrn_ge10:
            reasons_r.append("CDR_N>=10")
        if reduce and both_bad:
            reasons_r.append("DL and UL bad quality>2%")
        reasons_a = []
        if add and load_ge70:
            reasons_a.append("Voice load>=70%")
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
        if daily is None:
            notes.append("No matching daily row")
        if reduce and load_after is not None and load_after > 100:
            notes.append("Projected BH voice load exceeds 100% after cutting to 2 TRX")
        elif reduce and load_after is not None and load_after > 60:
            notes.append("Projected BH voice load exceeds 60% after cutting to 2 TRX")
        if days < 7:
            notes.append("BH sample is shorter than 7 days")
        if daily is not None and daily_days < 7:
            notes.append("Daily sample is shorter than 7 days")

        if reduce and add:
            action = "Reduce G900 TRX and add G1800 cell"
        elif reduce:
            action = "Reduce G900 TRX"
        elif add:
            action = "Add G1800 cell"
        else:
            action = "No action"

        if load_ge70:
            load_band = "High >=70%"
        elif load_lt70:
            load_band = "Low <70%"
        else:
            load_band = ""

        item.update({
            "k3014_avg": k3014_avg,
            "k3034_avg": k3034_avg,
            "k3004_avg": k3004_avg,
            "carried": carried,
            "tchh": tchh,
            "sdcch_erl": k3004_avg * TRAFFIC_HOURS,
            "avail": avail,
            "configured_tch": item["K3016:Configured TCHs"] / days,
            "half_rate": half,
            "capacity": capacity,
            "voice": voice,
            "tch_cong": tch_cong,
            "sdcch_cong": sdcch_cong,
            "tchn_avg": item["TCH Congestion_N"] / days,
            "tchd_avg": item["TCH Congestion_D"] / days,
            "sdn_avg": item["K3001:Failed SDCCH Seizures due to Busy SDCCH"] / days,
            "sdd_avg": item["K3000:SDCCH Seizure Requests"] / days,
            "a900_avg": item["R3027A:Number of Available Channels (TCHF) (900/850/810 band)"] / days,
            "b1800_avg": item["R3027B:Number of Available Channels (TCHF) (1800/1900 Band)"] / days,
            "h900_avg": item["R3028A:Number of Available Channels (TCHH) (900/850/810 band)"] / days,
            "h1800_avg": item["R3028B:Number of Available Channels (TCHH) (1800/1900 Band)"] / days,
            "daily_days": daily_days,
            "trx": trx,
            "band": band,
            "daily": shown,
            "cdr_n": cdr_n,
            "dl": dl,
            "ul": ul,
            "cdr": cdr,
            "hosr": hosr,
            "icm": icm,
            "ihosr": ihosr,
            "cssr": cssr,
            "daily_tch": daily_tch,
            "reduce": reduce,
            "add": add,
            "load_after": load_after,
            "reason_r": ", ".join(reasons_r),
            "reason_a": ", ".join(reasons_a),
            "notes": "; ".join(notes),
            "action": action,
            "load_band": load_band,
            "load_le60": load_le60,
            "load_ge70": load_ge70,
            "cdrn_ge10": cdrn_ge10,
            "both_bad": both_bad,
            "cong_ok": cong_ok,
            "hosr_le90": hosr_le90,
            "icm_pair": icm_pair,
        })
        records.append(item)
    records.sort(key=lambda row: (row["gbsc"], row["name"], row["ci"]))
    return records


def blank(value):
    return "" if value is None else value


def build(records):
    last = len(records) + 1
    workbook = xlsxwriter.Workbook(OUT, {"strings_to_urls": False, "nan_inf_to_errors": True})
    workbook.set_properties({
        "title": "GSM Capacity Enhancement",
        "subject": "G900 TRX reduction and G1800 cell addition",
        "category": "Radio capacity",
    })

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
    int_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11, "num_format": "#,##0", "align": "center"})
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
    input_num = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0.00",
        "align": "center", "bg_color": pale_gray, "valign": "vcenter",
    })
    input_int = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "num_format": "#,##0",
        "align": "center", "bg_color": pale_gray, "valign": "vcenter",
    })
    input_text = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "bg_color": pale_gray, "valign": "vcenter"})
    input_center = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "align": "center",
        "bg_color": pale_gray, "valign": "vcenter",
    })
    param_value = workbook.add_format({
        "font_name": "Calibri", "font_size": 12, "bold": True, "border": 1,
        "bg_color": pale_yellow, "align": "center", "num_format": "0.0",
    })
    param_int = workbook.add_format({
        "font_name": "Calibri", "font_size": 12, "bold": True, "border": 1,
        "bg_color": pale_yellow, "align": "center", "num_format": "0",
    })
    wrap = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "text_wrap": True, "valign": "vcenter"})
    total_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "bold": True, "border": 1,
        "bg_color": pale_green, "num_format": "#,##0", "align": "center",
    })
    total_text = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "bold": True, "border": 1, "bg_color": pale_green,
    })
    note_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "italic": True, "font_color": "#595959",
        "text_wrap": True, "valign": "top",
    })
    kpi_label = workbook.add_format({"font_name": "Calibri", "font_size": 11, "border": 1, "valign": "vcenter"})
    kpi_orange = workbook.add_format({
        "font_name": "Calibri", "font_size": 14, "bold": True, "border": 1,
        "bg_color": pale_orange, "num_format": "#,##0", "align": "center",
    })
    kpi_blue = workbook.add_format({
        "font_name": "Calibri", "font_size": 14, "bold": True, "border": 1,
        "bg_color": pale_blue, "num_format": "#,##0", "align": "center",
    })
    kpi_label_o = workbook.add_format({"font_name": "Calibri", "font_size": 11, "border": 1, "bg_color": pale_orange, "valign": "vcenter"})
    kpi_label_b = workbook.add_format({"font_name": "Calibri", "font_size": 11, "border": 1, "bg_color": pale_blue, "valign": "vcenter"})
    flag_fmt = workbook.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1, "align": "center",
        "num_format": "0", "valign": "vcenter",
    })

    columns = [
        ("Source", "input_text", 10),
        ("GBSC", "input_text", 12),
        ("Cell CI", "input_text", 12),
        ("Cell Name", "input_text", 16),
        ("Cell Index", "input_int", 12),
        ("Band", "input_center", 12),
        ("BH Days", "input_int", 11),
        ("Daily Days", "input_int", 12),
        ("TRX Configured", "input_int", 12),
        ("BH K3014 avg", "input_num", 14),
        ("BH Carried TCH, Erl", "formula_num", 16),
        ("BH Carried formula", "formula_show", 36),
        ("BH K3034 avg", "input_num", 13),
        ("BH TCHH, Erl", "formula_num", 14),
        ("BH TCHH formula", "formula_show", 36),
        ("Half Rate Traffic Ratio", "formula_ratio", 14),
        ("Half Rate formula", "formula_show", 42),
        ("Avail TCH Mean", "input_num", 14),
        ("Capacity, Erl", "formula_num", 13),
        ("Capacity formula", "formula_show", 42),
        ("Voice Load, %", "formula_num", 13),
        ("Voice Load formula", "formula_show", 46),
        ("BH TCH Cong_N avg", "counter", 14),
        ("BH TCH Cong_D avg", "counter", 14),
        ("TCH Cong, %", "formula_num", 12),
        ("TCH Cong formula", "formula_show", 42),
        ("BH SDCCH Busy Fail avg", "counter", 16),
        ("BH SDCCH Request avg", "counter", 16),
        ("SDCCH Cong, %", "formula_num", 13),
        ("SDCCH Cong formula", "formula_show", 46),
        ("BH K3004 avg", "counter", 14),
        ("BH SDCCH, Erl", "formula_num", 14),
        ("BH SDCCH formula", "formula_show", 36),
        ("Daily K3014 avg", "counter", 14),
        ("Daily avg TCH, Erl", "formula_num", 14),
        ("Daily TCH formula", "formula_show", 36),
        ("Daily K3034 avg", "counter", 14),
        ("Daily avg TCHH, Erl", "formula_num", 14),
        ("Daily TCHH formula", "formula_show", 36),
        ("Daily K3004 avg", "counter", 14),
        ("Daily avg SDCCH, Erl", "formula_num", 15),
        ("Daily SDCCH formula", "formula_show", 36),
        ("File CDR_N", "counter", 12),
        ("CDR_N daily avg", "formula_num", 13),
        ("CDR_N formula", "formula_show", 28),
        ("File CDR_D", "counter", 12),
        ("CDR, %", "formula_num", 12),
        ("CDR formula", "formula_show", 36),
        ("File DL Bad_N", "counter", 13),
        ("File DL Qual_D", "counter", 14),
        ("DL Bad Qual(6~7), %", "formula_num", 14),
        ("DL Bad Qual formula", "formula_show", 48),
        ("File UL Bad_N", "counter", 13),
        ("File UL Qual_D", "counter", 14),
        ("UL Bad Qual(6~7), %", "formula_num", 14),
        ("UL Bad Qual formula", "formula_show", 48),
        ("File HOSR_N", "counter", 12),
        ("File HOSR_D", "counter", 12),
        ("HOSR_Out, %", "formula_num", 13),
        ("HOSR formula", "formula_show", 36),
        ("File ICM 4-5", "counter", 13),
        ("File ICM_D", "counter", 13),
        ("ICM_B4~5, %", "formula_num", 13),
        ("ICM formula", "formula_show", 42),
        ("File Incoming HOSR_N", "counter", 16),
        ("File Incoming HOSR_D", "counter", 16),
        ("Incoming HOSR, %", "formula_num", 14),
        ("Incoming HOSR formula", "formula_show", 46),
        ("File CSSR_N", "counter", 12),
        ("File CSSR_D", "counter", 12),
        ("CSSR, %", "formula_num", 12),
        ("CSSR formula", "formula_show", 36),
        ("BH Configured TCH", "counter", 14),
        ("R3027A TCHF 900 avg", "counter", 16),
        ("R3027B TCHF 1800 avg", "counter", 16),
        ("R3028A TCHH 900 avg", "counter", 16),
        ("R3028B TCHH 1800 avg", "counter", 16),
        ("G900 TRX Reduction", "formula_center", 14),
        ("TRX Keep", "formula_int", 11),
        ("TRX Remove", "formula_int", 12),
        ("Voice Load after 2 TRX, %", "formula_num", 16),
        ("Voice Load after formula", "formula_show", 42),
        ("Reduction Reason", "formula_wrap", 36),
        ("G1800 Cell Addition", "formula_center", 14),
        ("Voice Load Band", "formula_center", 14),
        ("Addition Reason", "formula_wrap", 46),
        ("Action", "formula_wrap", 32),
        ("Review", "formula_wrap", 42),
        ("F TRX>2", "flag", 11),
        ("F Cong<1", "flag", 11),
        ("F Load<=60", "flag", 12),
        ("F CDR_N>=10", "flag", 12),
        ("F DL and UL>2", "flag", 13),
        ("F Load>=70", "flag", 12),
        ("F HOSR<=90", "flag", 12),
        ("F CDR and ICM", "flag", 13),
    ]
    letter = {name: col_letter(pos) for pos, (name, _kind, _width) in enumerate(columns, start=1)}

    def ref(name, row):
        return f"{letter[name]}{row}"

    parameters = workbook.add_worksheet("Parameters")
    parameters.set_tab_color(teal)
    parameters.set_column("A:A", 22)
    parameters.set_column("B:B", 14)
    parameters.set_column("C:C", 16)
    parameters.set_column("D:D", 92)
    parameters.write("A1", "Thresholds used by the All Cells formulas", title_fmt)
    parameters.write("A2", "Yellow cells are the live inputs. Edit one and Excel recalculates All Cells and the Summary counts. The two action sheets are the saved result at these defaults.", note_fmt)
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
    parameters.set_row(1, 36)

    all_cells = workbook.add_worksheet("All Cells")
    all_cells.set_tab_color("#7F7F7F")
    all_cells.freeze_panes(1, 4)
    all_cells.autofilter(0, 0, len(records), len(columns) - 1)
    all_cells.set_row(0, 42)

    decision_orange = {
        "G900 TRX Reduction", "TRX Keep", "TRX Remove",
        "Voice Load after 2 TRX, %", "Reduction Reason",
    }
    decision_blue = {"G1800 Cell Addition", "Voice Load Band", "Addition Reason"}
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
        "formula_show": header,
        "flag": header,
    }
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
        "formula_show": wrap,
        "flag": flag_fmt,
    }
    for pos, (name, kind, width) in enumerate(columns):
        if name in decision_orange:
            fmt = header_orange
        elif name in decision_blue:
            fmt = header_blue
        elif name in {"Action", "Review"}:
            fmt = header
        else:
            fmt = kind_header[kind]
        all_cells.set_column(pos, pos, width)
        all_cells.write(0, pos, name, fmt)

    for offset, item in enumerate(records):
        row = offset + 2
        daily = item["daily"] or {}

        def F(name, r=row):
            return ref(name, r)

        formulas = {
            "BH Carried TCH, Erl": f'={F("BH K3014 avg")}*pTrafficHours',
            "BH Carried formula": f'=FORMULATEXT({F("BH Carried TCH, Erl")})',
            "BH TCHH, Erl": f'={F("BH K3034 avg")}*pTrafficHours',
            "BH TCHH formula": f'=FORMULATEXT({F("BH TCHH, Erl")})',
            "BH SDCCH, Erl": f'={F("BH K3004 avg")}*pTrafficHours',
            "BH SDCCH formula": f'=FORMULATEXT({F("BH SDCCH, Erl")})',
            "Half Rate Traffic Ratio": f'=IF({F("BH Carried TCH, Erl")}=0,0,{F("BH TCHH, Erl")}/{F("BH Carried TCH, Erl")})',
            "Half Rate formula": f'=FORMULATEXT({F("Half Rate Traffic Ratio")})',
            "Capacity, Erl": f'={F("Avail TCH Mean")}*(1+{F("Half Rate Traffic Ratio")})',
            "Capacity formula": f'=FORMULATEXT({F("Capacity, Erl")})',
            "Voice Load, %": f'=IF({F("Capacity, Erl")}=0,"",{F("BH Carried TCH, Erl")}/{F("Capacity, Erl")}*100)',
            "Voice Load formula": f'=FORMULATEXT({F("Voice Load, %")})',
            "TCH Cong, %": f'=IF({F("BH TCH Cong_D avg")}=0,0,{F("BH TCH Cong_N avg")}/{F("BH TCH Cong_D avg")}*100)',
            "TCH Cong formula": f'=FORMULATEXT({F("TCH Cong, %")})',
            "SDCCH Cong, %": f'=IF({F("BH SDCCH Request avg")}=0,0,{F("BH SDCCH Busy Fail avg")}/{F("BH SDCCH Request avg")}*100)',
            "SDCCH Cong formula": f'=FORMULATEXT({F("SDCCH Cong, %")})',
            "CDR_N daily avg": f'={F("File CDR_N")}',
            "CDR_N formula": f'=FORMULATEXT({F("CDR_N daily avg")})',
            "DL Bad Qual(6~7), %": f'=IF({F("File DL Qual_D")}=0,"",{F("File DL Bad_N")}/{F("File DL Qual_D")}*100)',
            "DL Bad Qual formula": f'=FORMULATEXT({F("DL Bad Qual(6~7), %")})',
            "UL Bad Qual(6~7), %": f'=IF({F("File UL Qual_D")}=0,"",{F("File UL Bad_N")}/{F("File UL Qual_D")}*100)',
            "UL Bad Qual formula": f'=FORMULATEXT({F("UL Bad Qual(6~7), %")})',
            "CDR, %": f'=IF({F("File CDR_D")}=0,"",{F("File CDR_N")}/{F("File CDR_D")}*100)',
            "CDR formula": f'=FORMULATEXT({F("CDR, %")})',
            "HOSR_Out, %": f'=IF({F("File HOSR_D")}=0,"",{F("File HOSR_N")}/{F("File HOSR_D")}*100)',
            "HOSR formula": f'=FORMULATEXT({F("HOSR_Out, %")})',
            "ICM_B4~5, %": f'=IF({F("File ICM_D")}=0,"",{F("File ICM 4-5")}/{F("File ICM_D")}*100)',
            "ICM formula": f'=FORMULATEXT({F("ICM_B4~5, %")})',
            "Incoming HOSR, %": f'=IF({F("File Incoming HOSR_D")}=0,"",{F("File Incoming HOSR_N")}/{F("File Incoming HOSR_D")}*100)',
            "Incoming HOSR formula": f'=FORMULATEXT({F("Incoming HOSR, %")})',
            "CSSR, %": f'=IF({F("File CSSR_D")}=0,"",{F("File CSSR_N")}/{F("File CSSR_D")}*100)',
            "CSSR formula": f'=FORMULATEXT({F("CSSR, %")})',
            "Daily avg TCH, Erl": f'={F("Daily K3014 avg")}*pTrafficHours',
            "Daily TCH formula": f'=FORMULATEXT({F("Daily avg TCH, Erl")})',
            "Daily avg TCHH, Erl": f'={F("Daily K3034 avg")}*pTrafficHours',
            "Daily TCHH formula": f'=FORMULATEXT({F("Daily avg TCHH, Erl")})',
            "Daily avg SDCCH, Erl": f'={F("Daily K3004 avg")}*pTrafficHours',
            "Daily SDCCH formula": f'=FORMULATEXT({F("Daily avg SDCCH, Erl")})',
            "Voice Load after formula": f'=FORMULATEXT({F("Voice Load after 2 TRX, %")})',
            "F TRX>2": f'=IF({F("TRX Configured")}>pTRXAbove,1,0)',
            "F Cong<1": f'=IF(AND({F("TCH Cong, %")}<pCongMax,{F("SDCCH Cong, %")}<pCongMax),1,0)',
            "F Load<=60": f'=IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}<=pVoiceLoadMax),1,0)',
            "F CDR_N>=10": f'=IF({F("CDR_N daily avg")}>=pCDRnMin,1,0)',
            "F DL and UL>2": (
                f'=IF(AND(ISNUMBER({F("DL Bad Qual(6~7), %")}),{F("DL Bad Qual(6~7), %")}>pBadQualMin,'
                f'ISNUMBER({F("UL Bad Qual(6~7), %")}),{F("UL Bad Qual(6~7), %")}>pBadQualMin),1,0)'
            ),
            "F Load>=70": f'=IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}>=pVoiceLoadHigh),1,0)',
            "F HOSR<=90": f'=IF(AND(ISNUMBER({F("HOSR_Out, %")}),{F("HOSR_Out, %")}<=pHOSRMax),1,0)',
            "F CDR and ICM": (
                f'=IF(AND(ISNUMBER({F("CDR, %")}),{F("CDR, %")}>pCDRMin,'
                f'ISNUMBER({F("ICM_B4~5, %")}),{F("ICM_B4~5, %")}>pICMMin),1,0)'
            ),
            "G900 TRX Reduction": (
                f'=IF(AND({F("Band")}="G900",{F("F TRX>2")}=1,{F("F Cong<1")}=1,'
                f'OR({F("F Load<=60")}=1,{F("F CDR_N>=10")}=1,{F("F DL and UL>2")}=1)),"Yes","No")'
            ),
            "TRX Keep": f'=IF({F("G900 TRX Reduction")}="Yes",pTRXKeep,{F("TRX Configured")})',
            "TRX Remove": f'=IF({F("G900 TRX Reduction")}="Yes",MAX(0,{F("TRX Configured")}-pTRXKeep),0)',
            "Voice Load after 2 TRX, %": (
                f'=IF(AND({F("G900 TRX Reduction")}="Yes",pTRXKeep>0,ISNUMBER({F("Voice Load, %")})),'
                f'{F("Voice Load, %")}*{F("TRX Configured")}/pTRXKeep,"")'
            ),
            "Reduction Reason": (
                f'=IF({F("G900 TRX Reduction")}<>"Yes","",MID('
                f'IF({F("F Load<=60")}=1,", Voice load<="&pVoiceLoadMax&"%","")'
                f'&IF({F("F CDR_N>=10")}=1,", CDR_N>="&pCDRnMin,"")'
                f'&IF({F("F DL and UL>2")}=1,", DL and UL bad quality>"&pBadQualMin&"%","")'
                f',3,300))'
            ),
            "G1800 Cell Addition": (
                f'=IF(AND({F("Band")}="G900",{F("TRX Configured")}>=pAddTRXMin,OR('
                f'{F("F Load>=70")}=1,{F("F HOSR<=90")}=1,{F("F DL and UL>2")}=1,'
                f'{F("F CDR_N>=10")}=1,{F("F CDR and ICM")}=1)),"Yes","No")'
            ),
            "Voice Load Band": (
                f'=IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}>=pVoiceLoadHigh),"High >=70%",'
                f'IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}<pVoiceLoadHigh),"Low <70%",""))'
            ),
            "Addition Reason": (
                f'=IF({F("G1800 Cell Addition")}<>"Yes","",MID('
                f'IF({F("F Load>=70")}=1,", Voice load>="&pVoiceLoadHigh&"%","")'
                f'&IF({F("F HOSR<=90")}=1,", HOSR_Out<="&pHOSRMax&"%","")'
                f'&IF({F("F DL and UL>2")}=1,", DL and UL bad quality>"&pBadQualMin&"%","")'
                f'&IF({F("F CDR_N>=10")}=1,", CDR_N>="&pCDRnMin,"")'
                f'&IF({F("F CDR and ICM")}=1,", CDR>"&pCDRMin&"% and ICM_B4~5>"&pICMMin&"%","")'
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
                f'&IF(AND({F("G900 TRX Reduction")}="Yes",ISNUMBER({F("Voice Load after 2 TRX, %")}),{F("Voice Load after 2 TRX, %")}>100),'
                f'"; Projected BH voice load exceeds 100% after cutting to "&pTRXKeep&" TRX","")'
                f'&IF(AND({F("G900 TRX Reduction")}="Yes",ISNUMBER({F("Voice Load after 2 TRX, %")}),'
                f'{F("Voice Load after 2 TRX, %")}>pVoiceLoadMax,{F("Voice Load after 2 TRX, %")}<=100),'
                f'"; Projected BH voice load exceeds "&pVoiceLoadMax&"% after cutting to "&pTRXKeep&" TRX","")'
                f'&IF({F("BH Days")}<pDailyDays,"; BH sample is shorter than 7 days","")'
                f'&IF({F("Daily Days")}<pDailyDays,"; Daily sample is shorter than 7 days","")'
                f',3,400)'
            ),
        }

        values = {
            "Source": item["source"],
            "GBSC": item["gbsc"],
            "Cell CI": item["ci"],
            "Cell Name": item["name"],
            "Cell Index": item["cell_index"] if isinstance(item["cell_index"], int) else 0,
            "Band": item["band"],
            "BH Days": item["days"],
            "Daily Days": item["daily_days"],
            "TRX Configured": item["trx"],
            "BH K3014 avg": item["k3014_avg"],
            "BH K3034 avg": item["k3034_avg"],
            "Avail TCH Mean": item["avail"],
            "BH TCH Cong_N avg": item["tchn_avg"],
            "BH TCH Cong_D avg": item["tchd_avg"],
            "BH SDCCH Busy Fail avg": item["sdn_avg"],
            "BH SDCCH Request avg": item["sdd_avg"],
            "File CDR_N": daily.get("cdrn", 0.0),
            "File CDR_D": daily.get("cdrd", 0.0),
            "File HOSR_N": daily.get("hosrn", 0.0),
            "File HOSR_D": daily.get("hosrd", 0.0),
            "File DL Bad_N": daily.get("dln", 0.0),
            "File DL Qual_D": daily.get("dld", 0.0),
            "File UL Bad_N": daily.get("uln", 0.0),
            "File UL Qual_D": daily.get("uld", 0.0),
            "File ICM 4-5": daily.get("icmn", 0.0),
            "File ICM_D": daily.get("icmd", 0.0),
            "File Incoming HOSR_N": daily.get("ihosrn", 0.0),
            "File Incoming HOSR_D": daily.get("ihosrd", 0.0),
            "File CSSR_N": daily.get("cssrn", 0.0),
            "File CSSR_D": daily.get("cssrd", 0.0),
            "Daily K3014 avg": daily.get("k3014", 0.0),
            "Daily K3034 avg": daily.get("k3034", 0.0),
            "Daily K3004 avg": daily.get("k3004", 0.0),
            "BH K3004 avg": item["k3004_avg"],
            "BH Configured TCH": item["configured_tch"],
            "R3027A TCHF 900 avg": item["a900_avg"],
            "R3027B TCHF 1800 avg": item["b1800_avg"],
            "R3028A TCHH 900 avg": item["h900_avg"],
            "R3028B TCHH 1800 avg": item["h1800_avg"],
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

    reduce_col = letter["G900 TRX Reduction"]
    add_col = letter["G1800 Cell Addition"]
    all_cells.conditional_format(1, 0, len(records), len(columns) - 1, {
        "type": "formula",
        "criteria": f'=${reduce_col}2="Yes"',
        "format": workbook.add_format({"bg_color": "#FFF4EC"}),
    })
    all_cells.conditional_format(1, 0, len(records), len(columns) - 1, {
        "type": "formula",
        "criteria": f'=AND(${add_col}2="Yes",${reduce_col}2<>"Yes")',
        "format": workbook.add_format({"bg_color": "#F3F8FD"}),
    })

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

    reductions = [item for item in records if item["reduce"]]
    additions = [item for item in records if item["add"]]
    reductions.sort(key=lambda item: (item["gbsc"], -item["trx"], item["name"]))
    additions.sort(key=lambda item: (item["gbsc"], -(item["voice"] or -1), item["name"]))

    reduction_headers = [
        ("Source", 10, cell_center),
        ("GBSC", 12, cell_text),
        ("Cell CI", 12, cell_text),
        ("Cell Name", 16, cell_text),
        ("BH Days", 11, cell_int),
        ("TRX now", 11, cell_int),
        ("TRX keep", 11, cell_int),
        ("TRX remove", 12, cell_int),
        ("BH Carried TCH, Erl", 16, cell_num),
        ("Half Rate Traffic Ratio", 16, cell_ratio),
        ("Avail TCH Mean", 14, cell_num),
        ("Capacity, Erl", 14, cell_num),
        ("Voice Load, %", 14, cell_num),
        ("Voice Load after 2 TRX, %", 18, cell_num),
        ("TCH Cong, %", 13, cell_num),
        ("SDCCH Cong, %", 14, cell_num),
        ("CDR_N daily avg", 14, cell_num),
        ("DL Bad Qual(6~7), %", 16, cell_num),
        ("UL Bad Qual(6~7), %", 16, cell_num),
        ("CDR, %", 12, cell_num),
        ("HOSR_Out, %", 13, cell_num),
        ("Projected load check", 22, cell_center),
        ("Reduction Reason", 42, wrap),
        ("Review", 46, wrap),
    ]
    reduction_rows = []
    for item in reductions:
        if item["load_after"] is None:
            check = ""
        elif item["load_after"] > 100:
            check = "Above 100%"
        elif item["load_after"] > 60:
            check = "Above 60%"
        else:
            check = "Within 60%"
        reduction_rows.append([
            item["source"], item["gbsc"], item["ci"], item["name"], item["days"],
            item["trx"], 2, item["trx"] - 2, item["carried"], item["half_rate"],
            item["avail"], item["capacity"], blank(item["voice"]), blank(item["load_after"]),
            item["tch_cong"], item["sdcch_cong"], blank(item["cdr_n"]),
            blank(item["dl"]), blank(item["ul"]), blank(item["cdr"]), blank(item["hosr"]),
            check, item["reason_r"], item["notes"],
        ])
    reduction_sheet = workbook.add_worksheet("G900 TRX Reduction")
    reduction_sheet.set_tab_color(orange)
    write_table(reduction_sheet, reduction_headers, reduction_rows, header_orange)
    if reduction_rows:
        reduction_sheet.conditional_format(1, 13, len(reduction_rows), 13, {
            "type": "cell", "criteria": ">", "value": 60,
            "format": workbook.add_format({
                "bg_color": pale_red, "font_name": "Calibri", "font_size": 10,
                "border": 1, "num_format": "#,##0.00", "align": "center",
            }),
        })

    addition_headers = [
        ("Source", 10, cell_center),
        ("GBSC", 12, cell_text),
        ("Cell CI", 12, cell_text),
        ("Cell Name", 16, cell_text),
        ("BH Days", 11, cell_int),
        ("TRX now", 11, cell_int),
        ("Voice Load Band", 16, cell_center),
        ("BH Carried TCH, Erl", 16, cell_num),
        ("Voice Load, %", 14, cell_num),
        ("HOSR_Out, %", 13, cell_num),
        ("DL Bad Qual(6~7), %", 16, cell_num),
        ("UL Bad Qual(6~7), %", 16, cell_num),
        ("CDR_N daily avg", 14, cell_num),
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
            item["load_band"], item["carried"], blank(item["voice"]), blank(item["hosr"]),
            blank(item["dl"]), blank(item["ul"]), blank(item["cdr_n"]), blank(item["cdr"]),
            blank(item["icm"]), item["tch_cong"], item["sdcch_cong"],
            "Yes" if item["reduce"] else "No", item["reason_a"],
        ])
    addition_sheet = workbook.add_worksheet("G1800 Addition")
    addition_sheet.set_tab_color(blue)
    write_table(addition_sheet, addition_headers, addition_rows, header_blue)

    # Summary counts are formulas on All Cells, so a threshold edit updates them.
    summary = workbook.add_worksheet("Summary")
    summary.set_tab_color(navy)
    summary.set_column("A:A", 3)
    summary.set_column("B:B", 78)
    summary.set_column("C:C", 16)
    summary.set_column("D:D", 18)
    summary.set_column("E:E", 22)
    summary.set_column("F:L", 16)
    summary.hide_gridlines(2)
    summary.write("B1", "GSM capacity enhancement", title_fmt)
    summary.write("B2", "G900 TRX reduction and G1800 cell addition. Busy hour 26 Oct 2026 to 1 Nov 2026.", section_fmt)
    summary.write(
        "B3",
        "Voice load, available TCH, half-rate ratio, TCH congestion, and SDCCH congestion come from the busy-hour files. "
        "CDR_N, HOSR, DL and UL bad quality, ICM band 4-5, incoming HOSR, and CSSR come from the daily files. "
        "Every KPI is the average of the dates available for that cell. Traffic Erlang from both files is then multiplied by 24. Voice load uses the busy-hour traffic after that multiplication.",
        note_fmt,
    )
    summary.set_row(2, 48)

    data_first = 2
    data_last = last
    sheet = "'All Cells'"

    def countifs(*pairs):
        parts = []
        for name, criteria in pairs:
            parts.append(f"{sheet}!{letter[name]}${data_first}:{letter[name]}${data_last}")
            parts.append(criteria)
        return "=COUNTIFS(" + ",".join(parts) + ")"

    def sumifs(sum_name, *pairs):
        parts = [f"{sheet}!{letter[sum_name]}${data_first}:{letter[sum_name]}${data_last}"]
        for name, criteria in pairs:
            parts.append(f"{sheet}!{letter[name]}${data_first}:{letter[name]}${data_last}")
            parts.append(criteria)
        return "=SUMIFS(" + ",".join(parts) + ")"

    summary.write("B5", "G900 TRX reduction", header_orange)
    summary.write("C5", "Cells / TRX", header_orange)
    reduction_formulas = [
        ("G900 cells", countifs(("Band", '"G900"'))),
        ("G900 cells with more than 2 TRX", countifs(("Band", '"G900"'), ("F TRX>2", "1"))),
        ("Of those, SDCCH and TCH congestion both under 1%", countifs(("Band", '"G900"'), ("F TRX>2", "1"), ("F Cong<1", "1"))),
        ("Cells reduced to 2 TRX", countifs(("G900 TRX Reduction", '"Yes"'))),
        ("G900 TRX removed", sumifs("TRX Remove", ("G900 TRX Reduction", '"Yes"'))),
        ("Reduced cells that also get a G1800 cell", countifs(("G900 TRX Reduction", '"Yes"'), ("G1800 Cell Addition", '"Yes"'))),
    ]
    for index, (label, formula) in enumerate(reduction_formulas):
        summary.write(5 + index, 1, label, kpi_label_o)
        summary.write_formula(5 + index, 2, formula, kpi_orange)

    summary.merge_range("E5:F5", "G1800 cell addition", header_blue)
    summary.write("G5", "Cells", header_blue)
    addition_formulas = [
        ("G900 cells proposed for a new G1800 cell", countifs(("G1800 Cell Addition", '"Yes"'))),
        ("Triggered by voice load >= 70%", countifs(("G1800 Cell Addition", '"Yes"'), ("F Load>=70", "1"))),
        ("Triggered by HOSR_Out <= 90%", countifs(("G1800 Cell Addition", '"Yes"'), ("F HOSR<=90", "1"))),
        ("Triggered by DL and UL bad quality > 2%", countifs(("G1800 Cell Addition", '"Yes"'), ("F DL and UL>2", "1"))),
        ("Triggered by daily-average CDR_N >= 10", countifs(("G1800 Cell Addition", '"Yes"'), ("F CDR_N>=10", "1"))),
        ("Triggered by CDR > 0.2% and ICM band 4-5 > 50%", countifs(("G1800 Cell Addition", '"Yes"'), ("F CDR and ICM", "1"))),
        ("High-load band, voice load >= 70%", countifs(("G1800 Cell Addition", '"Yes"'), ("Voice Load Band", '"High >=70%"'))),
    ]
    for index, (label, formula) in enumerate(addition_formulas):
        summary.merge_range(5 + index, 4, 5 + index, 5, label, kpi_label_b)
        summary.write_formula(5 + index, 6, formula, kpi_blue)

    summary.write("B13", "Footprint", header)
    summary.write("C13", "Cells", header)
    footprint = [
        ("All cells", f"=COUNTA({sheet}!{letter['Cell Name']}${data_first}:{letter['Cell Name']}${data_last})"),
        ("G900", countifs(("Band", '"G900"'))),
        ("G1800 already on air", countifs(("Band", '"G1800"'))),
        ("Both 900 and 1800 TCH", countifs(("Band", '"Both"'))),
        ("No 900 or 1800 TCH samples", countifs(("Band", '"NoTCH"'))),
        ("Cells with fewer than 7 busy-hour days", f'=COUNTIF({sheet}!{letter["BH Days"]}${data_first}:{letter["BH Days"]}${data_last},"<7")'),
    ]
    for index, (label, formula) in enumerate(footprint):
        summary.write(13 + index, 1, label, kpi_label)
        summary.write_formula(13 + index, 2, formula, int_fmt)

    bsc_header_row = 21
    bsc_headers = [
        "GBSC", "Cells", "G900", "G1800 already", "Both",
        "Cells reduced", "TRX removed", "G1800 cells to add", "Both actions",
    ]
    for col, text in enumerate(bsc_headers):
        summary.write(bsc_header_row, 1 + col, text, header)
    for index, gbsc in enumerate(BSC_ORDER):
        row = bsc_header_row + 1 + index
        summary.write(row, 1, gbsc, cell_text)
        bsc_formulas = [
            countifs(("GBSC", f'"{gbsc}"')),
            countifs(("GBSC", f'"{gbsc}"'), ("Band", '"G900"')),
            countifs(("GBSC", f'"{gbsc}"'), ("Band", '"G1800"')),
            countifs(("GBSC", f'"{gbsc}"'), ("Band", '"Both"')),
            countifs(("GBSC", f'"{gbsc}"'), ("G900 TRX Reduction", '"Yes"')),
            sumifs("TRX Remove", ("GBSC", f'"{gbsc}"'), ("G900 TRX Reduction", '"Yes"')),
            countifs(("GBSC", f'"{gbsc}"'), ("G1800 Cell Addition", '"Yes"')),
            countifs(("GBSC", f'"{gbsc}"'), ("G900 TRX Reduction", '"Yes"'), ("G1800 Cell Addition", '"Yes"')),
        ]
        for col, formula in enumerate(bsc_formulas, start=1):
            summary.write_formula(row, 1 + col, formula, cell_int)
    total_row = bsc_header_row + 1 + len(BSC_ORDER)
    summary.write(total_row, 1, "Total", total_text)
    first_bsc = bsc_header_row + 2
    last_bsc = bsc_header_row + 1 + len(BSC_ORDER)
    for col in range(1, len(bsc_headers)):
        excel_col = col_letter(2 + col)
        summary.write_formula(total_row, 1 + col, f"=SUM({excel_col}{first_bsc}:{excel_col}{last_bsc})", total_fmt)

    note_row = total_row + 2
    summary.write(note_row, 1, "How to read the result", section_fmt)
    notes = [
        "G900 TRX Reduction is the list of G900 cells that go down to 2 TRX. TRX removed is configured TRX minus 2. A red voice-load-after-cut value is the busy-hour load if the remaining 2 TRX carry today's traffic and available TCH scales with the TRX count.",
        "G1800 Addition is the list of G900 cells proposed for a new DCS1800 cell. A cell that already has 1800 TCH is not on this list. The voice-load band says whether the cell is at or above 70% or below 70%. Voice load below 70% does not by itself add a cell.",
        "All Cells has one row per cell. Gray columns are the date averages from the files. Navy, orange, and blue columns are formulas. BH Carried TCH multiplies the busy-hour K3014 average by 24, and that figure is the carried traffic in voice load. Daily avg TCH does the same multiplication for the daily file and is not used in voice load.",
        "Parameters holds the thresholds. Changing a yellow cell recalculates All Cells and the counts on this Summary. Filter All Cells on G900 TRX Reduction or G1800 Cell Addition after a threshold change. The two action sheets stay at the thresholds saved with this file.",
        "Rules records the file split, the formulas, and the reading of each OR gate.",
    ]
    for index, text in enumerate(notes):
        summary.merge_range(note_row + 1 + index, 1, note_row + 1 + index, 8, text, text_fmt)
        summary.set_row(note_row + 1 + index, 36)

    chart = workbook.add_chart({"type": "column"})
    chart.set_title({"name": "G900 TRX removed and G1800 cells to add"})
    chart.set_y_axis({"name": "Count"})
    chart.set_size({"width": 720, "height": 320})
    chart.add_series({
        "name": "TRX removed",
        "categories": ["Summary", first_bsc - 1, 1, last_bsc - 1, 1],
        "values": ["Summary", first_bsc - 1, 7, last_bsc - 1, 7],
        "fill": {"color": orange},
    })
    chart.add_series({
        "name": "G1800 cells to add",
        "categories": ["Summary", first_bsc - 1, 1, last_bsc - 1, 1],
        "values": ["Summary", first_bsc - 1, 8, last_bsc - 1, 8],
        "fill": {"color": blue},
    })
    chart.set_legend({"position": "bottom"})
    summary.insert_chart(note_row + 7, 1, chart)

    rules = workbook.add_worksheet("Rules")
    rules.set_tab_color("#548235")
    rules.set_column("A:A", 3)
    rules.set_column("B:B", 128)
    rule_lines = [
        "",
        "Purpose",
        "G900 TRX comes down on cells that can release TRX while SDCCH and TCH congestion stay under 1%. A G1800 cell is proposed on a G900 cell whose busy-hour load, handover, quality, drops, or interference meet the addition rule.",
        "",
        "Files",
        "Busy-hour files: 2G Capacity File_V1_BH and 2G Capacity File_V2_BH. Daily files: 2G Capacity File_V1_Daily and 2G Capacity File_V2_Daily. Each file has one row per cell per day, 26 Oct 2026 through 1 Nov 2026. V1 is MYMH05, MYMH06, MYMH07, and SYLH04. V2 is TNGH12.",
        "A cell is averaged over the dates that cell has. A cell with 7 dates uses 7. A cell with fewer dates uses only those dates. The average is not filled up to 7 and is not divided by 7 again.",
        "The join key is source, GBSC, cell CI, and cell name.",
        "",
        "What comes from the busy hour",
        "Voice load, carried TCH Erlang (K3014), half-rate TCH Erlang (K3034), available TCH (K3015), configured TRX (S3655), configured TCH (K3016), SDCCH Erlang (K3004), TCH congestion, and SDCCH congestion.",
        "Each of those starts as the average of the cell's available busy-hour dates. A ratio uses average numerator over average denominator, which is the same as the sum over the sum. Traffic Erlang (K3014, K3034, K3004) is then multiplied by 24. Available TCH, configured TRX, and the congestion counters are not multiplied by 24.",
        "TCH Cong, % = TCH Congestion_N / TCH Congestion_D x 100. No denominator is 0%.",
        "SDCCH Cong, % = K3001 Failed SDCCH seizures due to busy SDCCH / K3000 SDCCH seizure requests x 100. The capacity file has no SDCCH congestion column, so this is the pair used for the under-1% gate. No denominator is 0%.",
        "",
        "What comes from the daily file",
        "CDR_N, CDR %, HOSR_Out, DL bad quality, UL bad quality, ICM band 4-5, incoming HOSR, and CSSR. These are the rest of the KPIs. Each is averaged over the dates available for that cell.",
        "CDR_N daily avg is that date average. It is the number compared with 10. It is not divided by 7 again.",
        "A ratio is the average numerator over the average denominator. ICM_B4~5, % = Interference Band_4-5_Samples / Interference Band Samples_D x 100.",
        "DL Bad Qual(6~7), % = DL Quality(6~7)_N / DL Quality_D x 100. The file names the denominator DL Quality_D. UL uses UL Quality(6~7)_N / UL Quality_D.",
        "HOSR_Out, % = HOSR_N / HOSR_D x 100. CDR, % = CDR_N / CDR_D x 100.",
        "Daily traffic is averaged across the cell's dates, then multiplied by 24. Daily avg TCH, Erl = daily K3014 average x 24. The same x24 is applied to daily K3034 and K3004. These daily traffic columns are reference only. Voice load does not use them.",
        "Where a denominator is zero, voice load, bad quality, CDR, HOSR, ICM, incoming HOSR, and CSSR are blank, and a greater-than or less-than test on that KPI is not met.",
        "",
        "Voice load",
        "Half Rate Traffic Ratio = BH K3034 / BH K3014. Multiplying both by 24 leaves the ratio unchanged. The ratio is 0 when busy-hour K3014 is 0.",
        "Capacity, Erl = Avail TCH Mean x (1 + Half Rate Traffic Ratio). Avail TCH Mean is the busy-hour average of K3015 and is not multiplied by 24.",
        "Voice Load, % = BH carried TCH / Capacity x 100. BH carried TCH is the busy-hour average of K3014 multiplied by 24.",
        "",
        "Band",
        "A cell is G900 when the 900 TCH counters (R3027A and R3028A) are above zero across the week and the 1800 counters (R3027B and R3028B) are zero. The reverse is G1800. Both means the cell already carries 900 and 1800 TCH. Those counters are busy-hour sample counters, used only to tell the band. Actions apply to G900 cells.",
        "",
        "G900 TRX reduction",
        "All of these are required: the cell is G900, configured TRX is above 2, SDCCH congestion is below 1%, and TCH congestion is below 1%.",
        "The OR gate is then any one of: busy-hour voice load <= 60%, daily-average CDR_N >= 10, or both DL and UL bad quality above 2%.",
        "DL and UL bad quality above 2% means both directions. One direction above 2% does not meet that term. The two percentages are on the sheet.",
        "Action: keep 2 TRX. TRX removed = configured TRX - 2. A cell at 2 TRX or below is not changed.",
        "Voice load after the cut = current voice load x configured TRX / 2. This assumes available TCH scales with TRX count and the half-rate ratio stays as it is. It is a check beside the action. The action still keeps 2 TRX where the projected load goes above 60% or above 100%. Those rows are marked in Review and highlighted on the reduction sheet.",
        "",
        "G1800 cell addition",
        "The cell is G900. TRX count is not a filter. Trx<>=2 (All scenario) is read as every TRX count, including 1 TRX and more than 2 TRX. pAddTRXMin is 0 for that reading. Set it to 2 on Parameters if addition should start only at 2 TRX.",
        "A new G1800 cell is proposed when any one of these is true: busy-hour voice load >= 70%, HOSR_Out <= 90%, both DL and UL bad quality above 2%, daily-average CDR_N >= 10, or CDR above 0.2% and ICM band 4-5 above 50%.",
        "The low-load line (voice load < 70%) uses the same handover, quality, and drop tests. Voice load below 70% is the group label on Voice Load Band. It is not itself a reason to add a cell, because that comparison is true for almost every cell. A low-load cell is added only when HOSR, both-direction bad quality, CDR_N, or the CDR and ICM pair is true.",
        "A high-load cell is added because voice load is at or above 70%, and also when any of those problem KPIs are true.",
        "CDR > 0.2 is 0.2 percent. CDR and ICM is an AND pair, and that pair is one term in the OR.",
        "A cell that already has 1800 TCH stays off the addition list.",
        "",
        "What this file does not do",
        "It does not plan the L900 bandwidth change from 5 MHz to 10 MHz. It does not use the cell database, TRX parameter, co-MPT, or inventory files. It does not choose a G1800 ARFCN, the number of G1800 TRX, or the BCCH. The addition list says which G900 cell gets a G1800 cell.",
    ]
    rules.write("B1", "Rules applied in this workbook", title_fmt)
    headings = {
        "Purpose", "Files", "What comes from the busy hour", "What comes from the daily file",
        "Voice load", "Band", "G900 TRX reduction", "G1800 cell addition", "What this file does not do",
    }
    for index, line in enumerate(rule_lines):
        if line in headings:
            rules.write(1 + index, 1, line, section_fmt)
        elif line == "":
            rules.write(1 + index, 1, "", text_fmt)
        else:
            rules.write(1 + index, 1, line, text_fmt)
            rules.set_row(1 + index, 18 + 14 * (len(line) // 150))

    order = ["Summary", "G900 TRX Reduction", "G1800 Addition", "All Cells", "Parameters", "Rules"]
    workbook.worksheets_objs.sort(key=lambda sheet: order.index(sheet.name))
    workbook.close()
    return reductions, additions


def main():
    ensure_csvs()
    bh_cells = load_bh()
    daily_cells = load_daily()
    missing_daily = set(bh_cells) - set(daily_cells)
    missing_bh = set(daily_cells) - set(bh_cells)
    records = prepare(bh_cells, daily_cells)
    reductions, additions = build(records)
    dates = set()
    for item in records:
        dates.update(item["dates"])
    print("BH dates", sorted(dates))
    print("cells", len(records), "daily", len(daily_cells), "bh only", len(missing_daily), "daily only", len(missing_bh))
    print("reduce", len(reductions), "trx_remove", sum(item["trx"] - 2 for item in reductions))
    print("add", len(additions), "both actions", sum(1 for item in records if item["reduce"] and item["add"]))
    print("bands", dict(Counter(item["band"] for item in records)))
    print("short BH", sum(1 for item in records if item["days"] < 7))
    for name in ("BMBJN55C", "D50AR9B", "D50AR9A"):
        hit = next(item for item in records if item["name"] == name)
        print(
            name,
            "band", hit["band"],
            "trx", hit["trx"],
            "voice", None if hit["voice"] is None else round(hit["voice"], 2),
            "tchcong", round(hit["tch_cong"], 3),
            "sdcong", round(hit["sdcch_cong"], 3),
            "cdrn", None if hit["cdr_n"] is None else round(hit["cdr_n"], 2),
            "dl", None if hit["dl"] is None else round(hit["dl"], 2),
            "ul", None if hit["ul"] is None else round(hit["ul"], 2),
            "hosr", None if hit["hosr"] is None else round(hit["hosr"], 2),
            "icm", None if hit["icm"] is None else round(hit["icm"], 2),
            "cdr", None if hit["cdr"] is None else round(hit["cdr"], 3),
            "reduce", hit["reduce"],
            "add", hit["add"],
            "reasonR", hit["reason_r"],
            "reasonA", hit["reason_a"],
        )


if __name__ == "__main__":
    main()
