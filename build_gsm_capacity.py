#!/usr/bin/env python3
"""GSM capacity enhancement from the four renamed capacity files.

Busy-hour files supply voice load and the capacity KPIs.
Daily files supply CDR, HOSR, quality, interference, and the other KPIs.
Both files have one row per cell per day. Each cell is averaged over the
dates present for that cell. Busy-hour Erlang is already one hour, so voice
load uses that average with no x24. Daily Erlang is an hourly average, so
daily traffic is that average multiplied by 24.

Capacity Erlang is the Erlang B offered traffic at 5% blocking. The channel
count is available TCH times 1.60, the fixed 60% half-rate design. It is not
available TCH times (1 + the measured half-rate ratio).

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
# Bump VERSION on every workbook change. Keep the earlier version files.
VERSION = "v16"
# Rule v12 is the decision in this file.
# Each version sheet is G900 reduction first, then G1800 addition.
# GSM_Rules_Store_v8.xlsx is the same store.
# Capacity is read from the Half Rate 5% sheet, 60% column.
ACTIVE_RULE = "v12"
ACTIVE_RULE_FILE = "GSM_Rules_Store_v8.xlsx"
OUT = f"GSM_Capacity_Enhancement_{VERSION}.xlsx"
GOS = 0.05
ERLANG_MAX = 200
G1800_MAX_TRX = 12
# The load rule stops at 12 TRX. CDR and bad quality can each add one more.
G1800_TABLE_TRX = 16
# Physical TCH rows on the Half Rate 5% sheet. 120 TCH at 60% half rate is 192 channels.
HR_PHYSICAL_MAX = 120
HR_RATIOS = [
    (0.0, "0", "0% half rate"),
    (0.2, "1/5", "20% half rate"),
    (0.4, "2/5", "40% half rate"),
    (0.5, "1/2", "50% half rate"),
    (0.6, "3/5", "60% half rate"),
    (0.8, "4/5", "80% half rate"),
    (1.0, "1", "100% half rate"),
]
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
    "RR307:TCH Availability(%)",
]

# Yellow cells on Parameters. All Cells formulas read these names.
# pICMMin stays at 50 because the new rule says "high ICM band 4-5" and does not give a new percent.
# pTRXKeep stays at 2: the AND gate says which cells are reduced, and the reduction keeps 2 TRX.
PARAMS = [
    ("pTRXAbove", 2, "TRX", "G900 reduction AND gate: configured TRX must be above this"),
    ("pTRXKeep", 2, "TRX", "TRX kept when the G900 reduction AND gate is met"),
    ("pTCHCongMax", 1, "%", "G900 reduction AND gate: TCH congestion must be below this"),
    ("pSDCCHCongMax", 0.5, "%", "G900 reduction AND gate: SDCCH congestion must be below this"),
    ("pVoiceLoadMax", 40, "%", "G900 reduction AND gate: busy-hour voice load must be at or below this"),
    ("pRR307Ignore", 95, "%", "If busy-hour RR307 is below this, ignore the TCH and SDCCH congestion gates. RR307 above 99% is not a gate"),
    ("pHR", 0.6, "ratio", "Design half-rate ratio for Erlang B capacity. Equivalent channels = available TCH x (1 + this). The measured K3034/K3014 ratio is not used"),
    ("pVoiceLoadHigh", 80, "%", "After the G900 reduction step, voice load at or above this starts the G1800 TRX calculation"),
    ("pVoiceLoadTarget", 60, "%", "G1800 TRX are added until busy-hour voice load is below this"),
    ("pCDRnMin", 10, "drops/day", "Daily-average CDR_N at or above this is a G1800 problem line. CDR, quality, and high ICM share one G1800 TRX"),
    ("pBadQualMin", 2, "%", "Both DL and UL bad quality above this is a G1800 problem line. Problem lines share one G1800 TRX"),
    ("pICMMin", 50, "%", "High ICM band 4-5 means ICM_B4~5 above this. Used only when reduction, load, CDR, and bad quality did not trigger. It shares the one problem-line TRX"),
    ("pMaxG1800TRX", G1800_MAX_TRX, "TRX", "Upper limit of the G1800 TRX table used to bring voice load under the target"),
    ("pDailyDays", 7, "days", "Flag a cell when fewer than this many dates are available. Averages use the dates present"),
    ("pTrafficHours", 24, "hours", "Multiply daily traffic Erlang by this after the date average. Busy-hour Erlang is already one hour and is not multiplied."),
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


def defaults():
    return {name: value for name, value, _unit, _note in PARAMS}


def erlang_b(channels, offered):
    """Share of calls cleared because every channel is busy."""
    if channels <= 0:
        return 1.0
    if offered <= 0:
        return 0.0
    blocking = 1.0
    for k in range(1, channels + 1):
        blocking = offered * blocking / (k + offered * blocking)
    return blocking


def offered_traffic(channels, gos=GOS):
    """Offered Erlang at the target blocking."""
    low = 1e-12
    high = max(float(channels) * 2.0, 1.0)
    for _ in range(80):
        mid = (low + high) / 2.0
        if erlang_b(channels, mid) < gos:
            low = mid
        else:
            high = mid
    return (low + high) / 2.0


def build_erlang_table():
    table = {0: 0.0}
    for channels in range(1, ERLANG_MAX + 1):
        table[channels] = offered_traffic(channels)
    return table


def offered_at(channels, table):
    """Integer Erlang B row, or a straight line between the two nearest rows."""
    if channels <= 0:
        return 0.0
    if channels >= ERLANG_MAX:
        return table[ERLANG_MAX]
    lower = int(channels)
    if lower < 1:
        return table[1] * channels
    fraction = channels - lower
    if fraction < 1e-12:
        return table[lower]
    return table[lower] + fraction * (table[lower + 1] - table[lower])


def g1800_tch(trx):
    """TCH on a new G1800 cell: 8 timeslots per TRX, one BCCH, one SDCCH per two TRX."""
    if trx <= 0:
        return 0
    sdcch = (trx + 1) // 2
    return 8 * trx - 1 - sdcch


def g1800_caps(table, hr):
    caps = []
    for trx in range(0, G1800_TABLE_TRX + 1):
        caps.append(offered_at(g1800_tch(trx) * (1.0 + hr), table))
    return caps


def capacity_of(tch, hr, table):
    """Offered Erl from the Half Rate 5% sheet: physical TCH at the design half rate.

    Integer TCH reads that row. A fractional TCH is the straight line between
    the two nearest physical-TCH rows of the offered column.
    """
    if tch <= 0:
        return 0.0
    if tch >= HR_PHYSICAL_MAX:
        return offered_at(HR_PHYSICAL_MAX * (1.0 + hr), table)
    lower = int(tch)
    if lower < 1:
        return offered_at(1.0 + hr, table) * tch
    low = offered_at(lower * (1.0 + hr), table)
    fraction = tch - lower
    if fraction < 1e-12:
        return low
    high = offered_at((lower + 1) * (1.0 + hr), table)
    return low + fraction * (high - low)


def tch_for_trx(avail, configured, kept):
    if configured <= 0:
        return 0.0
    return avail * kept / configured


def trx_for_deficit(deficit, caps):
    """Smallest G1800 TRX count whose Erlang capacity is strictly above the deficit.

    MATCH on an ascending table returns the last row that is still at or below the
    deficit. The next row is the first one that takes voice load under 60%.
    """
    if deficit <= 0:
        return 0
    position = 0
    for index, capacity in enumerate(caps):
        if capacity <= deficit + 1e-9:
            position = index
        else:
            break
    return min(G1800_MAX_TRX, position + 1)


def decide(band, trx, voice, tch_cong, sdcch_cong, rr307, cdr_n, dl, ul, icm, carried, avail, table, caps, cfg):
    """Mirror of the All Cells formulas at the saved parameter values."""
    g900 = band == "G900"
    load_ok = voice is not None and voice <= cfg["pVoiceLoadMax"]
    # RR307 above 99% is not a gate. Below 95%, both congestion gates are skipped.
    ignore_cong = rr307 < cfg["pRR307Ignore"]
    tch_ok = ignore_cong or tch_cong < cfg["pTCHCongMax"]
    sdcch_ok = ignore_cong or sdcch_cong < cfg["pSDCCHCongMax"]
    trx_ok = trx > cfg["pTRXAbove"]
    reduce = g900 and trx_ok and tch_ok and sdcch_ok and load_ok
    trx_after = min(cfg["pTRXKeep"], trx) if reduce else trx
    tch_after = tch_for_trx(avail, trx, trx_after)
    cap_after = capacity_of(tch_after, cfg["pHR"], table)
    voice_after = (carried / cap_after * 100.0) if cap_after > 0 else None
    load_high = voice_after is not None and voice_after >= cfg["pVoiceLoadHigh"]
    cdr_hit = cdr_n is not None and cdr_n >= cfg["pCDRnMin"]
    qual_hit = dl is not None and ul is not None and dl > cfg["pBadQualMin"] and ul > cfg["pBadQualMin"]
    icm_hit = icm is not None and icm > cfg["pICMMin"]

    n_load = 0
    if g900 and load_high:
        deficit = carried / (cfg["pVoiceLoadTarget"] / 100.0) - cap_after
        n_load = min(cfg["pMaxG1800TRX"], trx_for_deficit(deficit, caps))
    n_cdr = 1 if g900 and cdr_hit else 0
    n_qual = 1 if g900 and qual_hit else 0
    n_icm = 1 if g900 and (not reduce) and (not load_high) and (not cdr_hit) and (not qual_hit) and icm_hit else 0

    # The extra removal is judged on the TRX left after the reduction step.
    # It removes one G900 TRX, once, and only while 3 or more TRX are still there.
    # CDR, quality, and ICM share one G1800 TRX. They do not each remove a G900 TRX.
    still_three = (not reduce) and trx_after >= 3
    removed_cdr = 1 if g900 and still_three and cdr_hit else 0
    removed_qual = 1 if g900 and still_three and qual_hit and not cdr_hit else 0
    removed_icm = 1 if n_icm == 1 and still_three else 0
    removed = removed_cdr + removed_qual + removed_icm
    trx_final = trx_after if reduce else max(0, trx - removed)
    tch_final = tch_for_trx(avail, trx, trx_final)
    g900_cap = capacity_of(tch_final, cfg["pHR"], table)
    # Problem lines share one TRX. The load count already brings voice under 60%,
    # so a problem line does not add another TRX on top of it. If one shared TRX
    # still leaves voice at or above 60%, add the smallest number that goes under.
    n_problem = 0
    if g900 and n_load == 0 and (n_cdr or n_qual or n_icm):
        n_problem = 1
        while n_problem < cfg["pMaxG1800TRX"]:
            cap = g900_cap + caps[n_problem]
            voice_try = (carried / cap * 100.0) if cap > 0 else None
            if voice_try is None or voice_try < cfg["pVoiceLoadTarget"]:
                break
            n_problem += 1
    n_add = n_load + n_problem
    if n_add > G1800_TABLE_TRX:
        raise SystemExit(f"G1800 TRX {n_add} is past the capacity table")
    cap_final = g900_cap + (caps[n_add] if n_add <= G1800_MAX_TRX else caps[-1])
    voice_final = (carried / cap_final * 100.0) if cap_final > 0 else None
    return {
        "reduce": reduce,
        "trx_after": trx_after,
        "cap_after": cap_after,
        "voice_after": voice_after,
        "load_high": load_high,
        "cdr_hit": cdr_hit,
        "qual_hit": qual_hit,
        "icm_hit": icm_hit,
        "n_load": n_load,
        "n_cdr": n_cdr,
        "n_qual": n_qual,
        "n_icm": n_icm,
        "n_problem": n_problem,
        "removed_cdr": removed_cdr,
        "removed_qual": removed_qual,
        "removed_icm": removed_icm,
        "trx_final": trx_final,
        "trx_remove": max(0, trx - trx_final),
        "n_add": n_add,
        "cap_final": cap_final,
        "voice_final": voice_final,
    }


def self_check(table, caps):
    published = {6: 2.960, 7: 3.738, 14: 9.730, 21: 16.19, 29: 23.83, 36: 30.66}
    for channels, expected in published.items():
        if abs(table[channels] - expected) > 0.006:
            raise SystemExit(f"Erlang B {channels} at 5% is {table[channels]:.4f}, Huawei prints {expected}")
    cfg = defaults()
    # TRX, voice, and both congestion gates. Low carried traffic stays under 80% after the cut.
    quiet = decide("G900", 4, 20, 0.2, 0.1, 99.5, 1, 0.4, 0.4, 10, 2.0, 30, table, caps, cfg)
    if not (quiet["reduce"] and quiet["trx_final"] == 2 and quiet["n_add"] == 0 and quiet["trx_remove"] == 2):
        raise SystemExit(f"quiet reduction failed: {quiet}")
    # RR307 below 95 ignores congestion, even when both congestion values are high.
    skipped = decide("G900", 4, 20, 5.0, 3.0, 94.5, 1, 0.4, 0.4, 10, 2.0, 30, table, caps, cfg)
    if not (skipped["reduce"] and skipped["trx_final"] == 2 and skipped["n_add"] == 0):
        raise SystemExit(f"low RR307 should ignore congestion: {skipped}")
    # At 95 the congestion gates still apply, so high TCH congestion blocks the cut.
    held = decide("G900", 4, 20, 5.0, 0.1, 95, 1, 0.4, 0.4, 10, 2.0, 30, table, caps, cfg)
    if held["reduce"]:
        raise SystemExit(f"RR307 at 95 should still test congestion: {held}")
    # Between 95 and 98, congestion is no longer ignored.
    between = decide("G900", 4, 20, 5.0, 0.1, 96, 1, 0.4, 0.4, 10, 2.0, 30, table, caps, cfg)
    if between["reduce"]:
        raise SystemExit(f"RR307 at 96 should still test congestion: {between}")
    # RR307 between 95 and 99 is no longer a block when congestion is low.
    mid = decide("G900", 4, 20, 0.2, 0.1, 96, 1, 0.4, 0.4, 10, 2.0, 30, table, caps, cfg)
    if not (mid["reduce"] and mid["trx_final"] == 2):
        raise SystemExit(f"RR307 between 95 and 99 should reduce: {mid}")
    # TRX is 2, so the reduction gate fails. Load is already high, so G1800 is sized from the table.
    hot = decide("G900", 2, 90, 0.1, 0.1, 100, 0, 0.1, 0.1, 1, 16, 14, table, caps, cfg)
    if hot["reduce"] or hot["n_load"] < 1 or hot["trx_final"] != 2:
        raise SystemExit(f"high load on 2 TRX failed: {hot}")
    if hot["n_load"] < G1800_MAX_TRX and not (hot["voice_final"] < cfg["pVoiceLoadTarget"]):
        raise SystemExit(f"G1800 add did not take load under 60%: {hot}")
    # No full reduction. Four TRX are still there, so one more G900 TRX comes off, not two.
    # CDR and bad quality share one G1800 TRX.
    both = decide("G900", 4, 50, 0.2, 0.2, 100, 12, 3, 3, 80, 8, 30, table, caps, cfg)
    if both["reduce"] or both["n_cdr"] != 1 or both["n_qual"] != 1 or both["n_icm"] != 0 or both["n_add"] != 1 or both["trx_final"] != 3 or both["trx_remove"] != 1:
        raise SystemExit(f"CDR and quality path failed: {both}")
    # ICM only when the other lines are quiet. High TCH congestion with RR307 at 99 blocks the cut. 4 TRX loses 1.
    icm_only = decide("G900", 4, 20, 2.0, 0.2, 99, 1, 0.2, 0.2, 70, 3, 30, table, caps, cfg)
    if icm_only["reduce"] or icm_only["n_icm"] != 1 or icm_only["n_add"] != 1 or icm_only["trx_final"] != 3:
        raise SystemExit(f"ICM-only path failed: {icm_only}")
    # Already reduced to 2, so CDR adds a G1800 TRX and does not remove more G900.
    locked = decide("G900", 4, 20, 0.2, 0.1, 100, 15, 0.2, 0.2, 10, 2, 30, table, caps, cfg)
    if not locked["reduce"] or locked["trx_final"] != 2 or locked["n_cdr"] != 1 or locked["removed_cdr"] != 0:
        raise SystemExit(f"locked reduction failed: {locked}")
    # 3 TRX can lose only one. CDR and quality share one G1800 TRX.
    short = decide("G900", 3, 50, 0.2, 0.2, 100, 12, 3, 4, 10, 6, 21, table, caps, cfg)
    if short["reduce"] or short["trx_final"] != 2 or short["n_add"] != 1 or short["trx_remove"] != 1:
        raise SystemExit(f"3 TRX sequential remove failed: {short}")
    other = decide("G1800", 4, 10, 0, 0, 100, 20, 5, 5, 90, 4, 20, table, caps, cfg)
    if other["reduce"] or other["n_add"] or other["trx_remove"]:
        raise SystemExit(f"non-G900 cell was changed: {other}")


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


def prepare(bh_cells, daily_cells, table, caps, cfg):
    records = []
    for key, item in bh_cells.items():
        days = item["days"]
        daily = daily_cells.get(key)
        k3014_avg = item["K3014:Traffic Volume on TCH(Erl)"] / days
        k3034_avg = item["K3034:TCHH Traffic Volume(Erl)"] / days
        k3004_avg = item["K3004:Traffic Volume on SDCCH(Erl)"] / days
        carried = k3014_avg
        tchh = k3034_avg
        avail = item["K3015:Available TCHs"] / days
        half = (k3034_avg / k3014_avg) if k3014_avg > 0 else 0.0
        equivalent = avail * (1.0 + cfg["pHR"])
        capacity = capacity_of(avail, cfg["pHR"], table)
        voice = (carried / capacity * 100.0) if capacity > 0 else None
        tch_cong = (item["TCH Congestion_N"] / item["TCH Congestion_D"] * 100.0) if item["TCH Congestion_D"] > 0 else 0.0
        sdcch_cong = (
            item["K3001:Failed SDCCH Seizures due to Busy SDCCH"] / item["K3000:SDCCH Seizure Requests"] * 100.0
        ) if item["K3000:SDCCH Seizure Requests"] > 0 else 0.0
        rr307 = item["RR307:TCH Availability(%)"] / days
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

        plan = decide(
            band, trx, voice, tch_cong, sdcch_cong, rr307, cdr_n, dl, ul, icm,
            carried, avail, table, caps, cfg,
        )
        reduce = plan["reduce"]
        add = plan["n_add"] > 0

        reasons_r = []
        if reduce:
            if rr307 < cfg["pRR307Ignore"]:
                reasons_r.append(
                    f"TRX>{cfg['pTRXAbove']}, voice load<={cfg['pVoiceLoadMax']}%, "
                    f"RR307<{cfg['pRR307Ignore']}% so TCH and SDCCH congestion are ignored"
                )
            else:
                reasons_r.append(
                    f"TRX>{cfg['pTRXAbove']}, TCH cong<{cfg['pTCHCongMax']}%, "
                    f"SDCCH cong<{cfg['pSDCCHCongMax']}%, voice load<={cfg['pVoiceLoadMax']}%"
                )
        reasons_a = []
        problem_bits = []
        if plan["n_cdr"]:
            problem_bits.append(f"Daily CDR_N>={cfg['pCDRnMin']}")
        if plan["n_qual"]:
            problem_bits.append(f"DL and UL bad quality>{cfg['pBadQualMin']}%")
        if plan["n_icm"]:
            problem_bits.append(f"ICM band 4-5>{cfg['pICMMin']}% and no other trigger")
        if plan["n_load"]:
            reasons_a.append(
                f"Voice load after G900 step {plan['voice_after']:.1f}% >= {cfg['pVoiceLoadHigh']}%, "
                f"add {plan['n_load']} G1800 TRX toward <{cfg['pVoiceLoadTarget']}%"
            )
            if problem_bits:
                reasons_a.append("A problem line also triggered. No extra G1800 TRX, because voice load is already brought under 60%")
        elif plan["n_problem"]:
            shared = "; ".join(problem_bits)
            if plan["n_problem"] == 1:
                reasons_a.append(f"{shared}. Problem lines share 1 G1800 TRX")
            else:
                reasons_a.append(
                    f"{shared}. One G1800 TRX leaves voice load at or above {cfg['pVoiceLoadTarget']}%, "
                    f"so add {plan['n_problem']} G1800 TRX"
                )
        if plan["trx_remove"] and not reduce:
            reasons_a.append("After the G900 step at least 3 TRX remain, so 1 more G900 TRX is removed with the G1800 add")
        elif add and reduce:
            reasons_a.append("G900 already reduced to the keep value, so it is not reduced further")

        notes = []
        if band == "G1800":
            notes.append("G1800 cell already on air")
        elif band == "Both":
            notes.append("Cell already has 900 and 1800 TCH")
        elif band == "NoTCH":
            notes.append("No 900 or 1800 TCH samples in the busy hour")
        if daily is None:
            notes.append("No matching daily row")
        if plan["n_load"] and plan["voice_final"] is not None and plan["voice_final"] >= cfg["pVoiceLoadTarget"]:
            notes.append(
                f"Voice load stays at or above {cfg['pVoiceLoadTarget']}% after the G1800 TRX on this row"
            )
        if days < 7:
            notes.append("BH sample is shorter than 7 days")
        if daily is not None and daily_days < 7:
            notes.append("Daily sample is shorter than 7 days")

        if reduce and add:
            action = f"Reduce G900 to {plan['trx_final']} TRX and add {plan['n_add']} G1800 TRX"
        elif reduce:
            action = f"Reduce G900 to {plan['trx_final']} TRX"
        elif add and plan["trx_remove"]:
            action = f"Add {plan['n_add']} G1800 TRX and remove {plan['trx_remove']} G900 TRX"
        elif add:
            action = f"Add {plan['n_add']} G1800 TRX"
        else:
            action = "No action"

        item.update({
            "k3014_avg": k3014_avg,
            "k3034_avg": k3034_avg,
            "k3004_avg": k3004_avg,
            "carried": carried,
            "tchh": tchh,
            "sdcch_erl": k3004_avg,
            "avail": avail,
            "configured_tch": item["K3016:Configured TCHs"] / days,
            "half_rate": half,
            "equivalent": equivalent,
            "capacity": capacity,
            "voice": voice,
            "rr307": rr307,
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
            "plan": plan,
            "reason_r": ", ".join(reasons_r),
            "reason_a": "; ".join(reasons_a),
            "notes": "; ".join(notes),
            "action": action,
        })
        records.append(item)
    records.sort(key=lambda row: (row["gbsc"], row["name"], row["ci"]))
    return records


def blank(value):
    return "" if value is None else value


def build(records, table):
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
        "bg_color": pale_yellow, "align": "center", "num_format": "0.00",
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
        ("Traffic Volume on TCH_K3014_Avg", "input_num", 28),
        ("TCHH Traffic Volume_K3034_Avg", "input_num", 26),
        ("Half Rate Traffic Ratio", "formula_ratio", 14),
        ("Available TCHs_K3015_Avg", "input_num", 22),
        ("Equivalent channels", "formula_num", 16),
        ("Capacity, Erl", "formula_num", 14),
        ("Voice Load, %", "formula_num", 13),
        ("TCH Congestion_N_Avg", "counter", 20),
        ("TCH Congestion_D_Avg", "counter", 20),
        ("TCH Cong, %", "formula_num", 12),
        ("Failed SDCCH Seizures due to Busy SDCCH_K3001_Avg", "counter", 32),
        ("SDCCH Seizure Requests_K3000_Avg", "counter", 26),
        ("SDCCH Cong, %", "formula_num", 13),
        ("TCH Availability_RR307_Avg", "counter", 24),
        ("Traffic Volume on SDCCH_K3004_Avg", "counter", 28),
        ("Daily Traffic Volume on TCH_K3014_Avg", "counter", 30),
        ("Daily avg TCH, Erl", "formula_num", 14),
        ("Daily TCHH Traffic Volume_K3034_Avg", "counter", 28),
        ("Daily avg TCHH, Erl", "formula_num", 14),
        ("Daily Traffic Volume on SDCCH_K3004_Avg", "counter", 30),
        ("Daily avg SDCCH, Erl", "formula_num", 15),
        ("CDR_N_Avg", "counter", 14),
        ("CDR_D_Avg", "counter", 14),
        ("CDR, %", "formula_num", 12),
        ("DL Quality(6~7)_N_Avg", "counter", 22),
        ("DL Quality_D_Avg", "counter", 16),
        ("DL Bad Qual(6~7), %", "formula_num", 14),
        ("UL Quality(6~7)_N_Avg", "counter", 22),
        ("UL Quality_D_Avg", "counter", 16),
        ("UL Bad Qual(6~7), %", "formula_num", 14),
        ("HOSR_N_Avg", "counter", 14),
        ("HOSR_D_Avg", "counter", 14),
        ("HOSR_Out, %", "formula_num", 13),
        ("Interference Band_4-5_Samples_Avg", "counter", 28),
        ("Interference Band Samples_D_Avg", "counter", 28),
        ("ICM_B4~5, %", "formula_num", 13),
        ("Incoming HOSR_N_Avg", "counter", 20),
        ("Incoming HOSR_D_Avg", "counter", 20),
        ("Incoming HOSR, %", "formula_num", 14),
        ("CSSR_N_Avg", "counter", 14),
        ("CSSR_D_Avg", "counter", 14),
        ("CSSR, %", "formula_num", 12),
        ("Configured TCHs_K3016_Avg", "counter", 22),
        ("Available Channels TCHF 900_R3027A_Avg", "counter", 28),
        ("Available Channels TCHF 1800_R3027B_Avg", "counter", 28),
        ("Available Channels TCHH 900_R3028A_Avg", "counter", 28),
        ("Available Channels TCHH 1800_R3028B_Avg", "counter", 28),
        ("F TRX>2", "flag", 11),
        ("F TCH Cong<1", "flag", 13),
        ("F SDCCH Cong<0.5", "flag", 15),
        ("F Load<=40", "flag", 12),
        ("F RR307<95", "flag", 13),
        ("G900 TRX Reduction", "formula_center", 16),
        ("TRX after reduction", "formula_int", 14),
        ("TCH after reduction", "formula_num", 16),
        ("Capacity after reduction, Erl", "formula_num", 18),
        ("Voice Load after reduction, %", "formula_num", 18),
        ("F Load after>=80", "flag", 14),
        ("G1800 TRX for load", "formula_int", 14),
        ("F CDR_N>=10", "flag", 12),
        ("F DL and UL>2", "flag", 13),
        ("F High ICM", "flag", 12),
        ("CDR line", "formula_int", 12),
        ("Quality line", "formula_int", 14),
        ("ICM line", "formula_int", 12),
        ("G1800 TRX for problem lines", "formula_int", 18),
        ("G900 removed for CDR", "formula_int", 14),
        ("G900 removed for quality", "formula_int", 16),
        ("G900 removed for ICM", "formula_int", 14),
        ("G900 TRX final", "formula_int", 13),
        ("TRX Remove", "formula_int", 12),
        ("G1800 TRX Add", "formula_int", 13),
        ("G1800 Addition", "formula_center", 14),
        ("TCH final", "formula_num", 12),
        ("Capacity final, Erl", "formula_num", 16),
        ("Voice Load final, %", "formula_num", 16),
        ("Reduction Reason", "formula_wrap", 42),
        ("Addition Reason", "formula_wrap", 55),
        ("Action", "formula_wrap", 42),
        ("Review", "formula_wrap", 42),
    ]
    letter = {name: col_letter(pos) for pos, (name, _kind, _width) in enumerate(columns, start=1)}

    def ref(name, row):
        return f"{letter[name]}{row}"

    def half_rate_lookup(physical):
        """60% half-rate offered Erlang. physical is a TCH count, not equivalent channels."""
        rows = "'Half Rate 5%'!$A$7:$A$126"
        offered = "'Half Rate 5%'!$K$7:$K$126"
        return (
            f"IF({physical}<=0,0,"
            f"IF({physical}<1,INDEX({offered},1)*{physical},"
            f"IF({physical}>=120,INDEX({offered},120),"
            f"IF(ABS({physical}-INT({physical}))<1E-9,INDEX({offered},MATCH(INT({physical}),{rows},0)),"
            f"INDEX({offered},MATCH(INT({physical}),{rows},0))+({physical}-INT({physical}))*"
            f"(INDEX({offered},MATCH(INT({physical})+1,{rows},0))-INDEX({offered},MATCH(INT({physical}),{rows},0)))))))"
        )

    parameters = workbook.add_worksheet("Parameters")
    parameters.set_tab_color(teal)
    parameters.set_column("A:A", 22)
    parameters.set_column("B:B", 14)
    parameters.set_column("C:C", 16)
    parameters.set_column("D:D", 110)
    parameters.write("A1", "Thresholds used by the All Cells formulas", title_fmt)
    parameters.write("A2", "Yellow cells are the live inputs. Edit one and Excel recalculates All Cells, the Erlang B lookup, and the Summary counts. The two action sheets are the saved result at these defaults.", note_fmt)
    parameters.set_row(1, 36)
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

    half = workbook.add_worksheet("Half Rate 5%")
    half.set_tab_color("#C65911")
    half.freeze_panes(6, 1)
    half.set_column("A:A", 16)
    half.write("A1", "Half-rate capacity read from the 5% Erlang B table", title_fmt)
    half.write("A2", "Equivalent channels = physical TCH x (1 + half-rate traffic ratio). Capacity, Erl reads the 60% half rate column, Offered Erl at 5%. A fractional TCH count is the straight line between the two nearest physical TCH rows.", note_fmt)
    half.write("A3", "10 physical TCH at full rate carry 6.22 Erl. The same 10 TCH at 60% half rate are 16 channels and carry 11.54 Erl. Huawei 5% checks: 14 TCH = 9.73, 21 TCH = 16.19, 29 TCH = 23.83, 36 TCH = 30.66.", note_fmt)
    half.set_row(1, 36)
    half.set_row(2, 32)
    half.set_row(4, 22)
    half.set_row(5, 30)
    header_hr = workbook.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "font_color": "white", "bg_color": "#C65911", "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
    half.merge_range(4, 0, 5, 0, "Physical TCH", header)
    for index, (_ratio, _expr, ratio_label) in enumerate(HR_RATIOS):
        col = 1 + index * 2
        half.set_column(col, col + 1, 16)
        group = header_hr if _expr == "3/5" else header
        half.merge_range(4, col, 4, col + 1, ratio_label, group)
        half.write(5, col, "Equivalent channels", group)
        half.write(5, col + 1, "Offered Erl at 5%", group)
    marked_tch = {6, 7, 10, 13, 14, 21, 29, 36}
    hi_n = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "num_format": "0", "bg_color": pale_yellow, "bold": True})
    hi_e = workbook.add_format({"font_name": "Calibri", "font_size": 10, "border": 1, "align": "center", "num_format": "0.00", "bg_color": pale_yellow, "bold": True})
    for physical in range(1, HR_PHYSICAL_MAX + 1):
        row = 6 + physical - 1
        excel_row = row + 1
        marked = physical in marked_tch
        half.write_number(row, 0, physical, hi_n if marked else cell_int)
        for index, (ratio, expr, _label) in enumerate(HR_RATIOS):
            col = 1 + index * 2
            highlight = marked or expr == "3/5"
            number_fmt = hi_e if highlight else cell_num
            half.write_formula(row, col, f"=$A{excel_row}*(1+{expr})", number_fmt, physical * (1.0 + ratio))
            half.write_number(row, col + 1, offered_at(physical * (1.0 + ratio), table), number_fmt)
    half.write(4, 16, "10 TCH example", header)
    half.write(5, 16, "Full rate, offered Erl at 5%", cell_text)
    half.write_number(5, 17, table[10], cell_num)
    half.write(6, 16, "60% half rate, equivalent channels", cell_text)
    half.write_number(6, 17, 16, cell_num)
    half.write(7, 16, "60% half rate, offered Erl at 5%", cell_text)
    half.write_number(7, 17, table[16], cell_num)

    g1800 = workbook.add_worksheet("G1800 TRX Capacity")
    g1800.set_tab_color(blue)
    g1800.set_column("A:G", 22)
    g1800.write("A1", "New G1800 cell capacity at 5% blocking and the design half-rate ratio", title_fmt)
    g1800.write("A2", "TCH = 8 x TRX - 1 BCCH - one SDCCH for every two TRX. Capacity is the 60% half-rate column of the Half Rate 5% sheet. All Cells uses this column to choose how many G1800 TRX bring voice load under pVoiceLoadTarget.", note_fmt)
    g1800.set_row(1, 36)
    for col, text in enumerate(["TRX", "Timeslots", "BCCH", "SDCCH", "TCH", "Equivalent channels", "Capacity, Erl"]):
        g1800.write(3, col, text, header)
    for trx in range(0, G1800_TABLE_TRX + 1):
        row = 4 + trx
        excel_row = row + 1
        g1800.write_number(row, 0, trx, cell_int)
        g1800.write_formula(row, 1, f"=A{excel_row}*8", cell_int, trx * 8)
        g1800.write_formula(row, 2, f"=IF(A{excel_row}=0,0,1)", cell_int, 0 if trx == 0 else 1)
        g1800.write_formula(row, 3, f"=IF(A{excel_row}=0,0,ROUNDUP(A{excel_row}/2,0))", cell_int, 0 if trx == 0 else (trx + 1) // 2)
        g1800.write_formula(row, 4, f"=IF(A{excel_row}=0,0,B{excel_row}-C{excel_row}-D{excel_row})", cell_int, g1800_tch(trx))
        g1800.write_formula(row, 5, f"=E{excel_row}*(1+pHR)", cell_num, g1800_tch(trx) * (1 + defaults()["pHR"]))
        g1800.write_formula(
            row, 6, "=" + half_rate_lookup(f"E{excel_row}"), cell_num,
            offered_at(g1800_tch(trx) * (1 + defaults()["pHR"]), table),
        )
    workbook.define_name("G1800Cap", f"='G1800 TRX Capacity'!$G$5:$G${5 + G1800_TABLE_TRX}")

    all_cells = workbook.add_worksheet("All Cells")
    all_cells.set_tab_color("#7F7F7F")
    all_cells.freeze_panes(1, 4)
    all_cells.autofilter(0, 0, len(records), len(columns) - 1)
    all_cells.set_row(0, 48)
    decision_orange = {
        "G900 TRX Reduction", "TRX after reduction", "TCH after reduction",
        "Capacity after reduction, Erl", "Voice Load after reduction, %", "Reduction Reason",
        "G900 removed for CDR", "G900 removed for quality", "G900 removed for ICM",
        "G900 TRX final", "TRX Remove",
    }
    decision_blue = {
        "F Load after>=80", "G1800 TRX for load", "CDR line", "Quality line", "ICM line",
        "G1800 TRX for problem lines", "G1800 TRX Add", "G1800 Addition", "Addition Reason",
        "TCH final", "Capacity final, Erl", "Voice Load final, %",
    }
    kind_header = {
        "input_text": header_gray, "input_int": header_gray, "input_center": header_gray,
        "input_num": header_gray, "counter": header_gray, "formula_ratio": header,
        "formula_num": header, "formula_center": header, "formula_int": header,
        "formula_wrap": header, "flag": header,
    }
    value_formats = {
        "input_text": input_text, "input_int": input_int, "input_center": input_center,
        "input_num": input_num, "counter": input_num, "formula_ratio": cell_ratio,
        "formula_num": cell_num, "formula_center": cell_center, "formula_int": cell_int,
        "formula_wrap": wrap, "flag": flag_fmt,
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
        plan = item["plan"]

        def F(name, r=row):
            return ref(name, r)

        formulas = {
            "Half Rate Traffic Ratio": f'=IF({F("Traffic Volume on TCH_K3014_Avg")}=0,0,{F("TCHH Traffic Volume_K3034_Avg")}/{F("Traffic Volume on TCH_K3014_Avg")})',
            "Equivalent channels": f'={F("Available TCHs_K3015_Avg")}*(1+pHR)',
            "Capacity, Erl": "=" + half_rate_lookup(F("Available TCHs_K3015_Avg")),
            "Voice Load, %": f'=IF({F("Capacity, Erl")}=0,"",{F("Traffic Volume on TCH_K3014_Avg")}/{F("Capacity, Erl")}*100)',
            "TCH Cong, %": f'=IF({F("TCH Congestion_D_Avg")}=0,0,{F("TCH Congestion_N_Avg")}/{F("TCH Congestion_D_Avg")}*100)',
            "SDCCH Cong, %": f'=IF({F("SDCCH Seizure Requests_K3000_Avg")}=0,0,{F("Failed SDCCH Seizures due to Busy SDCCH_K3001_Avg")}/{F("SDCCH Seizure Requests_K3000_Avg")}*100)',
            "DL Bad Qual(6~7), %": f'=IF({F("DL Quality_D_Avg")}=0,"",{F("DL Quality(6~7)_N_Avg")}/{F("DL Quality_D_Avg")}*100)',
            "UL Bad Qual(6~7), %": f'=IF({F("UL Quality_D_Avg")}=0,"",{F("UL Quality(6~7)_N_Avg")}/{F("UL Quality_D_Avg")}*100)',
            "CDR, %": f'=IF({F("CDR_D_Avg")}=0,"",{F("CDR_N_Avg")}/{F("CDR_D_Avg")}*100)',
            "HOSR_Out, %": f'=IF({F("HOSR_D_Avg")}=0,"",{F("HOSR_N_Avg")}/{F("HOSR_D_Avg")}*100)',
            "ICM_B4~5, %": f'=IF({F("Interference Band Samples_D_Avg")}=0,"",{F("Interference Band_4-5_Samples_Avg")}/{F("Interference Band Samples_D_Avg")}*100)',
            "Incoming HOSR, %": f'=IF({F("Incoming HOSR_D_Avg")}=0,"",{F("Incoming HOSR_N_Avg")}/{F("Incoming HOSR_D_Avg")}*100)',
            "CSSR, %": f'=IF({F("CSSR_D_Avg")}=0,"",{F("CSSR_N_Avg")}/{F("CSSR_D_Avg")}*100)',
            "Daily avg TCH, Erl": f'={F("Daily Traffic Volume on TCH_K3014_Avg")}*pTrafficHours',
            "Daily avg TCHH, Erl": f'={F("Daily TCHH Traffic Volume_K3034_Avg")}*pTrafficHours',
            "Daily avg SDCCH, Erl": f'={F("Daily Traffic Volume on SDCCH_K3004_Avg")}*pTrafficHours',
            "F TRX>2": f'=IF({F("TRX Configured")}>pTRXAbove,1,0)',
            "F TCH Cong<1": f'=IF({F("TCH Cong, %")}<pTCHCongMax,1,0)',
            "F SDCCH Cong<0.5": f'=IF({F("SDCCH Cong, %")}<pSDCCHCongMax,1,0)',
            "F Load<=40": f'=IF(AND(ISNUMBER({F("Voice Load, %")}),{F("Voice Load, %")}<=pVoiceLoadMax),1,0)',
            "F RR307<95": f'=IF({F("TCH Availability_RR307_Avg")}<pRR307Ignore,1,0)',
            "G900 TRX Reduction": (
                f'=IF(AND({F("Band")}="G900",{F("F TRX>2")}=1,{F("F Load<=40")}=1,'
                f'OR({F("F RR307<95")}=1,{F("F TCH Cong<1")}=1),'
                f'OR({F("F RR307<95")}=1,{F("F SDCCH Cong<0.5")}=1)),"Yes","No")'
            ),
            "TRX after reduction": f'=IF({F("G900 TRX Reduction")}="Yes",MIN(pTRXKeep,{F("TRX Configured")}),{F("TRX Configured")})',
            "TCH after reduction": f'=IF({F("TRX Configured")}=0,0,{F("Available TCHs_K3015_Avg")}*{F("TRX after reduction")}/{F("TRX Configured")})',
            "Capacity after reduction, Erl": "=" + half_rate_lookup(F("TCH after reduction")),
            "Voice Load after reduction, %": f'=IF({F("Capacity after reduction, Erl")}=0,"",{F("Traffic Volume on TCH_K3014_Avg")}/{F("Capacity after reduction, Erl")}*100)',
            "F Load after>=80": f'=IF(AND(ISNUMBER({F("Voice Load after reduction, %")}),{F("Voice Load after reduction, %")}>=pVoiceLoadHigh),1,0)',
            "G1800 TRX for load": (
                f'=IF(OR({F("Band")}<>"G900",{F("F Load after>=80")}<>1),0,'
                f'IF({F("Traffic Volume on TCH_K3014_Avg")}/(pVoiceLoadTarget/100)-{F("Capacity after reduction, Erl")}<=0,0,'
                f'MIN(pMaxG1800TRX,MATCH({F("Traffic Volume on TCH_K3014_Avg")}/(pVoiceLoadTarget/100)-{F("Capacity after reduction, Erl")},G1800Cap,1))))'
            ),
            "F CDR_N>=10": f'=IF({F("CDR_N_Avg")}>=pCDRnMin,1,0)',
            "F DL and UL>2": (
                f'=IF(AND(ISNUMBER({F("DL Bad Qual(6~7), %")}),{F("DL Bad Qual(6~7), %")}>pBadQualMin,'
                f'ISNUMBER({F("UL Bad Qual(6~7), %")}),{F("UL Bad Qual(6~7), %")}>pBadQualMin),1,0)'
            ),
            "F High ICM": f'=IF(AND(ISNUMBER({F("ICM_B4~5, %")}),{F("ICM_B4~5, %")}>pICMMin),1,0)',
            "CDR line": f'=IF(AND({F("Band")}="G900",{F("F CDR_N>=10")}=1),1,0)',
            "Quality line": f'=IF(AND({F("Band")}="G900",{F("F DL and UL>2")}=1),1,0)',
            "ICM line": (
                f'=IF(AND({F("Band")}="G900",{F("G900 TRX Reduction")}<>"Yes",{F("F Load after>=80")}<>1,'
                f'{F("F CDR_N>=10")}<>1,{F("F DL and UL>2")}<>1,{F("F High ICM")}=1),1,0)'
            ),
            "G1800 TRX for problem lines": (
                f'=IF(OR({F("Band")}<>"G900",{F("G1800 TRX for load")}>0,AND({F("CDR line")}=0,{F("Quality line")}=0,{F("ICM line")}=0)),0,'
                f'IF({F("Traffic Volume on TCH_K3014_Avg")}/(pVoiceLoadTarget/100)-(' + half_rate_lookup(F("TCH final")) + f')<=0,1,'
                f'MIN(pMaxG1800TRX,MAX(1,MATCH({F("Traffic Volume on TCH_K3014_Avg")}/(pVoiceLoadTarget/100)-(' + half_rate_lookup(F("TCH final")) + f'),G1800Cap,1)))))'
            ),
            "G900 removed for CDR": (
                f'=IF(AND({F("Band")}="G900",{F("G900 TRX Reduction")}<>"Yes",{F("TRX after reduction")}>=3,{F("F CDR_N>=10")}=1),1,0)'
            ),
            "G900 removed for quality": (
                f'=IF(AND({F("Band")}="G900",{F("G900 TRX Reduction")}<>"Yes",{F("TRX after reduction")}>=3,'
                f'{F("F DL and UL>2")}=1,{F("F CDR_N>=10")}<>1),1,0)'
            ),
            "G900 removed for ICM": (
                f'=IF(AND({F("ICM line")}=1,{F("TRX after reduction")}>=3),1,0)'
            ),
            "G900 TRX final": (
                f'=IF({F("G900 TRX Reduction")}="Yes",MIN(pTRXKeep,{F("TRX Configured")}),'
                f'MAX(0,{F("TRX Configured")}-{F("G900 removed for CDR")}-{F("G900 removed for quality")}-{F("G900 removed for ICM")}))'
            ),
            "TRX Remove": f'=MAX(0,{F("TRX Configured")}-{F("G900 TRX final")})',
            "G1800 TRX Add": f'={F("G1800 TRX for load")}+{F("G1800 TRX for problem lines")}',
            "G1800 Addition": f'=IF(AND({F("Band")}="G900",{F("G1800 TRX Add")}>0),"Yes","No")',
            "TCH final": f'=IF({F("TRX Configured")}=0,0,{F("Available TCHs_K3015_Avg")}*{F("G900 TRX final")}/{F("TRX Configured")})',
            "Capacity final, Erl": (
                "=" + half_rate_lookup(F("TCH final"))
                + f'+IF({F("G1800 TRX Add")}=0,0,INDEX(G1800Cap,{F("G1800 TRX Add")}+1))'
            ),
            "Voice Load final, %": f'=IF({F("Capacity final, Erl")}=0,"",{F("Traffic Volume on TCH_K3014_Avg")}/{F("Capacity final, Erl")}*100)',
            "Reduction Reason": (
                f'=IF({F("G900 TRX Reduction")}<>"Yes","",'
                f'IF({F("F RR307<95")}=1,'
                f'"TRX>"&pTRXAbove&", voice load<="&pVoiceLoadMax&"%, RR307<"&pRR307Ignore&"% so TCH and SDCCH congestion are ignored",'
                f'"TRX>"&pTRXAbove&", TCH cong<"&pTCHCongMax&"%, SDCCH cong<"&pSDCCHCongMax&"%, voice load<="&pVoiceLoadMax&"%"))'
            ),
            "Addition Reason": (
                f'=IF({F("G1800 TRX Add")}=0,"",MID('
                f'IF({F("G1800 TRX for load")}>0,"; Voice load after G900 step "&TEXT({F("Voice Load after reduction, %")},"0.0")&"% >= "&pVoiceLoadHigh&"%, add "&{F("G1800 TRX for load")}&" G1800 TRX toward <"&pVoiceLoadTarget&"%","")'
                f'&IF(AND({F("G1800 TRX for load")}>0,OR({F("CDR line")}=1,{F("Quality line")}=1,{F("ICM line")}=1)),"; A problem line also triggered. No extra G1800 TRX, because voice load is already brought under "&pVoiceLoadTarget&"%","")'
                f'&IF(AND({F("G1800 TRX for load")}=0,{F("G1800 TRX for problem lines")}>0),'
                f'IF({F("CDR line")}=1,"; Daily CDR_N>="&pCDRnMin,"")'
                f'&IF({F("Quality line")}=1,"; DL and UL bad quality>"&pBadQualMin&"%","")'
                f'&IF({F("ICM line")}=1,"; ICM band 4-5>"&pICMMin&"% and no other trigger","")'
                f'&IF({F("G1800 TRX for problem lines")}=1,"; Problem lines share 1 G1800 TRX",'
                f'"; One G1800 TRX leaves voice load at or above "&pVoiceLoadTarget&"%, so add "&{F("G1800 TRX for problem lines")}&" G1800 TRX"),"")'
                f'&IF(AND({F("TRX Remove")}>0,{F("G900 TRX Reduction")}<>"Yes"),"; After the G900 step at least 3 TRX remain, so 1 more G900 TRX is removed with the G1800 add","")'
                f'&IF(AND({F("G1800 TRX Add")}>0,{F("G900 TRX Reduction")}="Yes"),"; G900 already reduced to the keep value, so it is not reduced further","")'
                f',3,500))'
            ),
            "Action": (
                f'=IF(AND({F("G900 TRX Reduction")}="Yes",{F("G1800 TRX Add")}>0),"Reduce G900 to "&{F("G900 TRX final")}&" TRX and add "&{F("G1800 TRX Add")}&" G1800 TRX",'
                f'IF({F("G900 TRX Reduction")}="Yes","Reduce G900 to "&{F("G900 TRX final")}&" TRX",'
                f'IF(AND({F("G1800 TRX Add")}>0,{F("TRX Remove")}>0),"Add "&{F("G1800 TRX Add")}&" G1800 TRX and remove "&{F("TRX Remove")}&" G900 TRX",'
                f'IF({F("G1800 TRX Add")}>0,"Add "&{F("G1800 TRX Add")}&" G1800 TRX","No action"))))'
            ),
            "Review": (
                f'=MID('
                f'IF({F("Band")}="G1800","; G1800 cell already on air","")'
                f'&IF({F("Band")}="Both","; Cell already has 900 and 1800 TCH","")'
                f'&IF({F("Band")}="NoTCH","; No 900 or 1800 TCH samples in the busy hour","")'
                f'&IF({F("Daily Days")}=0,"; No matching daily row","")'
                f'&IF(AND({F("G1800 TRX for load")}>0,ISNUMBER({F("Voice Load final, %")}),{F("Voice Load final, %")}>=pVoiceLoadTarget),"; Voice load stays at or above "&pVoiceLoadTarget&"% after the G1800 TRX on this row","")'
                f'&IF({F("BH Days")}<pDailyDays,"; BH sample is shorter than "&pDailyDays&" days","")'
                f'&IF(AND({F("Daily Days")}>0,{F("Daily Days")}<pDailyDays),"; Daily sample is shorter than "&pDailyDays&" days","")'
                f',3,500)'
            ),
        }
        cache = {
            "Half Rate Traffic Ratio": item["half_rate"],
            "Equivalent channels": item["equivalent"],
            "Capacity, Erl": item["capacity"],
            "Voice Load, %": blank(item["voice"]),
            "TCH Cong, %": item["tch_cong"],
            "SDCCH Cong, %": item["sdcch_cong"],
            "DL Bad Qual(6~7), %": blank(item["dl"]),
            "UL Bad Qual(6~7), %": blank(item["ul"]),
            "CDR, %": blank(item["cdr"]),
            "HOSR_Out, %": blank(item["hosr"]),
            "ICM_B4~5, %": blank(item["icm"]),
            "Incoming HOSR, %": blank(item["ihosr"]),
            "CSSR, %": blank(item["cssr"]),
            "Daily avg TCH, Erl": blank(item["daily_tch"]),
            "Daily avg TCHH, Erl": (daily.get("k3034", 0.0) or 0.0) * TRAFFIC_HOURS,
            "Daily avg SDCCH, Erl": (daily.get("k3004", 0.0) or 0.0) * TRAFFIC_HOURS,
            "F TRX>2": 1 if item["trx"] > defaults()["pTRXAbove"] else 0,
            "F TCH Cong<1": 1 if item["tch_cong"] < defaults()["pTCHCongMax"] else 0,
            "F SDCCH Cong<0.5": 1 if item["sdcch_cong"] < defaults()["pSDCCHCongMax"] else 0,
            "F Load<=40": 1 if item["voice"] is not None and item["voice"] <= defaults()["pVoiceLoadMax"] else 0,
            "F RR307<95": 1 if item["rr307"] < defaults()["pRR307Ignore"] else 0,
            "G900 TRX Reduction": "Yes" if plan["reduce"] else "No",
            "TRX after reduction": plan["trx_after"],
            "TCH after reduction": tch_for_trx(item["avail"], item["trx"], plan["trx_after"]),
            "Capacity after reduction, Erl": plan["cap_after"],
            "Voice Load after reduction, %": blank(plan["voice_after"]),
            "F Load after>=80": 1 if plan["load_high"] else 0,
            "G1800 TRX for load": plan["n_load"],
            "F CDR_N>=10": 1 if plan["cdr_hit"] else 0,
            "F DL and UL>2": 1 if plan["qual_hit"] else 0,
            "F High ICM": 1 if plan["icm_hit"] else 0,
            "CDR line": plan["n_cdr"],
            "Quality line": plan["n_qual"],
            "ICM line": plan["n_icm"],
            "G1800 TRX for problem lines": plan["n_problem"],
            "G900 removed for CDR": plan["removed_cdr"],
            "G900 removed for quality": plan["removed_qual"],
            "G900 removed for ICM": plan["removed_icm"],
            "G900 TRX final": plan["trx_final"],
            "TRX Remove": plan["trx_remove"],
            "G1800 TRX Add": plan["n_add"],
            "G1800 Addition": "Yes" if plan["n_add"] else "No",
            "TCH final": tch_for_trx(item["avail"], item["trx"], plan["trx_final"]),
            "Capacity final, Erl": plan["cap_final"],
            "Voice Load final, %": blank(plan["voice_final"]),
            "Reduction Reason": item["reason_r"],
            "Addition Reason": item["reason_a"],
            "Action": item["action"],
            "Review": item["notes"],
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
            "Traffic Volume on TCH_K3014_Avg": item["k3014_avg"],
            "TCHH Traffic Volume_K3034_Avg": item["k3034_avg"],
            "Available TCHs_K3015_Avg": item["avail"],
            "TCH Congestion_N_Avg": item["tchn_avg"],
            "TCH Congestion_D_Avg": item["tchd_avg"],
            "Failed SDCCH Seizures due to Busy SDCCH_K3001_Avg": item["sdn_avg"],
            "SDCCH Seizure Requests_K3000_Avg": item["sdd_avg"],
            "TCH Availability_RR307_Avg": item["rr307"],
            "CDR_N_Avg": daily.get("cdrn", 0.0),
            "CDR_D_Avg": daily.get("cdrd", 0.0),
            "HOSR_N_Avg": daily.get("hosrn", 0.0),
            "HOSR_D_Avg": daily.get("hosrd", 0.0),
            "DL Quality(6~7)_N_Avg": daily.get("dln", 0.0),
            "DL Quality_D_Avg": daily.get("dld", 0.0),
            "UL Quality(6~7)_N_Avg": daily.get("uln", 0.0),
            "UL Quality_D_Avg": daily.get("uld", 0.0),
            "Interference Band_4-5_Samples_Avg": daily.get("icmn", 0.0),
            "Interference Band Samples_D_Avg": daily.get("icmd", 0.0),
            "Incoming HOSR_N_Avg": daily.get("ihosrn", 0.0),
            "Incoming HOSR_D_Avg": daily.get("ihosrd", 0.0),
            "CSSR_N_Avg": daily.get("cssrn", 0.0),
            "CSSR_D_Avg": daily.get("cssrd", 0.0),
            "Daily Traffic Volume on TCH_K3014_Avg": daily.get("k3014", 0.0),
            "Daily TCHH Traffic Volume_K3034_Avg": daily.get("k3034", 0.0),
            "Daily Traffic Volume on SDCCH_K3004_Avg": daily.get("k3004", 0.0),
            "Traffic Volume on SDCCH_K3004_Avg": item["k3004_avg"],
            "Configured TCHs_K3016_Avg": item["configured_tch"],
            "Available Channels TCHF 900_R3027A_Avg": item["a900_avg"],
            "Available Channels TCHF 1800_R3027B_Avg": item["b1800_avg"],
            "Available Channels TCHH 900_R3028A_Avg": item["h900_avg"],
            "Available Channels TCHH 1800_R3028B_Avg": item["h1800_avg"],
        }
        for pos, (name, kind, _width) in enumerate(columns):
            fmt = value_formats[kind]
            if name in formulas:
                all_cells.write_formula(row - 1, pos, formulas[name], fmt, cache[name])
            else:
                value = values[name]
                if isinstance(value, str):
                    all_cells.write_string(row - 1, pos, value, fmt)
                else:
                    all_cells.write_number(row - 1, pos, value, fmt)

    reduce_col = letter["G900 TRX Reduction"]
    add_col = letter["G1800 Addition"]
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
    additions.sort(key=lambda item: (item["gbsc"], -item["plan"]["n_add"], -(item["voice"] or -1), item["name"]))

    reduction_headers = [
        ("Source", 10, cell_center), ("GBSC", 12, cell_text), ("Cell CI", 12, cell_text),
        ("Cell Name", 16, cell_text), ("TRX now", 11, cell_int), ("TRX keep", 11, cell_int),
        ("TRX remove", 12, cell_int), ("Traffic Volume on TCH_K3014_Avg", 16, cell_num),
        ("Equivalent channels", 16, cell_num), ("Capacity, Erl", 14, cell_num),
        ("Voice Load, %", 14, cell_num), ("TCH after reduction", 16, cell_num),
        ("Capacity after reduction, Erl", 18, cell_num), ("Voice Load after reduction, %", 18, cell_num),
        ("TCH Cong, %", 12, cell_num), ("SDCCH Cong, %", 13, cell_num),
        ("TCH Availability_RR307_Avg", 18, cell_num), ("G1800 TRX Add", 14, cell_int),
        ("Voice Load final, %", 16, cell_num), ("Reduction Reason", 46, wrap), ("Review", 42, wrap),
    ]
    reduction_rows = []
    for item in reductions:
        plan = item["plan"]
        reduction_rows.append([
            item["source"], item["gbsc"], item["ci"], item["name"], item["trx"], plan["trx_final"],
            plan["trx_remove"], item["carried"], item["equivalent"], item["capacity"], blank(item["voice"]),
            tch_for_trx(item["avail"], item["trx"], plan["trx_after"]), plan["cap_after"], blank(plan["voice_after"]),
            item["tch_cong"], item["sdcch_cong"], item["rr307"], plan["n_add"], blank(plan["voice_final"]),
            item["reason_r"], item["notes"],
        ])
    reduction_sheet = workbook.add_worksheet("G900 TRX Reduction")
    reduction_sheet.set_tab_color(orange)
    write_table(reduction_sheet, reduction_headers, reduction_rows, header_orange)

    addition_headers = [
        ("Source", 10, cell_center), ("GBSC", 12, cell_text), ("Cell CI", 12, cell_text),
        ("Cell Name", 16, cell_text), ("TRX now", 11, cell_int), ("G900 TRX final", 14, cell_int),
        ("G900 removed", 13, cell_int), ("Voice Load, %", 14, cell_num),
        ("Voice Load after reduction, %", 18, cell_num), ("G1800 for load", 14, cell_int),
        ("CDR line", 12, cell_int), ("Quality line", 14, cell_int), ("ICM line", 12, cell_int),
        ("G1800 for problem lines", 18, cell_int), ("G1800 TRX Add", 14, cell_int),
        ("Capacity final, Erl", 16, cell_num), ("Voice Load final, %", 16, cell_num),
        ("CDR_N_Avg", 12, cell_num), ("DL Bad Qual(6~7), %", 16, cell_num),
        ("UL Bad Qual(6~7), %", 16, cell_num), ("ICM_B4~5, %", 13, cell_num),
        ("Full G900 reduction", 16, cell_center), ("Addition Reason", 55, wrap), ("Review", 42, wrap),
    ]
    addition_rows = []
    for item in additions:
        plan = item["plan"]
        addition_rows.append([
            item["source"], item["gbsc"], item["ci"], item["name"], item["trx"], plan["trx_final"],
            plan["trx_remove"], blank(item["voice"]), blank(plan["voice_after"]), plan["n_load"],
            plan["n_cdr"], plan["n_qual"], plan["n_icm"], plan["n_problem"], plan["n_add"], plan["cap_final"],
            blank(plan["voice_final"]), blank(item["cdr_n"]), blank(item["dl"]), blank(item["ul"]),
            blank(item["icm"]), "Yes" if plan["reduce"] else "No", item["reason_a"], item["notes"],
        ])
    addition_sheet = workbook.add_worksheet("G1800 Addition")
    addition_sheet.set_tab_color(blue)
    write_table(addition_sheet, addition_headers, addition_rows, header_blue)
    if addition_rows:
        addition_sheet.conditional_format(1, 16, len(addition_rows), 16, {
            "type": "cell", "criteria": ">=", "value": 60,
            "format": workbook.add_format({
                "bg_color": pale_red, "font_name": "Calibri", "font_size": 10,
                "border": 1, "num_format": "#,##0.00", "align": "center",
            }),
        })

    summary = workbook.add_worksheet("Summary")
    summary.set_tab_color(navy)
    summary.set_column("A:A", 3)
    summary.set_column("B:B", 78)
    summary.set_column("C:C", 16)
    summary.set_column("D:D", 18)
    summary.set_column("E:E", 24)
    summary.set_column("F:I", 16)
    summary.hide_gridlines(2)
    summary.write("B1", f"GSM capacity enhancement {VERSION}", title_fmt)
    summary.write("B2", f"{VERSION}. Rule v12. G900 TRX reduction is the first step. G1800 TRX addition is the second step. CDR, quality, and high ICM share one G1800 TRX.", section_fmt)
    summary.set_row(1, 32)
    summary.write(
        "B3",
        "Capacity Erlang is the 60% half-rate column of the Half Rate 5% sheet, read from physical TCH. "
        "The measured half-rate ratio is still on the sheet and is not used in that capacity. "
        "Voice load is busy-hour K3014 divided by that capacity. Busy-hour Erlang is not multiplied by 24. "
        "Daily traffic is still the hourly average multiplied by 24 and is not used in voice load.",
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
        ("Reduced to 2 TRX", countifs(("G900 TRX Reduction", '"Yes"'))),
        ("TRX removed by that reduction", sumifs("TRX Remove", ("G900 TRX Reduction", '"Yes"'))),
        ("Of those, voice load after the cut is at least 80%", countifs(("G900 TRX Reduction", '"Yes"'), ("F Load after>=80", "1"))),
    ]
    for index, (label, formula) in enumerate(reduction_formulas):
        summary.write(5 + index, 1, label, kpi_label_o)
        summary.write_formula(5 + index, 2, formula, kpi_orange)

    summary.merge_range("E5:F5", "G1800 TRX addition", header_blue)
    summary.write("G5", "Cells / TRX", header_blue)
    addition_formulas = [
        ("G900 cells that receive G1800 TRX", countifs(("G1800 Addition", '"Yes"'))),
        ("G1800 TRX to add", f"=SUM({sheet}!{letter['G1800 TRX Add']}${data_first}:{letter['G1800 TRX Add']}${data_last})"),
        ("Cells sized because voice load after the G900 step is at least 80%", countifs(("G1800 TRX for load", '">=1"'))),
        ("Cells with daily CDR_N at least 10", countifs(("CDR line", "1"))),
        ("Cells with DL and UL bad quality above 2%", countifs(("Quality line", "1"))),
        ("Cells with high ICM only", countifs(("ICM line", "1"))),
        ("Extra G900 TRX removed outside the full reduction", sumifs("TRX Remove", ("G900 TRX Reduction", '"No"'))),
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
        "GBSC", "Cells", "G900", "Reduced to 2", "TRX removed", "G1800 cells", "G1800 TRX", "Both actions",
    ]
    for col, text in enumerate(bsc_headers):
        summary.write(bsc_header_row, 1 + col, text, header)
    for index, gbsc in enumerate(BSC_ORDER):
        row = bsc_header_row + 1 + index
        summary.write(row, 1, gbsc, cell_text)
        bsc_formulas = [
            countifs(("GBSC", f'"{gbsc}"')),
            countifs(("GBSC", f'"{gbsc}"'), ("Band", '"G900"')),
            countifs(("GBSC", f'"{gbsc}"'), ("G900 TRX Reduction", '"Yes"')),
            sumifs("TRX Remove", ("GBSC", f'"{gbsc}"')),
            countifs(("GBSC", f'"{gbsc}"'), ("G1800 Addition", '"Yes"')),
            sumifs("G1800 TRX Add", ("GBSC", f'"{gbsc}"')),
            countifs(("GBSC", f'"{gbsc}"'), ("G900 TRX Reduction", '"Yes"'), ("G1800 Addition", '"Yes"')),
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
        "G900 TRX Reduction lists every G900 cell that meets the v12 gates and is cut to 2 TRX. Voice load after that cut is recalculated from the Half Rate 5% sheet, with available TCH scaled by 2 / configured TRX.",
        "G1800 Addition lists every G900 cell that receives one or more DCS1800 TRX. G1800 for load is the count that takes the pooled voice load under 60%. CDR, quality, and high ICM share one G1800 TRX. If the load count is already there, those lines add nothing more. If one shared TRX still leaves voice load at or above 60%, the problem-line count is the smallest number that goes under.",
        "A red voice-load-final cell is still at or above 60% after the TRX on that row. Capacity final is the remaining G900 Erlang plus the new G1800 Erlang.",
        "All Cells holds the live formulas. Gray columns are date averages. Capacity, Erl reads the 60% column of the Half Rate 5% sheet. Thresholds are on Parameters.",
        "The two action sheets stay at the thresholds saved with this file. Changing a yellow parameter recalculates All Cells and this Summary.",
    ]
    for index, text in enumerate(notes):
        summary.merge_range(note_row + 1 + index, 1, note_row + 1 + index, 8, text, text_fmt)
        summary.set_row(note_row + 1 + index, 36)

    chart = workbook.add_chart({"type": "column"})
    chart.set_title({"name": "G900 TRX removed and G1800 TRX to add"})
    chart.set_y_axis({"name": "TRX"})
    chart.set_size({"width": 720, "height": 320})
    chart.add_series({
        "name": "TRX removed",
        "categories": ["Summary", first_bsc - 1, 1, last_bsc - 1, 1],
        "values": ["Summary", first_bsc - 1, 5, last_bsc - 1, 5],
        "fill": {"color": orange},
    })
    chart.add_series({
        "name": "G1800 TRX",
        "categories": ["Summary", first_bsc - 1, 1, last_bsc - 1, 1],
        "values": ["Summary", first_bsc - 1, 7, last_bsc - 1, 7],
        "fill": {"color": blue},
    })
    chart.set_legend({"position": "bottom"})
    summary.insert_chart(note_row + 7, 1, chart)

    rules = workbook.add_worksheet("Rules")
    rules.set_tab_color("#548235")
    rules.set_column("A:A", 3)
    rules.set_column("B:B", 140)
    rule_lines = [
        "",
        "Version",
        f"This workbook is {VERSION}: GSM_Capacity_Enhancement_{VERSION}.xlsx. Its decision is rule v12. Open the v12 sheet: G900 reduction is the first step, with the gate table, and G1800 addition is the second step. Capacity is the 60% column of the Half Rate 5% sheet.",
        "",
        "Capacity",
        "Capacity, Erl is the Offered Erl at 5% on the Half Rate 5% sheet, 60% half-rate column, for the physical TCH count. A fractional TCH is the straight line between the two nearest rows.",
        "Half Rate Traffic Ratio is still K3034 / K3014 from the busy hour. It is not used in Capacity, Erl. The design half rate is 60%.",
        "Voice Load, % = busy-hour K3014 / Capacity x 100. Busy-hour K3014 is already one hour and is not multiplied by 24.",
        "G1800 TRX Capacity reads the same 60% column for the new cell's TCH count.",
        "10 physical TCH at full rate carry 6.22 Erl. The same 10 TCH at 60% half rate carry 11.54 Erl.",
        "",
        "1. G900 TRX reduction, first step",
        "Configured TRX above 2. Busy-hour voice load at or below 40%. TCH congestion below 1%, and SDCCH congestion below 0.5%, unless busy-hour RR307 is below 95%. When RR307 is below 95%, both congestion checks are ignored. Busy-hour RR307 above 99% is not a gate. The gate table is on the v12 sheet.",
        "The cell then keeps 2 TRX. TCH after the cut = available TCH x 2 / configured TRX, and voice load is recalculated from that.",
        "If 3 or more TRX are still there after this step, and the G1800 step is adding a TRX for CDR, bad quality, or high ICM, remove exactly 1 more G900 TRX. A cell already at 2 loses nothing more.",
        "",
        "2. G1800 TRX addition, second step",
        "Voice load for these lines is the load after the reduction. A cell that was not reduced uses its current load.",
        "1. Load. If that voice load is at least 80%, add the smallest number of G1800 TRX that brings pooled voice load under 60%. Cap is 12 TRX. This is the capacity count.",
        "2. CDR. Daily CDR_N average at least 10 is one problem line.",
        "3. Quality. DL bad quality above 2% and UL bad quality above 2% is one problem line.",
        "4. High ICM. ICM band 4-5 above 50% is one problem line, and only when the reduction, the load line, CDR, and quality did not trigger.",
        "CDR, quality, and high ICM share one G1800 TRX. A second line does not add another TRX. If the load line already added TRX, these lines add nothing more. If the load line added nothing, add 1 G1800 TRX. If that 1 TRX still leaves voice load at or above 60%, add the smallest number that brings it under 60%.",
        "Actions apply to G900 cells. A cell that already has 1800 TCH is not given another G1800 plan by these rules.",
        "",
        "What did not change",
        "Congestion, CDR, HOSR, bad quality, ICM, incoming HOSR, CSSR, and the daily x24 traffic columns use the same formulas as v7. Daily traffic is reference only.",
        "CDR_N, bad quality, and ICM still come from the daily file. Voice load, available TCH, congestion, TRX, and RR307 come from the busy-hour file. Each cell is averaged over the dates that cell has.",
        "",
        "What this file does not do",
        "It does not plan the L900 move from 5 MHz to 10 MHz. It does not choose a G1800 ARFCN or BCCH.",
    ]
    rules.write("B1", f"Rules applied in this workbook ({VERSION})", title_fmt)
    headings = {
        "Version", "Capacity", "1. G900 TRX reduction, first step", "2. G1800 TRX addition, second step", "What did not change", "What this file does not do",
    }
    for index, line in enumerate(rule_lines):
        if line in headings:
            rules.write(1 + index, 1, line, section_fmt)
        elif line == "":
            rules.write(1 + index, 1, "", text_fmt)
        else:
            rules.write(1 + index, 1, line, text_fmt)
            rules.set_row(1 + index, 18 + 14 * (len(line) // 160))

    from build_gsm_rules import ACTIVE_CAPACITY, ACTIVE_RULE as STORE_RULE, add_rule_sheets
    if STORE_RULE != ACTIVE_RULE or ACTIVE_CAPACITY != VERSION:
        raise SystemExit(f"Rule store {STORE_RULE}/{ACTIVE_CAPACITY} does not match this workbook {ACTIVE_RULE}/{VERSION}")
    rule_sheets = add_rule_sheets(workbook)
    order = [
        "Summary", "Store",
        "G900 TRX Reduction", "G1800 Addition", "All Cells",
        "Half Rate 5%", "G1800 TRX Capacity", "Parameters", "Rules",
        *[name for name in rule_sheets if name != "Store"],
    ]
    workbook.worksheets_objs.sort(key=lambda item: order.index(item.name))
    workbook.close()
    return reductions, additions


def main():
    table = build_erlang_table()
    cfg = defaults()
    caps = g1800_caps(table, cfg["pHR"])
    self_check(table, caps)
    ensure_csvs()
    bh_cells = load_bh()
    daily_cells = load_daily()
    records = prepare(bh_cells, daily_cells, table, caps, cfg)
    failures = []
    still_high = 0
    for item in records:
        plan = item["plan"]
        if plan["reduce"]:
            ignore_cong = item["rr307"] < cfg["pRR307Ignore"]
            ok = (
                item["band"] == "G900"
                and item["trx"] > cfg["pTRXAbove"]
                and (ignore_cong or item["tch_cong"] < cfg["pTCHCongMax"])
                and (ignore_cong or item["sdcch_cong"] < cfg["pSDCCHCongMax"])
                and item["voice"] is not None
                and item["voice"] <= cfg["pVoiceLoadMax"]
                and plan["trx_final"] == min(cfg["pTRXKeep"], item["trx"])
            )
            if not ok:
                failures.append(("reduce", item["name"]))
        if item["band"] != "G900" and (plan["n_add"] or plan["trx_remove"]):
            failures.append(("band", item["name"]))
        lone_load = (
            plan["n_load"] and plan["n_load"] < cfg["pMaxG1800TRX"]
            and plan["n_cdr"] == 0 and plan["n_qual"] == 0 and plan["n_icm"] == 0
        )
        if lone_load and (plan["voice_final"] is None or plan["voice_final"] >= cfg["pVoiceLoadTarget"] + 0.05):
            failures.append(("load", item["name"], plan["voice_final"], plan["n_load"]))
        if plan["n_load"] and plan["voice_final"] is not None and plan["voice_final"] >= cfg["pVoiceLoadTarget"]:
            still_high += 1
        if plan["removed_cdr"] + plan["removed_qual"] + plan["removed_icm"] != (0 if plan["reduce"] else plan["trx_remove"]):
            failures.append(("remove", item["name"]))
    if failures:
        print(failures[:8])
        raise SystemExit(f"{len(failures)} decision checks failed")
    reductions, additions = build(records, table)
    print("workbook", OUT)
    print("cells", len(records))
    print("reduce", len(reductions), "trx_removed_by_reduction", sum(item["plan"]["trx_remove"] for item in reductions))
    print("trx_removed_all", sum(item["plan"]["trx_remove"] for item in records))
    print("g1800_cells", len(additions), "g1800_trx", sum(item["plan"]["n_add"] for item in records))
    print(
        "load_rule", sum(1 for item in records if item["plan"]["n_load"]),
        "cdr_rule", sum(1 for item in records if item["plan"]["n_cdr"]),
        "quality_rule", sum(1 for item in records if item["plan"]["n_qual"]),
        "icm_only", sum(1 for item in records if item["plan"]["n_icm"]),
        "still_at_or_above_60", still_high,
    )
    print("bands", dict(Counter(item["band"] for item in records)))
    sample = next((item for item in records if item["reduce"] and item["plan"]["n_load"]), None)
    if sample:
        plan = sample["plan"]
        print(
            "sample", sample["name"], "trx", sample["trx"], "avail", round(sample["avail"], 2),
            "carried", round(sample["carried"], 2), "voice", round(sample["voice"], 2),
            "voice_after", None if plan["voice_after"] is None else round(plan["voice_after"], 2),
            "g1800", plan["n_add"], "final_trx", plan["trx_final"],
            "voice_final", None if plan["voice_final"] is None else round(plan["voice_final"], 2),
        )


if __name__ == "__main__":
    main()
