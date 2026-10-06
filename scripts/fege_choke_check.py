#!/usr/bin/env python3
"""Find FEGE ports whose RxMaxSpeed trace is choked (stuck flat).

Counters
--------
RxMaxSpeed (Mbit/s) = VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000
Tx Total BW (Mbit/s) = VS.FEGE.TxTotalBW(kbit/s) / 1000

A site is listed only when RxMaxSpeed is pinned to one hard ceiling the way
the sample charts show for DHAPT35 and DHAPT48: after the night dip the line
sits flat, and it cannot climb above that level.

Two shapes are accepted. "Crowded level": most hours sit on one flat level
and nothing climbs far above it (DHAPT35 / DHAPT48). "Hard ceiling": the
trace moves up and down during the day but every rise stops at the same
top value and never goes above it (the clipped ~500 Mbit/s sites such as
DHSVRS5 / GPSDR01). Jagged traces whose peaks keep changing are not listed.

The rule is applied to every site over 2–4 Oct 2026 only (the window
requested for this source). Earlier days in the combined Dhaka file
are ignored. Every hour is scored, including night and other off-peak
hours. A site is listed when at least three of those days have three
hours on the cap, or when the last day has two hours on the cap.
Each snapshot chart shows the same three days.
"""

from __future__ import annotations

import argparse
import subprocess
from collections import defaultdict
from datetime import time
from pathlib import Path

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

# Busy window used by the sample snaps: the line is flat through the day
# and only leaves the cap in the night valley.
BUSY_START = 8
BUSY_END = 22  # inclusive
NIGHT_START = 2
NIGHT_END = 6  # inclusive, the valley on the sample charts

# Flat-cap test on 2–4 Oct 2026 only (combined FEG_KPI_DHAKA_5Oct source).
# A day is on-cap when at least 3 hours (busy or off-peak) sit on the ceiling.
# A site is listed when that happens on at least 3 of those days.
# Separately, the last day is flagged when at least 2 hours sit on the ceiling.
KPI_YEAR = 2026
ANALYSIS_START = pd.Timestamp(f"{KPI_YEAR}-10-02")
ANALYSIS_END = pd.Timestamp(f"{KPI_YEAR}-10-04")
TOL_MBPS = 4.0
TOL_FRAC = 0.015
# Hard ceiling. A few hours may sit a little above the crowded level.
# 40 Mbit/s (or 12% of the level) still reads as a cap. If the 99th
# percentile climbs past that, the crowded level is a busy cluster with
# spikes (the DHGULN2 / DHDHN40 / DHBDD32 / DHDHN09 shape), not a choke.
OVERSHOOT_MBPS = 40.0
OVERSHOOT_FRAC = 0.12
# Second shape: the ceiling is the top of the trace. The level is searched
# only in the top 10% of the samples, and the line may not go above it at
# all (one glitch hour excepted). A trace that keeps climbing to new peaks
# has no ceiling and fails here even if a busy cluster exists lower down.
CEIL_TOP_Q = 0.90
CEIL_OVER_MBPS = 8.0
CEIL_OVER_FRAC = 0.03
CEIL_MAX_ABOVE = 1
CAP_CROWDED = "Crowded level"
CAP_CEILING = "Hard ceiling"
MIN_RUN_HOURS = 3
STUCK_DAY_HOURS = 3
MIN_DAYS_ON_CAP = 3
LAST_DAY_HOURS = 2
ANALYSIS_DAYS = int((ANALYSIS_END - ANALYSIS_START).days) + 1
STD_MBPS = 2.5
STD_FRAC = 0.01
MIN_CENTER_MBPS = 20.0

# Severity is hours on the cap as a share of the 2–4 Oct window.
# RAN-side Tx Total BW is shown as a column only; it is not used to judge.
SEV_SEVERE_PCT = 50.0
SEV_HIGH_PCT = 25.0
SEV_MODERATE_PCT = 10.0
SEV_SEVERE = "Severe"
SEV_HIGH = "High"
SEV_MODERATE = "Moderate"
SEV_LOW = "Low"
REPORT_TITLE = "Transmission Link Health Review"
SNAP_SHEET = "2. HourlyChartOfIssueSites"
GEO_SHEET = "5. GeoPlot"
GEO_TITLE = "Transmission Link Health Monitoring"
GEO_DATA_SHEET = "_GeoData"
GEO_FILE = "Physical_Site_Database_24Sep26.xlsx"
TECH_FILE = "Site list_Dhaka_Tech_5oct.xlsx"
TECH_4G = "4G"
TECH_4G5G = "4G+5G"
DARK_RED = "#8B0000"
GEO_ISSUE = "#FF1744"
GEO_4G = "#00E5FF"
GEO_5G = "#FFD600"

# Sites on the two sample charts. The report must contain both.
REFERENCE_SITES = ("DHAPT35", "DHAPT48")

RX_COL = "VS.FEGE.RxMaxSpeed(bit/s)"
TXBW_COL = "VS.FEGE.TxTotalBW(kbit/s)"

NAVY = "#1F4E79"
BLUE = "#2F5496"
RED = "#C00000"
GREY = "#595959"
LINE = "#BFBFBF"
ZEBRA = "#F7F9FB"
AMBER = "#FFF2CC"
AMBER_FONT = "#7F6000"
ORANGE = "#FCE4D6"
ORANGE_FONT = "#C65911"
GREEN = "#E2EFDA"
GREEN_FONT = "#375623"
PALE = "#D6E3F0"


def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, unit="D", origin="1899-12-30")


def parse_kpi_date(series: pd.Series) -> pd.Series:
    """Parse Excel serials or text dates such as 2-Oct / 02-Oct-26."""
    if pd.api.types.is_numeric_dtype(series):
        return excel_serial_to_datetime(series)
    text = series.astype(str).str.strip()
    numeric = pd.to_numeric(text, errors="coerce")
    if numeric.notna().all():
        return excel_serial_to_datetime(numeric)
    parsed = pd.to_datetime(text + f"-{KPI_YEAR}", format="%d-%b-%Y", errors="coerce")
    if parsed.isna().any():
        parsed = parsed.fillna(pd.to_datetime(text, errors="coerce", dayfirst=True))
    if parsed.isna().any():
        raise SystemExit("Could not parse the Date column")
    if (parsed.dt.year != KPI_YEAR).any():
        parsed = parsed.map(lambda ts: ts.replace(year=KPI_YEAR))
    return parsed


def parse_hour(series: pd.Series) -> pd.Series:
    """Parse 0:00 / 00:00 / 8:00 into an integer hour."""
    text = series.astype(str).str.strip()
    hour = text.str.split(":", n=1).str[0]
    values = pd.to_numeric(hour, errors="coerce")
    if values.isna().any():
        raise SystemExit("Could not parse the Time column")
    return values.astype(int)


def resolve_source(path: Path) -> Path:
    """Use a CSV, workbook, or the first volume of a split RAR as the KPI source."""
    name = path.name.lower()
    if name.endswith(".part2.rar"):
        path = path.with_name(path.name.replace(".part2.rar", ".part1.rar"))
    if path.suffix.lower() == ".rar":
        dest = Path("/tmp/dhaka_kpi")
        dest.mkdir(parents=True, exist_ok=True)
        subprocess.check_call(["unrar", "x", "-o+", str(path), f"{dest}/"])
        csvs = sorted(dest.glob("*.csv"))
        if not csvs:
            raise SystemExit(f"No CSV extracted from {path.name}")
        return csvs[0]
    return path


def load_kpi(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix == ".xlsb":
        return pd.read_excel(path, engine="pyxlsb")
    return pd.read_excel(path)


def find_plateau(rx: np.ndarray) -> tuple[float, float, np.ndarray]:
    """Return (center Mbit/s, tolerance Mbit/s, boolean mask on the cap).

    The center is the level in the upper half of the trace that captures
    the most samples inside the tolerance. It is then recentered on the
    median of those samples so a slightly off grid point does not split
    a flat top.
    """
    p95 = float(np.quantile(rx, 0.95))
    tol = max(TOL_MBPS, TOL_FRAC * p95)
    hi = float(np.max(rx))
    lo = float(np.quantile(rx, 0.55))
    if not np.isfinite(hi) or hi <= lo:
        band = np.ones(len(rx), dtype=bool)
        return hi, tol, band

    grid = np.arange(lo, hi + 1e-6, 0.5)
    diff = np.abs(rx[None, :] - grid[:, None])
    counts = np.count_nonzero(diff <= tol, axis=1)
    center0 = float(grid[int(np.argmax(counts))])
    band = np.abs(rx - center0) <= tol
    center = float(np.median(rx[band]))
    band = np.abs(rx - center) <= tol
    return center, tol, band


def find_ceiling(rx: np.ndarray) -> tuple[float, float, np.ndarray, int]:
    """Return (ceiling Mbit/s, tolerance, mask on the ceiling, hours above it).

    Same grid search as find_plateau, but only over the top 10% of the
    samples, so the level is the top the trace keeps returning to. The
    last value counts the hours that climb clearly above that top; a
    choked port has none.
    """
    p95 = float(np.quantile(rx, 0.95))
    tol = max(TOL_MBPS, TOL_FRAC * p95)
    hi = float(np.max(rx))
    lo = float(np.quantile(rx, CEIL_TOP_Q))
    if not np.isfinite(hi) or hi <= lo:
        center = hi
    else:
        grid = np.arange(lo, hi + 1e-6, 0.5)
        diff = np.abs(rx[None, :] - grid[:, None])
        counts = np.count_nonzero(diff <= tol, axis=1)
        center0 = float(grid[int(np.argmax(counts))])
        band = np.abs(rx - center0) <= tol
        center = float(np.median(rx[band]))
    band = np.abs(rx - center) <= tol
    over = max(CEIL_OVER_MBPS, CEIL_OVER_FRAC * center)
    above = int(np.count_nonzero(rx > center + over))
    return center, tol, band, above


def longest_run(ts: np.ndarray, band: np.ndarray) -> int:
    """Longest run of hourly samples that stay on the cap without a gap."""
    longest = 0
    current = 0
    previous = None
    for flag, stamp in zip(band, ts):
        if flag and previous is not None:
            gap = (stamp - previous) / np.timedelta64(1, "m")
            current = current + 1 if gap <= 90 else 1
        elif flag:
            current = 1
        else:
            current = 0
        if current > longest:
            longest = current
        previous = stamp if flag else None
    return int(longest)


def severity(hours_pct: float) -> str:
    """How hard the port is sitting on the cap, from hours-on-cap share."""
    if hours_pct >= SEV_SEVERE_PCT:
        return SEV_SEVERE
    if hours_pct >= SEV_HIGH_PCT:
        return SEV_HIGH
    if hours_pct >= SEV_MODERATE_PCT:
        return SEV_MODERATE
    return SEV_LOW


SEVERITY_ORDER = {SEV_SEVERE: 0, SEV_HIGH: 1, SEV_MODERATE: 2, SEV_LOW: 3}
SEVERITY_NOTES = (
    (
        SEV_SEVERE,
        f"Severe: the port sat on the cap at least {SEV_SEVERE_PCT:.0f}% of hours "
        f"in the {ANALYSIS_DAYS}-day window.",
    ),
    (
        SEV_HIGH,
        f"High: {SEV_HIGH_PCT:.0f}% to {SEV_SEVERE_PCT:.0f}% of hours on the cap. "
        "The choke is clear on several days, but the port is not locked all window.",
    ),
    (
        SEV_MODERATE,
        f"Moderate: {SEV_MODERATE_PCT:.0f}% to {SEV_HIGH_PCT:.0f}% of hours on the cap. "
        "Enough repeating hours to list the site, but the wall is not all-day.",
    ),
    (
        SEV_LOW,
        f"Low: under {SEV_MODERATE_PCT:.0f}% of hours on the cap. Usually a hard "
        "ceiling (clipped top) or a last-day flag. Still choked — just fewer hours.",
    ),
)
SNAP_DAYS = 3
MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def cap_window(busy_on: int, off_on: int) -> str:
    """Where the flat ceiling sits across the day."""
    if busy_on > 0 and off_on > 0:
        return "Busy + off-peak"
    if busy_on > 0:
        return "Busy hours"
    if off_on > 0:
        return "Off-peak"
    return "Partial"


def last_day_text(rec: dict) -> str:
    """Label for the last-day flag column. Always includes the date when flagged."""
    if not rec.get("last_flag"):
        return "—"
    return (
        f"{rec['last_day_label']} · {rec['last_on']}h "
        f"(busy {rec['last_busy']}, off-peak {rec['last_off']})"
    )


def listed_by_text(week_flag: bool, last_flag: bool) -> str:
    if week_flag and last_flag:
        return "≥3 days + last day"
    if week_flag:
        return "≥3 days"
    return "Last day"


def analyse(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """Return the hourly frame (with KPI columns) and the choked-site records."""
    work = df.copy()
    work["Date"] = parse_kpi_date(work["Date"])
    work["Rx"] = work[RX_COL] / 1000.0 / 1000.0
    work["TxBW"] = work[TXBW_COL] / 1000.0
    work["Hour"] = parse_hour(work["Time"])
    work["ts"] = work["Date"] + pd.to_timedelta(work["Hour"], unit="h")
    work = work.sort_values(["eNodeB Name", "ts"])
    work = work.drop_duplicates(["eNodeB Name", "ts"], keep="last")

    work = work[
        (work["Date"] >= ANALYSIS_START) & (work["Date"] <= ANALYSIS_END)
    ].copy()
    if work.empty:
        raise SystemExit(
            f"No rows in the requested window "
            f"{ANALYSIS_START.date()} – {ANALYSIS_END.date()}"
        )
    work = work.reset_index(drop=True)
    period_end = ANALYSIS_END.normalize()
    last_day = period_end.date()
    last_day_label = period_end.strftime("%-d-%b-%y")

    records: list[dict] = []
    on_cap = np.zeros(len(work), dtype=bool)
    stuck_level = np.full(len(work), np.nan)

    for site, idx in work.groupby("eNodeB Name", sort=False).groups.items():
        s = work.loc[idx]
        rx = s["Rx"].to_numpy(dtype=float)
        if len(rx) < 24 or float(np.nanmax(rx)) < MIN_CENTER_MBPS:
            continue
        hours = s["Hour"].to_numpy()
        busy = (hours >= BUSY_START) & (hours <= BUSY_END)
        offpeak = ~busy
        night = (hours >= NIGHT_START) & (hours <= NIGHT_END)
        days = s["Date"].dt.date.to_numpy()

        # Two ways to find the cap. The crowded level (sample-snap shape) is
        # tried first; if the trace is not crowded on one level, the top of
        # the trace is tried as a hard ceiling. The first one that passes
        # every test decides the site.
        candidates = []
        center, tol, band = find_plateau(rx)
        overshoot = float(np.quantile(rx, 0.99) - center)
        if center >= MIN_CENTER_MBPS and overshoot <= max(
            OVERSHOOT_MBPS, OVERSHOOT_FRAC * center
        ):
            candidates.append((CAP_CROWDED, center, tol, band))
        c_center, c_tol, c_band, c_above = find_ceiling(rx)
        if c_center >= MIN_CENTER_MBPS and c_above <= CEIL_MAX_ABOVE:
            candidates.append((CAP_CEILING, c_center, c_tol, c_band))

        chosen = None
        for cap_type, center, tol, band in candidates:
            in_band_std = float(np.std(rx[band])) if band.any() else 999.0
            if in_band_std > max(STD_MBPS, STD_FRAC * center):
                continue
            day_on: dict = defaultdict(int)
            day_busy: dict = defaultdict(int)
            day_off: dict = defaultdict(int)
            for day, flag, is_busy in zip(days, band, busy):
                if not flag:
                    continue
                day_on[day] += 1
                if is_busy:
                    day_busy[day] += 1
                else:
                    day_off[day] += 1
            days_on_cap = sum(v >= STUCK_DAY_HOURS for v in day_on.values())
            week_flag = days_on_cap >= MIN_DAYS_ON_CAP
            last_on = int(day_on.get(last_day, 0))
            last_flag = last_on >= LAST_DAY_HOURS
            if week_flag or last_flag:
                chosen = (cap_type, center, tol, band, in_band_std)
                break
        if chosen is None:
            continue
        cap_type, center, tol, band, in_band_std = chosen
        overshoot = float(np.quantile(rx, 0.99) - center)
        last_busy = int(day_busy.get(last_day, 0))
        last_off = int(day_off.get(last_day, 0))
        days_total = int(s["Date"].dt.normalize().nunique())

        hours_pct = 100.0 * float(band.mean())
        run = longest_run(s["ts"].to_numpy(), band)

        bw = float(np.median(s["TxBW"].to_numpy()))
        if bw <= 0:
            continue
        util = center / bw
        night_rx = float(np.median(rx[night])) if night.any() else float("nan")
        busy_on = int(band[busy].sum()) if busy.any() else 0
        off_on = int(band[offpeak].sum()) if offpeak.any() else 0
        busy_n = int(busy.sum())
        off_n = int(offpeak.sum())
        pos = work.index.get_indexer(s.index)
        on_cap[pos] = band
        stuck_level[pos] = center

        records.append(
            {
                "site": site,
                "severity": severity(hours_pct),
                "cap_type": cap_type,
                "bw": bw,
                "center": center,
                "tol": tol,
                "max_rx": float(rx.max()),
                "overshoot": overshoot,
                "in_band_std": in_band_std,
                "util": 100.0 * util,
                "hours_on": int(band.sum()),
                "hours_n": int(len(band)),
                "hours_pct": hours_pct,
                "busy_on": busy_on,
                "busy_n": busy_n,
                "busy_pct": 100.0 * busy_on / busy_n if busy_n else 0.0,
                "off_on": off_on,
                "off_n": off_n,
                "off_pct": 100.0 * off_on / off_n if off_n else 0.0,
                "cap_when": cap_window(busy_on, off_on),
                "all_pct": hours_pct,
                "days_on_cap": days_on_cap,
                "days_total": days_total,
                "week_flag": week_flag,
                "last_flag": last_flag,
                "last_day": last_day,
                "last_day_label": last_day_label,
                "last_on": last_on,
                "last_busy": last_busy,
                "last_off": last_off,
                "listed_by": listed_by_text(week_flag, last_flag),
                "longest": run,
                "night_rx": night_rx,
                "reference": site in REFERENCE_SITES,
                "n": int(len(rx)),
                "start": s["Date"].min(),
                "end": s["Date"].max(),
            }
        )

    work["OnCap"] = on_cap
    work["StuckLevel"] = stuck_level
    listed_order = {"≥3 days + last day": 0, "≥3 days": 1, "Last day": 2}
    records.sort(
        key=lambda r: (
            SEVERITY_ORDER[r["severity"]],
            listed_order[r["listed_by"]],
            -r["hours_pct"],
            r["site"],
        )
    )
    return work, records


def _require_references(records: list[dict]) -> None:
    found = {r["site"] for r in records}
    missing = [site for site in REFERENCE_SITES if site not in found]
    if missing:
        print(
            "Note: sample snap sites were not classified as choked: "
            + ", ".join(missing)
        )


def _styles(book):
    def fmt(**kwargs):
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        base.update(kwargs)
        return book.add_format(base)

    return {
        "title": fmt(font_size=20, bold=True, font_color=NAVY),
        "subtitle": fmt(font_size=11, font_color=GREY),
        "section": fmt(
            bold=True, font_size=12, font_color="white", bg_color=NAVY, text_wrap=True
        ),
        "body": fmt(font_size=11, text_wrap=True, valign="top"),
        "label": fmt(bold=True, font_color=NAVY, font_size=10),
        "meta": fmt(font_size=10, font_color="black"),
        "header": fmt(
            bold=True,
            font_size=10,
            font_color="white",
            bg_color=NAVY,
            align="center",
            text_wrap=True,
            border=1,
            border_color=NAVY,
        ),
        "text": fmt(border=1, border_color=LINE),
        "text_z": fmt(border=1, border_color=LINE, bg_color=ZEBRA),
        "int": fmt(border=1, border_color=LINE, align="center", num_format="0"),
        "int_z": fmt(
            border=1, border_color=LINE, align="center", num_format="0", bg_color=ZEBRA
        ),
        "num": fmt(border=1, border_color=LINE, align="center", num_format="0.00"),
        "num_z": fmt(
            border=1,
            border_color=LINE,
            align="center",
            num_format="0.00",
            bg_color=ZEBRA,
        ),
        "pct": fmt(border=1, border_color=LINE, align="center", num_format="0.0"),
        "pct_z": fmt(
            border=1, border_color=LINE, align="center", num_format="0.0", bg_color=ZEBRA
        ),
        "center": fmt(border=1, border_color=LINE, align="center"),
        "center_z": fmt(border=1, border_color=LINE, align="center", bg_color=ZEBRA),
        "link": fmt(
            border=1, border_color=LINE, font_color=BLUE, underline=1, bold=True
        ),
        "link_z": fmt(
            border=1,
            border_color=LINE,
            font_color=BLUE,
            underline=1,
            bold=True,
            bg_color=ZEBRA,
        ),
        "at": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=ORANGE,
            font_color=ORANGE_FONT,
        ),
        "at_z": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=ORANGE,
            font_color=ORANGE_FONT,
        ),
        "severe": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color="#F4CCCC",
            font_color="#990000",
        ),
        "high": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=ORANGE,
            font_color=ORANGE_FONT,
        ),
        "moderate": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=AMBER,
            font_color=AMBER_FONT,
        ),
        "low": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=GREEN,
            font_color=GREEN_FONT,
        ),
        "above": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color="#F8CBAD",
            font_color="#843C0C",
        ),
        "below": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=AMBER,
            font_color=AMBER_FONT,
        ),
        "yes": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=GREEN,
            font_color=GREEN_FONT,
        ),
        "yes_z": fmt(
            border=1,
            border_color=LINE,
            align="center",
            bold=True,
            bg_color=GREEN,
            font_color=GREEN_FONT,
        ),
        "banner": fmt(
            bold=True, font_size=16, font_color="white", bg_color=NAVY, valign="vcenter"
        ),
        "kpi_label": fmt(bold=True, font_size=9, font_color=GREY, bg_color=PALE),
        "kpi_value": fmt(bold=True, font_size=12, font_color=NAVY, bg_color=PALE),
        "note": fmt(font_size=9, font_color=GREY, italic=True, text_wrap=True),
        "method_head": fmt(
            bold=True, font_color="white", bg_color=NAVY, border=1, border_color=NAVY
        ),
        "method_key": fmt(bold=True, border=1, border_color=LINE, valign="top"),
        "method_val": fmt(border=1, border_color=LINE, valign="top", text_wrap=True),
        "date": fmt(border=1, border_color=LINE, num_format="dd-mmm-yyyy", align="center"),
        "date_z": fmt(
            border=1,
            border_color=LINE,
            num_format="dd-mmm-yyyy",
            align="center",
            bg_color=ZEBRA,
        ),
        "time": fmt(border=1, border_color=LINE, num_format="hh:mm", align="center"),
        "time_z": fmt(
            border=1,
            border_color=LINE,
            num_format="hh:mm",
            align="center",
            bg_color=ZEBRA,
        ),
        "formula": fmt(border=1, border_color=LINE, align="center", num_format="0.00"),
        "formula_z": fmt(
            border=1,
            border_color=LINE,
            align="center",
            num_format="0.00",
            bg_color=ZEBRA,
        ),
        "raw": fmt(border=1, border_color=LINE, align="center", num_format="#,##0"),
        "raw_z": fmt(
            border=1,
            border_color=LINE,
            align="center",
            num_format="#,##0",
            bg_color=ZEBRA,
        ),
    }


def _severity_format(styles, level: str, zebra: bool):
    if level == SEV_SEVERE:
        return styles["severe"]
    if level == SEV_HIGH:
        return styles["high"]
    if level == SEV_MODERATE:
        return styles["moderate"]
    return styles["low"]


def _page(ws, title: str, *, fit_width: bool = True):
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.45, right=0.45, top=0.6, bottom=0.5)
    # Snapshots must not be scaled onto the page width: that crushes the
    # two-row date/time axis into one overlapping band.
    if fit_width:
        ws.fit_to_pages(1, 0)
    ws.set_header(f"&L&8&K1F4E79{title}&R&8{REPORT_TITLE}  |  DHK")
    ws.set_footer("&L&8Choked = RxMaxSpeed stuck flat&R&8Page &P of &N")
    ws.hide_gridlines(2)
    ws.set_tab_color(NAVY)



def _chart_y_scale(values: list[float]) -> tuple[int, int]:
    """Pick Y max and tick size so a low stuck level is not lost on a 0–100 axis.

    Example: DHDRS99 at 33.31 Mbit/s uses 0–50 with ticks of 10, not 0–100 / 50.
    """
    finite = [float(v) for v in values if np.isfinite(v)]
    peak = max(finite) if finite else 50.0
    target = max(peak * 1.15, peak + 4.0)
    brackets = (
        (25, 5),
        (50, 10),
        (75, 25),
        (100, 25),
        (125, 25),
        (150, 25),
        (200, 50),
        (250, 50),
        (300, 50),
        (350, 50),
        (400, 50),
        (500, 100),
        (600, 100),
        (800, 100),
        (1000, 200),
        (1500, 250),
        (2000, 250),
    )
    for y_max, unit in brackets:
        if target <= y_max:
            return y_max, unit
    y_max = int(np.ceil(target / 200.0) * 200)
    return max(y_max, 200), 200


def _write_snapshots(book, styles, work, records, period_txt, chart_start, chart_end) -> dict[str, int]:
    ws = book.add_worksheet(SNAP_SHEET)
    _page(ws, "Hourly chart of issue sites", fit_width=False)
    ws.set_zoom(100)
    ws.set_print_scale(100)
    ws.set_tab_color(BLUE)
    data = book.add_worksheet("_ChartData")
    data.hide()

    # The chart is about one screen wide, so it does not need a long sheet.
    for col, width in enumerate([18, 16, 18, 16, 18, 16, 18, 16]):
        ws.set_column(col, col, width)
    ws.set_column(8, 16, 11)

    anchors: dict[str, int] = {}
    # Fixed stride so the site list can link to row i * BLOCK_ROWS + 1.
    block = BLOCK_ROWS

    cat_fmt = book.add_format({"font_name": "Calibri", "font_size": 8, "align": "center"})

    for i, rec in enumerate(records):
        top = i * block
        anchors[rec["site"]] = top
        site = rec["site"]
        s = work.loc[work["eNodeB Name"] == site].sort_values("ts")
        s = s[(s["Date"] >= chart_start) & (s["Date"] <= chart_end)]

        # Hourly line, compact 2-level category axis matching the target snap.
        # Outer level: Date shown only once per day (blank for hours 1..23).
        # Inner level: All 24 hourly timestamps (00:00 to 23:00) rotated vertically.
        c0 = i * 5
        data.write(0, c0, "Date")
        data.write(0, c0 + 1, "Time")
        data.write(0, c0 + 2, "Rx")
        data.write(0, c0 + 3, "Peak")
        data.write(0, c0 + 4, "AvgPeak")
        rx_by_key = {
            (row["Date"].normalize(), int(row["Hour"])): float(row["Rx"])
            for _, row in s.iterrows()
        }
        chart_rx = [v for v in rx_by_key.values() if np.isfinite(v)]
        peak_72 = float(max(chart_rx)) if chart_rx else float("nan")
        # Typical high: mean of each day's peak (not 5% below the single max).
        daily_peaks = []
        for day in pd.date_range(chart_start, chart_end, freq="D"):
            day_vals = [
                rx_by_key[(day.normalize(), hour)]
                for hour in range(24)
                if (day.normalize(), hour) in rx_by_key
                and np.isfinite(rx_by_key[(day.normalize(), hour)])
            ]
            if day_vals:
                daily_peaks.append(max(day_vals))
        avg_peak = float(sum(daily_peaks) / len(daily_peaks)) if daily_peaks else float("nan")
        dates: list[str] = []
        times: list[str] = []
        plotted: list[float] = [v for v in (peak_72, avg_peak) if np.isfinite(v)]
        r = 0
        for day in pd.date_range(chart_start, chart_end, freq="D"):
            date_label = f"{int(day.day)}/{MONTHS[int(day.month) - 1]}"
            for hour in range(24):
                r += 1
                hour_label = f"{hour:02d}:00"
                # Write date only on the first hour (hour 00:00) of each day.
                # In Excel, leaving subsequent hours blank merges/groups the outer date
                # category across the 24 hours of that day in a clean box.
                d_val = date_label if hour == 0 else ""
                dates.append(d_val)
                times.append(hour_label)
                if hour == 0:
                    data.write_string(r, c0, date_label, cat_fmt)
                else:
                    data.write_blank(r, c0, None)
                data.write_string(r, c0 + 1, hour_label, cat_fmt)
                value = rx_by_key.get((day.normalize(), hour))
                if value is None:
                    data.write_blank(r, c0 + 2, None)
                else:
                    data.write_number(r, c0 + 2, value)
                    plotted.append(value)
                if np.isfinite(peak_72):
                    data.write_number(r, c0 + 3, peak_72)
                if np.isfinite(avg_peak):
                    data.write_number(r, c0 + 4, avg_peak)
        n = r

        ws.set_row(top, 26)
        banner = (
            f"{i + 1:02d}    {site}    ·    {rec['severity']}"
            + ("    ·    sample snap" if rec["reference"] else "")
        )
        ws.merge_range(top, 0, top, 7, banner, styles["banner"])

        labels = [
            (0, "Tx Total BW (Mbit/s)", f"{rec['bw']:.2f}"),
            (2, "Stuck RxMaxSpeed (Mbit/s)", f"{rec['center']:.2f}"),
            (4, "Hours on cap (all 24h)", f"{rec['hours_on']} / {rec['hours_n']}  ({rec['hours_pct']:.1f}%)"),
            (6, "Days on cap (≥3h)", f"{rec['days_on_cap']} / {rec['days_total']}"),
        ]
        ws.set_row(top + 1, 16)
        ws.set_row(top + 2, 20)
        for col, label, _value in labels:
            ws.merge_range(top + 1, col, top + 1, col + 1, label, styles["kpi_label"])
        for col, _label, value in labels:
            ws.merge_range(top + 2, col, top + 2, col + 1, value, styles["kpi_value"])

        ws.set_row(top + 3, 8)

        if n == 0:
            ws.merge_range(
                top + 5,
                0,
                top + 5,
                7,
                "No hourly samples in the latest 3 days.",
                styles["note"],
            )
            for r in range(top + 4, top + block):
                ws.set_row(r, 15)
            continue

        chart = book.add_chart({"type": "line"})
        chart.show_blanks_as("gap")
        chart.set_title(
            {
                "name": site,
                "name_font": {"name": "Calibri", "size": 16, "color": "#595959"},
            }
        )
        chart.add_series(
            {
                "name": "Sum of RxMaxSpeed,Mbps",
                "categories": ["_ChartData", 1, c0, n, c0 + 1],
                "categories_data": [dates, times],
                "values": ["_ChartData", 1, c0 + 2, n, c0 + 2],
                "line": {"color": "#5B9BD5", "width": 2.25},
                "marker": {"type": "none"},
            }
        )
        if np.isfinite(peak_72):
            chart.add_series(
                {
                    "name": f"Peak ({peak_72:.2f} Mbit/s)",
                    "categories": ["_ChartData", 1, c0, n, c0 + 1],
                    "categories_data": [dates, times],
                    "values": ["_ChartData", 1, c0 + 3, n, c0 + 3],
                    "line": {"color": "#C00000", "width": 1.5, "dash_type": "dash"},
                    "marker": {"type": "none"},
                }
            )
        if np.isfinite(avg_peak):
            chart.add_series(
                {
                    "name": f"Avg peak ({avg_peak:.2f} Mbit/s)",
                    "categories": ["_ChartData", 1, c0, n, c0 + 1],
                    "categories_data": [dates, times],
                    "values": ["_ChartData", 1, c0 + 4, n, c0 + 4],
                    "line": {"color": "#ED7D31", "width": 1.5, "dash_type": "dash"},
                    "marker": {"type": "none"},
                }
            )
        chart.set_x_axis(
            {
                "num_font": {
                    "name": "Calibri",
                    "size": 8,
                    "color": "#595959",
                    "rotation": -90,
                },
                "label_position": "nextTo",
                "label_align": "center",
                "position_axis": "on_tick",
                "major_tick_mark": "none",
                "minor_tick_mark": "none",
                "major_gridlines": {"visible": False},
                "minor_gridlines": {"visible": False},
                "line": {"color": "#B0B0B0"},
            }
        )
        y_max, y_unit = _chart_y_scale(plotted)
        chart.set_y_axis(
            {
                "num_font": {"name": "Calibri", "size": 9, "color": "#595959"},
                "min": 0,
                "max": y_max,
                "major_unit": y_unit,
                "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}},
                "line": {"none": True},
            }
        )
        chart.set_legend({"position": "top", "font": {"name": "Calibri", "size": 9}})
        chart.set_chartarea({"border": {"color": "#D0D0D0"}, "fill": {"color": "white"}})
        # Compact, precise chart layout matching the sample snap.
        chart.set_plotarea(
            {
                "border": {"none": True},
                "fill": {"color": "white"},
                "layout": {"x": 0.06, "y": 0.16, "width": 0.91, "height": 0.58},
            }
        )
        chart.set_size({"width": 1200, "height": 410})
        # Don't let the chart shrink when the sheet is scaled or columns move.
        ws.insert_chart(
            top + 4,
            0,
            chart,
            {"x_offset": 6, "y_offset": 6, "object_position": 2},
        )

        for r in range(top + 4, top + block):
            ws.set_row(r, 16)

    if records:
        ws.set_h_pagebreaks([i * block for i in range(1, len(records))])
    return anchors




def _write_hourly(book, styles, work, records):
    ws = book.add_worksheet("3. Hourly KPI")
    _page(ws, "Hourly FEGE KPI — choked sites")
    ws.set_tab_color("#5B9BD5")
    headers = [
        "Date",
        "Time",
        "eNodeB Name",
        "VS.FEGE.RxMaxSpeed(bit/s)",
        "RxMaxSpeed (Mbit/s)",
        "VS.FEGE.TxTotalBW(kbit/s)",
        "Tx Total BW (Mbit/s)",
        "On cap",
        "Stuck level (Mbit/s)",
    ]
    widths = [14, 10, 16, 30, 22, 30, 22, 12, 20]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    ws.set_row(0, 30)
    for col, text in enumerate(headers):
        ws.write(0, col, text, styles["header"])

    sites = [r["site"] for r in records]
    hourly = work.loc[work["eNodeB Name"].isin(sites)].sort_values(
        ["eNodeB Name", "ts"]
    )
    # Zebra by site, not by row, so a site reads as one block.
    site_parity = {site: i % 2 for i, site in enumerate(sites)}

    for excel_row, (_, row) in enumerate(hourly.iterrows(), start=1):
        zebra = site_parity[row["eNodeB Name"]] == 1
        ws.write_datetime(
            excel_row,
            0,
            row["Date"].to_pydatetime(),
            styles["date_z"] if zebra else styles["date"],
        )
        hh, mm = (int(p) for p in str(row["Time"]).split(":")[:2])
        ws.write_datetime(
            excel_row,
            1,
            time(hh, mm),
            styles["time_z"] if zebra else styles["time"],
        )
        ws.write_string(
            excel_row, 2, row["eNodeB Name"], styles["text_z"] if zebra else styles["text"]
        )
        ws.write_number(
            excel_row, 3, float(row[RX_COL]), styles["raw_z"] if zebra else styles["raw"]
        )
        # Excel rows are 1-based; xlsxwriter row 1 is Excel row 2.
        formula_row = excel_row + 1
        ws.write_formula(
            excel_row,
            4,
            f"=D{formula_row}/1000/1000",
            styles["formula_z"] if zebra else styles["formula"],
        )
        ws.write_number(
            excel_row, 5, float(row[TXBW_COL]), styles["raw_z"] if zebra else styles["raw"]
        )
        ws.write_formula(
            excel_row,
            6,
            f"=F{formula_row}/1000",
            styles["formula_z"] if zebra else styles["formula"],
        )
        ws.write_string(
            excel_row,
            7,
            "Yes" if bool(row["OnCap"]) else "No",
            styles["center_z"] if zebra else styles["center"],
        )
        ws.write_number(
            excel_row,
            8,
            float(row["StuckLevel"]),
            styles["num_z"] if zebra else styles["num"],
        )

    last = len(hourly)
    ws.autofilter(0, 0, max(last, 1), len(headers) - 1)
    ws.repeat_rows(0, 0)
    ws.set_zoom(110)


def _write_method(book, styles, source_name, period_txt, n_sites, records):
    ws = book.add_worksheet("4. Method")
    _page(ws, "How a site was judged choked")
    ws.set_tab_color(GREY)
    ws.set_column(0, 0, 36)
    ws.set_column(1, 1, 88)
    ws.set_row(0, 28)
    ws.merge_range("A1:B1", "How RxMaxSpeed was judged stuck", styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:B2",
        f"{source_name}    ·    {period_txt}    ·    {n_sites} sites, "
        f"{len(records)} listed",
        styles["subtitle"],
    )

    ws.set_row(3, 20)
    ws.merge_range("A4:B4", "Counters", styles["section"])
    rows = [
        (
            "RxMaxSpeed (Mbit/s)",
            "VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000. "
            "This is the trace on the sample charts. A flat top means the maximum "
            "rate is no longer moving: the transmission path is limiting it.",
        ),
        (
            "Tx Total BW (Mbit/s)",
            "VS.FEGE.TxTotalBW(kbit/s) / 1000. "
            "Configured FEGE transmit bandwidth for the same hour. "
            "The site value in the list is the median across the period.",
        ),
    ]
    for i, (key, val) in enumerate(rows):
        r = 4 + i
        ws.set_row(r, 48)
        ws.write(r, 0, key, styles["method_key"])
        ws.write(r, 1, val, styles["method_val"])

    ws.set_row(7, 20)
    ws.merge_range("A8:B8", "What the sample snaps show", styles["section"])
    ws.set_row(8, 72)
    ws.merge_range(
        "A9:B9",
        "DHAPT35 and DHAPT48 were used as the pattern, not as a filter by name. "
        "On both charts RxMaxSpeed falls in the early morning and then locks to "
        "one level from about 08:00 through the evening (about 270 Mbit/s on "
        "DHAPT35, about 241 Mbit/s on DHAPT48). Hours on that level differ by "
        "only 1–3 Mbit/s, and the line never rises above it. A normal site in "
        "this file reaches a different maximum every hour, often tens of Mbit/s "
        "apart, so it is not listed.",
        styles["method_val"],
    )

    ws.set_row(10, 20)
    ws.merge_range("A11:B11", "Rule applied to every site", styles["section"])
    rules = [
        (
            f"Cap shape 1 — {CAP_CROWDED}",
            "The level in the upper half of the trace that contains the most "
            f"hourly samples inside ±{TOL_MBPS:.1f} Mbit/s "
            f"(or ±{TOL_FRAC:.0%} of the 95th percentile, when that is wider). "
            "The level is the median of those samples. The 99th percentile of "
            f"RxMaxSpeed may sit at most {OVERSHOOT_MBPS:.0f} Mbit/s above it "
            f"(or {OVERSHOOT_FRAC:.0%} of the level, when that is wider). "
            "This is the DHAPT35 / DHAPT48 shape: the line locks to one level.",
        ),
        (
            f"Cap shape 2 — {CAP_CEILING}",
            "Tried when shape 1 does not fit. The same search is run only over "
            f"the top {1 - CEIL_TOP_Q:.0%} of the samples, so the level is the top "
            "the trace keeps returning to. The line may not go above that top "
            f"by more than {CEIL_OVER_MBPS:.0f} Mbit/s (or {CEIL_OVER_FRAC:.0%}); "
            f"at most {CEIL_MAX_ABOVE} glitch hour is excused. This catches ports "
            "that move up and down during the day but are clipped at one value "
            "every time they rise (the ~500 Mbit/s group such as DHSVRS5, "
            "GPSDR01, DHKKT87, and DHKKTE2 whose cap moved from 270 to 352). "
            "Cap shape is shown on the site list.",
        ),
        (
            "Not listed (outliers)",
            "A trace whose peaks keep changing — spikes of tens to hundreds of "
            "Mbit/s above the crowded level and no fixed top (DHGULN2, DHDHN40, "
            "DHBDD32, DHDHN09) — fails both shapes. It is busy, not choked, and is "
            "not listed.",
        ),
        (
            "Analysis window",
            f"Only {ANALYSIS_START.strftime('%-d %b %Y')} – "
            f"{ANALYSIS_END.strftime('%-d %b %Y')} ({ANALYSIS_DAYS} days) is scored. "
            "Earlier days in the combined Dhaka source (28 Sep–1 Oct) are ignored.",
        ),
        (
            "Hours on the cap (all 24h)",
            "Every hour is scored — busy hours "
            f"({BUSY_START:02d}:00–{BUSY_END:02d}:00) and off-peak / night hours. "
            "A site that is only flat at night is still listed.",
        ),
        (
            "Repeats across days",
            f"A day counts as on-cap when at least {STUCK_DAY_HOURS} hours "
            "(any hour of the day, busy or off-peak) sit on the stuck level. "
            f"A site is listed when this happens on at least {MIN_DAYS_ON_CAP} of the "
            f"{ANALYSIS_DAYS} days.",
        ),
        (
            "Last day flag",
            f"Independently, the last day of the window is flagged when at least "
            f"{LAST_DAY_HOURS} hours that day (busy and/or off-peak) sit on the "
            "stuck level. The date is written in the Last day cap column. "
            "A site can be listed for the last-day flag even if it has fewer than "
            f"{MIN_DAYS_ON_CAP} on-cap days.",
        ),
        (
            "Continuous flat run",
            f"Longest consecutive run on the cap is reported. A day still counts "
            f"when {STUCK_DAY_HOURS} hours sit on the level even if they are not "
            "adjacent.",
        ),
        (
            "Tightness",
            f"Samples on the cap have a standard deviation of at most {STD_MBPS:.1f} Mbit/s "
            f"(or {STD_FRAC:.0%} of the stuck level, when that is wider).",
        ),
        (
            "Severity",
            f"Severe ≥ {SEV_SEVERE_PCT:.0f}% of hours on the cap; "
            f"High {SEV_HIGH_PCT:.0f}–{SEV_SEVERE_PCT:.0f}%; "
            f"Moderate {SEV_MODERATE_PCT:.0f}–{SEV_HIGH_PCT:.0f}%; "
            f"Low under {SEV_MODERATE_PCT:.0f}%. "
            "This is the share of hours sitting on the wall, not a comparison "
            "with the RAN Tx Total BW counter.",
        ),
        (
            "Tx Total BW column",
            "VS.FEGE.TxTotalBW is a RAN-side counter, not the transmission "
            "engineered bandwidth. It is shown for reference only and is not "
            "used to list or grade a site.",
        ),
        (
            "Night Rx",
            f"Median RxMaxSpeed from {NIGHT_START:02d}:00 to {NIGHT_END:02d}:00. "
            "Shown so the night dip on the sample charts can be compared. "
            "A site that stays on the cap through the night is still listed: "
            "that link is full all day.",
        ),
    ]
    for i, (key, val) in enumerate(rules):
        r = 11 + i
        ws.set_row(r, 48)
        ws.write(r, 0, key, styles["method_key"])
        ws.write(r, 1, val, styles["method_val"])

    foot = 11 + len(rules) + 1
    ws.set_row(foot, 20)
    ws.merge_range(foot, 0, foot, 1, "Check against the sample snaps", styles["section"])
    ws.set_row(foot + 1, 36)
    ws.merge_range(
        foot + 1,
        0,
        foot + 1,
        1,
        "DHAPT35 and DHAPT48 are the sample-snap pattern. Scoring uses "
        f"{ANALYSIS_START.strftime('%-d-%b-%y')} to "
        f"{ANALYSIS_END.strftime('%-d-%b-%y')} and every hour. "
        "Each snapshot is a clean chart of those three days: "
        "hourly timestamps 00:00 to 23:00, "
        "and the date is shown once under that day. The chart title is the site name.",
        styles["method_val"],
    )

    last_day_hdr = records[0]["last_day_label"] if records else "last day"
    kept_head = foot + 3
    ws.set_row(kept_head, 20)
    ws.merge_range(
        kept_head,
        0,
        kept_head,
        1,
        "Notes kept from Dashboard and Site List",
        styles["section"],
    )
    kept = [
        (
            "Dashboard listing note",
            f"Scored on {period_txt} (latest {ANALYSIS_DAYS} days). "
            f"Listed when ≥{MIN_DAYS_ON_CAP} days have ≥{STUCK_DAY_HOURS} hours on the cap "
            f"(busy and off-peak), or the last day has ≥{LAST_DAY_HOURS} hours on the cap. "
            "Two cap shapes count: a crowded flat level (DHAPT35 shape) or a hard ceiling "
            "the trace is clipped at and never rises above. Jagged traces with changing "
            "peaks are not listed. Each hourly chart is the latest 3 days only. "
            "Open a site name to see it.",
        ),
        (
            "Site List counters",
            "RxMaxSpeed (Mbit/s)  =  VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000"
            "          Tx Total BW (Mbit/s)  =  VS.FEGE.TxTotalBW(kbit/s) / 1000  "
            f"(RAN-side, reference only)          Source: {source_name}",
        ),
        (
            "Site List scoring note",
            f"Scored on {period_txt} "
            f"({ANALYSIS_START.strftime('%-d-%b-%y')} to "
            f"{ANALYSIS_END.strftime('%-d-%b-%y')}). "
            f"Listed when ≥{MIN_DAYS_ON_CAP} of {ANALYSIS_DAYS} days have ≥{STUCK_DAY_HOURS} hours on the cap "
            f"(busy 08:00–22:00 and off-peak / night). "
            f"Last day cap column flags {last_day_hdr} when ≥{LAST_DAY_HOURS} hours that day "
            "sit on the cap. "
            f"Severity is hours on the cap: Severe ≥{SEV_SEVERE_PCT:.0f}%, "
            f"High ≥{SEV_HIGH_PCT:.0f}%, Moderate ≥{SEV_MODERATE_PCT:.0f}%, Low below that. "
            f"Cap shape: '{CAP_CROWDED}' or '{CAP_CEILING}'. "
            "The hourly chart is the latest 3 days only. Open a site name to jump to it.",
        ),
    ]
    for i, (key, val) in enumerate(kept):
        r = kept_head + 1 + i
        ws.set_row(r, 56)
        ws.write(r, 0, key, styles["method_key"])
        ws.write(r, 1, val, styles["method_val"])



def load_site_geo(path: Path) -> pd.DataFrame:
    """Read the physical site database (Site Code, Latitude, Longitude)."""
    raw = pd.read_excel(path)
    cols = {str(c).strip().lower(): c for c in raw.columns}

    def pick(*names):
        for name in names:
            if name in cols:
                return cols[name]
        raise SystemExit(
            f"{path.name} is missing a required column. Need one of: {', '.join(names)}"
        )

    site_col = pick("site code", "sitecode", "site", "enodeb name")
    lat_col = pick("latitude", "lat")
    lon_col = pick("longitude", "lon", "long")
    district_col = cols.get("district")
    region_col = cols.get("region")
    tech_col = cols.get("tech")
    geo = pd.DataFrame(
        {
            "site": raw[site_col].astype(str).str.strip(),
            "lat": pd.to_numeric(raw[lat_col], errors="coerce"),
            "lon": pd.to_numeric(raw[lon_col], errors="coerce"),
            "district": raw[district_col].astype(str) if district_col else "",
            "region": raw[region_col].astype(str) if region_col else "",
            "tech": raw[tech_col].astype(str).str.strip() if tech_col else "",
        }
    )
    geo = geo.dropna(subset=["lat", "lon"])
    geo = geo[geo["site"] != ""]
    return geo.drop_duplicates("site", keep="last")


def _find_input(name: str, source: Path | None = None) -> Path | None:
    candidates = []
    if source is not None:
        candidates.append(source.parent / name)
    candidates.extend((Path(name), Path("/workspace") / name))
    return next((path for path in candidates if path.exists()), None)


def _norm_list_tech(value: object) -> str:
    key = str(value or "").strip().upper().replace(" ", "")
    if key in ("4G+5G", "4G5G"):
        return TECH_4G5G
    if key == "4G":
        return TECH_4G
    return str(value or "").strip()


def load_site_tech(path: Path) -> dict[str, str]:
    """Site Name / Tech list (4G vs 4G+5G) used only for the Dashboard summary."""
    raw = pd.read_excel(path)
    cols = {str(c).strip().lower(): c for c in raw.columns}
    site_col = cols.get("site name") or cols.get("sitename") or cols.get("site")
    tech_col = cols.get("tech")
    if site_col is None or tech_col is None:
        raise SystemExit(f"{path.name} needs Site Name and Tech columns")
    out: dict[str, str] = {}
    for site, tech in zip(raw[site_col], raw[tech_col]):
        name = str(site).strip()
        if name and name.lower() != "nan":
            out[name] = _norm_list_tech(tech)
    return out


def summarise_tech(records: list[dict], tech_by_site: dict[str, str]) -> dict[str, int]:
    checked_4g = sum(1 for tech in tech_by_site.values() if tech == TECH_4G)
    checked_4g5g = sum(1 for tech in tech_by_site.values() if tech == TECH_4G5G)
    choked_4g = 0
    choked_4g5g = 0
    unmapped = 0
    for rec in records:
        tech = tech_by_site.get(rec["site"])
        if tech == TECH_4G:
            choked_4g += 1
        elif tech == TECH_4G5G:
            choked_4g5g += 1
        else:
            unmapped += 1
    return {
        "checked_4g": checked_4g,
        "checked_4g5g": checked_4g5g,
        "choked_4g": choked_4g,
        "choked_4g5g": choked_4g5g,
        "unmapped": unmapped,
    }


CAP_WHEN_COLS = ("Busy + off-peak", "Busy hours", "Off-peak")


def cap_when_matrix(records: list[dict], tech_by_site: dict[str, str]) -> dict:
    """Severity × cap-window counts for all sites, 4G+5G, and 4G."""
    sevs = (SEV_SEVERE, SEV_HIGH, SEV_MODERATE, SEV_LOW)
    groups = ("total", TECH_4G5G, TECH_4G)
    matrix = {
        group: {sev: {when: 0 for when in CAP_WHEN_COLS} for sev in sevs}
        for group in groups
    }
    for rec in records:
        sev = rec["severity"]
        when = rec["cap_when"]
        if sev not in matrix["total"] or when not in matrix["total"][sev]:
            continue
        matrix["total"][sev][when] += 1
        tech = tech_by_site.get(rec["site"])
        if tech in (TECH_4G, TECH_4G5G):
            matrix[tech][sev][when] += 1
    return matrix


def _write_geoplot(
    book,
    styles,
    work,
    records,
    geo: pd.DataFrame,
    output: Path,
    tech_by_site: dict[str, str] | None = None,
) -> None:
    """Google-satellite GeoPlot: 4G / 5G coverage plus 4G / 4G+5G issue legends."""
    from geoplot_satellite import (
        KIND_ISSUE_4G,
        KIND_ISSUE_4G5G,
        classify_rows,
        cluster_bounds,
        cluster_issue_sites,
        cluster_label,
        render_html,
        render_map_images,
        thana_focus_specs,
    )

    ws = book.add_worksheet(GEO_SHEET)
    _page(ws, "GeoPlot — 4G / 5G / issue sites")
    ws.set_tab_color(DARK_RED)
    ws.hide_gridlines(2)

    data_ws = book.add_worksheet(GEO_DATA_SHEET)
    data_ws.hide()

    issue_order = [rec["site"] for rec in records]
    issue_set = set(issue_order)
    sev_by_site = {rec["site"]: rec["severity"] for rec in records}
    by_code = geo.set_index("site")

    def lookup(site: str) -> dict | None:
        if site not in by_code.index:
            return None
        row = by_code.loc[site]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        return {
            "site": site,
            "lat": float(row["lat"]),
            "lon": float(row["lon"]),
            "district": str(row["district"]),
            "region": str(row["region"]),
            "tech": str(row.get("tech", "")),
            "severity": sev_by_site.get(site, "—"),
        }

    issue_rows = [r for r in (lookup(site) for site in issue_order) if r]
    other_rows = []
    for row in geo.itertuples(index=False):
        site = str(row.site)
        if site in issue_set:
            continue
        other_rows.append(
            {
                "site": site,
                "lat": float(row.lat),
                "lon": float(row.lon),
                "district": str(row.district),
                "region": str(row.region),
                "tech": str(getattr(row, "tech", "")),
                "severity": "—",
            }
        )

    headers = ["Site", "Latitude", "Longitude", "Status", "Tech", "Issue Tech", "District", "Region", "Severity"]
    for col, text in enumerate(headers):
        data_ws.write_string(0, col, text)
    excel_row = 1
    for row in other_rows + issue_rows:
        status = "Issue" if row["site"] in issue_set else "Non-issue"
        listed = (tech_by_site or {}).get(row["site"], "")
        data_ws.write_string(excel_row, 0, row["site"])
        data_ws.write_number(excel_row, 1, row["lat"])
        data_ws.write_number(excel_row, 2, row["lon"])
        data_ws.write_string(excel_row, 3, status)
        data_ws.write_string(excel_row, 4, row["tech"])
        data_ws.write_string(excel_row, 5, listed if status == "Issue" else "")
        data_ws.write_string(excel_row, 6, row["district"])
        data_ws.write_string(excel_row, 7, row["region"])
        data_ws.write_string(excel_row, 8, row["severity"])
        excel_row += 1

    classified = classify_rows(other_rows + issue_rows, issue_set, tech_by_site)
    kind_by_site = {row["site"]: row["kind"] for row in classified}
    n4 = sum(1 for r in classified if r["tech"] == "4G")
    n5 = sum(1 for r in classified if r["tech"] == "5G")
    n_issue_4g = sum(1 for r in classified if r["kind"] == KIND_ISSUE_4G)
    n_issue_4g5g = sum(1 for r in classified if r["kind"] == KIND_ISSUE_4G5G)
    n_issue = n_issue_4g + n_issue_4g5g

    zoom_specs = thana_focus_specs(classified)
    for index, zoom_issues in enumerate(cluster_issue_sites(issue_rows), start=1):
        lat_min, lat_max, lon_min, lon_max = cluster_bounds(zoom_issues)
        issue_names = {row["site"] for row in zoom_issues}
        zoom_rows = []
        for row in classified:
            if not (lat_min <= row["lat"] <= lat_max and lon_min <= row["lon"] <= lon_max):
                continue
            # Keep this cluster's issue-legend counts exact: neighbouring
            # issue sites stay as 4G/5G coverage, not extra issue dots.
            if row["kind"] in (KIND_ISSUE_4G, KIND_ISSUE_4G5G) and row["site"] not in issue_names:
                zoom_rows.append({**row, "kind": row["tech"] if row["tech"] in ("4G", "5G") else "4G"})
            else:
                zoom_rows.append(row)
        label = cluster_label(zoom_issues, index)
        issue_4g = sum(1 for row in zoom_issues if kind_by_site.get(row["site"]) == KIND_ISSUE_4G)
        issue_4g5g = sum(1 for row in zoom_issues if kind_by_site.get(row["site"]) == KIND_ISSUE_4G5G)
        zoom_specs.append(
            {
                "label": label,
                "slug": f"Cluster_{index:02d}",
                "issue_count": len(zoom_issues),
                "issue_4g_count": issue_4g,
                "issue_4g5g_count": issue_4g5g,
                "issue_sites": [row["site"] for row in zoom_issues],
                "rows": zoom_rows,
                "lat_min": lat_min,
                "lat_max": lat_max,
                "lon_min": lon_min,
                "lon_max": lon_max,
            }
        )

    html_path = output.with_name(output.stem + "_GeoPlot.html")
    maps = render_map_images(classified, zoom_specs, output.parent, output.stem + "_GeoPlot")
    render_html(classified, zoom_specs, html_path, GEO_TITLE)

    ws.set_column(0, 18, 12)
    ws.set_row(0, 28)
    ws.merge_range("A1:R1", GEO_TITLE, styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:R2",
        f"Google satellite    ·    Physical sites: {len(classified)}"
        f"    ·    4G: {n4}    ·    5G: {n5}"
        f"    ·    Issue 4G: {n_issue_4g}    ·    Issue 4G+5G: {n_issue_4g5g}"
        f"    ·    Source: {GEO_FILE} + {TECH_FILE}",
        styles["subtitle"],
    )
    ws.set_row(2, 20)
    ws.merge_range(
        "A3:R3",
        "National map (left) is the 4G / 5G / issue-tech overview. The map beside it "
        "colours the same issue sites by severity (Severe / High / Moderate / Low). "
        "Then Gulshan, Banani, Dhanmondi, and Tejgaon thana maps (high 5G footprint). "
        "Issue clusters after that are built dynamically until every issue site is "
        f"covered ({sum(1 for z in zoom_specs if z.get('focus') != 'thana')} maps). "
        "Urban clusters come first (Dhaka Metro, Gazipur, surrounding), then maps with "
        "more issue sites. Coverage: Cyan = 4G, Yellow = 5G. Issue sites: Red = 4G, "
        "Orange = 4G+5G. Open the HTML file for interactive maps.",
        styles["note"],
    )
    row = 5
    paired = {slot: (title, path) for title, path, slot in maps if slot in ("national", "severity")}
    stacked = [(title, path) for title, path, slot in maps if slot == "stack"]
    if "national" in paired and "severity" in paired:
        nat_title, nat_path = paired["national"]
        sev_title, sev_path = paired["severity"]
        ws.merge_range(row, 0, row, 8, nat_title, styles["section"])
        ws.merge_range(row, 9, row, 17, sev_title, styles["section"])
        pair_opts = {"x_scale": 0.42, "y_scale": 0.42, "object_position": 2}
        ws.insert_image(row + 1, 0, str(nat_path), pair_opts)
        ws.insert_image(row + 1, 9, str(sev_path), pair_opts)
        row += 34
    for title, path in stacked:
        ws.merge_range(row, 0, row, 17, title, styles["section"])
        ws.insert_image(row + 1, 0, str(path), {"x_scale": 0.50, "y_scale": 0.50, "object_position": 2})
        row += 34
    ws.set_zoom(100)


# Banner + KPI strip + a compact chart (about 1120 × 400 px).
BLOCK_ROWS = 32


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "source",
        nargs="?",
        default="FEG_KPI_DHAKA_5Oct.part1.rar",
        help="FEGE KPI source (.csv, .xlsb, .xlsx, or split .part1.rar)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="FEGE_Choked_Flat_Sites_v42.xlsx",
        help="Report workbook to write",
    )
    args = parser.parse_args()
    source = resolve_source(Path(args.source))
    output = Path(args.output)

    raw = load_kpi(source)
    required = {"Date", "Time", "eNodeB Name", RX_COL, TXBW_COL}
    missing = required - set(raw.columns)
    if missing:
        raise SystemExit(f"Missing columns: {', '.join(sorted(missing))}")

    work, records = analyse(raw)
    _require_references(records)

    geo_path = _find_input(GEO_FILE, source)
    if geo_path is None:
        raise SystemExit(f"Physical site database not found: {GEO_FILE}")
    geo = load_site_geo(geo_path)

    tech_path = _find_input(TECH_FILE, source)
    if tech_path is None:
        raise SystemExit(f"Dhaka tech site list not found: {TECH_FILE}")
    tech_by_site = load_site_tech(tech_path)
    tech_summary = summarise_tech(records, tech_by_site)

    # Snapshots are laid out at a fixed stride so the site-list links can be
    # written in the same pass: site i starts at Excel row i * BLOCK_ROWS + 1.
    _write_workbook(output, source.name, work, records, geo, tech_summary, tech_by_site)

    by_sev = defaultdict(int)
    for rec in records:
        by_sev[rec["severity"]] += 1
    print(f"Window: {work['Date'].min().date()} – {work['Date'].max().date()} ({ANALYSIS_DAYS} days, all hours)")
    print(f"Sites checked: {work['eNodeB Name'].nunique()}")
    print(f"Issue sites: {len(records)}")
    print(
        f"4G choked: {tech_summary['choked_4g']} of {tech_summary['checked_4g']}"
    )
    print(
        f"4G+5G choked: {tech_summary['choked_4g5g']} of {tech_summary['checked_4g5g']}"
    )
    print("Severity (hours on cap):")
    for name in (SEV_SEVERE, SEV_HIGH, SEV_MODERATE, SEV_LOW):
        print(f"  {name}: {by_sev[name]}")
    by_when = defaultdict(int)
    by_listed = defaultdict(int)
    last_n = 0
    for rec in records:
        by_when[rec["cap_when"]] += 1
        by_listed[rec["listed_by"]] += 1
        last_n += int(rec["last_flag"])
    print("Listed because:")
    for name in ("≥3 days + last day", "≥3 days", "Last day"):
        print(f"  {name}: {by_listed[name]}")
    print(f"Last-day cap flag ({records[0]['last_day_label'] if records else 'n/a'}): {last_n}")
    print("Where capped:")
    for name in ("Busy + off-peak", "Busy hours", "Off-peak", "Partial"):
        print(f"  {name}: {by_when[name]}")
    by_shape = defaultdict(int)
    for rec in records:
        by_shape[rec["cap_type"]] += 1
    print("Cap shape:")
    for name in (CAP_CROWDED, CAP_CEILING):
        print(f"  {name}: {by_shape[name]}")
    print("Sample snaps:")
    for rec in records:
        if rec["reference"]:
            print(
                f"  {rec['site']}: stuck {rec['center']:.2f} Mbit/s, "
                f"Tx BW {rec['bw']:.0f}, hours {rec['hours_pct']:.1f}%, "
                f"days {rec['days_on_cap']}/{rec['days_total']}, "
                f"{last_day_text(rec)}"
            )
    print(f"Wrote {output}")


def _write_summary(
    book,
    styles,
    source_name,
    period_txt,
    n_sites,
    records,
    tech_summary=None,
    tech_by_site=None,
):
    """One sheet that shows the result without opening the other pages."""
    ws = book.add_worksheet("Dashboard")
    _page(ws, REPORT_TITLE)
    ws.set_tab_color("#C65911")

    widths = [16, 16, 16, 16, 16, 20, 22, 16, 16, 16]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)

    big = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 22,
            "bold": True,
            "font_color": NAVY,
            "align": "center",
            "valign": "vcenter",
            "bg_color": PALE,
        }
    )
    big_alert = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 22,
            "bold": True,
            "font_color": ORANGE_FONT,
            "align": "center",
            "valign": "vcenter",
            "bg_color": ORANGE,
        }
    )
    tile = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 9,
            "bold": True,
            "font_color": GREY,
            "align": "center",
            "valign": "vcenter",
            "bg_color": PALE,
        }
    )
    tile_alert = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 9,
            "bold": True,
            "font_color": ORANGE_FONT,
            "align": "center",
            "valign": "vcenter",
            "bg_color": ORANGE,
        }
    )

    counts = defaultdict(int)
    for rec in records:
        counts[rec["severity"]] += 1

    ws.set_row(0, 28)
    ws.merge_range("A1:J1", REPORT_TITLE, styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:J2",
        f"DHK transmission    ·    {period_txt}    ·    hourly    ·    {source_name}",
        styles["subtitle"],
    )

    tiles = [
        (0, "Sites checked", n_sites, False),
        (2, "Issue sites", len(records), True),
        (4, "Severe choke", counts[SEV_SEVERE], True),
        (
            6,
            "High / Moderate / Low",
            f"{counts[SEV_HIGH]}  /  {counts[SEV_MODERATE]}  /  {counts[SEV_LOW]}",
            False,
        ),
    ]
    ws.set_row(3, 18)
    ws.set_row(4, 32)
    for col, label, value, alert in tiles:
        label_fmt = tile_alert if alert else tile
        value_fmt = big_alert if alert else big
        ws.merge_range(3, col, 3, col + 1, label, label_fmt)
        if isinstance(value, int):
            ws.merge_range(4, col, 4, col + 1, value, value_fmt)
        else:
            ws.merge_range(4, col, 4, col + 1, value, value_fmt)

    tech_summary = tech_summary or {}
    tech_by_site = tech_by_site or {}
    group_title = book.add_format(
        {
            "font_name": "Calibri",
            "font_size": 14,
            "bold": True,
            "font_color": NAVY,
            "valign": "vcenter",
        }
    )
    row = 6
    if tech_summary:
        ws.set_row(row, 22)
        ws.merge_range(row, 0, row, 9, "Technology-wise Link Health", group_title)
        row += 1
        ws.set_row(row, 20)
        ws.merge_range(
            row,
            0,
            row,
            9,
            f"4G / 4G+5G choke summary    ·    source: {TECH_FILE}",
            styles["section"],
        )
        row += 1
        tech_tiles = [
            (0, "4G sites checked", tech_summary["checked_4g"], False),
            (2, "4G sites choked", tech_summary["choked_4g"], True),
            (4, "4G+5G sites checked", tech_summary["checked_4g5g"], False),
            (6, "4G+5G sites choked", tech_summary["choked_4g5g"], True),
        ]
        ws.set_row(row, 18)
        ws.set_row(row + 1, 32)
        for col, label, value, alert in tech_tiles:
            label_fmt = tile_alert if alert else tile
            value_fmt = big_alert if alert else big
            ws.merge_range(row, col, row, col + 1, label, label_fmt)
            ws.merge_range(row + 1, col, row + 1, col + 1, value, value_fmt)
        row += 3

    if tech_by_site:
        matrix = cap_when_matrix(records, tech_by_site)
        ws.set_row(row, 22)
        ws.merge_range(row, 0, row, 9, "Cap Window Observation", group_title)
        row += 1
        ws.set_row(row, 20)
        ws.merge_range(row, 0, row + 1, 0, "Category", styles["header"])
        ws.merge_range(row, 1, row, 3, "Total Summary", styles["header"])
        ws.merge_range(row, 4, row, 6, "4G+5G Sites Summary", styles["header"])
        ws.merge_range(row, 7, row, 9, "4G Sites Summary", styles["header"])
        row += 1
        ws.set_row(row, 28)
        subheads = ("Busy + off-peak", "Busy hours", "Off-peak")
        for base in (1, 4, 7):
            for offset, text in enumerate(subheads):
                ws.write(row, base + offset, text, styles["header"])
        row += 1
        total_fmt = book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": PALE,
                "border": 1,
                "border_color": LINE,
            }
        )
        total_label = book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "bg_color": PALE,
                "border": 1,
                "border_color": LINE,
            }
        )
        sevs = (SEV_SEVERE, SEV_HIGH, SEV_MODERATE, SEV_LOW)
        groups = ("total", TECH_4G5G, TECH_4G)
        for sev in sevs:
            ws.set_row(row, 18)
            ws.write(row, 0, sev, styles["center"])
            col = 1
            for group in groups:
                for when in CAP_WHEN_COLS:
                    value = matrix[group][sev][when]
                    if value:
                        ws.write_number(row, col, value, styles["int"])
                    else:
                        ws.write_blank(row, col, None, styles["center"])
                    col += 1
            row += 1
        ws.set_row(row, 18)
        ws.write(row, 0, "Grand Total", total_label)
        col = 1
        for group in groups:
            for when in CAP_WHEN_COLS:
                value = sum(matrix[group][sev][when] for sev in sevs)
                if value:
                    ws.write_number(row, col, value, total_fmt)
                else:
                    ws.write_blank(row, col, None, total_fmt)
                col += 1
        row += 2

    headers = [
        "No.",
        "eNodeB Name",
        "Stuck Rx (Mbit/s)",
        "Hours on cap (%)",
        "Days on cap (≥3h)",
        "Last day cap",
        "Listed because",
    ]
    indexed = list(enumerate(records))
    for name, note in SEVERITY_NOTES:
        group = [(i, rec) for i, rec in indexed if rec["severity"] == name]
        ws.set_row(row, 22)
        ws.merge_range(
            row,
            0,
            row,
            9,
            f"{name}    ·    {len(group)} site{'s' if len(group) != 1 else ''}",
            styles["section"],
        )
        row += 1
        ws.set_row(row, 28)
        ws.merge_range(row, 0, row, 9, note, styles["body"])
        row += 1
        ws.set_row(row, 22)
        for col, text in enumerate(headers):
            ws.write(row, col, text, styles["header"])
        row += 1
        for n_in_group, (i, rec) in enumerate(group):
            zebra = n_in_group % 2 == 1
            ws.set_row(row, 18)
            nfmt = styles["num_z"] if zebra else styles["num"]
            cfmt = styles["center_z"] if zebra else styles["center"]
            link = styles["link_z"] if zebra else styles["link"]
            excel_anchor = i * BLOCK_ROWS + 1
            ws.write_number(row, 0, i + 1, styles["int_z"] if zebra else styles["int"])
            ws.write_url(
                row,
                1,
                f"internal:'{SNAP_SHEET}'!A{excel_anchor}",
                link,
                string=rec["site"],
            )
            ws.write_number(row, 2, rec["center"], nfmt)
            ws.write_number(row, 3, rec["hours_pct"], styles["pct_z"] if zebra else styles["pct"])
            ws.write_string(row, 4, f"{rec['days_on_cap']}/{rec['days_total']}", cfmt)
            ws.write_string(row, 5, last_day_text(rec), cfmt)
            ws.write_string(row, 6, rec["listed_by"], cfmt)
            row += 1
        row += 1

    ws.set_zoom(110)


def _write_workbook(
    path: Path,
    source_name: str,
    work: pd.DataFrame,
    records: list[dict],
    geo: pd.DataFrame | None = None,
    tech_summary: dict[str, int] | None = None,
    tech_by_site: dict[str, str] | None = None,
) -> None:
    """Write the report. List links use the fixed snapshot block stride."""
    import xlsxwriter

    book = xlsxwriter.Workbook(str(path))
    styles = _styles(book)
    period_start = work["Date"].min()
    period_end = work["Date"].max()
    n_sites = int(work["eNodeB Name"].nunique())
    period_txt = (
        f"{period_start.strftime('%d %b %Y')} – {period_end.strftime('%d %b %Y')}"
    )

    # Dashboard is the first sheet so the file opens on the report view.
    chart_end = period_end.normalize()
    chart_start = chart_end - pd.Timedelta(days=SNAP_DAYS - 1)
    _write_summary(
        book, styles, source_name, period_txt, n_sites, records, tech_summary, tech_by_site
    )
    _write_list_linked(book, styles, source_name, period_txt, n_sites, records)
    _write_snapshots(book, styles, work, records, period_txt, chart_start, chart_end)
    _write_hourly(book, styles, work, records)
    _write_method(book, styles, source_name, period_txt, n_sites, records)
    if geo is not None and len(geo):
        _write_geoplot(book, styles, work, records, geo, path, tech_by_site)
    book.close()


def _write_list_linked(book, styles, source_name, period_txt, n_sites, records):
    """Same list as _write_list, with each site name linked to its snapshot."""
    ws = book.add_worksheet("1. Site List")
    _page(ws, REPORT_TITLE)

    widths = [5, 14, 12, 14, 16, 16, 16, 18, 16, 12, 32, 18, 12, 14, 12]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)

    last_day_hdr = records[0]["last_day_label"] if records else "last day"
    last_col_letter = "O"
    ws.set_row(0, 28)
    ws.merge_range(
        f"A1:{last_col_letter}1",
        f"{REPORT_TITLE} — issue sites ({ANALYSIS_START.strftime('%-d %b')}–{ANALYSIS_END.strftime('%-d %b %Y')})",
        styles["title"],
    )
    ws.set_row(1, 18)
    ws.merge_range(
        f"A2:{last_col_letter}2",
        f"DHK transmission    ·    {period_txt}    ·    hourly    ·    "
        f"{n_sites} sites checked    ·    {len(records)} issue sites",
        styles["subtitle"],
    )

    headers = [
        "No.",
        "eNodeB Name",
        "Severity",
        "Cap shape",
        "Tx Total BW (Mbit/s)",
        "Stuck RxMaxSpeed (Mbit/s)",
        "Max RxMaxSpeed (Mbit/s)",
        "Hours on cap (all 24h)",
        "Where capped",
        "Days on cap (≥3h)",
        f"Last day cap ({last_day_hdr})",
        "Listed because",
        "Longest flat run (h)",
        "Night Rx 02–06 (Mbit/s)",
        "Sample snap",
    ]
    header_row = 3
    ws.set_row(header_row, 36)
    for col, text in enumerate(headers):
        ws.write(header_row, col, text, styles["header"])

    for i, rec in enumerate(records):
        row = header_row + 1 + i
        zebra = i % 2 == 1
        ws.set_row(row, 18)
        n = styles["num_z"] if zebra else styles["num"]
        c = styles["center_z"] if zebra else styles["center"]
        link = styles["link_z"] if zebra else styles["link"]
        excel_anchor = i * BLOCK_ROWS + 1  # 1-based row on the snapshot sheet
        ws.write_number(row, 0, i + 1, styles["int_z"] if zebra else styles["int"])
        ws.write_url(
            row,
            1,
            f"internal:'{SNAP_SHEET}'!A{excel_anchor}",
            link,
            string=rec["site"],
        )
        ws.write_string(row, 2, rec["severity"], _severity_format(styles, rec["severity"], zebra))
        ws.write_string(row, 3, rec["cap_type"], c)
        ws.write_number(row, 4, rec["bw"], n)
        ws.write_number(row, 5, rec["center"], n)
        ws.write_number(row, 6, rec["max_rx"], n)
        ws.write_string(
            row,
            7,
            f"{rec['hours_on']}/{rec['hours_n']} ({rec['hours_pct']:.1f}%)",
            c,
        )
        ws.write_string(row, 8, rec["cap_when"], c)
        ws.write_string(row, 9, f"{rec['days_on_cap']}/{rec['days_total']}", c)
        last_txt = last_day_text(rec)
        ws.write_string(row, 10, last_txt, styles["yes"] if rec["last_flag"] else c)
        ws.write_string(row, 11, rec["listed_by"], c)
        ws.write_number(row, 12, rec["longest"], styles["int_z"] if zebra else styles["int"])
        if np.isfinite(rec["night_rx"]):
            ws.write_number(row, 13, rec["night_rx"], n)
        else:
            ws.write_string(row, 13, "—", c)
        if rec["reference"]:
            ws.write_string(row, 14, "Yes", styles["yes"])
        else:
            ws.write_string(row, 14, "—", c)

    last = header_row + len(records)
    ws.autofilter(header_row, 0, last, len(headers) - 1)
    ws.repeat_rows(header_row, header_row)

    last_col = 14
    note_row = last + 2
    ws.set_row(note_row, 20)
    ws.merge_range(note_row, 0, note_row, last_col, "How to read Severity", styles["section"])
    for offset, (name, text) in enumerate(SEVERITY_NOTES):
        r = note_row + 1 + offset
        ws.set_row(r, 32)
        ws.merge_range(r, 0, r, 1, name, _severity_format(styles, name, False))
        ws.merge_range(r, 2, r, last_col, text, styles["body"])

    where_row = note_row + 1 + len(SEVERITY_NOTES) + 1
    ws.set_row(where_row, 20)
    ws.merge_range(where_row, 0, where_row, last_col, "How to read Where capped", styles["section"])
    where_notes = (
        (
            "Busy + off-peak",
            "The ceiling appears in both 08:00–22:00 and night / off-peak hours. "
            "The port is limited around the clock.",
        ),
        (
            "Busy hours",
            "The ceiling is clear in 08:00–22:00. Night hours leave the cap.",
        ),
        (
            "Off-peak",
            "The ceiling is clear in night / off-peak hours even if daytime "
            "busy hours are less flat. These sites are still listed.",
        ),
    )
    for offset, (name, text) in enumerate(where_notes):
        r = where_row + 1 + offset
        ws.set_row(r, 28)
        ws.merge_range(r, 0, r, 1, name, styles["center"])
        ws.merge_range(r, 2, r, last_col, text, styles["body"])

    last_help = where_row + 5
    ws.set_row(last_help, 20)
    ws.merge_range(last_help, 0, last_help, last_col, "How to read Last day cap", styles["section"])
    ws.set_row(last_help + 1, 36)
    ws.merge_range(
        last_help + 1,
        0,
        last_help + 1,
        last_col,
        f"Green cell with a date means {last_day_hdr} had at least {LAST_DAY_HOURS} hours "
        "on the same stuck level (busy and/or off-peak). Example: "
        f"'{last_day_hdr} · 6h (busy 4, off-peak 2)'. "
        "A dash means that last day was not on the cap. "
        "Listed because = Last day means the site is in this file only for that last-day flag.",
        styles["body"],
    )

    counts = defaultdict(int)
    listed = defaultdict(int)
    when = defaultdict(int)
    shapes = defaultdict(int)
    last_n = 0
    for rec in records:
        counts[rec["severity"]] += 1
        listed[rec["listed_by"]] += 1
        when[rec["cap_when"]] += 1
        shapes[rec["cap_type"]] += 1
        last_n += int(rec["last_flag"])
    count_row = last_help + 3
    ws.write(count_row, 0, "Severity", styles["label"])
    ws.merge_range(count_row, 1, count_row, 2, f"Severe: {counts[SEV_SEVERE]}", styles["meta"])
    ws.merge_range(
        count_row, 3, count_row, 4, f"High: {counts[SEV_HIGH]}", styles["meta"]
    )
    ws.merge_range(
        count_row, 5, count_row, 6,
        f"Moderate: {counts[SEV_MODERATE]}    ·    Low: {counts[SEV_LOW]}",
        styles["meta"],
    )
    ws.merge_range(
        count_row,
        7,
        count_row,
        last_col,
        f"Sample snaps in the list: {', '.join(REFERENCE_SITES)}",
        styles["meta"],
    )
    ws.set_row(count_row + 1, 18)
    ws.write(count_row + 1, 0, "Listed", styles["label"])
    ws.merge_range(
        count_row + 1, 1, count_row + 1, 4,
        f"≥3 days + last day: {listed['≥3 days + last day']}",
        styles["meta"],
    )
    ws.merge_range(
        count_row + 1, 5, count_row + 1, 8,
        f"≥3 days: {listed['≥3 days']}",
        styles["meta"],
    )
    ws.merge_range(
        count_row + 1, 9, count_row + 1, last_col,
        f"Last day only: {listed['Last day']}    ·    last-day flag: {last_n}",
        styles["meta"],
    )
    ws.set_row(count_row + 2, 18)
    ws.write(count_row + 2, 0, "Where", styles["label"])
    ws.merge_range(
        count_row + 2, 1, count_row + 2, 4,
        f"Busy + off-peak: {when['Busy + off-peak']}",
        styles["meta"],
    )
    ws.merge_range(
        count_row + 2, 5, count_row + 2, 8,
        f"Busy hours: {when['Busy hours']}",
        styles["meta"],
    )
    ws.merge_range(
        count_row + 2, 9, count_row + 2, last_col,
        f"Off-peak: {when['Off-peak']}    ·    Partial: {when['Partial']}",
        styles["meta"],
    )
    ws.set_row(count_row + 3, 18)
    ws.write(count_row + 3, 0, "Shape", styles["label"])
    ws.merge_range(
        count_row + 3, 1, count_row + 3, 4,
        f"{CAP_CROWDED}: {shapes[CAP_CROWDED]}",
        styles["meta"],
    )
    ws.merge_range(
        count_row + 3, 5, count_row + 3, last_col,
        f"{CAP_CEILING}: {shapes[CAP_CEILING]}",
        styles["meta"],
    )
    ws.set_row(count_row + 4, 20)
    ws.merge_range(
        count_row + 4,
        0,
        count_row + 4,
        last_col,
        f"Window is {period_txt}. "
        f"{ANALYSIS_DAYS}-day rule: ≥{MIN_DAYS_ON_CAP} days with ≥{STUCK_DAY_HOURS} hours on the cap. "
        f"Last-day rule: {last_day_hdr} with ≥{LAST_DAY_HOURS} hours on the cap "
        "(busy and/or off-peak).",
        styles["note"],
    )




if __name__ == "__main__":
    main()
