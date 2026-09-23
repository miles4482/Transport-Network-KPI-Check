#!/usr/bin/env python3
"""Find FEGE ports whose RxMaxSpeed trace is choked (stuck flat).

Counters
--------
RxMaxSpeed (Mbit/s) = VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000
Tx Total BW (Mbit/s) = VS.FEGE.TxTotalBW(kbit/s) / 1000

A site is listed only when RxMaxSpeed is pinned to one hard ceiling the way
the sample charts show for DHAPT35 and DHAPT48: after the night dip the line
sits flat, and it cannot climb above that level.

The rule is applied to every site over the latest seven days in the file
(16–22 Sep 2026 in the current source). Every hour is scored, including
night and other off-peak hours. A site is listed when at least three of
those days have three hours on the cap, or when the last day has two
hours on the cap. Each snapshot chart shows only the latest three days.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import time
from pathlib import Path

import numpy as np
import pandas as pd

# Busy window used by the sample snaps: the line is flat through the day
# and only leaves the cap in the night valley.
BUSY_START = 8
BUSY_END = 22  # inclusive
NIGHT_START = 2
NIGHT_END = 6  # inclusive, the valley on the sample charts

# Flat-cap test on the latest 7 days (16–22 Sep in this file).
# A day is on-cap when at least 3 hours (busy or off-peak) sit on the ceiling.
# A site is listed when that happens on at least 3 of the 7 days.
# Separately, the last day is flagged when at least 2 hours sit on the ceiling.
TOL_MBPS = 4.0
TOL_FRAC = 0.015
# Hard ceiling. A few hours may sit a little above the crowded level.
# 40 Mbit/s (or 12% of the level) still reads as a cap. If the 99th
# percentile climbs past that, the crowded level is a busy cluster with
# spikes (the DHGULN2 / DHDHN40 / DHBDD32 / DHDHN09 shape), not a choke.
OVERSHOOT_MBPS = 40.0
OVERSHOOT_FRAC = 0.12
MIN_RUN_HOURS = 3
STUCK_DAY_HOURS = 3
MIN_DAYS_ON_CAP = 3
LAST_DAY_HOURS = 2
ANALYSIS_DAYS = 7
STD_MBPS = 2.5
STD_FRAC = 0.01
MIN_CENTER_MBPS = 20.0

# How the stuck level compares with the configured FEGE Tx bandwidth.
AT_BW_LOW = 0.85
AT_BW_HIGH = 1.20

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


def classify(util: float) -> str:
    if util >= AT_BW_HIGH:
        return "Above Tx BW"
    if util >= AT_BW_LOW:
        return "At Tx BW"
    return "Below Tx BW"


FINDING_ORDER = {"At Tx BW": 0, "Above Tx BW": 1, "Below Tx BW": 2}
FINDING_NOTES = (
    (
        "At Tx BW",
        "Stuck level is 85% to 120% of the FEGE Tx Total BW counter. "
        "This is the sample-snap case: nominal bandwidth is 250 Mbit/s and the "
        "measured cap sits around 220–280 Mbit/s. The port is full.",
    ),
    (
        "Above Tx BW",
        "The trace is flat, and the cap is more than 20% above the Tx Total BW "
        "counter. Speed is still stuck; the bandwidth counter and the real limit "
        "do not agree.",
    ),
    (
        "Below Tx BW",
        "The trace is flat at less than 85% of Tx Total BW. RxMaxSpeed is stuck, "
        "so transmission is limited below the configured FEGE bandwidth.",
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
    work["Date"] = excel_serial_to_datetime(work["Date"])
    work["Rx"] = work[RX_COL] / 1000.0 / 1000.0
    work["TxBW"] = work[TXBW_COL] / 1000.0
    work["Hour"] = work["Time"].str.slice(0, 2).astype(int)
    work["ts"] = work["Date"] + pd.to_timedelta(work["Hour"], unit="h")
    work = work.sort_values(["eNodeB Name", "ts"])
    work = work.drop_duplicates(["eNodeB Name", "ts"], keep="last")

    period_end = work["Date"].max().normalize()
    window_start = period_end - pd.Timedelta(days=ANALYSIS_DAYS - 1)
    work = work[(work["Date"] >= window_start) & (work["Date"] <= period_end)].copy()
    work = work.reset_index(drop=True)
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
        center, tol, band = find_plateau(rx)
        if center < MIN_CENTER_MBPS:
            continue
        hours = s["Hour"].to_numpy()
        busy = (hours >= BUSY_START) & (hours <= BUSY_END)
        offpeak = ~busy
        night = (hours >= NIGHT_START) & (hours <= NIGHT_END)
        overshoot = float(np.quantile(rx, 0.99) - center)
        overshoot_limit = max(OVERSHOOT_MBPS, OVERSHOOT_FRAC * center)
        if overshoot > overshoot_limit:
            continue
        in_band_std = float(np.std(rx[band])) if band.any() else 999.0
        if in_band_std > max(STD_MBPS, STD_FRAC * center):
            continue

        hours_pct = 100.0 * float(band.mean())
        run = longest_run(s["ts"].to_numpy(), band)

        day_on: dict = defaultdict(int)
        day_busy: dict = defaultdict(int)
        day_off: dict = defaultdict(int)
        for day, flag, is_busy in zip(s["Date"].dt.date, band, busy):
            if not flag:
                continue
            day_on[day] += 1
            if is_busy:
                day_busy[day] += 1
            else:
                day_off[day] += 1
        days_on_cap = sum(v >= STUCK_DAY_HOURS for v in day_on.values())
        days_total = int(s["Date"].dt.normalize().nunique())
        week_flag = days_on_cap >= MIN_DAYS_ON_CAP
        last_on = int(day_on.get(last_day, 0))
        last_busy = int(day_busy.get(last_day, 0))
        last_off = int(day_off.get(last_day, 0))
        last_flag = last_on >= LAST_DAY_HOURS
        if not week_flag and not last_flag:
            continue

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
                "finding": classify(util),
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
            FINDING_ORDER[r["finding"]],
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
        raise SystemExit(
            "Sample snap sites were not classified as choked: " + ", ".join(missing)
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


def _finding_format(styles, finding: str, zebra: bool):
    if finding == "At Tx BW":
        return styles["at"]
    if finding == "Above Tx BW":
        return styles["above"]
    return styles["below"]


def _page(ws, title: str, *, fit_width: bool = True):
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.45, right=0.45, top=0.6, bottom=0.5)
    # Snapshots must not be scaled onto the page width: that crushes the
    # two-row date/time axis into one overlapping band.
    if fit_width:
        ws.fit_to_pages(1, 0)
    ws.set_header(f"&L&8&K1F4E79{title}&R&8FEGE port  |  DHK")
    ws.set_footer("&L&8Choked = RxMaxSpeed stuck flat&R&8Page &P of &N")
    ws.hide_gridlines(2)
    ws.set_tab_color(NAVY)



def _chart_y_max(values: list[float]) -> int:
    """Round the visible peak up to the next 100, with a little headroom."""
    peak = max(values) if values else 100.0
    return max(100, int(np.ceil(peak * 1.12 / 100.0) * 100))


def _write_snapshots(book, styles, work, records, period_txt, chart_start, chart_end) -> dict[str, int]:
    ws = book.add_worksheet("2. Snapshots")
    _page(ws, "RxMaxSpeed snapshots", fit_width=False)
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
        c0 = i * 3
        data.write(0, c0, "Date")
        data.write(0, c0 + 1, "Time")
        data.write(0, c0 + 2, "Rx")
        rx_by_key = {
            (row["Date"].normalize(), int(row["Hour"])): float(row["Rx"])
            for _, row in s.iterrows()
        }
        dates: list[str] = []
        times: list[str] = []
        plotted: list[float] = []
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
        n = r

        ws.set_row(top, 26)
        banner = (
            f"{i + 1:02d}    {site}    ·    {rec['finding']}"
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

        ws.set_row(top + 3, 28)
        ws.merge_range(
            top + 3,
            0,
            top + 3,
            7,
            f"Max {rec['max_rx']:.2f} Mbit/s"
            f"   ·   vs Tx BW {rec['util']:.1f}%"
            f"   ·   longest flat run {rec['longest']} h"
            f"   ·   night Rx 02:00–06:00 {rec['night_rx']:.2f} Mbit/s"
            f"   ·   scored on {rec['start'].strftime('%-d-%b-%y')} – {rec['end'].strftime('%-d-%b-%y')}, all hours"
            f"   ·   {rec['cap_when']}"
            f"   ·   {rec['listed_by']}"
            f"   ·   last day: {last_day_text(rec)}"
            f"   ·   snapshot is the latest 3 days "
            f"({chart_start.strftime('%-d/%b')} – {chart_end.strftime('%-d/%b %Y')}), "
            f"hourly 00:00–23:00 with date grouped by day",
            styles["note"],
        )

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
        chart.set_y_axis(
            {
                "num_font": {"name": "Calibri", "size": 9, "color": "#595959"},
                "min": 0,
                "max": _chart_y_max(plotted),
                "major_unit": 50,
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
    ws.freeze_panes(1, 0)
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
            "Stuck level",
            "The level in the upper half of the trace that contains the most "
            f"hourly samples inside ±{TOL_MBPS:.1f} Mbit/s "
            f"(or ±{TOL_FRAC:.0%} of the 95th percentile, when that is wider). "
            "The level is the median of those samples.",
        ),
        (
            "Hard ceiling",
            "The 99th percentile of hourly RxMaxSpeed may sit at most "
            f"{OVERSHOOT_MBPS:.0f} Mbit/s above the stuck level "
            f"(or {OVERSHOOT_FRAC:.0%} of the stuck level, when that is wider). "
            "One odd hour above the cap is ignored. If the trace still climbs "
            "well above the crowded level (spikes of tens to hundreds of Mbit/s, "
            "as on DHGULN2, DHDHN40, DHBDD32, DHDHN09), that level is an outlier "
            "busy cluster, not a choke, and the site is not listed.",
        ),
        (
            "Analysis window",
            f"Only the latest {ANALYSIS_DAYS} days in the source file are scored. "
            "Earlier days are ignored so a site that recovered, or one that only "
            "became flat recently, is judged on the current week.",
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
            "At Tx BW",
            f"Stuck level is between {AT_BW_LOW:.0%} and {AT_BW_HIGH:.0%} of "
            "median Tx Total BW. Same region as the two sample snaps.",
        ),
        (
            "Above Tx BW",
            f"Stuck level is above {AT_BW_HIGH:.0%} of median Tx Total BW. "
            "The speed is stuck, but higher than the bandwidth counter.",
        ),
        (
            "Below Tx BW",
            f"Stuck level is below {AT_BW_LOW:.0%} of median Tx Total BW. "
            "The speed is stuck under the configured FEGE bandwidth.",
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

    foot = 24
    ws.set_row(foot, 20)
    ws.merge_range(foot, 0, foot, 1, "Check against the sample snaps", styles["section"])
    ws.set_row(foot + 1, 36)
    ws.merge_range(
        foot + 1,
        0,
        foot + 1,
        1,
        "DHAPT35 and DHAPT48 both pass this rule and are marked Sample snap = Yes "
        "on the site list. Scoring uses 16-Sep-26 to 22-Sep-26 and every hour. "
        "Each snapshot is a clean chart of the latest 3 days: "
        "hourly timestamps 00:00 to 23:00, "
        "and the date is shown once under that day. The chart title is the site name.",
        styles["method_val"],
    )



# Banner + KPI strip + a compact chart (about 1120 × 400 px).
BLOCK_ROWS = 32


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "source",
        nargs="?",
        default="FEG_KPI_5G Site_DHK.xlsb",
        help="FEGE KPI workbook (.xlsb or .xlsx)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="FEGE_Choked_Flat_Sites_v13.xlsx",
        help="Report workbook to write",
    )
    args = parser.parse_args()
    source = Path(args.source)
    output = Path(args.output)

    raw = pd.read_excel(source, engine="pyxlsb" if source.suffix.lower() == ".xlsb" else None)
    required = {"Date", "Time", "eNodeB Name", RX_COL, TXBW_COL}
    missing = required - set(raw.columns)
    if missing:
        raise SystemExit(f"Missing columns: {', '.join(sorted(missing))}")

    work, records = analyse(raw)
    _require_references(records)

    # Snapshots are laid out at a fixed stride so the site-list links can be
    # written in the same pass: site i starts at Excel row i * BLOCK_ROWS + 1.
    _write_workbook(output, source.name, work, records)

    by_finding = defaultdict(int)
    for rec in records:
        by_finding[rec["finding"]] += 1
    print(f"Window: {work['Date'].min().date()} – {work['Date'].max().date()} ({ANALYSIS_DAYS} days, all hours)")
    print(f"Sites checked: {work['eNodeB Name'].nunique()}")
    print(f"Choked / flat: {len(records)}")
    for name in ("At Tx BW", "Above Tx BW", "Below Tx BW"):
        print(f"  {name}: {by_finding[name]}")
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


def _write_summary(book, styles, source_name, period_txt, n_sites, records):
    """One sheet that shows the result without opening the other pages."""
    ws = book.add_worksheet("Summary")
    _page(ws, "FEGE choke check — at a glance")
    ws.set_tab_color("#C65911")

    widths = [6, 14, 14, 14, 12, 14, 32, 18]
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
        counts[rec["finding"]] += 1

    ws.set_row(0, 28)
    ws.merge_range("A1:H1", "FEGE choke check — at a glance", styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:H2",
        f"DHK 5G    ·    {period_txt}    ·    hourly    ·    {source_name}",
        styles["subtitle"],
    )

    tiles = [
        (0, "Sites checked", n_sites, False),
        (2, "Choked / flat", len(records), True),
        (4, "At Tx BW", counts["At Tx BW"], True),
        (6, "Above / below Tx BW", f"{counts['Above Tx BW']}  /  {counts['Below Tx BW']}", False),
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

    ws.set_row(5, 32)
    ws.merge_range(
        5,
        0,
        5,
        7,
        f"Scored on {period_txt} (latest {ANALYSIS_DAYS} days). "
        f"Listed when ≥{MIN_DAYS_ON_CAP} days have ≥{STUCK_DAY_HOURS} hours on the cap "
        f"(busy and off-peak), or the last day has ≥{LAST_DAY_HOURS} hours on the cap. "
        "Jagged traces that still spike well above the crowded level are outliers and are not listed. "
        "Each snapshot is the latest 3 days only. Open a site name to see it.",
        styles["note"],
    )

    headers = [
        "No.",
        "eNodeB Name",
        "Tx BW (Mbit/s)",
        "Stuck Rx (Mbit/s)",
        "vs Tx BW (%)",
        "Days on cap (≥3h)",
        "Last day cap",
        "Listed because",
    ]
    row = 7
    indexed = list(enumerate(records))
    for name, note in FINDING_NOTES:
        group = [(i, rec) for i, rec in indexed if rec["finding"] == name]
        ws.set_row(row, 22)
        ws.merge_range(
            row,
            0,
            row,
            7,
            f"{name}    ·    {len(group)} site{'s' if len(group) != 1 else ''}",
            styles["section"],
        )
        row += 1
        ws.set_row(row, 36)
        ws.merge_range(row, 0, row, 7, note, styles["body"])
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
                f"internal:'2. Snapshots'!A{excel_anchor}",
                link,
                string=rec["site"],
            )
            ws.write_number(row, 2, rec["bw"], nfmt)
            ws.write_number(row, 3, rec["center"], nfmt)
            ws.write_number(row, 4, rec["util"], styles["pct_z"] if zebra else styles["pct"])
            ws.write_string(row, 5, f"{rec['days_on_cap']}/{rec['days_total']}", cfmt)
            ws.write_string(row, 6, last_day_text(rec), cfmt)
            ws.write_string(row, 7, rec["listed_by"], cfmt)
            row += 1
        row += 1

    ws.freeze_panes(6, 0)
    ws.set_zoom(110)


def _write_workbook(path: Path, source_name: str, work: pd.DataFrame, records: list[dict]) -> None:
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

    # Summary is the first sheet so the file opens on the at-a-glance view.
    chart_end = period_end.normalize()
    chart_start = chart_end - pd.Timedelta(days=SNAP_DAYS - 1)
    _write_summary(book, styles, source_name, period_txt, n_sites, records)
    _write_list_linked(book, styles, source_name, period_txt, n_sites, records)
    _write_snapshots(book, styles, work, records, period_txt, chart_start, chart_end)
    _write_hourly(book, styles, work, records)
    _write_method(book, styles, source_name, period_txt, n_sites, records)
    book.close()


def _write_list_linked(book, styles, source_name, period_txt, n_sites, records):
    """Same list as _write_list, with each site name linked to its snapshot."""
    ws = book.add_worksheet("1. Site List")
    _page(ws, "FEGE choked / flat sites")

    widths = [5, 14, 14, 12, 14, 13, 11, 16, 16, 12, 32, 18, 12, 14, 12]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)

    last_day_hdr = records[0]["last_day_label"] if records else "last day"
    ws.set_row(0, 28)
    ws.merge_range("A1:O1", "FEGE transmission — choked and flat sites (latest 7 days)", styles["title"])
    ws.set_row(1, 18)
    ws.merge_range(
        "A2:O2",
        f"DHK 5G sites    ·    {period_txt}    ·    hourly    ·    "
        f"{n_sites} sites checked    ·    {len(records)} with RxMaxSpeed stuck flat",
        styles["subtitle"],
    )
    ws.set_row(2, 32)
    ws.merge_range(
        "A3:O3",
        "RxMaxSpeed (Mbit/s)  =  VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000"
        "          Tx Total BW (Mbit/s)  =  VS.FEGE.TxTotalBW(kbit/s) / 1000"
        f"          Source: {source_name}",
        styles["meta"],
    )
    ws.set_row(3, 32)
    ws.merge_range(
        "A4:O4",
        f"Scored on {period_txt} (16-Sep-26 to 22-Sep-26). "
        f"Listed when ≥{MIN_DAYS_ON_CAP} of 7 days have ≥{STUCK_DAY_HOURS} hours on the cap "
        f"(busy 08:00–22:00 and off-peak / night). "
        f"Last day cap column flags {last_day_hdr} when ≥{LAST_DAY_HOURS} hours that day sit on the cap. "
        "Outlier traces that keep climbing above the crowded level (DHGULN2-type spikes) are removed. "
        "The snapshot is the latest 3 days only. Open a site name to jump to it. Full rule on sheet 4.",
        styles["note"],
    )

    headers = [
        "No.",
        "eNodeB Name",
        "Finding",
        "Tx Total BW (Mbit/s)",
        "Stuck RxMaxSpeed (Mbit/s)",
        "Max RxMaxSpeed (Mbit/s)",
        "vs Tx BW (%)",
        "Hours on cap (all 24h)",
        "Where capped",
        "Days on cap (≥3h)",
        f"Last day cap ({last_day_hdr})",
        "Listed because",
        "Longest flat run (h)",
        "Night Rx 02–06 (Mbit/s)",
        "Sample snap",
    ]
    header_row = 5
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
            f"internal:'2. Snapshots'!A{excel_anchor}",
            link,
            string=rec["site"],
        )
        ws.write_string(row, 2, rec["finding"], _finding_format(styles, rec["finding"], zebra))
        ws.write_number(row, 3, rec["bw"], n)
        ws.write_number(row, 4, rec["center"], n)
        ws.write_number(row, 5, rec["max_rx"], n)
        ws.write_number(row, 6, rec["util"], styles["pct_z"] if zebra else styles["pct"])
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
    ws.freeze_panes(header_row + 1, 0)
    ws.repeat_rows(header_row, header_row)

    last_col = 14
    note_row = last + 2
    ws.set_row(note_row, 20)
    ws.merge_range(note_row, 0, note_row, last_col, "How to read Finding", styles["section"])
    for offset, (name, text) in enumerate(FINDING_NOTES):
        r = note_row + 1 + offset
        ws.set_row(r, 32)
        ws.merge_range(r, 0, r, 1, name, _finding_format(styles, name, False))
        ws.merge_range(r, 2, r, last_col, text, styles["body"])

    where_row = note_row + 5
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
            "The ceiling is clear in 08:00–22:00. Night hours leave the cap "
            "(same shape as DHAPT35 / DHAPT48).",
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
    last_n = 0
    for rec in records:
        counts[rec["finding"]] += 1
        listed[rec["listed_by"]] += 1
        when[rec["cap_when"]] += 1
        last_n += int(rec["last_flag"])
    count_row = last_help + 3
    ws.write(count_row, 0, "Count", styles["label"])
    ws.merge_range(count_row, 1, count_row, 2, f"At Tx BW: {counts['At Tx BW']}", styles["meta"])
    ws.merge_range(
        count_row, 3, count_row, 4, f"Above Tx BW: {counts['Above Tx BW']}", styles["meta"]
    )
    ws.merge_range(
        count_row, 5, count_row, 6, f"Below Tx BW: {counts['Below Tx BW']}", styles["meta"]
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
    ws.set_row(count_row + 3, 20)
    ws.merge_range(
        count_row + 3,
        0,
        count_row + 3,
        last_col,
        f"Window is {period_txt}. "
        f"7-day rule: ≥{MIN_DAYS_ON_CAP} days with ≥{STUCK_DAY_HOURS} hours on the cap. "
        f"Last-day rule: {last_day_hdr} with ≥{LAST_DAY_HOURS} hours on the cap "
        "(busy and/or off-peak).",
        styles["note"],
    )




if __name__ == "__main__":
    main()
