#!/usr/bin/env python3
"""Management PPT: one slide per site for comparison_v1 findings.

Pre window is 5–7 Oct. Post window is 6–8 Nov. comparison_v1 keeps
sites that were already flat, then the ceiling rose. November is either
released (solved) or flat again on the new ceiling (choked again).
Choked again keeps every site. Each other finding keeps the 10
largest daytime throughput gains. Within a finding, the largest gain
comes first.
"""

from __future__ import annotations

import pickle
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fege_choke_check as fege

OUT = Path("Comparison_v1_Findings_5Oct_vs_6Nov_v20.pptx")
CHART_DIR = Path("/tmp/ppt_charts")
ROW_CACHE = Path("/tmp/compare_v1_rows.pkl")

# Significant Data BusyHour 1H move. Absolute and percent both have to clear
# the bar so a tiny base rate does not look like a large gain.
SIG_TP_MBPS = 1.0
SIG_TP_PCT = 10.0
SIG_VOL_PCT = 10.0
# "Not much" / maintained volume: not a real drop, and not a significant rise.
HOLD_VOL_MIN = -5.0
HOLD_VOL_MAX = SIG_VOL_PCT
# Small throughput gain still has to be visible.
SMALL_TP_MBPS = 0.30

NAVY = "#1F4E79"
BLUE = "#5B9BD5"
GREEN = "#548235"
RED = "#C00000"
ORANGE = "#ED7D31"
PURPLE = "#7030A0"
GREY = "#595959"

# Highest daytime-throughput improvement first.
CATEGORY_ORDER = (
    fege.FINDING_UP_CHOKED,
    fege.FINDING_SOLVED_UP,
    fege.FINDING_UP,
    fege.FINDING_AGAIN,
    fege.FINDING_SOLVED,
)
BUCKET_BOTH = "both"
BUCKET_HOLD = "hold"
BUCKET_SMALL = "small"
BUCKET_OTHER = "other"
BUCKET_DAY = "day"
BUCKET_LABEL = {
    BUCKET_BOTH: "Both throughput and data volume improved",
    BUCKET_HOLD: "Throughput improved, data volume held",
    BUCKET_SMALL: "Small improvement, KPIs did not fall",
    BUCKET_DAY: "Daytime throughput improved",
    BUCKET_OTHER: "Included with every choked-again site",
}
# Choked again is shown in full. Each other finding shows this many,
# largest daytime gain first.
OTHER_LIMIT = 10
# Choked again opens the deck. The other findings follow.
DECK_ORDER = (
    fege.FINDING_AGAIN,
    fege.FINDING_SOLVED_UP,
    fege.FINDING_SOLVED,
)
BH_HOURS = (20, 21, 22)


def _num(value: object) -> float:
    try:
        out = float(value)
    except (TypeError, ValueError):
        return float("nan")
    return out if np.isfinite(out) else float("nan")


def _metrics(row: dict) -> dict | None:
    prev = row.get("prev_oss") or {}
    latest = row.get("latest_oss") or {}
    tp0, tp1 = _num(prev.get("tp_data_1h")), _num(latest.get("tp_data_1h"))
    vol0, vol1 = _num(prev.get("vol_data_1h")), _num(latest.get("vol_data_1h"))
    if not all(np.isfinite(v) for v in (tp0, tp1, vol0, vol1)):
        return None
    if tp0 == 0 or vol0 == 0:
        return None
    tp_delta = tp1 - tp0
    vol_delta = vol1 - vol0
    return {
        "tp0": tp0,
        "tp1": tp1,
        "tp_delta": tp_delta,
        "tp_pct": 100.0 * tp_delta / tp0,
        "vol0": vol0,
        "vol1": vol1,
        "vol_delta": vol_delta,
        "vol_pct": 100.0 * vol_delta / vol0,
    }


def _sig_tp(m: dict) -> bool:
    return m["tp_delta"] >= SIG_TP_MBPS and m["tp_pct"] >= SIG_TP_PCT


def _bucket(m: dict) -> str:
    """Which of the three management examples this site is, or blank."""
    if m["tp_delta"] < 0 and m["vol_delta"] < 0:
        return ""
    if _sig_tp(m) and m["vol_pct"] >= SIG_VOL_PCT and m["vol_delta"] > 0:
        return BUCKET_BOTH
    if _sig_tp(m) and HOLD_VOL_MIN <= m["vol_pct"] < HOLD_VOL_MAX:
        return BUCKET_HOLD
    if (
        m["tp_delta"] >= SMALL_TP_MBPS
        and not _sig_tp(m)
        and m["vol_pct"] >= HOLD_VOL_MIN
    ):
        return BUCKET_SMALL
    return ""


def _daytime(series: np.ndarray) -> float:
    """Median of 09:00–22:00. That is the part of the curve a reader compares."""
    hours = np.arange(series.size) % 24
    body = series[(hours >= 9) & (hours <= 22)]
    body = body[np.isfinite(body)]
    return float(np.median(body)) if body.size else float("nan")


def _min_day_hours(series: np.ndarray) -> int:
    counts = []
    for day in range(3):
        chunk = series[day * 24 : (day + 1) * 24]
        if chunk.size < 24:
            return 0
        counts.append(int(np.isfinite(chunk).sum()))
    return min(counts) if counts else 0


def _hour_mean(series: np.ndarray, hours: tuple[int, ...] | None = None) -> float:
    """Mean of the hourly samples. hours limits the clock hours, such as 20–22."""
    if series is None or series.size == 0:
        return float("nan")
    if hours is None:
        good = series[np.isfinite(series)]
    else:
        clock = np.arange(series.size) % 24
        good = series[np.isin(clock, list(hours)) & np.isfinite(series)]
    return float(np.mean(good)) if good.size else float("nan")


def _step(before: float, after: float) -> float:
    if np.isfinite(before) and np.isfinite(after):
        return after - before
    return float("nan")


def _profiles(pre: dict, post: dict) -> dict[str, dict]:
    out = {}
    for site in pre:
        if site not in post:
            continue
        tp0, tp1 = _daytime(pre[site]["tp"]), _daytime(post[site]["tp"])
        vol0, vol1 = _daytime(pre[site]["vol"]), _daytime(post[site]["vol"])
        tp24_0, tp24_1 = _hour_mean(pre[site]["tp"]), _hour_mean(post[site]["tp"])
        vol24_0, vol24_1 = _hour_mean(pre[site]["vol"]), _hour_mean(post[site]["vol"])
        tp_bh_0, tp_bh_1 = _hour_mean(pre[site]["tp"], BH_HOURS), _hour_mean(post[site]["tp"], BH_HOURS)
        vol_bh_0, vol_bh_1 = _hour_mean(pre[site]["vol"], BH_HOURS), _hour_mean(post[site]["vol"], BH_HOURS)
        out[site] = {
            "busy_tp0": tp0,
            "busy_tp1": tp1,
            "busy_tp_delta": tp1 - tp0 if np.isfinite(tp0) and np.isfinite(tp1) else float("nan"),
            "busy_tp_pct": (100.0 * (tp1 - tp0) / tp0) if tp0 else float("nan"),
            "busy_vol0": vol0,
            "busy_vol1": vol1,
            "busy_vol_pct": (100.0 * (vol1 - vol0) / vol0) if vol0 else float("nan"),
            "tp24": _step(tp24_0, tp24_1),
            "vol24": _step(vol24_0, vol24_1),
            "tp_bh": _step(tp_bh_0, tp_bh_1),
            "vol_bh": _step(vol_bh_0, vol_bh_1),
            "min_hours": min(_min_day_hours(pre[site]["tp"]), _min_day_hours(post[site]["tp"])),
        }
    return out


def _median_step(rows: list[dict], profiles: dict[str, dict], key: str, finding: str | None = None) -> float:
    values = []
    for row in rows:
        if finding is not None and row.get("finding") != finding:
            continue
        profile = profiles.get(row["site"])
        if profile is None:
            continue
        value = profile.get(key, float("nan"))
        if np.isfinite(value):
            values.append(value)
    return float(np.median(values)) if values else float("nan")


def _rank_key(site: dict) -> tuple[float, float]:
    delta = site.get("busy_tp_delta", float("nan"))
    pct = site.get("busy_tp_pct", float("nan"))
    return (
        delta if np.isfinite(delta) else -1e9,
        pct if np.isfinite(pct) else -1e9,
    )


def select_sites(rows: list[dict], profiles: dict[str, dict]) -> list[dict]:
    """Every choked-again site, and the 10 largest gains in each other finding.

    Order inside a finding is the daytime throughput step, largest first.
    """
    grouped: dict[str, list[dict]] = {name: [] for name in DECK_ORDER}
    for row in rows:
        finding = row.get("finding")
        if finding not in grouped:
            continue
        metrics = _metrics(row)
        profile = profiles.get(row["site"])
        if metrics is None or profile is None:
            continue
        bucket = _bucket(metrics)
        if not bucket:
            bucket = BUCKET_OTHER if finding == fege.FINDING_AGAIN else BUCKET_DAY
        grouped[finding].append({**row, "metrics": metrics, "bucket": bucket, **profile})

    picked: list[dict] = []
    for finding in DECK_ORDER:
        pool = sorted(grouped[finding], key=_rank_key, reverse=True)
        chosen = pool if finding == fege.FINDING_AGAIN else pool[:OTHER_LIMIT]
        for index, site in enumerate(chosen, start=1):
            site["pick_index"] = index
            site["pick_total"] = len(chosen)
            picked.append(site)
        print(f"{finding}: {len(chosen)} of {len(pool)} sites", flush=True)
        for site in chosen[:12]:
            m = site["metrics"]
            print(
                f"  {site['pick_index']}. {site['site']}  {site['bucket']}"
                f"  daytime {site['busy_tp0']:.1f}->{site['busy_tp1']:.1f}"
                f" ({site['busy_tp_pct']:+.0f}%)"
                f"  1H TP {m['tp0']:.2f}->{m['tp1']:.2f}",
                flush=True,
            )
        if len(chosen) > 12:
            print(f"  … {len(chosen) - 12} more", flush=True)
    return picked


def _hour_matrix(frame: pd.DataFrame, sites: set[str], value_cols: dict[str, str], start, end):
    """site -> {name: 72-length array} aligned to the three-day window."""
    days = list(pd.date_range(start, end, freq="D"))
    day_index = {day.normalize(): i for i, day in enumerate(days)}
    out = {
        site: {name: np.full(72, np.nan) for name in value_cols}
        for site in sites
    }
    if frame.empty:
        return out, days
    sub = frame.loc[frame["eNodeB Name"].astype(str).isin(sites)]
    sites_arr = sub["eNodeB Name"].astype(str).to_numpy()
    dates = pd.to_datetime(sub["Date"]).dt.normalize().to_numpy()
    hours = sub["Hour"].to_numpy()
    arrays = {name: sub[col].to_numpy(dtype=float) for name, col in value_cols.items()}
    for i, site in enumerate(sites_arr):
        day = pd.Timestamp(dates[i]).normalize()
        slot = day_index.get(day)
        hour = int(hours[i])
        if slot is None or not 0 <= hour <= 23 or site not in out:
            continue
        pos = slot * 24 + hour
        for name, values in arrays.items():
            value = float(values[i])
            if np.isfinite(value):
                out[site][name][pos] = value
    return out, days


def _prepare_feg(work: pd.DataFrame, sites: set[str], start, end):
    part = work.loc[
        work["eNodeB Name"].astype(str).isin(sites)
        & (work["Date"] >= start)
        & (work["Date"] <= end),
        ["eNodeB Name", "Date", "Hour", "Rx", "TxBW"],
    ].copy()
    return _hour_matrix(part, sites, {"rx": "Rx", "tx": "TxBW"}, start, end)


def _prepare_oss(path: Path, sites: set[str], start, end):
    raw = pd.read_csv(path)
    raw["Date"] = fege.parse_kpi_date(raw["Date"])
    raw["Hour"] = fege.parse_hour(raw["Time"])
    raw["eNodeB Name"] = raw["eNodeB Name"].astype(str).str.strip()
    raw = raw.loc[
        raw["eNodeB Name"].isin(sites)
        & (raw["Date"] >= start)
        & (raw["Date"] <= end)
    ].copy()
    num = raw["DL User Throughput_N"].astype(float)
    den = raw["DL User Throughput_D"].astype(float)
    raw["Throughput"] = np.where(den > 0, num / den, np.nan)
    raw["Volume"] = raw["Data Volume,GB"].astype(float)
    return _hour_matrix(
        raw, sites, {"tp": "Throughput", "vol": "Volume"}, start, end
    )


def _concat(pre: np.ndarray, post: np.ndarray) -> np.ndarray:
    return np.concatenate([pre, post])


def _style_time_axis(ax, n: int) -> None:
    labels = ["5 Oct", "6 Oct", "7 Oct", "6 Nov", "7 Nov", "8 Nov"]
    ax.set_xlim(-0.5, n - 0.5)
    ticks = []
    tick_labels = []
    for day in range(6):
        for hour, text in ((0, "00"), (12, "12")):
            ticks.append(day * 24 + hour)
            tick_labels.append(text)
    ax.set_xticks(ticks)
    ax.set_xticklabels(tick_labels, fontsize=7, color=GREY)
    ax.tick_params(axis="x", length=0, pad=2)
    for day, label in enumerate(labels):
        ax.annotate(
            label,
            xy=(day * 24 + 11.5, 0),
            xycoords=("data", "axes fraction"),
            xytext=(0, -18),
            textcoords="offset points",
            ha="center",
            va="top",
            fontsize=8,
            color=NAVY,
            fontweight="regular",
            annotation_clip=False,
        )
    ax.axvspan(-0.5, 72 - 0.5, color="#DEEBF7", alpha=0.55, zorder=0)
    ax.axvspan(72 - 0.5, n - 0.5, color="#E2EFDA", alpha=0.55, zorder=0)
    ax.grid(axis="y", color="#D9D9D9", linewidth=0.6, zorder=0)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_color("#BFBFBF")
    ax.tick_params(axis="y", labelsize=8, colors=GREY)


def _sharp_limits(values: np.ndarray) -> tuple[float, float]:
    """Axis close to the 09:00–22:00 body, so a pre/post step is not flattened.

    Night hours still plot. If they sit below the daytime cluster they meet
    the axis floor instead of pulling the scale back to zero.
    """
    hours = np.arange(values.size) % 24
    body = values[(hours >= 9) & (hours <= 22) & np.isfinite(values)]
    if body.size < 8:
        body = values[np.isfinite(values)]
    if body.size == 0:
        return 0.0, 1.0
    finite = values[np.isfinite(values)]
    low = float(np.percentile(body, 10))
    body_high = float(np.percentile(body, 90))
    # The top follows almost every hour, so off-peak peaks are not sliced flat.
    high = float(np.percentile(finite, 98))
    peak = float(np.max(finite))
    body_span = max(body_high - low, 1.0)
    ymin = low - 0.45 * body_span
    if ymin < 0:
        ymin = 0.0
    if high > 0:
        ymin = min(ymin, high * 0.22)
    ymax = high + 0.08 * max(high - ymin, 1.0)
    if peak <= ymax * 1.15:
        ymax = peak + 0.06 * max(peak - ymin, 1.0)
    if ymax <= ymin:
        ymax = ymin + 1.0
    return ymin, ymax


def _draw_feg(path: Path, rx: np.ndarray, pre_peak: float, pre_avg: float, post_peak: float) -> None:
    fig, ax = plt.subplots(figsize=(7.4, 3.55), dpi=140)
    x = np.arange(len(rx))
    _style_time_axis(ax, len(rx))
    pre = np.where(x < 72, rx, np.nan)
    post = np.where(x >= 72, rx, np.nan)
    ax.plot(x, pre, color=BLUE, linewidth=1.8, label="RxMaxSpeed 5–7 Oct", zorder=3)
    ax.plot(x, post, color=GREEN, linewidth=1.8, label="RxMaxSpeed 6–8 Nov", zorder=3)
    if np.isfinite(pre_peak):
        ax.plot(x[:72], np.full(72, pre_peak), color=RED, linewidth=1.1, linestyle="--", label=f"Peak 5–7 Oct ({pre_peak:.0f})", zorder=2)
    if np.isfinite(pre_avg):
        ax.plot(x[:72], np.full(72, pre_avg), color=ORANGE, linewidth=1.1, linestyle="--", label=f"Avg peak 5–7 Oct ({pre_avg:.0f})", zorder=2)
    if np.isfinite(post_peak):
        ax.plot(x[72:], np.full(len(rx) - 72, post_peak), color=PURPLE, linewidth=1.1, linestyle="--", label=f"Peak 6–8 Nov ({post_peak:.0f})", zorder=2)
    finite = rx[np.isfinite(rx)]
    ymax, _step = fege._chart_y_scale(finite.tolist() if finite.size else [50])
    ax.set_ylim(0, ymax)
    ax.set_ylabel("Mbit/s", fontsize=8, color=GREY)
    fig.text(0.08, 0.98, "FEGE RxMaxSpeed", ha="left", va="top", fontsize=12, color=NAVY, fontweight="bold")
    ax.legend(
        loc="lower left",
        bbox_to_anchor=(0.0, 1.02),
        ncol=3,
        frameon=False,
        fontsize=7,
        borderaxespad=0,
    )
    fig.subplots_adjust(left=0.08, right=0.98, top=0.70, bottom=0.22)
    fig.savefig(path, dpi=140)
    plt.close(fig)


def _daily_peak_low(tp: np.ndarray) -> tuple[float, float]:
    """Mean of each day's highest throughput hour, and of each day's lowest."""
    peaks, lows = [], []
    for day in range(tp.size // 24):
        chunk = tp[day * 24 : (day + 1) * 24]
        good = chunk[np.isfinite(chunk)]
        if good.size == 0:
            continue
        peaks.append(float(np.max(good)))
        lows.append(float(np.min(good)))
    peak = float(np.mean(peaks)) if peaks else float("nan")
    low = float(np.mean(lows)) if lows else float("nan")
    return peak, low


def _guide(ax, xspan: np.ndarray, level: float, color: str, style: str, width: float = 1.3):
    if not np.isfinite(level) or xspan.size == 0:
        return None
    line, = ax.plot(
        xspan,
        np.full(xspan.size, level),
        color=color,
        linewidth=width,
        linestyle=style,
        zorder=4,
    )
    return line


def _evening_volume_mean(vol: np.ndarray) -> float:
    """Mean data volume at 20:00, 21:00, 22:00 and 23:00 across this window."""
    picked = []
    for day in range(vol.size // 24):
        chunk = vol[day * 24 : (day + 1) * 24]
        for hour in (20, 21, 22, 23):
            value = chunk[hour]
            if np.isfinite(value):
                picked.append(float(value))
    return float(np.mean(picked)) if picked else float("nan")


def _draw_oss(path: Path, tp: np.ndarray, vol: np.ndarray, pre_level: float, post_level: float) -> None:
    """Hourly throughput, volume bars, and pre/post guides.

    Throughput guides are the mean daily peak and the mean daily low, each
    drawn only across its own days. Data volume has one mean line per window,
    the average of hours 20–23. The pre and post daytime throughput levels
    stay as a note.
    """
    fig, ax = plt.subplots(figsize=(7.4, 3.55), dpi=140)
    x = np.arange(len(tp))
    _style_time_axis(ax, len(tp))
    ax2 = ax.twinx()
    bars = ax2.bar(x, vol, width=0.85, color="#C65911", alpha=0.45, zorder=1)
    line_tp, = ax.plot(x, tp, color="#2F5496", linewidth=1.6, zorder=3)
    pre_peak, pre_low = _daily_peak_low(tp[:72])
    post_peak, post_low = _daily_peak_low(tp[72:])
    guides = (
        (_guide(ax, x[:72], pre_peak, "#548235", "--", 1.6), "TP peak pre"),
        (_guide(ax, x[72:], post_peak, "#203864", "--", 1.6), "TP peak post"),
        (_guide(ax, x[:72], pre_low, "#7F7F7F", ":", 1.6), "TP low pre"),
        (_guide(ax, x[72:], post_low, "#000000", ":", 1.6), "TP low post"),
        # Volume average of hours 20–23. Purple, not orange, so it is not the bars.
        (_guide(ax2, x[:72], _evening_volume_mean(vol[:72]), "#7030A0", "-.", 2.0), "Vol 20–23 pre"),
        (_guide(ax2, x[72:], _evening_volume_mean(vol[72:]), "#C00000", "-.", 2.0), "Vol 20–23 post"),
    )
    levels = [value for value in (pre_peak, post_peak, pre_low, post_low) if np.isfinite(value)]
    tp_min, tp_max = _sharp_limits(np.concatenate([tp, levels]) if levels else tp)
    vol_min, vol_max = _sharp_limits(vol)
    ax.set_ylim(tp_min, tp_max)
    ax2.set_ylim(vol_min, vol_max)
    ax.set_ylabel("Throughput (Mbps)", fontsize=8, color="#2F5496")
    ax2.set_ylabel("Data volume (GB)", fontsize=8, color="#C65911")
    ax.tick_params(axis="y", colors="#2F5496")
    ax2.tick_params(axis="y", labelsize=8, colors="#C65911")
    ax2.spines["right"].set_color("#C65911")
    fig.text(0.10, 0.985, "OSS throughput and data volume", ha="left", va="top", fontsize=11, color=NAVY, fontweight="bold")
    note = []
    if np.isfinite(pre_level):
        note.append(f"Pre {pre_level:.1f} Mbps")
    if np.isfinite(post_level):
        note.append(f"Post {post_level:.1f} Mbps")
    if note:
        fig.text(0.98, 0.985, "    ".join(note), ha="right", va="top", fontsize=8, color=GREY)
    handles = [line_tp, bars]
    labels = ["Throughput", "Data volume"]
    for line, label in guides:
        if line is not None:
            handles.append(line)
            labels.append(label)
    fig.legend(
        handles,
        labels,
        loc="upper left",
        bbox_to_anchor=(0.08, 0.955),
        ncol=4,
        frameon=False,
        fontsize=7.5,
        borderaxespad=0,
        columnspacing=1.0,
        handlelength=2.2,
        handletextpad=0.4,
    )
    fig.subplots_adjust(left=0.10, right=0.90, top=0.74, bottom=0.18)
    fig.savefig(path, dpi=140)
    plt.close(fig)


def _peaks(rx: np.ndarray, on_cap: np.ndarray | None, center: float) -> tuple[float, float]:
    hours = rx[:72] if rx.size >= 72 else rx
    finite = hours[np.isfinite(hours)]
    peak = float(np.max(finite)) if finite.size else float("nan")
    # Daily maximums approximate the chart's average peak when a stuck level
    # is not known. Prefer the official wall average when OnCap is present.
    if on_cap is not None and np.isfinite(center):
        collected = []
        for day in range(3):
            sl = slice(day * 24, day * 24 + 24)
            day_rx = hours[sl]
            day_on = on_cap[sl]
            good = day_rx[np.isfinite(day_rx)]
            if good.size == 0:
                continue
            day_max = float(np.max(good))
            for value, flag in zip(day_rx, day_on):
                if not np.isfinite(value):
                    continue
                if flag or value >= day_max or value >= center:
                    collected.append(float(value))
        avg = float(np.mean(collected)) if collected else float("nan")
    else:
        daily = []
        for day in range(3):
            sl = hours[day * 24 : day * 24 + 24]
            good = sl[np.isfinite(sl)]
            if good.size:
                daily.append(float(np.max(good)))
        avg = float(np.mean(daily)) if daily else float("nan")
    return peak, avg


def _sentence(site: dict) -> list[str]:
    m = site["metrics"]
    bucket = site["bucket"]
    rx0, rx1 = _num(site.get("old_rx")), _num(site.get("new_rx"))
    if site.get("bw_up") and np.isfinite(rx0) and np.isfinite(rx1):
        ceiling = (
            f"Tx ceiling rose from {rx0:.0f} to {rx1:.0f} Mbit/s "
            f"({rx1 - rx0:+.0f} Mbit/s, busy-hour RxMaxSpeed p90)."
        )
    elif np.isfinite(rx0) and np.isfinite(rx1):
        ceiling = (
            f"Busy-hour RxMaxSpeed p90 moved from {rx0:.0f} to {rx1:.0f} Mbit/s. "
            "That is below the Tx BW step-up test of 40 Mbit/s or 12%."
        )
    else:
        ceiling = "Busy-hour RxMaxSpeed ceiling could not be compared."
    daytime = (
        f" Daytime 09:00–22:00 median {site['busy_tp0']:.1f} → {site['busy_tp1']:.1f} Mbps."
    )
    tp = (
        f"Data BusyHour 1H throughput {m['tp0']:.2f} → {m['tp1']:.2f} Mbps "
        f"({m['tp_delta']:+.2f} Mbps, {m['tp_pct']:+.1f}%).{daytime}"
    )
    vol = (
        f"Data BusyHour 1H data volume {m['vol0']:.1f} → {m['vol1']:.1f} GB "
        f"({m['vol_delta']:+.1f} GB, {m['vol_pct']:+.1f}%)."
    )
    if site.get("prev_choked") and site.get("latest_choked"):
        choke = "Still choked after the window, so the higher throughput is already on the new ceiling."
    elif site.get("prev_choked") and not site.get("latest_choked"):
        choke = "Choked on 5–7 Oct. Off the choke list on 6–8 Nov."
    elif site.get("latest_choked"):
        choke = "Not choked on 5–7 Oct. Choked on 6–8 Nov."
    else:
        choke = "Not on the choke list in either window."
    day_clear = site.get("busy_tp_pct", 0) >= 8
    if site.get("bw_up") and bucket == BUCKET_BOTH and day_clear:
        lead = "After the Tx BW increase, throughput improved significantly and data volume increased with it."
    elif site.get("bw_up") and bucket == BUCKET_BOTH:
        lead = "After the Tx BW increase, the congested busy hour is faster and data volume increased. Throughput across the rest of the day is about level."
    elif site.get("bw_up") and bucket == BUCKET_HOLD and day_clear:
        lead = "After the Tx BW increase, throughput improved significantly. Data volume was held."
    elif site.get("bw_up") and bucket == BUCKET_HOLD:
        lead = "After the Tx BW increase, the congested busy hour is faster and data volume was held. Throughput across the rest of the day is about level."
    elif site.get("bw_up"):
        lead = "After the Tx BW increase, throughput is higher and data volume did not fall."
    elif bucket == BUCKET_BOTH and day_clear:
        lead = "Throughput and data volume both improved. The Tx ceiling did not meet the step-up test."
    elif bucket == BUCKET_BOTH:
        lead = "The congested busy hour is faster and data volume increased. The Tx ceiling did not meet the step-up test."
    elif bucket == BUCKET_HOLD and day_clear:
        lead = "Throughput improved significantly and data volume was held. The Tx ceiling did not meet the step-up test."
    elif bucket == BUCKET_HOLD:
        lead = "The congested busy hour is faster and data volume was held. The Tx ceiling did not meet the step-up test."
    elif site.get("busy_tp_pct", 0) >= 8 and site.get("bw_up"):
        lead = "After the Tx BW increase, daytime throughput improved."
    elif site.get("busy_tp_pct", 0) >= 8:
        lead = "Daytime throughput improved. The Tx ceiling did not meet the step-up test."
    elif site.get("finding") == fege.FINDING_AGAIN:
        lead = "Still choked on the new ceiling."
    else:
        lead = "A small throughput gain, and data volume did not fall. The Tx ceiling did not meet the step-up test."
    return [lead, ceiling, tp, vol, choke]


def _style_run(run, size: int, color: RGBColor, bold: bool) -> None:
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    # Force Calibri on the latin typeface so PowerPoint does not substitute.
    rpr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        node = rpr.find(qn(f"a:{tag}"))
        if node is None:
            node = rpr.makeelement(qn(f"a:{tag}"), {})
            rpr.append(node)
        node.set("typeface", "Calibri")


def _set_run(paragraph, text: str, size: int, color: RGBColor, bold: bool = False) -> None:
    paragraph.clear()
    run = paragraph.add_run()
    run.text = text
    _style_run(run, size, color, bold)


def _add_run(paragraph, text: str, size: int, color: RGBColor, bold: bool = False) -> None:
    run = paragraph.add_run()
    run.text = text
    _style_run(run, size, color, bold)


def _add_slide(prs: Presentation, site: dict, feg_png: Path, oss_png: Path, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy = RGBColor(0x1F, 0x4E, 0x79)
    white = RGBColor(0xFF, 0xFF, 0xFF)
    ink = RGBColor(0x33, 0x33, 0x33)
    accent = {
        BUCKET_BOTH: RGBColor(0x54, 0x82, 0x35),
        BUCKET_HOLD: RGBColor(0x2F, 0x54, 0x96),
        BUCKET_SMALL: RGBColor(0xC6, 0x59, 0x11),
        BUCKET_OTHER: RGBColor(0x83, 0x3C, 0x0C),
        BUCKET_DAY: RGBColor(0x1F, 0x4E, 0x79),
    }.get(site["bucket"], RGBColor(0x1F, 0x4E, 0x79))

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.78))
    bar.fill.solid()
    bar.fill.fore_color.rgb = navy
    bar.line.fill.background()
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.78), prs.slide_width, Inches(0.08))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent
    stripe.line.fill.background()

    title = slide.shapes.add_textbox(Inches(0.28), Inches(0.06), Inches(10.15), Inches(0.40))
    tf = title.text_frame
    tf.word_wrap = False
    _set_run(tf.paragraphs[0], f"{site['finding']}: {site['site']}", 22, white, True)

    meta = slide.shapes.add_textbox(Inches(0.30), Inches(0.44), Inches(9.6), Inches(0.30))
    mf = meta.text_frame
    mf.word_wrap = False
    _set_run(
        mf.paragraphs[0],
        BUCKET_LABEL[site["bucket"]],
        13,
        RGBColor(0xFF, 0xF2, 0xCC),
        False,
    )

    side = slide.shapes.add_textbox(Inches(10.45), Inches(0.16), Inches(2.65), Inches(0.46))
    sf = side.text_frame
    sf.word_wrap = False
    p = sf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    _set_run(
        p,
        f"{site['pick_index']} of {site['pick_total']} in this finding",
        12,
        white,
        False,
    )

    slide.shapes.add_picture(str(feg_png), Inches(0.18), Inches(0.98), Inches(6.45), Inches(4.15))
    slide.shapes.add_picture(str(oss_png), Inches(6.70), Inches(0.98), Inches(6.45), Inches(4.15))

    notes = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.22), Inches(5.22), Inches(12.90), Inches(1.95)
    )
    notes.fill.solid()
    notes.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFB)
    notes.line.color.rgb = RGBColor(0xD0, 0xD7, 0xE2)
    ntf = notes.text_frame
    ntf.word_wrap = True
    ntf.auto_size = None
    ntf.margin_left = Inches(0.14)
    ntf.margin_right = Inches(0.12)
    ntf.margin_top = Inches(0.08)
    ntf.margin_bottom = Inches(0.06)
    try:
        ntf._txBody.bodyPr.set("anchor", "t")
    except Exception:
        pass
    lines = _sentence(site)
    for i, line in enumerate(lines):
        paragraph = ntf.paragraphs[0] if i == 0 else ntf.add_paragraph()
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_after = Pt(2)
        _set_run(paragraph, line, 12, navy if i == 0 else ink, i == 0)

    foot = slide.shapes.add_textbox(Inches(0.28), Inches(7.20), Inches(10.5), Inches(0.26))
    ff = foot.text_frame
    _set_run(
        ff.paragraphs[0],
        "Pre 5–7 Oct    ·    Post 6–8 Nov    ·    Charts are hourly    ·    Notes use Data BusyHour 1H",
        10,
        RGBColor(0x80, 0x80, 0x80),
        False,
    )
    page_box = slide.shapes.add_textbox(Inches(11.4), Inches(7.20), Inches(1.65), Inches(0.26))
    pf = page_box.text_frame
    pp = pf.paragraphs[0]
    pp.alignment = PP_ALIGN.RIGHT
    _set_run(pp, f"{page} / {pages}", 10, RGBColor(0x80, 0x80, 0x80), False)


def _set_cell_fill(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag == qn("a:solidFill"):
            tc_pr.remove(child)
    solid = tc_pr.makeelement(qn("a:solidFill"), {})
    solid.append(solid.makeelement(qn("a:srgbClr"), {"val": hex_color}))
    tc_pr.append(solid)


def _set_cell_border(cell, color: str = "D0D7E2", width_pt: float = 0.75) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    width = str(int(width_pt * 12700))
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        for child in list(tc_pr):
            if child.tag == qn(f"a:{edge}"):
                tc_pr.remove(child)
        line = tc_pr.makeelement(
            qn(f"a:{edge}"), {"w": width, "cap": "flat", "cmpd": "sng", "algn": "ctr"}
        )
        fill = line.makeelement(qn("a:solidFill"), {})
        fill.append(fill.makeelement(qn("a:srgbClr"), {"val": color}))
        line.append(fill)
        tc_pr.append(line)


def _signed(value: float, unit: str, digits: int) -> str:
    if not np.isfinite(value):
        return "—"
    return f"{value:+.{digits}f} {unit}"


def _category_brief(rows: list[dict], profiles: dict[str, dict], finding: str) -> dict:
    """Median hourly gain for one finding, across every comparison site."""
    group = [row for row in rows if row.get("finding") == finding]
    return {
        "n": len(group),
        "tp_up": sum(1 for row in group if row.get("throughput_impact") == "Improved"),
        "vol_up": sum(1 for row in group if row.get("volume_impact") == "Improved"),
        "tp24": _median_step(group, profiles, "tp24"),
        "tp_bh": _median_step(group, profiles, "tp_bh"),
        "vol24": _median_step(group, profiles, "vol24"),
        "vol_bh": _median_step(group, profiles, "vol_bh"),
    }


def _chart_delta(group: dict, frame_key: str, column: str) -> float:
    """Post minus pre on the combined chart for one category."""
    before = _window_mean(group[frame_key], TREND_PRE, column)
    after = _window_mean(group[frame_key], TREND_POST, column)
    if not (np.isfinite(before) and np.isfinite(after)):
        return float("nan")
    return after - before


def _summary_box(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str,
    story: str,
    brief: dict,
    trend: dict | None,
) -> None:
    """One rounded finding box: median gain, then the combined-chart change."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFB)
    card.line.color.rgb = RGBColor(0x8F, 0xA8, 0xC4)
    card.line.width = Pt(1.25)
    body = slide.shapes.add_textbox(
        Inches(left + 0.22), Inches(top + 0.10), Inches(width - 0.40), Inches(height - 0.18)
    )
    frame = body.text_frame
    frame.word_wrap = True
    _set_run(frame.paragraphs[0], title, 16, RGBColor(0x1F, 0x4E, 0x79), True)
    lines: list[tuple[str, bool, int]] = [
        (story, False, 11),
        (f"Throughput, 24h avg    {_signed(brief['tp24'], 'Mbps', 1)}", False, 12),
        (f"Throughput, 20–22    {_signed(brief['tp_bh'], 'Mbps', 1)}", False, 12),
        (f"Data volume, 24h avg    {_signed(brief['vol24'], 'GB', 1)}", False, 12),
        (f"Data volume, 20–22    {_signed(brief['vol_bh'], 'GB', 1)}", False, 12),
    ]
    if trend is not None:
        prb = _chart_delta(trend, "daily", "prb")
        prb_text = f"{prb:+.1f} pp" if np.isfinite(prb) else "—"
        n = len(trend["sites"])
        lines.extend(
            (
                (f"Combined chart, {n} sites, 11–12 Oct to 8–9 Nov", True, 12),
                (
                    f"BH throughput {_signed(_chart_delta(trend, 'bh', 'tp'), 'Mbps', 1)}"
                    f"    ·    volume {_signed(_chart_delta(trend, 'bh', 'vol'), 'GB', 0)}",
                    False,
                    12,
                ),
                (
                    f"Daily throughput {_signed(_chart_delta(trend, 'daily', 'tp'), 'Mbps', 1)}"
                    f"    ·    volume {_signed(_chart_delta(trend, 'daily', 'vol'), 'GB', 0)}",
                    False,
                    12,
                ),
                (
                    f"Cell throughput {_signed(_chart_delta(trend, 'daily', 'cell'), 'Mbps', 1)}"
                    f"    ·    DL PRB {prb_text}",
                    False,
                    12,
                ),
            )
        )
    for text, strong, size in lines:
        paragraph = frame.add_paragraph()
        paragraph.space_before = Pt(1)
        color = RGBColor(0x1F, 0x4E, 0x79) if strong else RGBColor(0x33, 0x33, 0x33)
        _set_run(paragraph, text, size, color, strong)


def _combined_note(rows: list[dict], profiles: dict[str, dict]) -> str:
    """One line for the header: median gain at 24h and at 20–22."""
    tp24 = _signed(_median_step(rows, profiles, "tp24"), "Mbps", 1)
    tp_bh = _signed(_median_step(rows, profiles, "tp_bh"), "Mbps", 1)
    vol24 = _signed(_median_step(rows, profiles, "vol24"), "GB", 1)
    vol_bh = _signed(_median_step(rows, profiles, "vol_bh"), "GB", 1)
    return (
        f"Combined gain. Throughput {tp24} over 24h, {tp_bh} at 20–22. "
        f"Data volume {vol24} over 24h, {vol_bh} at 20–22."
    )


def _summary_table(slide, left: float, top: float, width: float, height: float, briefs: list[tuple[str, dict]]) -> None:
    """Combined count table, beside the third finding box."""
    headers = ("Finding", "Sites", "Throughput improved", "Data volume improved")
    shape = slide.shapes.add_table(
        1 + len(briefs), 4, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    table = shape.table
    for index, col_width in enumerate((2.28, 0.84, 1.60, 1.60)):
        table.columns[index].width = Inches(col_width)
    tbl_pr = table._tbl.tblPr
    if tbl_pr is not None:
        tbl_pr.set("bandRow", "0")
        tbl_pr.set("firstRow", "0")
    row_height = Inches(height / (1 + len(briefs)))
    table.rows[0].height = row_height
    for col, text in enumerate(headers):
        cell = table.cell(0, col)
        _set_cell_fill(cell, "1F4E79")
        _set_cell_border(cell, "1F4E79", 0.6)
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.06)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER if col else PP_ALIGN.LEFT
        _set_run(paragraph, text, 11, RGBColor(0xFF, 0xFF, 0xFF), True)
    for index, (title, brief) in enumerate(briefs, start=1):
        table.rows[index].height = row_height
        values = (title, str(brief["n"]), f"{brief['tp_up']} of {brief['n']}", f"{brief['vol_up']} of {brief['n']}")
        for col, text in enumerate(values):
            cell = table.cell(index, col)
            _set_cell_fill(cell, "F7F9FB" if index % 2 else "FFFFFF")
            _set_cell_border(cell, "D0D7E2", 0.5)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.06)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.word_wrap = True
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.CENTER if col else PP_ALIGN.LEFT
            _set_run(paragraph, text, 12, RGBColor(0x1F, 0x4E, 0x79) if col == 0 else RGBColor(0x33, 0x33, 0x33), col == 0)


def _add_summary_slide(
    prs: Presentation,
    rows: list[dict],
    selected: list[dict],
    profiles: dict[str, dict],
    trends: list[dict],
    pages: int,
) -> None:
    """Open summary: median gain, then each category's combined-chart change."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy = RGBColor(0x1F, 0x4E, 0x79)
    white = RGBColor(0xFF, 0xFF, 0xFF)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = white
    bg.line.fill.background()

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.92))
    bar.fill.solid()
    bar.fill.fore_color.rgb = navy
    bar.line.fill.background()
    title = slide.shapes.add_textbox(Inches(0.36), Inches(0.06), Inches(12.4), Inches(0.36))
    _set_run(title.text_frame.paragraphs[0], "Management summary", 24, white, True)
    sub = slide.shapes.add_textbox(Inches(0.36), Inches(0.42), Inches(12.5), Inches(0.44))
    sub_frame = sub.text_frame
    sub_frame.word_wrap = True
    _set_run(sub_frame.paragraphs[0], _combined_note(rows, profiles), 12, RGBColor(0xD6, 0xE3, 0xF0), False)

    trend_by = {group["finding"]: group for group in trends}
    panels = (
        (0.28, 1.04, 6.32, 2.72, fege.FINDING_AGAIN, "Choked again", "The ceiling rose, and the port is flat again on the new ceiling."),
        (6.72, 1.04, 6.32, 2.72, fege.FINDING_SOLVED_UP, "Solved — Tx BW increased", "The ceiling rose and the port is no longer choked."),
        (0.28, 3.88, 6.32, 2.72, fege.FINDING_SOLVED, "Solved", "Choked before, and clear now."),
    )
    briefs = []
    for left, top, width, height, finding, title_text, story in panels:
        brief = _category_brief(rows, profiles, finding)
        briefs.append((title_text, brief))
        _summary_box(slide, left, top, width, height, title_text, story, brief, trend_by.get(finding))
    _summary_table(slide, 6.72, 3.88, 6.32, 2.72, briefs)

    foot = slide.shapes.add_textbox(Inches(0.28), Inches(6.68), Inches(10.7), Inches(0.58))
    foot_frame = foot.text_frame
    foot_frame.word_wrap = True
    _set_run(
        foot_frame.paragraphs[0],
        "Median lines are 6–8 Nov versus 5–7 Oct. Combined chart sums the uploaded trend sites, 11–12 Oct to 8–9 Nov.",
        11,
        RGBColor(0x80, 0x80, 0x80),
        False,
    )
    page_box = slide.shapes.add_textbox(Inches(11.3), Inches(6.72), Inches(1.7), Inches(0.32))
    page_p = page_box.text_frame.paragraphs[0]
    page_p.alignment = PP_ALIGN.RIGHT
    _set_run(page_p, f"1 / {pages}", 12, RGBColor(0x80, 0x80, 0x80), False)


def _on_cap_vector(work: pd.DataFrame, site: str, start, end) -> np.ndarray:
    out = np.zeros(72, dtype=bool)
    if "OnCap" not in work.columns:
        return out
    part = work.loc[
        (work["eNodeB Name"].astype(str) == site)
        & (work["Date"] >= start)
        & (work["Date"] <= end)
    ]
    days = list(pd.date_range(start, end, freq="D"))
    day_index = {day.normalize(): i for i, day in enumerate(days)}
    for date, hour, flag in zip(part["Date"], part["Hour"], part["OnCap"]):
        slot = day_index.get(pd.Timestamp(date).normalize())
        hour = int(hour)
        if slot is None or not 0 <= hour <= 23:
            continue
        out[slot * 24 + hour] = bool(flag)
    return out


# Long trend uploaded on main. Daily is 15 Sep–9 Nov. Hourly is 10 Oct–9 Nov.
TREND_DAILY_CSV = Path("/tmp/kpi_trend/daily/OSS KPI trend_Daily_15Sep-09Nov.csv")
TREND_HOURLY_CSV = Path("/tmp/kpi_trend/hourly/OSS KPI trend_Hourly_10Oct-09Nov.csv")
TREND_PRE = (pd.Timestamp("2026-10-11"), pd.Timestamp("2026-10-12"))
TREND_POST = (pd.Timestamp("2026-11-08"), pd.Timestamp("2026-11-09"))
TREND_COLS = {
    "traffic": "Total Traffic (Mbit)",
    "tp_n": "User Throughput in DownLink (Mbps)_N",
    "tp_d": "User Throughput in DownLink (Mbps)_D",
    "cell_n": "DL PDCP Throughput (Mbit/s)_N",
    "cell_d": "DL PDCP Throughput (Mbit/s)_D",
    "prb_n": "DL PRB Utilization_N",
    "prb_d": "UL PRB Utilization_D",
}


def _trend_dates(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series.astype(str).str.strip(), format="%d-%b")
    return parsed.map(lambda stamp: stamp.replace(year=2026))


def _load_trend_frame(path: Path, hourly: bool) -> pd.DataFrame:
    if not path.exists():
        raise SystemExit(f"Trend KPI file not found: {path}")
    usecols = ["Date", "eNodeB Name", *TREND_COLS.values()]
    if hourly:
        usecols.insert(1, "Time")
    frame = pd.read_csv(path, usecols=usecols)
    frame = frame.rename(columns={"eNodeB Name": "site", **{v: k for k, v in TREND_COLS.items()}})
    frame["site"] = frame["site"].astype(str).str.strip()
    frame["day"] = _trend_dates(frame["Date"])
    for column in TREND_COLS:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    if hourly:
        frame["hour"] = frame["Time"].astype(str).str.split(":").str[0].astype(int)
    return frame


def _aggregate_kpi(frame: pd.DataFrame) -> pd.DataFrame:
    """One row per day. Rates are summed counters, then divided."""
    grouped = frame.groupby("day", sort=True)[list(TREND_COLS)].sum(min_count=1)
    denominator = grouped["tp_d"].where(grouped["tp_d"] > 0)
    cell_den = grouped["cell_d"].where(grouped["cell_d"] > 0)
    prb_den = grouped["prb_d"].where(grouped["prb_d"] > 0)
    return pd.DataFrame(
        {
            "vol": grouped["traffic"] / 8 / 1024 / 1024 / 1024,
            "tp": grouped["tp_n"] / denominator,
            "cell": grouped["cell_n"] / cell_den / 1000,
            "prb": grouped["prb_n"] / prb_den * 100,
        }
    )


def _window_mean(frame: pd.DataFrame, days: tuple[pd.Timestamp, pd.Timestamp], column: str) -> float:
    values = frame.reindex(list(days))[column].astype(float)
    values = values[np.isfinite(values)]
    if values.empty:
        return float("nan")
    return float(values.mean())


def _move(before: float, after: float, unit: str, digits: int, points: bool = False) -> str:
    if not (np.isfinite(before) and np.isfinite(after)):
        return "n/a"
    delta = after - before
    if points:
        return f"{before:.1f}% to {after:.1f}% ({delta:+.1f} pp)"
    pct = delta / before * 100 if before else float("nan")
    pct_text = f", {pct:+.0f}%" if np.isfinite(pct) else ""
    return f"{before:,.{digits}f} to {after:,.{digits}f} {unit} ({delta:+.{digits}f}{pct_text})"


def _trend_groups(rows: list[dict]) -> list[dict]:
    daily = _load_trend_frame(TREND_DAILY_CSV, hourly=False)
    hourly = _load_trend_frame(TREND_HOURLY_CSV, hourly=True)
    hourly = hourly[hourly["hour"].isin(BH_HOURS)]
    uploaded = set(daily["site"])
    groups = []
    for finding in DECK_ORDER:
        sites = sorted(row["site"] for row in rows if row.get("finding") == finding and row["site"] in uploaded)
        if not sites:
            continue
        wanted = set(sites)
        groups.append(
            {
                "finding": finding,
                "sites": sites,
                "daily": _aggregate_kpi(daily[daily["site"].isin(wanted)]),
                "bh": _aggregate_kpi(hourly[hourly["site"].isin(wanted)]),
            }
        )
    skipped = len(uploaded - {row["site"] for row in rows})
    print(f"Trend sites in a finding: {sum(len(g['sites']) for g in groups)} of {len(uploaded)}; {skipped} uploaded sites are outside the current findings", flush=True)
    return groups


def _trend_note(group: dict) -> list[str]:
    daily, bh = group["daily"], group["bh"]
    n = len(group["sites"])
    site_word = "site" if n == 1 else "sites"
    lines = [
        f"{group['finding']}: {n} {site_word} from the uploaded trend. Pre is the mean of 11–12 Oct. Post is the mean of 8–9 Nov.",
        "Busy hour (hours 20, 21 and 22): "
        f"user throughput {_move(_window_mean(bh, TREND_PRE, 'tp'), _window_mean(bh, TREND_POST, 'tp'), 'Mbps', 1)}. "
        f"Data volume {_move(_window_mean(bh, TREND_PRE, 'vol'), _window_mean(bh, TREND_POST, 'vol'), 'GB', 0)}.",
        "Daily: "
        f"user throughput {_move(_window_mean(daily, TREND_PRE, 'tp'), _window_mean(daily, TREND_POST, 'tp'), 'Mbps', 1)}. "
        f"Data volume {_move(_window_mean(daily, TREND_PRE, 'vol'), _window_mean(daily, TREND_POST, 'vol'), 'GB', 0)}.",
        "Daily cell throughput "
        f"{_move(_window_mean(daily, TREND_PRE, 'cell'), _window_mean(daily, TREND_POST, 'cell'), 'Mbps', 1)}. "
        f"DL PRB {_move(_window_mean(daily, TREND_PRE, 'prb'), _window_mean(daily, TREND_POST, 'prb'), '%', 1, points=True)}.",
    ]
    return lines


def _raised_limits(low: float, high: float) -> tuple[float, float]:
    """Place the axis floor just under the series so the movement fills the chart."""
    span = max(high - low, 0.4)
    ymin = low - 0.55 * span
    ymax = high + 0.35 * span
    if ymin < 0:
        ymin = 0.0
    if ymax <= ymin:
        ymax = ymin + 1.0
    return ymin, ymax


def _draw_trend_chart(
    path: Path,
    frame: pd.DataFrame,
    *,
    bar: str,
    line: str,
    bar_name: str,
    line_name: str,
    bar_color: str,
    left_label: str,
    pre_on: str,
) -> None:
    dates = frame.index.to_pydatetime()
    bars = frame[bar].to_numpy(dtype=float)
    line_values = frame[line].to_numpy(dtype=float)
    pre_days = list(TREND_PRE)
    pre_value = float(frame.reindex(pre_days)[line if pre_on == "line" else bar].mean())
    fig, ax = plt.subplots(figsize=(8.15, 4.25), dpi=140)
    fig.patch.set_facecolor("white")
    x = mdates.date2num(dates)
    ax.bar(x, bars, width=0.72, color=bar_color, label=bar_name, zorder=2)
    ax2 = ax.twinx()
    ax2.plot(x, line_values, color="#C459B0", lw=1.8, label=line_name, zorder=3)
    if np.isfinite(pre_value):
        host = ax2 if pre_on == "line" else ax
        host.axhline(
            pre_value,
            color="#222222",
            ls="--",
            lw=1.05,
            zorder=4,
            label="11–12 Oct throughput" if pre_on == "line" else "11–12 Oct PRB",
        )
    finite_bars = bars[np.isfinite(bars)]
    if len(finite_bars):
        low, high = float(np.min(finite_bars)), float(np.max(finite_bars))
        if pre_on == "bar" and np.isfinite(pre_value):
            low, high = min(low, pre_value), max(high, pre_value)
        ax.set_ylim(*_raised_limits(low, high))
    finite_line = line_values[np.isfinite(line_values)]
    if len(finite_line):
        low, high = float(np.min(finite_line)), float(np.max(finite_line))
        if pre_on == "line" and np.isfinite(pre_value):
            low, high = min(low, pre_value), max(high, pre_value)
        ax2.set_ylim(*_raised_limits(low, high))
    span_days = max((dates[-1] - dates[0]).days, 1) if len(dates) else 1
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=4 if span_days < 40 else 7))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
    ax.tick_params(axis="x", labelsize=8, labelrotation=40)
    ax.tick_params(axis="y", labelsize=8, colors="#595959")
    ax2.tick_params(axis="y", labelsize=8, colors="#C459B0")
    ax.set_ylabel(left_label, fontsize=8, color="#595959")
    ax2.set_ylabel("Mbps", fontsize=8, color="#C459B0")
    ax.yaxis.grid(True, color="#E6E6E6", zorder=0)
    ax.set_axisbelow(True)
    for spine in ("top",):
        ax.spines[spine].set_visible(False)
        ax2.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    ax2.spines["right"].set_color("#C459B0")
    handles, labels = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(handles + handles2, labels + labels2, loc="upper center", bbox_to_anchor=(0.5, 1.18), ncol=3, frameon=False, fontsize=7.5)
    fig.subplots_adjust(left=0.13, right=0.87, top=0.82, bottom=0.22)
    fig.savefig(path)
    plt.close(fig)


def _add_trend_slide(prs: Presentation, group: dict, chart_dir: Path, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.62))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    header.line.fill.background()
    title = slide.shapes.add_textbox(Inches(0.32), Inches(0.08), Inches(9.4), Inches(0.46))
    _set_run(title.text_frame.paragraphs[0], group["finding"], 26, RGBColor(0xFF, 0xFF, 0xFF), True)
    charts = (
        ("BH Chart", group["bh"], "vol", "tp", "Data volume, GB", "User throughput, Mbps", "#B0B0B0", "GB", "line"),
        ("Daily Chart", group["daily"], "vol", "tp", "Data volume, GB", "User throughput, Mbps", "#B0B0B0", "GB", "line"),
        ("Daily", group["daily"], "prb", "cell", "DL PRB, %", "Cell throughput, Mbps", "#7DCEA0", "DL PRB, %", "bar"),
    )
    lefts = (0.22, 4.52, 8.82)
    for left, spec in zip(lefts, charts):
        caption, frame, bar, line, bar_name, line_name, color, ylabel, pre_on = spec
        png = chart_dir / f"trend_{page:02d}_{caption.replace(' ', '_').lower()}.png"
        _draw_trend_chart(
            png,
            frame,
            bar=bar,
            line=line,
            bar_name=bar_name,
            line_name=line_name,
            bar_color=color,
            left_label=ylabel,
            pre_on=pre_on,
        )
        cap = slide.shapes.add_textbox(Inches(left), Inches(0.70), Inches(4.15), Inches(0.28))
        cap_p = cap.text_frame.paragraphs[0]
        cap_p.alignment = PP_ALIGN.CENTER
        _set_run(cap_p, caption, 13, RGBColor(0x1F, 0x4E, 0x79), True)
        slide.shapes.add_picture(str(png), Inches(left), Inches(0.98), Inches(4.22), Inches(4.15))

    notes = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.22), Inches(5.22), Inches(12.90), Inches(1.88)
    )
    notes.fill.solid()
    notes.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFB)
    notes.line.color.rgb = RGBColor(0xD0, 0xD7, 0xE2)
    box = notes.text_frame
    box.word_wrap = True
    box.margin_left = Inches(0.14)
    box.margin_right = Inches(0.12)
    box.margin_top = Inches(0.08)
    for index, line in enumerate(_trend_note(group)):
        paragraph = box.paragraphs[0] if index == 0 else box.add_paragraph()
        paragraph.space_after = Pt(3)
        _set_run(paragraph, line, 13, RGBColor(0x1F, 0x4E, 0x79) if index == 0 else RGBColor(0x33, 0x33, 0x33), index == 0)
    foot = slide.shapes.add_textbox(Inches(0.28), Inches(7.18), Inches(10.6), Inches(0.26))
    _set_run(
        foot.text_frame.paragraphs[0],
        "Dashed line is 11–12 Oct. BH sums hours 20, 21 and 22. Daily runs 15 Sep–9 Nov.",
        10,
        RGBColor(0x80, 0x80, 0x80),
        False,
    )
    page_box = slide.shapes.add_textbox(Inches(11.4), Inches(7.18), Inches(1.65), Inches(0.26))
    page_p = page_box.text_frame.paragraphs[0]
    page_p.alignment = PP_ALIGN.RIGHT
    _set_run(page_p, f"{page} / {pages}", 10, RGBColor(0x80, 0x80, 0x80), False)


def main() -> None:
    cached = pickle.loads(ROW_CACHE.read_bytes())
    rows = cached["rows"]
    latest_work = cached["latest_work"]
    pre_cache = pickle.loads(Path("/tmp/fege_report_cache.pkl").read_bytes())
    pre_work = pre_cache["work"]
    all_sites = {row["site"] for row in rows}
    print(f"Scoring hourly OSS for {len(all_sites)} comparison sites...", flush=True)
    pre_oss_all, _ = _prepare_oss(
        fege._previous_capacity_path(), all_sites, fege.ANALYSIS_START, fege.ANALYSIS_END
    )
    post_oss_all, _ = _prepare_oss(
        Path("/tmp/oss_nov_capacity.csv"), all_sites, fege.LATEST_START, fege.LATEST_END
    )
    profiles = _profiles(pre_oss_all, post_oss_all)
    selected = select_sites(rows, profiles)
    if not selected:
        raise SystemExit("No sites matched the finding selection")

    sites = {site["site"] for site in selected}
    print(f"Building charts for {len(selected)} slides...", flush=True)
    pre_feg, _ = _prepare_feg(pre_work, sites, fege.ANALYSIS_START, fege.ANALYSIS_END)
    post_feg, _ = _prepare_feg(latest_work, sites, fege.LATEST_START, fege.LATEST_END)
    pre_oss = {name: pre_oss_all[name] for name in sites}
    post_oss = {name: post_oss_all[name] for name in sites}

    CHART_DIR.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Tx BW increase: throughput and data volume"
    prs.core_properties.subject = "comparison_v1 site findings, 5–7 Oct versus 6–8 Nov"

    trends = _trend_groups(rows)
    pages = len(selected) + 1 + len(trends)
    _add_summary_slide(prs, rows, selected, profiles, trends, pages)
    for offset, group in enumerate(trends):
        page = 2 + offset
        _add_trend_slide(prs, group, CHART_DIR, page, pages)
        print(f"  trend {page}/{pages} {group['finding']} ({len(group['sites'])} sites)", flush=True)
        for line in _trend_note(group):
            print("   ", line, flush=True)
    for page, site in enumerate(selected, start=2 + len(trends)):
        name = site["site"]
        rx = _concat(pre_feg[name]["rx"], post_feg[name]["rx"])
        pre_on = _on_cap_vector(pre_work, name, fege.ANALYSIS_START, fege.ANALYSIS_END)
        post_on = _on_cap_vector(latest_work, name, fege.LATEST_START, fege.LATEST_END)
        pre_peak, pre_avg = _peaks(rx[:72], pre_on, _num(site.get("prev_center")))
        post_peak, _post_avg = _peaks(rx[72:], post_on, _num(site.get("latest_center")))
        feg_png = CHART_DIR / f"{page:02d}_{name}_feg.png"
        oss_png = CHART_DIR / f"{page:02d}_{name}_oss.png"
        _draw_feg(feg_png, rx, pre_peak, pre_avg, post_peak)
        _draw_oss(
            oss_png,
            _concat(pre_oss[name]["tp"], post_oss[name]["tp"]),
            _concat(pre_oss[name]["vol"], post_oss[name]["vol"]),
            site["busy_tp0"],
            site["busy_tp1"],
        )
        _add_slide(prs, site, feg_png, oss_png, page, pages)
        print(f"  slide {page}/{pages} {name}", flush=True)

    prs.save(OUT)
    print(f"Wrote {OUT} ({pages} slides)", flush=True)


if __name__ == "__main__":
    main()
