#!/usr/bin/env python3
"""Management PPT: one slide per site for comparison_v1 findings.

Pre window is 5–7 Oct. Post window is 6–8 Nov. Each finding category
contributes up to five sites, then those slides are ordered by the
daytime throughput step, largest first:

- throughput and data volume both improved significantly
- throughput improved significantly, data volume held or rose only a little
- a smaller gain, with neither KPI down

Sites where both throughput and data volume fell are left out.
"""

from __future__ import annotations

import pickle
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
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

OUT = Path("Comparison_v1_Findings_5Oct_vs_6Nov_v7.pptx")
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
BUCKET_LABEL = {
    BUCKET_BOTH: "Both throughput and data volume improved",
    BUCKET_HOLD: "Throughput improved, data volume held",
    BUCKET_SMALL: "Small improvement, KPIs did not fall",
}


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


def _profiles(pre: dict, post: dict) -> dict[str, dict]:
    out = {}
    for site in pre:
        if site not in post:
            continue
        tp0, tp1 = _daytime(pre[site]["tp"]), _daytime(post[site]["tp"])
        vol0, vol1 = _daytime(pre[site]["vol"]), _daytime(post[site]["vol"])
        out[site] = {
            "busy_tp0": tp0,
            "busy_tp1": tp1,
            "busy_tp_delta": tp1 - tp0 if np.isfinite(tp0) and np.isfinite(tp1) else float("nan"),
            "busy_tp_pct": (100.0 * (tp1 - tp0) / tp0) if tp0 else float("nan"),
            "busy_vol0": vol0,
            "busy_vol1": vol1,
            "busy_vol_pct": (100.0 * (vol1 - vol0) / vol0) if vol0 else float("nan"),
            "min_hours": min(_min_day_hours(pre[site]["tp"]), _min_day_hours(post[site]["tp"])),
        }
    return out


def select_sites(rows: list[dict], profiles: dict[str, dict]) -> list[dict]:
    """Up to five sites in each finding category, in the requested mix.

    A site is kept only when the hourly curve does not contradict the note:
    daytime throughput must not have fallen, and a very small data volume
    is left out so a percent change is not an empty-site artefact.
    Sites whose daytime throughput actually stepped up are taken first,
    because that is what the chart shows.
    """
    grouped: dict[str, list[dict]] = {name: [] for name in CATEGORY_ORDER}
    for row in rows:
        metrics = _metrics(row)
        profile = profiles.get(row["site"])
        if metrics is None or profile is None or row["finding"] not in grouped:
            continue
        bucket = _bucket(metrics)
        if not bucket:
            continue
        if metrics["vol0"] < 10 or metrics["vol1"] < 10:
            continue
        if profile["min_hours"] < 20:
            continue
        if not np.isfinite(profile["busy_tp_pct"]) or profile["busy_tp_pct"] <= -5:
            continue
        if bucket == BUCKET_BOTH and (
            not np.isfinite(profile["busy_vol_pct"]) or profile["busy_vol_pct"] <= -5
        ):
            continue
        grouped[row["finding"]].append(
            {**row, "metrics": metrics, "bucket": bucket, **profile}
        )

    quotas = ((BUCKET_BOTH, 2), (BUCKET_HOLD, 2), (BUCKET_SMALL, 1))
    picked: list[dict] = []
    for finding in CATEGORY_ORDER:
        pool = grouped[finding]
        chosen: list[dict] = []
        used: set[str] = set()
        for bucket, quota in quotas:
            options = [
                site for site in pool if site["bucket"] == bucket and site["site"] not in used
            ]
            options.sort(
                key=lambda site: (
                    site["busy_tp_delta"] >= 0.8 and site["busy_tp_pct"] >= 5,
                    site["busy_tp_delta"],
                    site["metrics"]["tp_delta"],
                    site["site"],
                ),
                reverse=True,
            )
            for site in options[:quota]:
                chosen.append(site)
                used.add(site["site"])
        # Largest visible daytime throughput step first, in every category.
        chosen.sort(
            key=lambda site: (
                site["busy_tp_delta"] if np.isfinite(site["busy_tp_delta"]) else -1e9,
                site["busy_tp_pct"] if np.isfinite(site["busy_tp_pct"]) else -1e9,
            ),
            reverse=True,
        )
        for index, site in enumerate(chosen, start=1):
            site["pick_index"] = index
            site["pick_total"] = len(chosen)
            picked.append(site)
        print(f"{finding}: {len(chosen)} sites", flush=True)
        for site in chosen:
            m = site["metrics"]
            print(
                f"  {site['pick_index']}. {site['site']}  {site['bucket']}"
                f"  1H TP {m['tp0']:.2f}->{m['tp1']:.2f} ({m['tp_pct']:+.1f}%)"
                f"  Vol {m['vol0']:.1f}->{m['vol1']:.1f} ({m['vol_pct']:+.1f}%)"
                f"  daytime {site['busy_tp0']:.1f}->{site['busy_tp1']:.1f}"
                f" ({site['busy_tp_pct']:+.0f}%)"
                f"  Rx {site['old_rx']:.0f}->{site['new_rx']:.0f}",
                flush=True,
            )
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
        ["eNodeB Name", "Date", "Hour", "Rx"],
    ].copy()
    return _hour_matrix(part, sites, {"rx": "Rx"}, start, end)


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
        (_guide(ax, x[:72], pre_peak, "#548235", "--"), "TP peak pre"),
        (_guide(ax, x[72:], post_peak, "#203864", "--"), "TP peak post"),
        (_guide(ax, x[:72], pre_low, "#7F7F7F", "--"), "TP low pre"),
        (_guide(ax, x[72:], post_low, "#000000", "--"), "TP low post"),
        # One volume average per window: mean of 20:00–23:00, not a peak and a low.
        (_guide(ax2, x[:72], _evening_volume_mean(vol[:72]), "#843C0C", (0, (6, 2)), 1.8), "Vol 20–23 pre"),
        (_guide(ax2, x[72:], _evening_volume_mean(vol[72:]), "#C65911", (0, (1, 1)), 2.0), "Vol 20–23 post"),
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
        fontsize=6.5,
        borderaxespad=0,
        columnspacing=0.9,
        handlelength=1.7,
        handletextpad=0.35,
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
    }[site["bucket"]]

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


def _add_summary_slide(prs: Presentation, selected: list[dict], pages: int) -> None:
    """Ranked table: finding, then sites from the largest daytime step down."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    navy = RGBColor(0x1F, 0x4E, 0x79)
    ink = RGBColor(0x33, 0x33, 0x33)
    grey = RGBColor(0x60, 0x60, 0x60)
    white = RGBColor(0xFF, 0xFF, 0xFF)
    readings = {
        fege.FINDING_UP_CHOKED: "Ceiling stepped up. Newly choked on 6–8 Nov.",
        fege.FINDING_SOLVED_UP: "Ceiling stepped up. Off the choke list.",
        fege.FINDING_UP: "Ceiling stepped up. Not choked in either window.",
        fege.FINDING_AGAIN: "Still choked. The gain is on the new ceiling.",
        fege.FINDING_SOLVED: "Off the choke list. No Tx BW step-up.",
    }
    rows = []
    for finding in CATEGORY_ORDER:
        sites = [site for site in selected if site["finding"] == finding]
        if sites:
            rows.append((finding, sites))

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.62))
    bar.fill.solid()
    bar.fill.fore_color.rgb = navy
    bar.line.fill.background()
    title = slide.shapes.add_textbox(Inches(0.32), Inches(0.10), Inches(10.5), Inches(0.42))
    _set_run(title.text_frame.paragraphs[0], "Management summary", 26, white, True)

    intro = slide.shapes.add_textbox(Inches(0.32), Inches(0.72), Inches(12.7), Inches(0.36))
    intro.text_frame.word_wrap = False
    _set_run(
        intro.text_frame.paragraphs[0],
        "Pre 5–7 Oct versus post 6–8 Nov. Sites in each row run from the largest daytime throughput step to the smallest.",
        14,
        navy,
        False,
    )

    row_heights = [0.36]
    for _, sites in rows:
        row_heights.append(0.18 + 0.22 * len(sites))
    table_height = sum(row_heights)
    table_shape = slide.shapes.add_table(
        1 + len(rows), 3, Inches(0.28), Inches(1.16), Inches(12.78), Inches(table_height)
    )
    table = table_shape.table
    table.columns[0].width = Inches(0.62)
    table.columns[1].width = Inches(4.35)
    table.columns[2].width = Inches(7.81)
    tbl_pr = table._tbl.tblPr
    if tbl_pr is not None:
        tbl_pr.set("bandRow", "0")
        tbl_pr.set("firstRow", "0")

    headers = ("#", "Finding", "Sites — largest daytime gain first")
    for col, text in enumerate(headers):
        cell = table.cell(0, col)
        _set_cell_fill(cell, "1F4E79")
        _set_cell_border(cell, "1F4E79", 0.75)
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.04)
        cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER if col == 0 else PP_ALIGN.LEFT
        _set_run(paragraph, text, 12, white, True)

    for index, ((finding, sites), height) in enumerate(zip(rows, row_heights[1:]), start=1):
        table.rows[index].height = Inches(height)
        rank = table.cell(index, 0)
        name = table.cell(index, 1)
        detail = table.cell(index, 2)
        fill = "F7F9FB" if index % 2 else "FFFFFF"
        for cell in (rank, name, detail):
            _set_cell_fill(cell, "E8EEF4" if cell is rank else fill)
            _set_cell_border(cell, "D0D7E2", 0.6)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.word_wrap = True
        rank_p = rank.text_frame.paragraphs[0]
        rank_p.alignment = PP_ALIGN.CENTER
        _set_run(rank_p, str(index), 18, navy, True)

        name_p = name.text_frame.paragraphs[0]
        name_p.space_after = Pt(1)
        _set_run(name_p, finding, 13, navy, True)
        meaning = name.text_frame.add_paragraph()
        meaning.space_before = Pt(1)
        _set_run(meaning, readings[finding], 11, grey, False)
        count = name.text_frame.add_paragraph()
        count.space_before = Pt(2)
        label = f"{len(sites)} site" + ("s" if len(sites) != 1 else "")
        _set_run(count, label, 11, ink, False)

        for site_index, site in enumerate(sites):
            paragraph = detail.text_frame.paragraphs[0] if site_index == 0 else detail.text_frame.add_paragraph()
            paragraph.space_after = Pt(0)
            paragraph.space_before = Pt(0)
            _set_run(paragraph, site["site"], 12, navy, True)
            _add_run(
                paragraph,
                f"    {site['busy_tp0']:.1f} → {site['busy_tp1']:.1f} Mbps"
                f"    ({site['busy_tp_delta']:+.1f})",
                12,
                ink,
                False,
            )
    table.rows[0].height = Inches(row_heights[0])

    foot_top = 1.16 + table_height + 0.10
    foot = slide.shapes.add_textbox(Inches(0.32), Inches(foot_top), Inches(10.4), Inches(0.28))
    _set_run(
        foot.text_frame.paragraphs[0],
        "Daytime throughput is the 09:00–22:00 median. The following pages are one site each, in this order.",
        11,
        RGBColor(0x80, 0x80, 0x80),
        False,
    )
    page_box = slide.shapes.add_textbox(Inches(11.4), Inches(foot_top), Inches(1.65), Inches(0.28))
    page_p = page_box.text_frame.paragraphs[0]
    page_p.alignment = PP_ALIGN.RIGHT
    _set_run(page_p, f"1 / {pages}", 11, RGBColor(0x80, 0x80, 0x80), False)


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
    selected = select_sites(rows, _profiles(pre_oss_all, post_oss_all))
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

    pages = len(selected) + 1
    _add_summary_slide(prs, selected, pages)
    for page, site in enumerate(selected, start=2):
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
