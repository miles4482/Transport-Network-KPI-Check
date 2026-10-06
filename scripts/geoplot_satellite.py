"""Google-satellite GeoPlot renderer for 4G / 5G coverage and 4G / 4G+5G issues."""

from __future__ import annotations

import json
import math
import time
import urllib.request
from collections import Counter
from io import BytesIO
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont

# Colours chosen to stay visible on Google satellite (dark greens / roofs).
COLOR_4G = "#00E5FF"
COLOR_5G = "#FFD600"
COLOR_ISSUE_4G = "#FF1744"
COLOR_ISSUE_4G5G = "#FF6D00"
KIND_ISSUE_4G = "Issue 4G"
KIND_ISSUE_4G5G = "Issue 4G+5G"
ISSUE_KINDS = (KIND_ISSUE_4G, KIND_ISSUE_4G5G)
SEV_ORDER = ("Severe", "High", "Moderate", "Low")
COLOR_SEV = {
    "Severe": "#FF1A1A",  # bright red — highest visibility
    "High": "#A31B1B",  # same family, less bright
    "Moderate": "#C47A12",  # muted amber — next step down
    "Low": "#1F7A3A",  # calm green — lowest emphasis
}
TILE_SIZE = 256
TILE_CACHE = Path("/tmp/google_sat_tiles")
UA = "Mozilla/5.0 (compatible; TXPortChokeGeoPlot/1.0)"


def _norm_tech(value: object) -> str:
    text = str(value or "").strip().upper().replace(" ", "")
    if "5G" in text:
        return "5G"
    if "4G" in text:
        return "4G"
    return text or "Unknown"


def _deg2num(lat: float, lon: float, zoom: int) -> tuple[float, float]:
    lat_rad = math.radians(lat)
    n = 2.0**zoom
    x = (lon + 180.0) / 360.0 * n
    y = (1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * n
    return x, y


def _fetch_tile(z: int, x: int, y: int) -> Image.Image:
    TILE_CACHE.mkdir(parents=True, exist_ok=True)
    path = TILE_CACHE / f"{z}_{x}_{y}.jpg"
    if path.exists() and path.stat().st_size > 0:
        return Image.open(path).convert("RGB")
    last_error = None
    hosts = ("mt1", "mt0", "mt2", "mt3")
    for attempt in range(6):
        host = hosts[attempt % len(hosts)]
        url = f"https://{host}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=25) as resp:
                data = resp.read()
            if data:
                path.write_bytes(data)
                return Image.open(BytesIO(data)).convert("RGB")
        except Exception as exc:  # noqa: BLE001 — tile fetch must not abort the report
            last_error = exc
            time.sleep(1.2 * (attempt + 1))
    print(f"Warning: satellite tile z{z}/{x}/{y} failed ({last_error}); using placeholder")
    return Image.new("RGB", (TILE_SIZE, TILE_SIZE), "#1a2228")


def _choose_zoom(lat_min: float, lat_max: float, lon_min: float, lon_max: float, target: int) -> int:
    for zoom in range(16, 6, -1):
        x0, y0 = _deg2num(lat_max, lon_min, zoom)
        x1, y1 = _deg2num(lat_min, lon_max, zoom)
        width = abs(x1 - x0) * TILE_SIZE
        height = abs(y1 - y0) * TILE_SIZE
        if width <= target * 1.35 and height <= target * 1.35:
            return zoom
    return 7


def _stitch(lat_min: float, lat_max: float, lon_min: float, lon_max: float, target: int) -> tuple[Image.Image, int, float, float]:
    zoom = _choose_zoom(lat_min, lat_max, lon_min, lon_max, target)
    x0, y0 = _deg2num(lat_max, lon_min, zoom)
    x1, y1 = _deg2num(lat_min, lon_max, zoom)
    ix0, iy0 = int(math.floor(min(x0, x1))), int(math.floor(min(y0, y1)))
    ix1, iy1 = int(math.floor(max(x0, x1))), int(math.floor(max(y0, y1)))
    mosaic = Image.new("RGB", ((ix1 - ix0 + 1) * TILE_SIZE, (iy1 - iy0 + 1) * TILE_SIZE))
    for ty in range(iy0, iy1 + 1):
        for tx in range(ix0, ix1 + 1):
            mosaic.paste(_fetch_tile(zoom, tx, ty), ((tx - ix0) * TILE_SIZE, (ty - iy0) * TILE_SIZE))
    left = (min(x0, x1) - ix0) * TILE_SIZE
    top = (min(y0, y1) - iy0) * TILE_SIZE
    right = (max(x0, x1) - ix0) * TILE_SIZE
    bottom = (max(y0, y1) - iy0) * TILE_SIZE
    crop = mosaic.crop((int(left), int(top), max(int(right), int(left) + 8), max(int(bottom), int(top) + 8)))
    return crop, zoom, min(x0, x1), min(y0, y1)


def _enhance_satellite(img: Image.Image) -> Image.Image:
    """Lift the dark Google satellite base so sites stay readable."""
    bright = ImageEnhance.Brightness(img).enhance(1.42)
    contrast = ImageEnhance.Contrast(bright).enhance(1.22)
    color = ImageEnhance.Color(contrast).enhance(1.12)
    return color.filter(ImageFilter.UnsharpMask(radius=1.4, percent=115, threshold=2))


def _scale_to_min_width(img: Image.Image, min_width: int) -> Image.Image:
    """Grow a map to fill its slot instead of leaving empty black bars."""
    if img.width >= min_width:
        return img
    height = max(1, int(round(img.height * (min_width / img.width))))
    return img.resize((min_width, height), Image.Resampling.LANCZOS)


def _scale_to_max_width(img: Image.Image, max_width: int) -> Image.Image:
    """Cap framed zoom maps so paired JPEGs stay compact in the workbook."""
    if img.width <= max_width:
        return img
    height = max(1, int(round(img.height * (max_width / img.width))))
    return img.resize((max_width, height), Image.Resampling.LANCZOS)


def _to_px(lat: float, lon: float, zoom: int, x_origin: float, y_origin: float) -> tuple[int, int]:
    x, y = _deg2num(lat, lon, zoom)
    return int((x - x_origin) * TILE_SIZE), int((y - y_origin) * TILE_SIZE)


def _font(size: int) -> ImageFont.ImageFont:
    for path in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ):
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def _hex_rgba(color: str, alpha: int = 255) -> tuple[int, int, int, int]:
    text = color.lstrip("#")
    return (int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16), alpha)


def _circle_stamp(radius: int, fill: str, outline: str, ring: int = 1) -> Image.Image:
    """Antialiased circle so small markers stay round instead of square."""
    scale = 4
    pad = max(ring, 1) + 2
    outer = radius + pad
    size = outer * 2 * scale
    stamp = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(stamp)
    cx = cy = size // 2
    r = radius * scale
    extra = max(ring, 1) * scale
    draw.ellipse((cx - r - extra, cy - r - extra, cx + r + extra, cy + r + extra), fill=_hex_rgba(outline))
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=_hex_rgba(fill))
    return stamp.resize((outer * 2, outer * 2), Image.Resampling.LANCZOS)


def _paste_circle(img: Image.Image, x: int, y: int, stamp: Image.Image) -> None:
    px = x - stamp.width // 2
    py = y - stamp.height // 2
    img.paste(stamp, (px, py), stamp)


def _is_issue_kind(kind: object) -> bool:
    return str(kind or "") in ISSUE_KINDS


def _issue_kind(site: str, physical_tech: str, tech_by_site: dict[str, str] | None) -> str:
    """Map an issue site to 4G or 4G+5G using the Dhaka tech list."""
    listed = str((tech_by_site or {}).get(site) or "").strip()
    if listed == "4G+5G":
        return KIND_ISSUE_4G5G
    if listed == "4G":
        return KIND_ISSUE_4G
    return KIND_ISSUE_4G5G if physical_tech == "5G" else KIND_ISSUE_4G


def _draw_points(
    img: Image.Image,
    rows: list[dict],
    zoom: int,
    x_origin: float,
    y_origin: float,
    *,
    labels: bool,
    emphasize_5g: bool = False,
) -> None:
    font = _font(9)
    draw_order = ("4G", "5G", KIND_ISSUE_4G, KIND_ISSUE_4G5G)
    groups = {kind: [r for r in rows if r["kind"] == kind] for kind in draw_order}
    # The national overview uses one compact marker size for all technologies.
    # Colour distinguishes 4G, 5G, and the two issue-tech legends.
    if emphasize_5g:
        styles = {
            "4G": (COLOR_4G, 3, "#0B0E12", 1),
            "5G": (COLOR_5G, 3, "#0B0E12", 1),
            KIND_ISSUE_4G: (COLOR_ISSUE_4G, 4, "#0B0E12", 1),
            KIND_ISSUE_4G5G: (COLOR_ISSUE_4G5G, 4, "#0B0E12", 1),
        }
    else:
        styles = {
            "4G": (COLOR_4G, 4, "#111111", 1),
            "5G": (COLOR_5G, 5, "#FFFFFF", 1),
            KIND_ISSUE_4G: (COLOR_ISSUE_4G, 7, "#111111", 1),
            KIND_ISSUE_4G5G: (COLOR_ISSUE_4G5G, 7, "#FFFFFF", 1),
        }
    stamps = {
        kind: _circle_stamp(radius, fill, outline, ring)
        for kind, (fill, radius, outline, ring) in styles.items()
    }
    if img.mode != "RGBA":
        base = img.convert("RGBA")
    else:
        base = img
    for kind in draw_order:
        _fill, radius, _outline, _ring = styles[kind]
        stamp = stamps[kind]
        for row in groups[kind]:
            x, y = _to_px(row["lat"], row["lon"], zoom, x_origin, y_origin)
            if not (0 <= x < base.width and 0 <= y < base.height):
                continue
            _paste_circle(base, x, y, stamp)
            if labels and _is_issue_kind(kind):
                draw = ImageDraw.Draw(base)
                draw.text((x + radius + 2, y - 5), row["site"], fill="#FFFFFF", font=font)
    finished = base.convert("RGB")
    img.paste(finished)


def _legend_strip(width: int, counts: dict[str, int] | None = None) -> Image.Image:
    """Two legends: coverage (4G / 5G) and issue tech (4G / 4G+5G)."""
    strip = Image.new("RGB", (width, 132), "#101418")
    draw = ImageDraw.Draw(strip)
    font = _font(26)
    group_font = _font(20)
    groups = [
        (
            "Coverage",
            18,
            [
                ("4G sites", COLOR_4G, 12, "4G"),
                ("5G sites", COLOR_5G, 11, "5G"),
            ],
        ),
        (
            "Issue sites",
            78,
            [
                ("Issue 4G", COLOR_ISSUE_4G, 12, KIND_ISSUE_4G),
                ("Issue 4G+5G", COLOR_ISSUE_4G5G, 12, KIND_ISSUE_4G5G),
            ],
        ),
    ]
    overlay = Image.new("RGBA", strip.size, (0, 0, 0, 0))
    for group_name, y, items in groups:
        draw.text((16, y - 2), group_name, fill="#90A4AE", font=group_font)
        x = 168
        for name, color, radius, key in items:
            count = (counts or {}).get(key)
            label = f"{name} ({count:,})" if count is not None else name
            stamp = _circle_stamp(radius, color, "#FFFFFF", 2)
            _paste_circle(overlay, x, y + 12, stamp)
            draw.text((x + radius + 12, y), label, fill="#F5F5F5", font=font)
            x += 310
    return Image.alpha_composite(strip.convert("RGBA"), overlay).convert("RGB")


def _legend_strip_severity(width: int, counts: dict[str, int] | None = None) -> Image.Image:
    """Severity-only legend for the national companion map."""
    strip = Image.new("RGB", (width, 132), "#101418")
    draw = ImageDraw.Draw(strip)
    font = _font(26)
    group_font = _font(20)
    draw.text((16, 48), "Severity", fill="#90A4AE", font=group_font)
    overlay = Image.new("RGBA", strip.size, (0, 0, 0, 0))
    rows = (SEV_ORDER[:2], SEV_ORDER[2:])
    for y, names in zip((18, 78), rows):
        x = 168
        for name in names:
            count = (counts or {}).get(name)
            label = f"{name} ({count:,})" if count is not None else name
            stamp = _circle_stamp(12, COLOR_SEV[name], "#FFFFFF", 2)
            _paste_circle(overlay, x, y + 12, stamp)
            draw.text((x + 20, y), label, fill="#F5F5F5", font=font)
            x += 320
    return Image.alpha_composite(strip.convert("RGBA"), overlay).convert("RGB")


def _draw_severity_points(
    img: Image.Image,
    rows: list[dict],
    zoom: int,
    x_origin: float,
    y_origin: float,
    *,
    labels: bool = False,
    compact: bool = True,
) -> None:
    """Companion map: faint coverage, issue sites coloured by severity."""
    coverage = [r for r in rows if r.get("kind") in ("4G", "5G")]
    issues = [r for r in rows if _is_issue_kind(r.get("kind"))]
    if compact:
        cov_r, sev_r, outline = 3, 5, "#0B0E12"
    else:
        cov_r, sev_r, outline = 4, 7, "#111111"
    cov_stamp = {
        "4G": _circle_stamp(cov_r, COLOR_4G, outline, 1),
        "5G": _circle_stamp(cov_r, COLOR_5G, outline, 1),
    }
    sev_stamps = {name: _circle_stamp(sev_r, COLOR_SEV[name], outline, 1) for name in SEV_ORDER}
    font = _font(9)
    if img.mode != "RGBA":
        base = img.convert("RGBA")
    else:
        base = img
    for row in coverage:
        x, y = _to_px(row["lat"], row["lon"], zoom, x_origin, y_origin)
        if 0 <= x < base.width and 0 <= y < base.height:
            _paste_circle(base, x, y, cov_stamp.get(row["kind"], cov_stamp["4G"]))
    by_sev = {name: [] for name in SEV_ORDER}
    other = []
    for row in issues:
        sev = str(row.get("severity") or "")
        if sev in by_sev:
            by_sev[sev].append(row)
        else:
            other.append(row)
    # Low first, Severe last so the worst sites sit on top.
    for name in reversed(SEV_ORDER):
        stamp = sev_stamps[name]
        for row in by_sev[name]:
            x, y = _to_px(row["lat"], row["lon"], zoom, x_origin, y_origin)
            if 0 <= x < base.width and 0 <= y < base.height:
                _paste_circle(base, x, y, stamp)
                if labels:
                    ImageDraw.Draw(base).text(
                        (x + sev_r + 2, y - 5), row["site"], fill="#FFFFFF", font=font
                    )
    for row in other:
        x, y = _to_px(row["lat"], row["lon"], zoom, x_origin, y_origin)
        if 0 <= x < base.width and 0 <= y < base.height:
            _paste_circle(base, x, y, sev_stamps["Low"])
            if labels:
                ImageDraw.Draw(base).text(
                    (x + sev_r + 2, y - 5), row["site"], fill="#FFFFFF", font=font
                )
    img.paste(base.convert("RGB"))


def severity_counts(rows: list[dict]) -> dict[str, int]:
    counts = {name: 0 for name in SEV_ORDER}
    for row in rows:
        if not _is_issue_kind(row.get("kind")):
            continue
        sev = str(row.get("severity") or "")
        if sev in counts:
            counts[sev] += 1
    return counts


def classify_rows(
    geo_rows: list[dict],
    issue_set: set[str],
    tech_by_site: dict[str, str] | None = None,
) -> list[dict]:
    classified = []
    for row in geo_rows:
        tech = _norm_tech(row.get("tech"))
        kind = _issue_kind(row["site"], tech, tech_by_site) if row["site"] in issue_set else tech
        classified.append({**row, "tech": tech, "kind": kind})
    return classified


def _frame_map(
    panel: Image.Image,
    title: str,
    counts: dict[str, int] | None = None,
    *,
    legend: str = "tech",
) -> Image.Image:
    panel = _scale_to_min_width(panel, 900)
    width = panel.width
    strip = _legend_strip_severity(width, counts) if legend == "severity" else _legend_strip(width, counts)
    title_h = 44
    canvas = Image.new("RGB", (width, panel.height + strip.height + title_h), "#0B0E12")
    canvas.paste(strip, (0, 0))
    ImageDraw.Draw(canvas).text((14, strip.height + 8), title, fill="#F5F5F5", font=_font(26))
    canvas.paste(panel, (0, strip.height + title_h))
    return canvas


def tech_counts(rows: list[dict]) -> dict[str, int]:
    return {
        "4G": sum(1 for r in rows if r.get("kind") == "4G"),
        "5G": sum(1 for r in rows if r.get("kind") == "5G"),
        KIND_ISSUE_4G: sum(1 for r in rows if r.get("kind") == KIND_ISSUE_4G),
        KIND_ISSUE_4G5G: sum(1 for r in rows if r.get("kind") == KIND_ISSUE_4G5G),
    }


def _centroid(rows: list[dict]) -> tuple[float, float]:
    return (
        sum(r["lat"] for r in rows) / len(rows),
        sum(r["lon"] for r in rows) / len(rows),
    )


def urban_rank(row: dict) -> int:
    """Lower is more urban: Dhaka Metro, then Gazipur and surrounding cities."""
    region = str(row.get("region") or "").strip().lower()
    district = str(row.get("district") or "").strip().lower()
    if region == "dhaka metro":
        return 0
    if district == "gazipur":
        return 1
    if region in ("dhaka north", "dhaka west", "dhaka south"):
        return 2
    if district == "narayanganj":
        return 3
    if district in ("narsingdi", "manikganj", "munshiganj"):
        return 4
    if district == "tangail":
        return 5
    return 6


def cluster_urban_rank(rows: list[dict]) -> int:
    ranks = [urban_rank(row) for row in rows]
    if not ranks:
        return 9
    counts = Counter(ranks)
    top = max(counts.values())
    return min(rank for rank, count in counts.items() if count == top)


def cluster_issue_sites(
    issue_rows: list[dict],
    *,
    target: int = 16,
    max_sites: int = 20,
    far: float = 0.04,
    hard: float = 0.055,
    min_merge: int = 7,
    max_merge: int = 22,
    merge_dist: float = 0.07,
) -> list[list[dict]]:
    """Group issue sites into compact neighbourhoods for medium-scale maps."""
    if not issue_rows:
        return []
    pts = np.array([[r["lat"], r["lon"]] for r in issue_rows], dtype=float)
    ranks = np.array([urban_rank(row) for row in issue_rows], dtype=int)
    unused = set(range(len(issue_rows)))
    raw: list[list[dict]] = []
    while unused:
        best_rank = min(ranks[i] for i in unused)
        candidates = [i for i in unused if ranks[i] == best_rank]
        start = candidates[0]
        best_near = -1
        for i in candidates:
            near = int(
                np.sum(
                    ((pts[list(unused)] - pts[i]) ** 2).sum(axis=1) <= far * far
                )
            )
            if near > best_near:
                start, best_near = i, near
        cluster_max = 22 if ranks[start] <= 2 else max_sites
        cluster = [start]
        unused.remove(start)
        while unused and len(cluster) < cluster_max:
            center = pts[cluster].mean(axis=0)
            rest = np.fromiter(unused, dtype=int)
            dist2 = ((pts[rest] - center) ** 2).sum(axis=1)
            nearest = int(rest[int(dist2.argmin())])
            dist = float(dist2.min()) ** 0.5
            local_target = 18 if ranks[start] <= 2 else target
            if len(cluster) >= local_target and dist > far:
                break
            if dist > hard and len(cluster) >= 8:
                break
            if dist > hard * 1.4:
                break
            cluster.append(nearest)
            unused.remove(nearest)
        raw.append([issue_rows[i] for i in cluster])

    pending = [list(group) for group in raw]
    locked: list[list[dict]] = []
    while pending:
        pending.sort(key=len)
        if len(pending[0]) >= min_merge:
            break
        small = pending.pop(0)
        sc = _centroid(small)
        best_j, best_d = None, 1e9
        for j, other in enumerate(pending):
            if len(other) + len(small) > max_merge:
                continue
            oc = _centroid(other)
            dist = ((sc[0] - oc[0]) ** 2 + (sc[1] - oc[1]) ** 2) ** 0.5
            if dist <= merge_dist and dist < best_d:
                best_d, best_j = dist, j
        if best_j is None:
            locked.append(small)
            continue
        pending[best_j].extend(small)
    clusters = pending + locked
    covered = {row["site"] for group in clusters for row in group}
    leftover = [row for row in issue_rows if row["site"] not in covered]
    if leftover:
        clusters.append(leftover)
    clusters.sort(key=lambda group: (cluster_urban_rank(group), -len(group), _centroid(group)[0]))
    return clusters


def cluster_label(rows: list[dict], index: int) -> str:
    districts = {}
    regions = {}
    for row in rows:
        districts[row.get("district", "")] = districts.get(row.get("district", ""), 0) + 1
        regions[row.get("region", "")] = regions.get(row.get("region", ""), 0) + 1
    district = max(districts, key=districts.get) if districts else "Dhaka"
    region = max(regions, key=regions.get) if regions else ""
    area = f"{district} / {region}".strip(" /")
    return f"Cluster {index:02d} — {area}"


def _norm_thana(name: object) -> str:
    text = str(name or "").strip().casefold()
    for suffix in (" thana", " than", " upazila"):
        if text.endswith(suffix):
            text = text[: -len(suffix)].strip()
    return text


def _thana_display(name: str) -> str:
    label = str(name or "").strip()
    if not label:
        return label
    if label.casefold().endswith("thana"):
        return label
    return f"{label} Thana"


def _thana_tech_counts(members: list[dict], tech_by_site: dict[str, str] | None) -> tuple[int, int]:
    n4 = 0
    n45 = 0
    listed = tech_by_site or {}
    for row in members:
        tech = listed.get(row["site"], "")
        if tech == "4G+5G":
            n45 += 1
        elif tech == "4G":
            n4 += 1
        elif "5G" in str(row.get("tech") or "").upper():
            n45 += 1
        else:
            n4 += 1
    return n4, n45


def area_bounds(rows: list[dict], pad: float = 0.008) -> tuple[float, float, float, float]:
    lats = [row["lat"] for row in rows]
    lons = [row["lon"] for row in rows]
    return min(lats) - pad, max(lats) + pad, min(lons) - pad, max(lons) + pad


def thana_focus_specs(
    classified: list[dict],
    top_thanas: list[str] | None = None,
    tech_by_site: dict[str, str] | None = None,
) -> list[dict]:
    """Thana maps listed in Top Thana.xlsx, matched to the physical Thana column."""
    specs = []
    for name in top_thanas or []:
        key = _norm_thana(name)
        if not key:
            continue
        members = [row for row in classified if _norm_thana(row.get("thana")) == key]
        if not members:
            continue
        lat_min, lat_max, lon_min, lon_max = area_bounds(members)
        member_sites = {row["site"] for row in members}
        member_issues = [row for row in members if row.get("kind") in ISSUE_KINDS]
        rows = []
        for row in classified:
            if not (lat_min <= row["lat"] <= lat_max and lon_min <= row["lon"] <= lon_max):
                continue
            if row.get("kind") in ISSUE_KINDS and row["site"] not in member_sites:
                rows.append({**row, "kind": row["tech"] if row["tech"] in ("4G", "5G") else "4G"})
            else:
                rows.append(row)
        issue_4g = sum(1 for row in member_issues if row["kind"] == KIND_ISSUE_4G)
        issue_4g5g = sum(1 for row in member_issues if row["kind"] == KIND_ISSUE_4G5G)
        n4, n45 = _thana_tech_counts(members, tech_by_site)
        title = f"{_thana_display(name)} (4G: {n4} 4G+5G: {n45})"
        slug = "Thana_" + "".join(ch if ch.isalnum() else "_" for ch in name).strip("_")
        specs.append(
            {
                "label": title,
                "title": title,
                "slug": slug,
                "focus": "thana",
                "issue_count": len(member_issues),
                "issue_4g_count": issue_4g,
                "issue_4g5g_count": issue_4g5g,
                "issue_sites": [row["site"] for row in member_issues],
                "rows": rows,
                "lat_min": lat_min,
                "lat_max": lat_max,
                "lon_min": lon_min,
                "lon_max": lon_max,
            }
        )
    return specs


def cluster_bounds(rows: list[dict]) -> tuple[float, float, float, float]:
    """Medium viewport: not a city-wide box, not a single-building crop."""
    lats = [row["lat"] for row in rows]
    lons = [row["lon"] for row in rows]
    lat_c = sum(lats) / len(lats)
    lon_c = sum(lons) / len(lons)
    lat_span = max(lats) - min(lats)
    lon_span = max(lons) - min(lons)
    half_lat = min(0.045, max(0.022, lat_span * 0.55 + 0.012))
    half_lon = min(0.050, max(0.024, lon_span * 0.55 + 0.014))
    return lat_c - half_lat, lat_c + half_lat, lon_c - half_lon, lon_c + half_lon


def render_map_images(rows: list[dict], zoom_specs: list[dict], dest_dir: Path, stem: str) -> list[tuple[str, Path]]:
    """Write one JPEG per map: national, then each zoom area."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[tuple[str, Path, str]] = []
    counts = tech_counts(rows)

    lats = [r["lat"] for r in rows]
    lons = [r["lon"] for r in rows]
    lat_pad = max(0.03, (max(lats) - min(lats)) * 0.05)
    lon_pad = max(0.03, (max(lons) - min(lons)) * 0.05)
    national, zoom, x0, y0 = _stitch(
        min(lats) - lat_pad,
        max(lats) + lat_pad,
        min(lons) - lon_pad,
        max(lons) + lon_pad,
        1200,
    )
    national = _enhance_satellite(national)
    blank = national.copy()
    _draw_points(national, rows, zoom, x0, y0, labels=False, emphasize_5g=True)
    national_path = dest_dir / f"{stem}_National.jpg"
    _frame_map(national, "National map", counts).save(national_path, format="JPEG", quality=92)
    outputs.append(("National map", national_path, "national"))

    severity_panel = blank.copy()
    _draw_severity_points(severity_panel, rows, zoom, x0, y0)
    severity_path = dest_dir / f"{stem}_Severity.jpg"
    _frame_map(
        severity_panel,
        "Issue severity map",
        severity_counts(rows),
        legend="severity",
    ).save(severity_path, format="JPEG", quality=95)
    outputs.append(("Issue severity map", severity_path, "severity"))

    for spec in zoom_specs:
        panel, z, zx, zy = _stitch(spec["lat_min"], spec["lat_max"], spec["lon_min"], spec["lon_max"], 900)
        blank = panel.copy()
        _draw_points(panel, spec["rows"], z, zx, zy, labels=True, emphasize_5g=False)
        slug = spec.get("slug") or spec["label"].replace(" ", "_").replace("/", "-")
        path = dest_dir / f"{stem}_{slug}.jpg"
        if spec.get("focus") == "thana":
            local = tech_counts(spec["rows"])
            title = spec.get("title") or spec["label"]
        else:
            local = {
                "4G": sum(1 for r in spec["rows"] if r.get("kind") == "4G"),
                "5G": sum(1 for r in spec["rows"] if r.get("kind") == "5G"),
                KIND_ISSUE_4G: spec.get("issue_4g_count", 0),
                KIND_ISSUE_4G5G: spec.get("issue_4g5g_count", 0),
            }
            title = f"{spec['label']} — {spec['issue_count']} issue sites"
        framed = _scale_to_max_width(_frame_map(panel, title, local), 900)
        framed.save(path, format="JPEG", quality=80)
        outputs.append((title, path, "zoom"))

        sev_panel = blank.copy()
        _draw_severity_points(sev_panel, spec["rows"], z, zx, zy, labels=True, compact=False)
        sev_title = "Issue severity map"
        sev_path = dest_dir / f"{stem}_{slug}_Severity.jpg"
        sev_framed = _scale_to_max_width(
            _frame_map(sev_panel, sev_title, severity_counts(spec["rows"]), legend="severity"),
            900,
        )
        sev_framed.save(sev_path, format="JPEG", quality=80)
        outputs.append((sev_title, sev_path, "zoom_sev"))
    return outputs


def render_html(rows: list[dict], zoom_specs: list[dict], dest: Path, title: str) -> Path:
    payload = {
        "points": [
            {
                "site": r["site"],
                "lat": r["lat"],
                "lon": r["lon"],
                "tech": r["tech"],
                "kind": r["kind"],
                "severity": r.get("severity") or "",
            }
            for r in rows
        ],
        "zooms": [
            {
                "label": spec["label"],
                "title": spec.get("title") or f"{spec['label']} — {spec['issue_count']} issue sites",
                "issueCount": spec["issue_count"],
                "issueSites": spec.get("issue_sites", []),
                "bounds": [spec["lat_min"], spec["lon_min"], spec["lat_max"], spec["lon_max"]],
            }
            for spec in zoom_specs
        ],
        "colors": {
            "4G": COLOR_4G,
            "5G": COLOR_5G,
            KIND_ISSUE_4G: COLOR_ISSUE_4G,
            KIND_ISSUE_4G5G: COLOR_ISSUE_4G5G,
        },
        "counts": tech_counts(rows),
        "severityCounts": severity_counts(rows),
        "severityColors": COLOR_SEV,
        "title": title,
    }
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <title>{title}</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <style>
    html, body {{ margin: 0; background: #0b0e12; color: #f5f5f5; font-family: Calibri, Arial, sans-serif; }}
    h1 {{ margin: 12px 16px 4px; font-size: 20px; }}
    .sub {{ margin: 0 16px 10px; color: #cfd8dc; font-size: 13px; }}
    nav {{ margin: 0 16px 12px; display: flex; flex-wrap: wrap; gap: 6px 12px; max-height: 22vh; overflow: auto; }}
    nav a {{ color: {COLOR_5G}; margin-right: 14px; text-decoration: none; font-size: 13px; }}
    .map-block {{ margin: 0 12px 28px; }}
    .map-block h2 {{ margin: 0 0 8px; font-size: 16px; }}
    .map-block h3 {{ margin: 0 0 6px; font-size: 14px; font-weight: 600; }}
    .map {{ height: 58vh; border: 1px solid #263238; }}
    .pair {{ display: flex; gap: 12px; align-items: stretch; }}
    .pair .pane {{ flex: 1; min-width: 0; }}
    .pair .map {{ height: 52vh; }}
    .legend {{
      background: rgba(16,20,24,.86); padding: 8px 12px; border-radius: 6px; font-size: 13px;
    }}
    .swatch {{ display: inline-block; border-radius: 50%; margin-right: 8px; border: 1px solid #fff; vertical-align: middle; }}
    .swatch-4g {{ width: 16px; height: 16px; }}
    .swatch-5g {{ width: 11px; height: 11px; }}
    .swatch-issue-4g {{ width: 13px; height: 13px; }}
    .swatch-issue-4g5g {{ width: 13px; height: 13px; }}
    .leaflet-tooltip.site-label {{ background: rgba(16,20,24,.75); color: #fff; border: none; font-size: 10px; box-shadow: none; }}
  </style>
</head>
<body>
  <h1>{title}</h1>
      <p class="sub">Google satellite · National + severity maps, then top thanas and issue clusters each with a severity map beside them · Coverage: cyan 4G / yellow 5G · Issue: red 4G / orange 4G+5G · Severity map: bright red Severe / deep red High / muted amber Moderate / calm green Low</p>
  <nav id="nav"></nav>
  <div id="maps"></div>
  <script>
    const DATA = {json.dumps(payload, separators=(",", ":"))};
    function addPoints(map, points, withLabels, national, severity) {{
      const layer = L.layerGroup();
      const ordered = [
        ...points.filter(p => p.kind === "4G"),
        ...points.filter(p => p.kind === "5G"),
        ...points.filter(p => p.kind === "Issue 4G"),
        ...points.filter(p => p.kind === "Issue 4G+5G"),
      ];
      for (const p of ordered) {{
        const isIssue = p.kind === "Issue 4G" || p.kind === "Issue 4G+5G";
        const color = severity && isIssue
          ? (DATA.severityColors[p.severity] || "#43A047")
          : (DATA.colors[p.kind] || "#ffffff");
        const radius = national
          ? (severity && isIssue ? 4 : 3)
          : (isIssue ? 7 : (p.kind === "5G" ? 6 : 5));
        const marker = L.circleMarker([p.lat, p.lon], {{
          radius, color: "#111", weight: 1,
          fillColor: color, fillOpacity: 0.98
        }});
        if (withLabels && isIssue) {{
          marker.bindTooltip(p.site, {{permanent: true, direction: "right", className: "site-label"}});
        }}
        marker.addTo(layer);
      }}
      layer.addTo(map);
    }}
    function kindCounts(points) {{
      return {{
        "4G": points.filter(p => p.kind === "4G").length,
        "5G": points.filter(p => p.kind === "5G").length,
        "Issue 4G": points.filter(p => p.kind === "Issue 4G").length,
        "Issue 4G+5G": points.filter(p => p.kind === "Issue 4G+5G").length,
      }};
    }}
    function sevCounts(points) {{
      const c = {{Severe: 0, High: 0, Moderate: 0, Low: 0}};
      for (const p of points) {{
        if ((p.kind === "Issue 4G" || p.kind === "Issue 4G+5G") && Object.prototype.hasOwnProperty.call(c, p.severity)) {{
          c[p.severity]++;
        }}
      }}
      return c;
    }}
    function makeMap(id, points, bounds, withLabels, national, severity, localSev) {{
      const map = L.map(id, {{ preferCanvas: true }});
      L.tileLayer("https://mt1.google.com/vt/lyrs=s&x={{x}}&y={{y}}&z={{z}}", {{
        maxZoom: 20, attribution: "Google Satellite"
      }}).addTo(map);
      addPoints(map, points, withLabels, national, severity);
      map.fitBounds(bounds, {{padding: [20, 20]}});
      const legend = L.control({{position: "topright"}});
      legend.onAdd = function () {{
        const div = L.DomUtil.create("div", "legend");
        if (severity) {{
          const c = localSev || DATA.severityCounts;
          const col = DATA.severityColors;
          div.innerHTML = "<div style='color:#90A4AE;margin-bottom:4px'>Severity</div>"
            + "<div><span class='swatch' style='background:" + col.Severe + ";width:13px;height:13px'></span>Severe (" + (c.Severe || 0).toLocaleString() + ")</div>"
            + "<div><span class='swatch' style='background:" + col.High + ";width:13px;height:13px'></span>High (" + (c.High || 0).toLocaleString() + ")</div>"
            + "<div><span class='swatch' style='background:" + col.Moderate + ";width:13px;height:13px'></span>Moderate (" + (c.Moderate || 0).toLocaleString() + ")</div>"
            + "<div><span class='swatch' style='background:" + col.Low + ";width:13px;height:13px'></span>Low (" + (c.Low || 0).toLocaleString() + ")</div>";
        }} else {{
          const c = kindCounts(points);
          div.innerHTML = "<div style='color:#90A4AE;margin-bottom:4px'>Coverage</div>"
            + "<div><span class='swatch swatch-4g' style='background:{COLOR_4G}'></span>4G sites (" + c["4G"].toLocaleString() + ")</div>"
            + "<div><span class='swatch swatch-5g' style='background:{COLOR_5G}'></span>5G sites (" + c["5G"].toLocaleString() + ")</div>"
            + "<div style='color:#90A4AE;margin:8px 0 4px'>Issue sites</div>"
            + "<div><span class='swatch swatch-issue-4g' style='background:{COLOR_ISSUE_4G}'></span>Issue 4G (" + (c["Issue 4G"] || 0).toLocaleString() + ")</div>"
            + "<div><span class='swatch swatch-issue-4g5g' style='background:{COLOR_ISSUE_4G5G}'></span>Issue 4G+5G (" + (c["Issue 4G+5G"] || 0).toLocaleString() + ")</div>";
        }}
        return div;
      }};
      legend.addTo(map);
    }}
    const maps = document.getElementById("maps");
    const nav = document.getElementById("nav");
    function addSingle(block) {{
      nav.innerHTML += "<a href='#" + block.id + "-block'>" + block.title + "</a>";
      const wrap = document.createElement("section");
      wrap.className = "map-block";
      wrap.id = block.id + "-block";
      wrap.innerHTML = "<h2>" + block.title + "</h2><div class='map' id='" + block.id + "'></div>";
      maps.appendChild(wrap);
      makeMap(block.id, block.points, block.bounds, block.labels, block.national, block.severity);
    }}
    addSingle({{
      id: "national",
      title: "National map",
      points: DATA.points,
      bounds: DATA.points.map(p => [p.lat, p.lon]),
      labels: false,
      national: true,
      severity: false
    }});
    addSingle({{
      id: "severity",
      title: "Issue severity map",
      points: DATA.points,
      bounds: DATA.points.map(p => [p.lat, p.lon]),
      labels: false,
      national: true,
      severity: true
    }});
    DATA.zooms.forEach((z, i) => {{
      const [s, w, n, e] = z.bounds;
      const title = z.title || (z.label + " — " + z.issueCount + " issue sites");
      const points = DATA.points.filter(p => p.lat >= s && p.lat <= n && p.lon >= w && p.lon <= e).map(p => {{
        if ((p.kind === "Issue 4G" || p.kind === "Issue 4G+5G") && z.issueSites.length && !z.issueSites.includes(p.site)) {{
          return Object.assign({{}}, p, {{kind: p.tech === "5G" ? "5G" : "4G"}});
        }}
        return p;
      }});
      const bounds = [[s, w], [n, e]];
      nav.innerHTML += "<a href='#zoom" + i + "-block'>" + title + "</a>";
      const wrap = document.createElement("section");
      wrap.className = "map-block";
      wrap.id = "zoom" + i + "-block";
      wrap.innerHTML = "<h2>" + title + "</h2><div class='pair'><div class='pane'><h3>" + title
        + "</h3><div class='map' id='zoom" + i + "'></div></div><div class='pane'><h3>Issue severity map</h3>"
        + "<div class='map' id='zoom" + i + "sev'></div></div></div>";
      maps.appendChild(wrap);
      makeMap("zoom" + i, points, bounds, true, false, false);
      makeMap("zoom" + i + "sev", points, bounds, true, false, true, sevCounts(points));
    }});
  </script>
</body>
</html>
"""
    dest.write_text(html, encoding="utf-8")
    return dest
