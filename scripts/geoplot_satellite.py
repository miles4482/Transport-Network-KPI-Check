"""Google-satellite GeoPlot renderer for 4G / 5G / issue sites."""

from __future__ import annotations

import json
import math
import urllib.request
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# Colours chosen to stay visible on Google satellite (dark greens / roofs).
COLOR_4G = "#00E5FF"
COLOR_5G = "#FFD600"
COLOR_ISSUE = "#FF1744"
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
    if path.exists():
        return Image.open(path).convert("RGB")
    url = f"https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}"
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = resp.read()
    path.write_bytes(data)
    return Image.open(BytesIO(data)).convert("RGB")


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
    draw = ImageDraw.Draw(img)
    font = _font(9)
    groups = {
        "4G": [r for r in rows if r["kind"] == "4G"],
        "5G": [r for r in rows if r["kind"] == "5G"],
        "Issue": [r for r in rows if r["kind"] == "Issue"],
    }
    # On the national map, 5G is a small set inside a dense 4G layer, so it
    # needs a larger marker and a white halo to stay visible.
    if emphasize_5g:
        styles = {
            "4G": (COLOR_4G, 4, "#0B0E12", 1),
            "5G": (COLOR_5G, 6, "#FFFFFF", 1),
            "Issue": (COLOR_ISSUE, 7, "#FFFFFF", 2),
        }
    else:
        styles = {
            "4G": (COLOR_4G, 4, "#111111", 1),
            "5G": (COLOR_5G, 5, "#FFFFFF", 1),
            "Issue": (COLOR_ISSUE, 7, "#111111", 1),
        }
    for kind in ("4G", "5G", "Issue"):
        fill, radius, outline, ring = styles[kind]
        for row in groups[kind]:
            x, y = _to_px(row["lat"], row["lon"], zoom, x_origin, y_origin)
            if not (0 <= x < img.width and 0 <= y < img.height):
                continue
            if ring > 1:
                draw.ellipse(
                    (x - radius - ring, y - radius - ring, x + radius + ring, y + radius + ring),
                    fill=outline,
                )
            draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill, outline=outline)
            if labels and kind == "Issue":
                draw.text((x + radius + 2, y - 5), row["site"], fill="#FFFFFF", font=font)


def _legend_strip(width: int, counts: dict[str, int] | None = None) -> Image.Image:
    strip = Image.new("RGB", (width, 48), "#101418")
    draw = ImageDraw.Draw(strip)
    font = _font(16)
    items = [
        ("4G sites", COLOR_4G, 9),
        ("5G sites", COLOR_5G, 6),
        ("Issue sites", COLOR_ISSUE, 7),
    ]
    x = 18
    for name, color, radius in items:
        count = (counts or {}).get(name.split()[0], None)
        label = f"{name} ({count:,})" if count is not None else name
        cy = 24
        draw.ellipse((x - radius, cy - radius, x + radius, cy + radius), fill=color, outline="#FFFFFF")
        draw.text((x + radius + 8, 14), label, fill="#F5F5F5", font=font)
        x += 220
    return strip


def classify_rows(geo_rows: list[dict], issue_set: set[str]) -> list[dict]:
    classified = []
    for row in geo_rows:
        tech = _norm_tech(row.get("tech"))
        kind = "Issue" if row["site"] in issue_set else tech
        classified.append({**row, "tech": tech, "kind": kind})
    return classified


def _frame_map(panel: Image.Image, title: str, counts: dict[str, int] | None = None) -> Image.Image:
    legend = _legend_strip(max(panel.width, 780), counts)
    canvas = Image.new("RGB", (max(panel.width, legend.width), panel.height + legend.height + 36), "#0B0E12")
    canvas.paste(legend, (0, 0))
    ImageDraw.Draw(canvas).text((12, legend.height + 6), title, fill="#F5F5F5", font=_font(18))
    canvas.paste(panel, (0, legend.height + 32))
    return canvas


def tech_counts(rows: list[dict]) -> dict[str, int]:
    return {
        "4G": sum(1 for r in rows if r.get("tech") == "4G"),
        "5G": sum(1 for r in rows if r.get("tech") == "5G"),
        "Issue": sum(1 for r in rows if r.get("kind") == "Issue"),
    }


def render_map_images(rows: list[dict], zoom_specs: list[dict], dest_dir: Path, stem: str) -> list[tuple[str, Path]]:
    """Write one JPEG per map: national, then each zoom area."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    outputs: list[tuple[str, Path]] = []
    counts = tech_counts(rows)

    lats = [r["lat"] for r in rows]
    lons = [r["lon"] for r in rows]
    lat_pad = max(0.04, (max(lats) - min(lats)) * 0.06)
    lon_pad = max(0.04, (max(lons) - min(lons)) * 0.06)
    national, zoom, x0, y0 = _stitch(
        min(lats) - lat_pad,
        max(lats) + lat_pad,
        min(lons) - lon_pad,
        max(lons) + lon_pad,
        1200,
    )
    _draw_points(national, rows, zoom, x0, y0, labels=False, emphasize_5g=True)
    national_path = dest_dir / f"{stem}_National.jpg"
    _frame_map(national, "National map", counts).save(national_path, format="JPEG", quality=92)
    outputs.append(("National map", national_path))

    for spec in zoom_specs:
        panel, z, zx, zy = _stitch(spec["lat_min"], spec["lat_max"], spec["lon_min"], spec["lon_max"], 1100)
        _draw_points(panel, spec["rows"], z, zx, zy, labels=True, emphasize_5g=True)
        slug = spec["label"].replace(" ", "_")
        path = dest_dir / f"{stem}_{slug}.jpg"
        _frame_map(panel, f"{spec['label']} — {spec['issue_count']} issue sites", counts).save(
            path, format="JPEG", quality=92
        )
        outputs.append((f"{spec['label']} — {spec['issue_count']} issue sites", path))
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
            }
            for r in rows
        ],
        "zooms": [
            {
                "label": spec["label"],
                "issueCount": spec["issue_count"],
                "bounds": [spec["lat_min"], spec["lon_min"], spec["lat_max"], spec["lon_max"]],
            }
            for spec in zoom_specs
        ],
        "colors": {"4G": COLOR_4G, "5G": COLOR_5G, "Issue": COLOR_ISSUE},
        "counts": tech_counts(rows),
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
    nav {{ margin: 0 16px 12px; }}
    nav a {{ color: {COLOR_5G}; margin-right: 14px; text-decoration: none; font-size: 14px; }}
    .map-block {{ margin: 0 12px 28px; }}
    .map-block h2 {{ margin: 0 0 8px; font-size: 16px; }}
    .map {{ height: 72vh; border: 1px solid #263238; }}
    .legend {{
      background: rgba(16,20,24,.86); padding: 8px 12px; border-radius: 6px; font-size: 13px;
    }}
    .swatch {{ display: inline-block; border-radius: 50%; margin-right: 8px; border: 1px solid #fff; vertical-align: middle; }}
    .swatch-4g {{ width: 16px; height: 16px; }}
    .swatch-5g {{ width: 11px; height: 11px; }}
    .swatch-issue {{ width: 13px; height: 13px; }}
    .leaflet-tooltip.site-label {{ background: rgba(16,20,24,.75); color: #fff; border: none; font-size: 10px; box-shadow: none; }}
  </style>
</head>
<body>
  <h1>{title}</h1>
  <p class="sub">Google satellite · each map is separate · 4G / 5G / issue legend uses database Tech counts</p>
  <nav id="nav"></nav>
  <div id="maps"></div>
  <script>
    const DATA = {json.dumps(payload, separators=(",", ":"))};
    function addPoints(map, points, withLabels, national) {{
      const layer = L.layerGroup();
      const ordered = [
        ...points.filter(p => p.kind === "4G"),
        ...points.filter(p => p.kind === "5G"),
        ...points.filter(p => p.kind === "Issue"),
      ];
      for (const p of ordered) {{
        const color = DATA.colors[p.kind] || "#ffffff";
        const radius = p.kind === "5G" ? 6 : (p.kind === "Issue" ? 7 : 5);
        const marker = L.circleMarker([p.lat, p.lon], {{
          radius, color: "#111", weight: 1,
          fillColor: color, fillOpacity: 0.98
        }});
        if (withLabels && p.kind === "Issue") {{
          marker.bindTooltip(p.site, {{permanent: true, direction: "right", className: "site-label"}});
        }}
        marker.addTo(layer);
      }}
      layer.addTo(map);
    }}
    function makeMap(id, points, bounds, withLabels, national) {{
      const map = L.map(id, {{ preferCanvas: true }});
      L.tileLayer("https://mt1.google.com/vt/lyrs=s&x={{x}}&y={{y}}&z={{z}}", {{
        maxZoom: 20, attribution: "Google Satellite"
      }}).addTo(map);
      addPoints(map, points, withLabels, national);
      map.fitBounds(bounds, {{padding: [20, 20]}});
      const legend = L.control({{position: "topright"}});
      legend.onAdd = function () {{
        const div = L.DomUtil.create("div", "legend");
        const c = DATA.counts;
        div.innerHTML = "<div><span class='swatch swatch-4g' style='background:{COLOR_4G}'></span>4G sites (" + c["4G"].toLocaleString() + ")</div>"
          + "<div><span class='swatch swatch-5g' style='background:{COLOR_5G}'></span>5G sites (" + c["5G"].toLocaleString() + ")</div>"
          + "<div><span class='swatch swatch-issue' style='background:{COLOR_ISSUE}'></span>Issue sites (" + c.Issue.toLocaleString() + ")</div>";
        return div;
      }};
      legend.addTo(map);
    }}
    const maps = document.getElementById("maps");
    const nav = document.getElementById("nav");
    const blocks = [
      {{
        id: "national",
        title: "National map",
        points: DATA.points,
        bounds: DATA.points.map(p => [p.lat, p.lon]),
        labels: false,
        national: true
      }},
      ...DATA.zooms.map((z, i) => {{
        const [s, w, n, e] = z.bounds;
        return {{
          id: "zoom" + i,
          title: z.label + " — " + z.issueCount + " issue sites",
          points: DATA.points.filter(p => p.lat >= s && p.lat <= n && p.lon >= w && p.lon <= e),
          bounds: [[s, w], [n, e]],
          labels: true,
          national: false
        }};
      }})
    ];
    blocks.forEach(block => {{
      nav.innerHTML += "<a href='#" + block.id + "-block'>" + block.title + "</a>";
      const wrap = document.createElement("section");
      wrap.className = "map-block";
      wrap.id = block.id + "-block";
      wrap.innerHTML = "<h2>" + block.title + "</h2><div class='map' id='" + block.id + "'></div>";
      maps.appendChild(wrap);
      makeMap(block.id, block.points, block.bounds, block.labels, block.national);
    }});
  </script>
</body>
</html>
"""
    dest.write_text(html, encoding="utf-8")
    return dest
