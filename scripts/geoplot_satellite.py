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
) -> None:
    draw = ImageDraw.Draw(img)
    font = _font(9)
    groups = {
        "4G": [r for r in rows if r["kind"] == "4G"],
        "5G": [r for r in rows if r["kind"] == "5G"],
        "Issue": [r for r in rows if r["kind"] == "Issue"],
    }
    styles = {
        "4G": (COLOR_4G, 3),
        "5G": (COLOR_5G, 4),
        "Issue": (COLOR_ISSUE, 6),
    }
    for kind in ("4G", "5G", "Issue"):
        fill, radius = styles[kind]
        for row in groups[kind]:
            x, y = _to_px(row["lat"], row["lon"], zoom, x_origin, y_origin)
            if not (0 <= x < img.width and 0 <= y < img.height):
                continue
            draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill, outline="#111111")
            if labels and kind == "Issue":
                draw.text((x + radius + 2, y - 5), row["site"], fill="#FFFFFF", font=font)


def _legend_strip(width: int) -> Image.Image:
    strip = Image.new("RGB", (width, 42), "#101418")
    draw = ImageDraw.Draw(strip)
    font = _font(16)
    items = [("4G sites", COLOR_4G), ("5G sites", COLOR_5G), ("Issue sites", COLOR_ISSUE)]
    x = 18
    for name, color in items:
        draw.ellipse((x, 12, x + 16, 28), fill=color, outline="#FFFFFF")
        draw.text((x + 24, 11), name, fill="#F5F5F5", font=font)
        x += 180
    return strip


def classify_rows(geo_rows: list[dict], issue_set: set[str]) -> list[dict]:
    classified = []
    for row in geo_rows:
        tech = _norm_tech(row.get("tech"))
        kind = "Issue" if row["site"] in issue_set else tech
        classified.append({**row, "tech": tech, "kind": kind})
    return classified


def render_jpeg(rows: list[dict], zoom_specs: list[dict], dest: Path) -> Path:
    lats = [r["lat"] for r in rows]
    lons = [r["lon"] for r in rows]
    lat_pad = max(0.04, (max(lats) - min(lats)) * 0.06)
    lon_pad = max(0.04, (max(lons) - min(lons)) * 0.06)
    national, zoom, x0, y0 = _stitch(
        min(lats) - lat_pad,
        max(lats) + lat_pad,
        min(lons) - lon_pad,
        max(lons) + lon_pad,
        980,
    )
    _draw_points(national, rows, zoom, x0, y0, labels=False)

    zoom_imgs = []
    for spec in zoom_specs:
        panel, z, zx, zy = _stitch(spec["lat_min"], spec["lat_max"], spec["lon_min"], spec["lon_max"], 520)
        _draw_points(panel, spec["rows"], z, zx, zy, labels=True)
        titled = Image.new("RGB", (panel.width, panel.height + 28), "#101418")
        titled.paste(panel, (0, 28))
        ImageDraw.Draw(titled).text(
            (10, 6),
            f"{spec['label']} — {spec['issue_count']} issue sites",
            fill="#F5F5F5",
            font=_font(15),
        )
        zoom_imgs.append(titled)

    cell_w = max((img.width for img in zoom_imgs), default=520)
    cell_h = max((img.height for img in zoom_imgs), default=400)
    right = Image.new("RGB", (cell_w * 2 + 16, cell_h * 2 + 16), "#0B0E12")
    for idx, img in enumerate(zoom_imgs):
        col, row = idx % 2, idx // 2
        right.paste(img, (col * (cell_w + 8) + 4, row * (cell_h + 8) + 4))

    legend = _legend_strip(national.width + 16 + right.width)
    canvas = Image.new("RGB", (legend.width, national.height + legend.height + 12), "#0B0E12")
    canvas.paste(legend, (0, 0))
    canvas.paste(national, (8, legend.height + 8))
    canvas.paste(right, (national.width + 16, legend.height + 8))
    dest.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(dest, format="JPEG", quality=92)
    return dest


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
    #national {{ height: 64vh; margin: 0 12px; border: 1px solid #263238; }}
    .legend {{
      position: absolute; z-index: 500; left: 24px; top: 86px;
      background: rgba(16,20,24,.86); padding: 8px 12px; border-radius: 6px; font-size: 13px;
    }}
    .swatch {{ display: inline-block; width: 12px; height: 12px; border-radius: 50%; margin-right: 6px; border: 1px solid #111; }}
    .zooms {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; padding: 12px; }}
    .zoom-wrap h2 {{ margin: 0 0 6px; font-size: 14px; color: #eceff1; }}
    .zoom {{ height: 420px; border: 1px solid #263238; }}
  </style>
</head>
<body>
  <h1>{title}</h1>
  <p class="sub">Google satellite · 4G / 5G / issue sites · zoom maps show site names only</p>
  <div style="position:relative">
    <div class="legend">
      <div><span class="swatch" style="background:{COLOR_4G}"></span>4G sites</div>
      <div><span class="swatch" style="background:{COLOR_5G}"></span>5G sites</div>
      <div><span class="swatch" style="background:{COLOR_ISSUE}"></span>Issue sites</div>
    </div>
    <div id="national"></div>
  </div>
  <div class="zooms" id="zooms"></div>
  <script>
    const DATA = {json.dumps(payload, separators=(",", ":"))};
    const googleSat = L.tileLayer("https://mt1.google.com/vt/lyrs=s&x={{x}}&y={{y}}&z={{z}}", {{
      maxZoom: 20, attribution: "Google Satellite"
    }});
    function addPoints(map, points, withLabels) {{
      const layer = L.layerGroup();
      for (const p of points) {{
        const color = DATA.colors[p.kind] || "#ffffff";
        const radius = p.kind === "Issue" ? 7 : (p.kind === "5G" ? 5 : 4);
        const marker = L.circleMarker([p.lat, p.lon], {{
          radius, color: "#111", weight: 1, fillColor: color, fillOpacity: 0.95
        }});
        if (withLabels && p.kind === "Issue") {{
          marker.bindTooltip(p.site, {{permanent: true, direction: "right", className: "site-label"}});
        }}
        marker.addTo(layer);
      }}
      layer.addTo(map);
    }}
    const national = L.map("national", {{ preferCanvas: true, zoomControl: true }});
    googleSat.addTo(national);
    addPoints(national, DATA.points, false);
    national.fitBounds(DATA.points.map(p => [p.lat, p.lon]), {{padding: [20, 20]}});
    const zooms = document.getElementById("zooms");
    DATA.zooms.forEach((z, i) => {{
      const wrap = document.createElement("div");
      wrap.className = "zoom-wrap";
      wrap.innerHTML = "<h2>" + z.label + " — " + z.issueCount + " issue sites</h2><div class=\\"zoom\\" id=\\"zoom"+i+"\\"></div>";
      zooms.appendChild(wrap);
      const map = L.map("zoom"+i, {{ preferCanvas: true }});
      L.tileLayer("https://mt1.google.com/vt/lyrs=s&x={{x}}&y={{y}}&z={{z}}", {{maxZoom: 20}}).addTo(map);
      const [s, w, n, e] = z.bounds;
      const subset = DATA.points.filter(p => p.lat >= s && p.lat <= n && p.lon >= w && p.lon <= e);
      addPoints(map, subset, true);
      map.fitBounds([[s, w], [n, e]]);
    }});
  </script>
  <style>.leaflet-tooltip.site-label {{ background: rgba(16,20,24,.75); color: #fff; border: none; font-size: 10px; box-shadow: none; }}</style>
</body>
</html>
"""
    dest.write_text(html, encoding="utf-8")
    return dest
