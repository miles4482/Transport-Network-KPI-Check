# FEGE choke check — V14 freeze

Saved so later work can continue from this version. The live implementation is `fege_choke_check.py`. Default output is `FEGE_Choked_Flat_Sites_v14.xlsx`.

## Intent

Flag 5G FEGE ports where `RxMaxSpeed` is **choked**: the trace is pinned and **does not go upward** above a fixed cap. Reference sites that must always appear: **DHAPT35**, **DHAPT48**.

## Counters

- `RxMaxSpeed (Mbit/s)` = `VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000`
- `Tx Total BW (Mbit/s)` = `VS.FEGE.TxTotalBW(kbit/s) / 1000`
- Finding: At Tx BW 85–120%, Above > 120%, Below < 85%

## Window and listing rules

- Latest **7 days** in the source (this file: 16–22 Sep 2026).
- Score **every hour** (busy 08:00–22:00 and off-peak / night).
- List if ≥ **3 of 7 days** have ≥ **3 hours** on the cap, **or** the last day has ≥ **2 hours** on the cap.
- Last-day flag is a separate dated column (`22-Sep-26 · Nh (busy X, off-peak Y)`).
- Snapshot charts show only the latest **3 days**.

## Two cap shapes (V14)

Tried in this order; first that passes tightness + day rules wins.

1. **Crowded level** (`find_plateau`)
   - Search the upper half of the trace (quantile 0.55 to max).
   - Tolerance ±4 Mbit/s or ±1.5% of p95.
   - 99th percentile may sit at most 40 Mbit/s or 12% of the level above it.
   - Sample-snap shape: line locks after the night dip.

2. **Hard ceiling** (`find_ceiling`)
   - Same search, but only over the top 10% of samples.
   - At most 1 hour may sit more than 8 Mbit/s or 3% above the top.
   - Catches traces that move during the day but are clipped at one value.

In-band standard deviation must be ≤ 2.5 Mbit/s or 1% of the level. Center ≥ 20 Mbit/s.

## Rejected outliers

Do not list traces whose peaks keep changing (spikes of tens to hundreds of Mbit/s, no fixed top). User examples: **DHGULN2, DHDHN40, DHBDD32, DHDHN09**.

## V14 result on `FEG_KPI_5G Site_DHK.xlsb`

78 listed of 356 checked: 66 Crowded level + 12 Hard ceiling.

Hard ceiling adds: DHDHN04, DHGUL16, DHGUL42, DHKKT87, DHKKTE2, DHSVRC0, DHSVRS5, DHTIA78, GPSDR01, GPSDR16, GPSDR3M, GPSDRL1.

## Chart (frozen)

2-level category axis: date only at 00:00 (blank 01:00–23:00 so Excel groups the day once); all 24 hours rotated −90. No axis title. Size 1200 × 410. Y major unit 50. Hidden `_ChartData` sheet.

## Later work

Keep this **detection rule** as the baseline. Presentation changed in **v15** (same 78 sites, same two shapes):

- Report title: **TX Port Choke Check**
- Do not grade sites as At / Above / Below Tx BW. RAN-side Tx Total BW is a column only.
- Severity from hours on cap: Severe ≥50%, High ≥25%, Moderate ≥10%, Low <10%.
- Sheet `2. Snapshots` renamed to `2. HourlyChartOfIssueSites`.
- The long note under each chart KPI strip is removed.
- File: `FEGE_Choked_Flat_Sites_v15.xlsx`
- Each hourly chart has a red dashed straight line at **Stuck RxMaxSpeed**.
- v17 adds sheet `5. GeoPlot`: Dark Red = issue sites, WhiteSmoke = non-issue sites, from `Physical_Site_Database_24Sep26.xlsx`.
- v18: `5. GeoPlot` is map-only (no visible table). The full 21,693-site physical database forms the Bangladesh map: 78 issue sites are Dark Red and the other 21,615 physical sites are WhiteSmoke, with exactly two legend entries and no connecting lines. Four area-level inset maps zoom Dhaka Metro, Dhaka North, Dhaka West, and Gazipur; inset legends are suppressed so the national map remains the only two-entry legend. Coordinate series live on hidden `_GeoData`. File: `FEGE_Choked_Flat_Sites_v18.xlsx`.
- v19 freezes that GeoPlot layout with no further map-logic change. File: `FEGE_Choked_Flat_Sites_v19.xlsx`.
- v20 keeps the same map and zooms, but changes marker colours: non-issue sites are mid-gray `#B0B0B0` (less white), issue sites are bright red `#FF2B2B` and slightly larger. File: `FEGE_Choked_Flat_Sites_v20.xlsx`.
- v21 adds site-name labels on the four zoom maps only, each showing `Site (Severity)`. The national map stays unlabeled. File: `FEGE_Choked_Flat_Sites_v21.xlsx`.
- v22 keeps zoom-only labels but drops severity from the text (`DHAPT35`) and uses a smaller 5-pt site-name font. File: `FEGE_Choked_Flat_Sites_v22.xlsx`.
- v23 uses the updated physical-site database (`Tech` column). GeoPlot is a Google-satellite map with a three-entry legend: cyan 4G, yellow 5G, red issue sites. Interactive HTML plus a JPEG embedded on `5. GeoPlot`. File: `FEGE_Choked_Flat_Sites_v23.xlsx`.
- v24 separates every map (national + four zooms) into individual JPEGs and full-width HTML sections. 5G markers are larger with a white halo so they stay visible on the national satellite map. File: `FEGE_Choked_Flat_Sites_v24.xlsx`.
- v25 adjusts marker/legend sizes (4G larger, 5G smaller) and reports the true Tech counts: 7,087 4G, 357 5G, 78 issue (all issue sites are 5G). File: `FEGE_Choked_Flat_Sites_v25.xlsx`.

## Jupyter (for local Windows use)

Copy both files into `D:\KPI Monitoring\Transmission\FEGEPortChockCheck` and Run All:

- `jupyter/TX_Port_Choke_Check.ipynb`
- `jupyter/tx_port_choke_engine.py`

Accepts CSV, XLS, XLSX, XLSB. Writes `TX_Port_Choke_Check.xlsx` in that same folder.
