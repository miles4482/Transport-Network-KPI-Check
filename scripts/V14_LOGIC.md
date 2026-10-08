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
- v26 keeps each zoom map's red issue count exact (Dhaka Metro 27, Dhaka North 17, Dhaka West 16, Gazipur 18). Neighbouring issue sites in the padded view stay yellow 5G, not extra red. File: `FEGE_Choked_Flat_Sites_v26.xlsx`.
- v27 shrinks 4G markers on the national map and enlarges 5G so the 357 5G sites stay visible against the dense 4G layer. File: `FEGE_Choked_Flat_Sites_v27.xlsx`.
- v28 uses the same compact marker size for 4G, 5G, and issue sites on the national map. Zoom-map marker sizes are unchanged. File: `FEGE_Choked_Flat_Sites_v28.xlsx`.
- v29 draws antialiased circular markers so 4G/5G/issue dots stay round, not square. File: `FEGE_Choked_Flat_Sites_v29.xlsx`.
- v30 is presentation only: Summary is renamed Dashboard; GeoPlot title is Transmission Link Health Monitoring; the Dashboard scored-on note and Site List formula/scoring lines move to 4. Method; freeze panes are removed. Detection stays at the v14/v17 78-site baseline. File: `FEGE_Choked_Flat_Sites_v30.xlsx`.
- v31 rebuilds the report from the split Dhaka upload (`FEG_KPI_DHAKA_5Oct.part1.rar` + `.part2.rar`, combined to `FEG_KPI_DHAKA_5Oct.csv`). Scoring is 2–4 Oct 2026 only. The 28 Sep–1 Oct rows in that file are ignored. Listing thresholds are unchanged. File: `FEGE_Choked_Flat_Sites_v31.xlsx`.
- v32 keeps that window and listing. Snapshot Y-axes are adaptive (DHDRS99 at 33.31 Mbit/s uses 0–50, not 0–100). File: `FEGE_Choked_Flat_Sites_v32.xlsx`.
- v33 adds a Dashboard-only 4G / 4G+5G choke summary from `Site list_Dhaka_Tech_5oct.xlsx`. Other sheets are unchanged. File: `FEGE_Choked_Flat_Sites_v33.xlsx`.
- v34 retitles the report Transmission Link Health Review (for transmission improvement planning, not blame) and drops DHAPT35/DHAPT48 case wording from Dashboard/Site List notes. File: `FEGE_Choked_Flat_Sites_v34.xlsx`.
- v35 rearranges Dashboard only: Technology-wise Link Health tiles, Cap Window Observation (Total / 4G+5G / 4G), and removes Tx BW from Dashboard site lists. File: `FEGE_Choked_Flat_Sites_v35.xlsx`.
- v36 keeps the national GeoPlot and replaces the four large area zooms with many medium issue-cluster maps (about 8–20 sites each). File: `FEGE_Choked_Flat_Sites_v36.xlsx`.
- v37 keeps GeoPlot clustering dynamic (maps until every issue site is covered) and orders clusters urban-first (Dhaka Metro, Gazipur, surrounding) then by site count. File: `FEGE_Choked_Flat_Sites_v37.xlsx`.
- v38 splits GeoPlot issue markers into two legends: red Issue 4G and orange Issue 4G+5G, from `Site list_Dhaka_Tech_5oct.xlsx`. Coverage 4G/5G colours stay. File: `FEGE_Choked_Flat_Sites_v38.xlsx`.
- v39 keeps the national map and urban-first issue clusters, and adds four thana maps after the national overview for the high-5G belt: Gulshan, Banani, Dhanmondi, Tejgaon. File: `FEGE_Choked_Flat_Sites_v39.xlsx`.
- v40 changes HourlyChartOfIssueSites only: keep the RxMaxSpeed series, replace the stuck line with a 72-hour peak line and a line 5% below that peak. File: `FEGE_Choked_Flat_Sites_v40.xlsx`.
- v41 keeps the 72-hour peak line and replaces the 5%-below line with the average of the three daily peaks. File: `FEGE_Choked_Flat_Sites_v41.xlsx`.
- v42 adds a national issue-severity map beside the national GeoPlot (Severe / High / Moderate / Low). Cluster maps are unchanged. File: `FEGE_Choked_Flat_Sites_v42.xlsx`.
- v43 fills the empty black gutter on the national pair: tighter crop, zoom-9 satellite, brighter/sharper tiles, and Excel scale that best-fits each half of the sheet. File: `FEGE_Choked_Flat_Sites_v43.xlsx`.
- v44 returns the national pair to the compact side-by-side size, reserves enough rows so thana/cluster maps do not overlap, and enlarges every map legend. File: `FEGE_Choked_Flat_Sites_v44.xlsx`.
- v45 pulls the national and severity maps together, leaving only a thin separator instead of a wide empty gutter. File: `FEGE_Choked_Flat_Sites_v45.xlsx`.
- v46 retunes GeoPlot severity colours: bright red Severe, less-bright deep red High, muted amber Moderate, calm green Low. File: `FEGE_Choked_Flat_Sites_v46.xlsx`.
- v47 keeps the 72-hour Peak line and redraws Avg peak as the mean of every 3-day hour that hits the wall, that day's high, or sits at/above the wall. File: `FEGE_Choked_Flat_Sites_v47.xlsx`.
- v48 plots GeoPlot top thanas from `Top Thana.xlsx` matched to the physical `Thana` column, shortens GeoPlot titles to `Thana (4G: n 4G+5G: m)`, and drops Night Rx / Sample snap from Site List. File: `FEGE_Choked_Flat_Sites_v48.xlsx`.
- v49 adds Site List **Double check** = Yes when Peak is ≥50% above Avg peak (DHGULAP type). Detection listing is unchanged. File: `FEGE_Choked_Flat_Sites_v49.xlsx`.
- v50 lowers Double check to Yes when Peak is ≥30% above Avg peak. Detection listing is unchanged. File: `FEGE_Choked_Flat_Sites_v50.xlsx`.
- v51 puts an issue-severity map beside every top-thana and cluster GeoPlot, matching the national pair. File: `FEGE_Choked_Flat_Sites_v51.xlsx`.
- v52 adds **Action plan** for urgent Tx BW increase: Severe/High (both cap shapes) with high Hours on cap and busy hours ≥50% on cap. Double check = Yes is left off. File: `FEGE_Choked_Flat_Sites_v52.xlsx`.
- v53 renames that sheet to **UrgentTxBWInc** and moves the How-to-read notes to **4. Method**. File: `FEGE_Choked_Flat_Sites_v53.xlsx`.
- v54 adds a 3-column Thana Level Summary above the UrgentTxBWInc site table. File: `FEGE_Choked_Flat_Sites_v54.xlsx`.
- v55 takes Tech (4G / 4G+5G) from `Physical_Site_Database_24Sep26.xlsx` (unmapped = 4G) and adds Site List **Site Type** from that file (unmapped = Not found). File: `FEGE_Choked_Flat_Sites_v55.xlsx`.
- v56 stretches the UrgentTxBWInc Thana Level Summary across 12 Name/Count pairs (columns B–Y) so ~75 thanas use 7 rows. File: `FEGE_Choked_Flat_Sites_v56.xlsx`.
- v57 fills that Thana summary down each column by count (highest first in column 1, then 2, …) and adds a **Tech** column on Dashboard, Site List, and UrgentTxBWInc. File: `FEGE_Choked_Flat_Sites_v57.xlsx`.
- v58 sizes the UrgentTxBWInc Thana summary from the thana count: more thanas add pairs to the right, fewer thanas shrink the grid. File: `FEGE_Choked_Flat_Sites_v58.xlsx`.
- v59 drops the “not every issue site can take a BW upgrade” line from Dashboard and UrgentTxBWInc. File: `FEGE_Choked_Flat_Sites_v59.xlsx`.
- v60 removes the Dashboard UrgentTxBWInc banner, the Tech-source note, and the UrgentTxBWInc Crowded/Hard ceiling footer. File: `FEGE_Choked_Flat_Sites_v60.xlsx`.
- v61 adds **URGENTBWINC_BusyHour** from `OSS KPI Hourly site  level.rar`: per day the 7 highest Data Volume hours and 7 highest Max User hours (any hour of the day), last-day tag, matched-hour %, and 7h vs 09:00–15:00 Data Volume / Max User / DL throughput. File: `FEGE_Choked_Flat_Sites_v61.xlsx`.
- v62 counts Hours matched as FEGE choked hours among the 21 DataBusyHours (7 × 3 days), % over 21, and Last day match as choked / 7 on 4 Oct only. File: `FEGE_Choked_Flat_Sites_v62.xlsx`.
- v63 splits Hours matched / % / Last day match / Listed because (and 7h DL throughput) into separate DataBusyHour and UserBusyHour blocks. File: `FEGE_Choked_Flat_Sites_v63.xlsx`.
- v64 makes **URGENTBWINC_BusyHour** independent of UrgentTxBWInc: it scans every OSS site with DataBusyHour and UserBusyHour (any day or last day), not the 560 UrgentTxBWInc list. Hours matched is still FEGE choke among those busy hours. File: `FEGE_Choked_Flat_Sites_v64.xlsx`.
- v65 drops **DL Throughput 09:00–15:00** from URGENTBWINC_BusyHour (7h DL throughput in each Data/User block remains). File: `FEGE_Choked_Flat_Sites_v65.xlsx`.
- v66 drops Data Volume 7h, Data Volume 09:00–15:00, Max User 7h, and Max User 09:00–15:00 from URGENTBWINC_BusyHour, and removes the 09:00–15:00 Method rule. File: `FEGE_Choked_Flat_Sites_v66.xlsx`.
- v67 keeps 7 BusyHour and adds matching **4 BusyHour** Data/User blocks (top 4 hours, Hours matched /12, last day choked / 4, DL Throughput 4h). File: `FEGE_Choked_Flat_Sites_v67.xlsx`.
- v68 adds matching **1 BusyHour** Data/User blocks to the right of 4 BusyHour (single highest hour, Hours matched /3, last day choked / 1, DL Throughput 1h). File: `FEGE_Choked_Flat_Sites_v68.xlsx`.
- v69 adds **Data Busy Hour** and **User Busy Hour** flag columns after Tech: Yes when any of that section’s 7h/4h/1h DL Throughput is ≤ 7 Mbps (OR). File: `FEGE_Choked_Flat_Sites_v69.xlsx`.
- v70 adds **Chocked %_Data Busy Hour** and **Chocked %_User Busy Hour** after Tech: Yes when any of that section’s Hours matched % (7h OR 4h OR 1h) is ≥ 30%. The ≤ 7 Mbps flags are unchanged. File: `FEGE_Choked_Flat_Sites_v70.xlsx`.
- v71 groups the four Yes flags under **Filtering(7H/4H/1H)** and colours each BusyHour section header differently. File: `FEGE_Choked_Flat_Sites_v71.xlsx`. BusyHour-only edits use `--fast` (skip snapshot charts, hourly KPI, and GeoPlot) plus `/tmp/fege_report_cache.pkl` so the 1,808 charts are not rebuilt.
- v72 adds **Filtering(4H only)** after Tech and before Filtering(7H/4H/1H): Chocked %_Data / Chocked %_User Yes when that 4h Hours matched % ≥ 30%; Data Busy Hour_TP / User Busy Hour_TP Yes when that 4h DL Throughput ≤ 7 Mbps. File: `FEGE_Choked_Flat_Sites_v72.xlsx`.
- v73 adds **Filtering(1H only)** after Tech and before Filtering(4H only), same Yes rules on the 1 BusyHour section (Hours matched % ≥ 30%; DL Throughput ≤ 7 Mbps). File: `FEGE_Choked_Flat_Sites_v73.xlsx`.
- v74 moves **Severity** to immediately after Tech on URGENTBWINC_BusyHour and adds **Site Type** from `Physical_Site_Database_24Sep26.xlsx` after Severity (unmapped = Not found). File: `FEGE_Choked_Flat_Sites_v74.xlsx`.
- v75 adds **4 BH Flag** and **1 BH Flag** after Site Type: Yes if any Filtering(4H only) / Filtering(1H only) column is Yes (OR). File: `FEGE_Choked_Flat_Sites_v75.xlsx`.
- v76 revises BH Flag: Yes when (Chocked %_Data OR Chocked %_User) **and** (Data Busy Hour_TP OR User Busy Hour_TP) in that 4H or 1H section. File: `FEGE_Choked_Flat_Sites_v76.xlsx`.
- v77 adds **Data Volume Nh (GB)** and **Max User Nh** to every 7H / 4H / 1H Data and User block (average over that block’s own busy hours). Not added to BH Flag. File: `FEGE_Choked_Flat_Sites_v77.xlsx`.
- v78 adds **4/1BH Flag** in the BH Flag block: Yes when 4 BH Flag OR 1 BH Flag is Yes. File: `FEGE_Choked_Flat_Sites_v78.xlsx`.
- v79 rebuilds the full report from **`FEG_KPI_DHAKA_5Oct.rar`** (single volume, replaces the part1/part2 split) and the updated **`OSS KPI Hourly site  level.rar`**. Scoring window is **5–7 Oct 2026**. File: `FEGE_Choked_Flat_Sites_v79.xlsx`.
- v80 adds **Comparison**: all sites, busy-hour RxMaxSpeed p90 on 2–4 Oct (part1/part2) vs 5–7 Oct (current RAR). **BW Increased** is Yes when the ceiling rises ≥ 40 Mbit/s or 12% (DHKKTE2 / GPSDR01 Tx step-up shape). File: `FEGE_Choked_Flat_Sites_v80.xlsx`.
- v81 Comparison uses **RxMaxSpeed only** (Tx Total BW columns removed). File: `FEGE_Choked_Flat_Sites_v81.xlsx`.
- v82 adds **DLPRBUtilization,%** to every 7H / 4H / 1H DataBusyHour and UserBusyHour block: `(DL PRB Utilization_N / DL PRB Utilization_D) × 100` over that block’s own busy hours. Not added to BH Flag. File: `FEGE_Choked_Flat_Sites_v82.xlsx`.
- v83 copies DataBusyHour / UserBusyHour **4H and 1H** DL Throughput, DLPRBUtilization,%, Data Volume, and Max User onto **1. Site List**, as a separate block immediately after Site Type. File: `FEGE_Choked_Flat_Sites_v83.xlsx`.
- v84 **stops using** `OSS KPI Hourly site  level.rar`. Cell-level parts (`OSS KPI Hourly Cel Level_05/06/07Oct.part1–3.rar`) are combined, **L900 (MID 8,3 = L09) is dropped**, remaining L1800/L2100/L2600 cells are rolled to site-hour (`OSS KPI Hourly site  level_capacity.csv`). BusyHour / Site List OSS KPIs use that file. Other sheets unchanged. File: `FEGE_Choked_Flat_Sites_v84.xlsx`.

## Jupyter (for local Windows use)

Copy both files into `D:\KPI Monitoring\Transmission\FEGEPortChockCheck` and Run All:

- `jupyter/TX_Port_Choke_Check.ipynb`
- `jupyter/tx_port_choke_engine.py`

Accepts CSV, XLS, XLSX, XLSB. Writes `TX_Port_Choke_Check.xlsx` in that same folder.
