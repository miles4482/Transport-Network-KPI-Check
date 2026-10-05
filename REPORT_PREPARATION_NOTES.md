# Transmission Link Health Review — comprehensive preparation notes

Latest published file: `FEGE_Choked_Flat_Sites_v39.xlsx`  
Window scored: **2–4 October 2026**  
Purpose: find FEGE ports whose **RxMaxSpeed is choked** (pinned to one level and cannot go upward) so transmission can plan improvement. This is a **health review**, not a blame list.

These notes record **where the report came from, how it was built, what was considered, and every rule that was kept**.

---

## 1. What the report is trying to catch

A site is an issue only when RxMaxSpeed looks **stuck on a wall**:

- After the night dip the line locks to one level, **or**
- The trace still moves during the day, but every rise **stops at the same top** and never goes above it.

A busy site whose peaks keep changing (tens to hundreds of Mbit/s apart) is **not** listed. User examples of that rejected shape: **DHGULN2, DHDHN40, DHBDD32, DHDHN09**.

The pattern sites used to design the rule (not a name filter): **DHAPT35** and **DHAPT48**. Both must always appear in the report.

The report is for **transmission link improvement planning**. Wording was kept neutral: title is **Transmission Link Health Review**. GeoPlot title is **Transmission Link Health Monitoring**.

---

## 2. Where the data came from

| Source | File | What it was used for |
|---|---|---|
| FEGE hourly KPI | `FEG_KPI_DHAKA_5Oct.part1.rar` + `.part2.rar` combined to `FEG_KPI_DHAKA_5Oct.csv` | Detection, listing, charts, Dashboard counts |
| Physical site database | `Physical_Site_Database_24Sep26.xlsx` | Latitude / longitude / district / region / radio Tech (4G or 5G) for GeoPlot |
| Dhaka tech list | `Site list_Dhaka_Tech_5oct.xlsx` | 4G vs **4G+5G** split on Dashboard and on GeoPlot issue legends |

How the KPI file was prepared:

1. The upload arrived as a **split RAR** (two parts).
2. The parts were combined and extracted to one CSV (~98 MB).
3. Dates such as `2-Oct` and hours such as `0:00` were parsed (not only Excel serials).
4. Year is taken as **2026**.
5. Duplicate site+timestamp rows keep the last value.

The combined CSV is a working file only. It is **not** committed to git.

---

## 3. What window was scored

- **Only 2, 3 and 4 October 2026** are scored.
- Days **28 September – 1 October** exist in the combined file and are **ignored**.
- Every hour is scored: busy **08:00–22:00** and off-peak / night.
- A site that is only flat at night can still be listed.
- Snapshot charts show the **same three days** (not a longer history).

This three-day window was requested for this Dhaka source. The **listing thresholds** (3 days / 3 hours, or last day 2 hours) were kept from the original V14 freeze; they were not loosened or tightened for the shorter window.

---

## 4. Counters

| Report name | Formula | Role |
|---|---|---|
| RxMaxSpeed (Mbit/s) | `VS.FEGE.RxMaxSpeed(bit/s) / 1000 / 1000` | **The only counter used to decide choke** |
| Tx Total BW (Mbit/s) | `VS.FEGE.TxTotalBW(kbit/s) / 1000` | RAN-side reference only |

**Tx Total BW is not used to list or grade a site.** It is a RAN counter, not transmission engineered bandwidth. It was removed from the Dashboard site lists. It remains a Site List / chart reference column only.

Old “At / Above / Below Tx BW” grading (85–120%) was **dropped** and must not be used to judge the list.

Night Rx (median RxMaxSpeed 02:00–06:00) is shown so the night dip can be compared. A site that stays on the cap through the night is still listed (the link is full all day).

---

## 5. Detection rules (frozen from V14)

Tried in this order. The **first shape that passes tightness + day rules** wins.

### Shape 1 — Crowded level (DHAPT35 / DHAPT48 shape)

- Search the **upper half** of the trace (quantile 0.55 to max).
- Hours on the level sit inside **±4 Mbit/s**, or **±1.5% of p95** if that is wider.
- The level is the **median** of those samples.
- 99th percentile may sit at most **40 Mbit/s** or **12% of the level** above it.
- This is the “line locks after the night dip” shape.

### Shape 2 — Hard ceiling (clipped top)

- Tried only if shape 1 does not fit.
- Same search, but only over the **top 10%** of samples.
- At most **1 hour** may sit more than **8 Mbit/s** or **3%** above the top.
- Catches traces that move during the day but are clipped at one value (examples: DHSVRS5, GPSDR01, DHKKT87, DHKKTE2).

### Tightness (both shapes)

- In-band standard deviation ≤ **2.5 Mbit/s** or **1% of the level**.
- Center (stuck level) ≥ **20 Mbit/s**.
- Fewer than 24 hourly samples, or a max below 20 Mbit/s: skip the site.

### What is rejected

Peaks that keep changing, with no fixed top. Those sites are busy, not choked.

---

## 6. Listing rules (who appears on the list)

A **day** is on-cap when at least **3 hours** that day (busy or off-peak, not necessarily consecutive) sit on the stuck level.

A site is listed if **either**:

1. **≥ 3 of the 3 scored days** are on-cap, **or**
2. The **last day (4 Oct 2026)** has ≥ **2 hours** on the cap.

A site can be listed for the last-day flag alone.

How the list is tagged:

- `≥3 days + last day`
- `≥3 days`
- `Last day`

The last-day column writes the date, hours, and busy / off-peak split, for example:  
`4-Oct-26 · 17h (busy 13, off-peak 4)`.

Longest consecutive run on the cap is reported as extra information. A day still counts when 3 hours sit on the level even if they are not adjacent.

**Reference check:** DHAPT35 and DHAPT48 must be in the list. Current result: both are listed (DHAPT35 stuck ~269 Mbit/s; DHAPT48 stuck ~243 Mbit/s).

---

## 7. Severity (not Tx BW)

Severity is **hours on the cap as a share of the 2–4 Oct window** (all hours):

| Severity | Hours on cap |
|---|---|
| Severe | ≥ 50% |
| High | 25% to < 50% |
| Moderate | 10% to < 25% |
| Low | < 10% |

Low still means choked — usually a hard ceiling or a last-day flag — just fewer hours.

---

## 8. Cap window observation

Where the flat ceiling sits across the day:

| Cap when | Meaning |
|---|---|
| Busy + off-peak | Hours on cap in both 08:00–22:00 and the rest of the day |
| Busy hours | Only busy hours |
| Off-peak | Only off-peak / night |
| Partial | Should not appear if listing rules fired |

Dashboard shows this as **Total / 4G+5G / 4G × severity × cap-when**.

---

## 9. Technology split (two different “tech” fields)

Two sources of technology were considered and **must not be mixed**:

### A. Physical radio Tech (GeoPlot coverage)

From `Physical_Site_Database_24Sep26.xlsx`:

- **4G** = 7,087 sites (cyan)
- **5G** = 357 sites (yellow)
- Total physical sites on the map = **7,444**

This is used for **coverage dots** on GeoPlot (the background network).

### B. Dhaka planning Tech (Dashboard + issue legends)

From `Site list_Dhaka_Tech_5oct.xlsx` (Site Name + Tech):

- **4G** = 4,515 checked
- **4G+5G** = 356 checked

This is used for:

- Dashboard “Technology-wise Link Health”
- Dashboard Cap Window Observation (Total / 4G+5G / 4G)
- GeoPlot **issue** legends (Issue 4G vs Issue 4G+5G)

If a listed site is missing from the tech list, GeoPlot falls back to physical Tech: physical 5G → Issue 4G+5G, otherwise Issue 4G.

---

## 10. Latest result (v39, this window)

| Item | Count |
|---|---|
| Sites checked (KPI, 2–4 Oct) | 4,871 |
| Issue sites listed | **1,808** |
| 4G choked / 4G checked | 1,729 / 4,515 |
| 4G+5G choked / 4G+5G checked | 79 / 356 |
| Severe / High / Moderate / Low | 439 / 434 / 743 / 192 |
| Listed ≥3 days + last day | 1,198 |
| Listed last day only | 610 |
| Listed ≥3 days only | 0 |
| Last-day cap flag (4 Oct) | 1,808 |
| Busy + off-peak / Busy / Off-peak | 1,616 / 188 / 4 |
| Crowded level / Hard ceiling | 1,613 / 195 |
| Issue sites with coordinates | **1,805** (1,726 Issue 4G + 79 Issue 4G+5G) |
| Issue sites without geo | 3 (4G issues; they stay on Site List / charts, not on maps) |

---

## 11. Workbook structure

| Sheet | What it contains | What was considered |
|---|---|---|
| **Dashboard** | Opening view. Technology-wise tiles. Cap Window Observation. Severity tiles. Short site lists **without Tx BW**. | Renamed from Summary. Neutral heading. No “scored on” footnote here (moved to Method). No DHAPT case-blame sentence. |
| **1. Site List** | Full issue list, sortable, each site name jumps to its chart. | Formula / scoring notes removed from this sheet (kept on Method). |
| **2. HourlyChartOfIssueSites** | One chart per issue site, 2–4 Oct, red dashed line at stuck RxMaxSpeed. | Adaptive Y-axis (example: DHDRS99 at 33.31 uses 0–50, not 0–100). Date shown once per day; hours 00:00–23:00. |
| **_ChartData** | Hidden source for charts | Not for reading |
| **3. Hourly KPI** | Hourly numbers behind the charts | Unchanged logic |
| **4. Method** | Full detection / listing / severity text | Holds notes that were removed from Dashboard and Site List |
| **5. GeoPlot** | Maps only (no visible table) | See section 12 |
| **_GeoData** | Hidden coordinates, status, tech, issue tech | Not for reading |

Excel freeze panes were removed so the file opens unfrozen.

---

## 12. GeoPlot — everything considered

### 12.1 National map (always first)

- Full Bangladesh physical-site layer from the database.
- Compact same-size markers so 4G / 5G / issues stay readable at country scale.
- No site-name labels (too dense).
- Google satellite background.

### 12.2 High-5G thana maps (after national, before clusters)

Requested because 5G footprint is high in these thanas. The physical database has **no Thana column**, so membership was built from **site-code prefix + coordinates**:

| Map | How a site is assigned | v39 thana 5G / issues |
|---|---|---|
| Gulshan Thana | Site code starts `DHGUL` **and** longitude ≥ 90.408 (east of Banani lake) | 47 5G · 29 issue |
| Banani Thana | `DHGUL` west of 90.408, plus `DHCNT` sites inside Banani box (23.784–23.808 N, 90.392–90.412 E) | 22 5G · 18 issue |
| Dhanmondi Thana | Site code starts `DHDHN` | 34 5G · 13 issue |
| Tejgaon Thana | Site code starts `DHTEJ` or `DHTIA` (Tejgaon + Tejgaon Industrial Area) | 28 5G · 37 issue |

Each site is assigned to **at most one** of these four thanas.

The map viewport is the member-site box plus a small pad (~0.008°) so the 5G footprint reads as an area. Neighbouring sites in the pad stay as coverage (cyan / yellow). Only that thana’s own issues stay red / orange.

These four maps are **fixed** (not dynamic clusters). Cluster logic after them is unchanged.

### 12.3 Dynamic issue clusters (after the thana maps)

Earlier four large zooms (Metro / North / West / Gazipur) were too messy (Metro alone had hundreds of issues). They were replaced by **neighbourhood clusters**.

Rules that were considered and kept:

1. **Cover every geo-located issue site.** Clusters are generated until 1,805 / 1,805 are on a map. Isolated rural sites get their own small maps rather than being dropped.
2. **Urban first.** Order is:
   - Dhaka Metro
   - Gazipur
   - Dhaka North / West / South
   - Narayanganj
   - Narsingdi / Manikganj / Munshiganj
   - Tangail
   - anything else
3. **Inside each urban band, larger packs first** (more issue sites on one map).
4. Cluster urban rank uses the **majority** of sites in the cluster (not the single most-urban site). That stops a Metro-majority map being labelled by one edge site.
5. Pack size: urban (rank ≤ 2) target **18**, max **22**. Other areas target **16**, max **20**. Tiny nearby leftovers may merge up to 22 if they are close enough.
6. Current run: **136** cluster maps. First maps are Dhaka Metro packs of 22. Gazipur follows. Tangail isolates sit last.
7. Labels: `Cluster NN — {majority district} / {majority region}`.

### 12.4 Two issue legends (GeoPlot only)

Coverage and issues are separate legends:

| Legend group | Colour | Meaning |
|---|---|---|
| Coverage | Cyan `#00E5FF` | Non-issue 4G (physical Tech) |
| Coverage | Yellow `#FFD600` | Non-issue 5G (physical Tech) |
| Issue sites | Red `#FF1744` | Issue **4G** (Dhaka tech list) |
| Issue sites | Orange `#FF6D00` | Issue **4G+5G** (Dhaka tech list) |

This split exists **only on GeoPlot**. Dashboard already had 4G vs 4G+5G tiles; maps needed the same distinction.

### 12.5 Exact issue counts on cluster / thana maps

A padded map always shows nearby sites. Those neighbours must **not** inflate the red / orange count.

Rule: neighbouring issue sites in the pad are drawn as coverage (4G cyan or 5G yellow), not as extra issue dots. The title / legend issue count is only that map’s own issues.

Site names are labelled on thana and cluster maps, not on the national map.

### 12.6 Interactive HTML

The same maps are also written as `FEGE_Choked_Flat_Sites_v39_GeoPlot.html` (Leaflet + Google satellite) for click / zoom. JPEG copies are embedded on the Excel GeoPlot sheet.

---

## 13. What was deliberately not done

- Detection rule was **not** changed after V14/V17. Presentation and maps changed; listing logic did not.
- Tx BW was **not** used as a pass/fail or severity input.
- Sites were **not** blamed. DHAPT35 / DHAPT48 case wording was removed from Dashboard / Site List.
- Four huge city zooms were **not** kept (unreadable).
- Cluster count is **not** a fixed number (128, 136, …). It is whatever is needed to cover all issue sites.
- Jupyter notebook was left as the local Windows copy; the live report is the Excel / HTML set on this branch.
- Combined KPI CSV and RAR parts were **not** committed (too large / working files).

---

## 14. How the report is produced (process)

1. Combine the two RAR parts → one KPI CSV.
2. Load KPI, restrict to 2–4 Oct 2026, parse dates and hours.
3. For every site, try crowded level, then hard ceiling.
4. Apply day / last-day listing rules.
5. Attach severity, cap-when, last-day flag, cap shape.
6. Load physical geo and Dhaka 4G / 4G+5G list.
7. Write Dashboard, Site List, charts, hourly KPI, Method.
8. Write GeoPlot: national → four thana maps → dynamic urban-first clusters, with two issue legends.
9. Publish Excel + HTML + JPEG maps.

Default output name follows the presentation version (`v39` at the time of these notes). Changing maps or Dashboard does **not** change the detection version; that stays V14.

---

## 15. Presentation decisions that shaped the current file

- Report title: **Transmission Link Health Review** (planning, not blame).
- GeoPlot title: **Transmission Link Health Monitoring**.
- Summary renamed **Dashboard**.
- Dashboard layout: technology tiles + Cap Window Observation; no Tx BW on Dashboard lists.
- Chart Y-axis adapts to the stuck level so a ~33 Mbit/s site is not crushed on a 0–100 axis.
- GeoPlot is map-only; the coordinate table is hidden.
- Urban / high-5G areas are shown first because that is where action and 5G footprint matter most.
- Every geo-located issue still gets a cluster map, including sparse Tangail sites.

---

## 16. How to read the published pack

- Open **Dashboard** for counts and 4G vs 4G+5G.
- Open **1. Site List** and click a site name for its hourly chart.
- Open **4. Method** for the same detection text that lives in this note.
- Open **5. GeoPlot**: national → Gulshan / Banani / Dhanmondi / Tejgaon → Cluster 01…136.
- Or open the HTML file for the same maps, interactive.

Direct files on the working branch:

- `FEGE_Choked_Flat_Sites_v39.xlsx`
- `FEGE_Choked_Flat_Sites_v39_GeoPlot.html`
- `FEGE_Choked_Flat_Sites_v39_GeoPlot_National.jpg`
- `FEGE_Choked_Flat_Sites_v39_GeoPlot_Thana_Gulshan.jpg` (and Banani, Dhanmondi, Tejgaon)
- `FEGE_Choked_Flat_Sites_v39_GeoPlot_Cluster_01.jpg` … `_136.jpg`
