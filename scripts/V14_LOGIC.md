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

Keep this rule as the baseline. If the rule or chart changes, write `FEGE_Choked_Flat_Sites_v15.xlsx` and leave v14 as the freeze point.
