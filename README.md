# Georgia OB/GYN Workforce — Rural Access Analysis

**Live map:** [jlmad20.github.io/Ga-OBGYN-Map](https://jlmad20.github.io/Ga-OBGYN-Map/)

An end-to-end data pipeline that streams 9.6 million provider records from the CMS/NPPES National Provider Identifier registry, filters and scores Georgia OB/GYN physicians, geocodes practice addresses, and visualizes rural workforce gaps on an interactive county-level map.

---

## Key Findings

| Metric | Value |
|---|---|
| OB/GYN physicians with confirmed GA practice address | **1,823** |
| HRSA FY2026 clinically active estimate | **~1,540** |
| Providers in rural counties | **121 (6.6%)** |
| Rural share of GA population | **21%** |
| Counties with zero OB/GYNs | **82 of 159 (52%)** |
| HRSA projected supply/demand ratio (FY2026) | **89%** |
| HRSA projected supply/demand ratio (FY2037) | **69%** |

Georgia's OB/GYN workforce is heavily concentrated in Atlanta, Augusta, and Savannah. Over half of the state's counties have no registered OB/GYN at all, and rural counties are served at roughly one-third the rate of their population share.

---

## Map Features

- **County choropleth** — provider count per county, color-scaled
- **Provider-to-population ratio** — OB/GYNs per 10,000 women aged 15–44 (2020 Census), benchmarked against HRSA's standard of 2.86 per 10,000
- **Rural status overlay** — HRSA rural classification for all 159 Georgia counties
- **Maternity care deserts** — counties with zero providers highlighted
- **Individual provider dots** — clustered, color-coded by rural status; hover for physician name and practice address; jittered so co-located providers don't stack

---

## Methodology

### Data Sources
| Source | Description |
|---|---|
| [CMS NPPES Full Replacement File](https://download.cms.gov/nppes/NPI_Files.html) | ~9.6M provider records, June 2026 vintage |
| [HRSA Health Workforce Simulation Model](https://data.hrsa.gov/topics/health-workforce/workforce-projections) | FY2022–2037 state-level supply/demand projections |
| [U.S. Census TIGER 2023 Shapefiles](https://www.census.gov/geographies/mapping-files/time-series/geo/tiger-line-file.html) | County boundary geometries |
| [Census County Age/Sex Estimates 2022](https://www.census.gov/data/tables/time-series/demo/popest/2020s-counties-detail.html) | Women aged 15–44 by county (2020 base) |
| [U.S. Census Geocoder API](https://geocoding.geo.census.gov/) | Address-to-coordinate resolution |

### Provider Identification
Providers were identified using the 207V OB/GYN taxonomy family (11 codes), filtering the full NPPES file to Georgia practice addresses. The pipeline streams the 9.6M-row file in 100,000-row chunks to avoid memory issues and caches results in Parquet format.

### Cohort Definitions (aligned with HRSA HWSM methodology)
- **Core General (n=1,823)** — MD/DO physicians with a primary OB/GYN taxonomy, excluding likely residents and non-physician providers
- **Umbrella (n=2,294)** — all providers with any 207V taxonomy code
- **APP Cohort (n=49)** — nurse midwives, nurse practitioners, and physician assistants with OB/GYN taxonomy
- **Resident Cohort (n=21)** — flagged by 2-of-3 signal: known residency program address, blank MD/DO credential, or mailstop/room number in address line

### Confidence Scoring
Each provider receives a 0–100 evidence score based on:
- State license match (Georgia)
- Taxonomy specificity (primary vs. secondary)
- Address geocode match quality
- Credential verification
- Single vs. multiple practice location

### Rural Classification
Counties are classified using HRSA's rural definition applied to all 159 Georgia counties: Rural / Partial rural / Not rural. Provider dot colors reflect the rural classification of their practice county.

### Known Limitations
- NPI registry reflects self-reported data and is not scrubbed for inactive providers — HRSA estimates approximately 283 fewer clinically active physicians than the raw NPI count
- Resident exclusion is heuristic (2-of-3 signals) rather than verified against residency program rosters
- Providers with unresolvable addresses (~210) appear in county-level counts but not as map dots
- The AMA Physician Masterfile would provide a more definitive count of clinically active, Georgia-licensed OB/GYNs but requires a licensing agreement

---

## Repository Structure

```
├── GA_OBGYN_Rural_NPI_Lookup_v6.py   # Main analysis pipeline
├── GA_OBGYN_Map_v1.py                # Interactive map builder
├── index.html                         # Self-contained interactive map (live site)
└── README.md
```

---

## How to Run

### Requirements
```bash
pip install pandas pyarrow duckdb openpyxl folium geopandas mapclassify requests numpy
```

### Step 1 — Run the analysis
```bash
# Download the NPPES Full Replacement File from:
# https://download.cms.gov/nppes/NPI_Files.html
# Place the ZIP in the project directory, then:

python GA_OBGYN_Rural_NPI_Lookup_v6.py
```
Outputs: Excel workbooks (CLEAN + AUDIT), CSV cohort files, geocode cache (~45 min on first run due to Census geocoder rate limiting; subsequent runs use cache and complete in ~2 min).

### Step 2 — Build the map
```bash
python GA_OBGYN_Map_v1.py
```
Output: `index.html` — fully self-contained interactive map with all JS libraries bundled inline (no CDN dependencies).

---

## Tools & Libraries

`pandas` · `pyarrow` · `duckdb` · `openpyxl` · `folium` · `geopandas` · `numpy` · `mapclassify` · `branca`

---

*Data current as of June 2026. HRSA projections from FY2022–2037 HWSM release.*
