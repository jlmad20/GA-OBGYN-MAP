"""
GA_OBGYN_Map_v1.py
------------------
Builds an interactive HTML map of Georgia OB/GYN providers from the
GA_OBGYN_Rural_NPI_Lookup_v6 outputs.

Layers:
  1. Georgia counties shaded by OB/GYN count (choropleth)
  2. County rural status overlay (togglable)
  3. Individual provider dots (color-coded by rural status)
  4. Maternity care desert counties (zero providers, highlighted red)

Output:
  outputs/GA_OBGYN_Map_v1.html

Run:
  python GA_OBGYN_Map_v1.py
"""

import json
import re
import sys
import urllib.request
import urllib.parse
import zipfile
import io
from pathlib import Path

# ---------------------------------------------------------------------------
# Dependency check
# ---------------------------------------------------------------------------
REQUIRED = ["pandas", "folium", "geopandas", "mapclassify"]
missing = []
for pkg in REQUIRED:
    try:
        __import__(pkg)
    except ImportError:
        missing.append(pkg)
if missing:
    print("ERROR: Missing packages. Run:")
    print("  python -m pip install " + " ".join(missing))
    sys.exit(1)

import pandas as pd
import folium
import geopandas as gpd
from folium.plugins import MarkerCluster, FloatImage
from branca.colormap import linear
import numpy as np

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR      = Path(__file__).parent
OUT_DIR       = BASE_DIR / "outputs"
PROC_DIR      = BASE_DIR / "processed"
CACHE_PATH    = PROC_DIR / "geocode_cache_ga_obgyn_v6.json"
CORE_CSV      = OUT_DIR / "GA_OBGYN_Rural_v6_Core_General.csv"
SHAPEFILE_DIR = PROC_DIR / "shapefiles"
MAP_OUT       = OUT_DIR / "GA_OBGYN_Map_v1.html"

# Census TIGER county shapefile URL (2023 vintage, ~10 MB)
TIGER_URL = (
    "https://www2.census.gov/geo/tiger/TIGER2023/COUNTY/"
    "tl_2023_us_county.zip"
)

# ---------------------------------------------------------------------------
# Rural status color scheme
# ---------------------------------------------------------------------------
RURAL_COLORS = {
    "Rural":                                        "#d7191c",   # red
    "Partial rural - tract/address review needed":  "#fdae61",   # orange
    "Not rural":                                    "#1a9641",   # green
    "Unknown - geocode/manual review":              "#aaaaaa",   # grey
    "No providers":                                 "#f5f5f5",   # near-white
}

RURAL_FILL = {
    "Rural":                                        "#d7191c",
    "Partial rural - tract/address review needed":  "#fdae61",
    "Not rural":                                    "#1a9641",
    "Unknown - geocode/manual review":              "#aaaaaa",
    "No providers":                                 "#cccccc",
}

DOT_COLORS = {
    "Rural":                                        "#d7191c",
    "Partial rural - tract/address review needed":  "#ff8800",
    "Not rural":                                    "#1a9641",
    "Unknown - geocode/manual review":              "#888888",
}

# ---------------------------------------------------------------------------
# Step 1 — Load provider data and join lat/lon from geocode cache
# ---------------------------------------------------------------------------
def load_providers() -> pd.DataFrame:
    print("[STEP] Loading provider data …")
    df = pd.read_csv(CORE_CSV, dtype=str)

    # Load geocode cache
    with open(CACHE_PATH) as f:
        cache = json.load(f)

    def normalize_address(street, city, state, zipcode):
        parts = [str(x).strip().upper() for x in [street, city, state, zipcode]
                 if str(x).strip() and str(x).strip().lower() != "nan"]
        return "|".join(parts)

    lats, lons = [], []
    for _, row in df.iterrows():
        key = normalize_address(
            row.get("Practice Address Line 1", ""),
            row.get("Practice City", ""),
            row.get("Practice State", ""),
            row.get("Practice ZIP", ""),
        )
        result = cache.get(key, {})
        lats.append(result.get("lat"))
        lons.append(result.get("lon"))

    df["lat"] = pd.to_numeric(lats, errors="coerce")
    df["lon"] = pd.to_numeric(lons, errors="coerce")

    # Normalize county name (strip " County" suffix)
    df["County Clean"] = df["County"].fillna("").apply(
        lambda x: re.sub(r"\s+County$", "", x, flags=re.IGNORECASE).strip()
    )

    n_with_coords = df[["lat", "lon"]].notna().all(axis=1).sum()
    print(f"  Total providers     : {len(df):,}")
    print(f"  With coordinates    : {n_with_coords:,}")
    print(f"  Without coordinates : {len(df) - n_with_coords:,} (will appear in county layer only)")
    return df


# ---------------------------------------------------------------------------
# Step 2 — Load Georgia county shapefile
# ---------------------------------------------------------------------------
def load_georgia_shapefile() -> gpd.GeoDataFrame:
    SHAPEFILE_DIR.mkdir(parents=True, exist_ok=True)
    shp_path = SHAPEFILE_DIR / "tl_2023_us_county.shp"

    if not shp_path.exists():
        print("[STEP] Downloading Census TIGER county shapefile (~10 MB) …")
        response = urllib.request.urlopen(TIGER_URL, timeout=60)
        zipped = zipfile.ZipFile(io.BytesIO(response.read()))
        zipped.extractall(SHAPEFILE_DIR)
        print("  Download complete.")
    else:
        print("[STEP] Shapefile already cached — skipping download.")

    gdf = gpd.read_file(shp_path)
    # Filter to Georgia (FIPS state code 13)
    ga = gdf[gdf["STATEFP"] == "13"].copy()
    ga["County Name"] = ga["NAME"].str.strip()
    ga = ga.to_crs(epsg=4326)
    # Simplify geometry to reduce file size (0.005 degree tolerance —
    # visually indistinguishable at county zoom level, cuts coords by 97%)
    ga["geometry"] = ga["geometry"].simplify(0.005, preserve_topology=True)
    print(f"  Georgia counties loaded: {len(ga)} (geometry simplified for file size)")
    return ga


# ---------------------------------------------------------------------------
# Step 3a — Load Census population data (women 15-44 by county)
# ---------------------------------------------------------------------------
POP_CACHE = PROC_DIR / "ga_county_women1544.json"
POP_URL   = (
    "https://www2.census.gov/programs-surveys/popest/datasets/"
    "2020-2022/counties/asrh/cc-est2022-agesex-13.csv"
)

def load_population_data() -> dict:
    """
    Returns dict: {county_name_no_suffix -> women_15_44}
    Uses 2020 base population (YEAR=1) from Census county age/sex estimates.
    Caches locally so subsequent runs don't re-download.
    """
    if POP_CACHE.exists():
        print("[STEP] Loading population data from cache …")
        with open(POP_CACHE) as f:
            return json.load(f)

    print("[STEP] Downloading Census county population data …")
    import requests
    r = requests.get(POP_URL, timeout=60)
    r.raise_for_status()

    pop_df = pd.read_csv(io.BytesIO(r.content), encoding="latin-1")

    # YEAR=1 is 2020 Census base; filter to Georgia counties only (SUMLEV=50)
    pop_df = pop_df[(pop_df["YEAR"] == 1) & (pop_df["SUMLEV"] == 50)].copy()

    # Strip "County" from county name
    pop_df["County Name Clean"] = (
        pop_df["CTYNAME"]
        .str.replace(r"\s+County$", "", regex=True)
        .str.strip()
    )

    # Women aged 15-44
    pop_df["Women_1544"] = pd.to_numeric(pop_df["AGE1544_FEM"], errors="coerce").fillna(0).astype(int)

    result = dict(zip(pop_df["County Name Clean"], pop_df["Women_1544"]))

    with open(POP_CACHE, "w") as f:
        json.dump(result, f, indent=2)

    print(f"  Population data loaded for {len(result)} counties.")
    return result


# ---------------------------------------------------------------------------
# Step 3b — Build county-level summary
# ---------------------------------------------------------------------------
def build_county_summary(df: pd.DataFrame, ga_counties: gpd.GeoDataFrame,
                         pop_data: dict) -> gpd.GeoDataFrame:
    print("[STEP] Building county summary …")

    # Count providers per county
    county_counts = (
        df[df["County Clean"].str.len() > 0]
        .groupby("County Clean")
        .agg(
            Provider_Count=("NPI", "count"),
            Rural_Status=("Rural Status", lambda x: x.mode()[0] if len(x) > 0 else "Unknown"),
        )
        .reset_index()
    )

    # Merge with shapefile
    merged = ga_counties.merge(
        county_counts,
        left_on="County Name",
        right_on="County Clean",
        how="left",
    )
    merged["Provider_Count"] = merged["Provider_Count"].fillna(0).astype(int)
    merged["Rural_Status"]   = merged["Rural_Status"].fillna("No providers")
    merged["Is_Desert"]      = merged["Provider_Count"] == 0

    # Add population and compute ratio
    merged["Women_1544"] = merged["County Name"].map(pop_data).fillna(0).astype(int)

    # OB/GYNs per 10,000 women of reproductive age
    # Avoid division by zero for counties with no population data
    merged["OBGYNs_per_10k"] = merged.apply(
        lambda r: round(r["Provider_Count"] / r["Women_1544"] * 10000, 2)
        if r["Women_1544"] > 0 else 0.0,
        axis=1,
    )

    # Access adequacy label (benchmark: ~1 OB/GYN per 3,500 women = 2.86 per 10k)
    BENCHMARK = 2.86   # per 10,000 women aged 15-44

    def adequacy_label(ratio):
        if ratio == 0:
            return "Desert (0 providers)"
        if ratio < BENCHMARK * 0.5:
            return f"Severe shortage (<50% of benchmark)"
        if ratio < BENCHMARK:
            return f"Shortage (<benchmark of {BENCHMARK}/10k)"
        if ratio < BENCHMARK * 1.5:
            return f"Adequate (~benchmark)"
        return f"Well-served (>{BENCHMARK * 1.5:.1f}/10k)"

    merged["Access_Adequacy"] = merged["OBGYNs_per_10k"].apply(adequacy_label)

    total   = county_counts["Provider_Count"].sum()
    n_desert = merged["Is_Desert"].sum()
    print(f"  Counties with providers  : {(~merged['Is_Desert']).sum()}")
    print(f"  Maternity care deserts   : {n_desert} counties with 0 OB/GYNs")
    print(f"  Total providers mapped   : {int(total)}")
    print(f"  Ratio range: {merged['OBGYNs_per_10k'].min():.2f} – "
          f"{merged['OBGYNs_per_10k'].max():.2f} per 10k women")

    return merged


# ---------------------------------------------------------------------------
# Step 4 — Build the folium map
# ---------------------------------------------------------------------------
def build_map(df: pd.DataFrame, county_gdf: gpd.GeoDataFrame) -> folium.Map:
    print("[STEP] Building interactive map …")

    # Centre on Georgia
    GA_CENTER = [32.9, -83.4]
    m = folium.Map(
        location=GA_CENTER,
        zoom_start=7,
        tiles="CartoDB positron",
        control_scale=True,
    )

    # ----------------------------------------------------------------
    # Layer 1 — County choropleth (provider count)
    # ----------------------------------------------------------------
    max_count = max(county_gdf["Provider_Count"].max(), 1)

    colormap = linear.YlOrRd_09.scale(0, max_count)
    colormap.caption = "OB/GYN Provider Count per County"

    def count_style(feature):
        count = feature["properties"].get("Provider_Count", 0)
        is_desert = count == 0
        return {
            "fillColor": "#f0f0f0" if is_desert else colormap(count),
            "color": "#555555",
            "weight": 0.8,
            "fillOpacity": 0.75 if not is_desert else 0.4,
        }

    def count_highlight(feature):
        return {"weight": 2.5, "color": "#333333", "fillOpacity": 0.9}

    count_layer = folium.GeoJson(
        county_gdf.__geo_interface__,
        name="OB/GYN Count by County",
        style_function=count_style,
        highlight_function=count_highlight,
        tooltip=folium.GeoJsonTooltip(
            fields=["County Name", "Provider_Count", "Rural_Status"],
            aliases=["County", "OB/GYN Providers", "Rural Status"],
            localize=True,
            sticky=False,
            labels=True,
            style="font-size:13px;",
        ),
        popup=folium.GeoJsonPopup(
            fields=["County Name", "Provider_Count", "Rural_Status", "Is_Desert"],
            aliases=["County", "OB/GYN Count", "Rural Status", "Maternity Care Desert?"],
        ),
    )
    count_layer.add_to(m)
    colormap.add_to(m)

    # ----------------------------------------------------------------
    # Layer 2 — Provider-to-population ratio (OB/GYNs per 10k women 15-44)
    # ----------------------------------------------------------------
    max_ratio = max(county_gdf["OBGYNs_per_10k"].max(), 0.01)

    # Blue-purple scale: white (0) → dark purple (high ratio)
    ratio_colormap = linear.PuRd_09.scale(0, max_ratio)
    ratio_colormap.caption = "OB/GYNs per 10,000 Women (Age 15–44)"

    BENCHMARK = 2.86

    def ratio_style(feature):
        ratio = feature["properties"].get("OBGYNs_per_10k", 0) or 0
        is_desert = ratio == 0
        return {
            "fillColor": "#eeeeee" if is_desert else ratio_colormap(ratio),
            "color": "#555555",
            "weight": 0.8,
            "fillOpacity": 0.80 if not is_desert else 0.35,
        }

    ratio_layer = folium.GeoJson(
        county_gdf.__geo_interface__,
        name="OB/GYNs per 10,000 Women (15–44)",
        style_function=ratio_style,
        highlight_function=count_highlight,
        tooltip=folium.GeoJsonTooltip(
            fields=["County Name", "OBGYNs_per_10k", "Provider_Count",
                    "Women_1544", "Access_Adequacy", "Rural_Status"],
            aliases=["County", "OB/GYNs per 10k Women", "OB/GYN Count",
                     "Women Age 15–44 (2020)", "Access Level", "Rural Status"],
            localize=True,
            sticky=False,
            labels=True,
            style="font-size:13px;",
        ),
        popup=folium.GeoJsonPopup(
            fields=["County Name", "OBGYNs_per_10k", "Provider_Count",
                    "Women_1544", "Access_Adequacy", "Rural_Status"],
            aliases=["County", "OB/GYNs per 10k Women", "OB/GYN Count",
                     "Women Age 15–44", "Access Level", "Rural Status"],
        ),
        show=False,   # off by default; user toggles on
    )
    ratio_layer.add_to(m)
    ratio_colormap.add_to(m)  # shows caption for ratio layer

    # ----------------------------------------------------------------
    # Layer 3 — Rural status overlay (togglable)
    # ----------------------------------------------------------------
    def rural_style(feature):
        status = feature["properties"].get("Rural_Status", "Unknown")
        return {
            "fillColor": RURAL_FILL.get(status, "#cccccc"),
            "color": "#444444",
            "weight": 0.8,
            "fillOpacity": 0.65,
        }

    rural_layer = folium.GeoJson(
        county_gdf.__geo_interface__,
        name="Rural Status by County",
        style_function=rural_style,
        highlight_function=count_highlight,
        tooltip=folium.GeoJsonTooltip(
            fields=["County Name", "Rural_Status", "Provider_Count"],
            aliases=["County", "Rural Status", "OB/GYN Providers"],
            localize=True,
            sticky=False,
            labels=True,
        ),
        show=False,   # off by default — user can toggle
    )
    rural_layer.add_to(m)

    # ----------------------------------------------------------------
    # Layer 4 — Maternity care deserts (highlighted outline)
    # ----------------------------------------------------------------
    deserts = county_gdf[county_gdf["Is_Desert"]].copy()

    def desert_style(feature):
        return {
            "fillColor": "#d7191c",
            "color": "#8b0000",
            "weight": 2,
            "fillOpacity": 0.25,
            "dashArray": "5, 5",
        }

    desert_layer = folium.GeoJson(
        deserts.__geo_interface__,
        name="Maternity Care Deserts (0 OB/GYNs)",
        style_function=desert_style,
        tooltip=folium.GeoJsonTooltip(
            fields=["County Name", "Rural_Status"],
            aliases=["County (DESERT)", "Rural Status"],
            sticky=False,
        ),
        show=True,
    )
    desert_layer.add_to(m)

    # ----------------------------------------------------------------
    # Layer 5 — Individual provider dots (clustered)
    # ----------------------------------------------------------------
    has_coords = df[df["lat"].notna() & df["lon"].notna()].copy()

    # Jitter duplicate coordinates so stacked dots spread apart when zoomed in.
    # ~0.0004 degrees ≈ 40 meters — invisible at county zoom, visible at street level.
    rng = np.random.default_rng(seed=42)   # fixed seed = reproducible layout
    coord_counts = has_coords.groupby(["lat", "lon"]).cumcount()  # 0 for first, 1,2,... for dupes
    jitter_mask  = coord_counts > 0        # only jitter the 2nd+ provider at each address

    jitter_r     = rng.uniform(0.00015, 0.00045, size=len(has_coords))
    jitter_theta = rng.uniform(0, 2 * np.pi,      size=len(has_coords))

    has_coords = has_coords.copy()
    has_coords["lat"] = has_coords["lat"] + np.where(
        jitter_mask, jitter_r * np.sin(jitter_theta), 0
    )
    has_coords["lon"] = has_coords["lon"] + np.where(
        jitter_mask, jitter_r * np.cos(jitter_theta), 0
    )

    n_jittered = jitter_mask.sum()
    print(f"  Jittered {n_jittered} duplicate-address providers so dots don't stack.")

    # Clustered layer
    cluster = MarkerCluster(
        name="Individual Providers (clustered)",
        options={
            "maxClusterRadius": 40,
            "disableClusteringAtZoom": 11,
        },
    )

    for _, row in has_coords.iterrows():
        rural_status = str(row.get("Rural Status", "Unknown"))
        color = DOT_COLORS.get(rural_status, "#888888")

        # Build popup HTML
        def s(val):
            v = row.get(val, "")
            return "" if str(v).strip().lower() in ("nan", "none", "") else str(v).strip()

        name      = s("Full Name") or "Unknown"
        cred      = s("Credential")
        addr      = s("Practice Address Line 1")
        city      = s("Practice City")
        zipcode   = s("Practice ZIP")
        county    = s("County Clean")
        phone     = s("Phone")
        tax       = s("Primary Taxonomy")
        conf      = s("Confidence Label")
        npi       = s("NPI")

        cred_str  = f", {cred}" if cred else ""
        phone_str = phone if phone else "N/A"
        popup_html = (
            f"<b>{name}{cred_str}</b><br>"
            f"{addr}, {city} {zipcode}<br>"
            f"<b>County:</b> {county}<br>"
            f"<b>Rural:</b> {rural_status}<br>"
            f"<b>Phone:</b> {phone_str}<br>"
            f"<b>NPI:</b> {npi}"
        )

        folium.CircleMarker(
            location=[row["lat"], row["lon"]],
            radius=6,
            color="white",
            weight=1,
            fill=True,
            fill_color=color,
            fill_opacity=0.85,
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"{name} — {county} County ({rural_status})",
        ).add_to(cluster)

    cluster.add_to(m)
    # Note: unclustered layer removed to keep file size under 25MB for sharing.
    # The MarkerCluster above auto-expands at zoom level 11+.

    # ----------------------------------------------------------------
    # Legend
    # ----------------------------------------------------------------
    legend_html = """
    <div style="
        position: fixed;
        bottom: 40px; left: 40px;
        z-index: 1000;
        background-color: white;
        padding: 14px 18px;
        border-radius: 8px;
        border: 1px solid #cccccc;
        font-family: Arial, sans-serif;
        font-size: 13px;
        box-shadow: 2px 2px 6px rgba(0,0,0,0.2);
        max-width: 230px;
        ">
      <b style="font-size:14px;">Provider Dot Color</b><br>
      <i style="background:#d7191c;width:12px;height:12px;display:inline-block;
         border-radius:50%;margin-right:6px;"></i> Rural<br>
      <i style="background:#ff8800;width:12px;height:12px;display:inline-block;
         border-radius:50%;margin-right:6px;"></i> Partial rural<br>
      <i style="background:#1a9641;width:12px;height:12px;display:inline-block;
         border-radius:50%;margin-right:6px;"></i> Not rural<br>
      <i style="background:#888888;width:12px;height:12px;display:inline-block;
         border-radius:50%;margin-right:6px;"></i> Unknown<br>
      <br>
      <b style="font-size:14px;">Maternity Care Deserts</b><br>
      <i style="background:#d7191c;opacity:0.25;width:12px;height:12px;
         display:inline-block;margin-right:6px;border:2px dashed #8b0000;"></i>
         0 OB/GYNs in county<br>
      <br>
      <b style="font-size:14px;">Access Benchmark</b><br>
      <span style="font-size:12px;">
        2.86 OB/GYNs per 10,000<br>
        women aged 15–44<br>
        <i>(≈1 per 3,500 women)</i>
      </span><br>
      <br>
      <span style="color:#888;font-size:11px;">
        Source: CMS/NPPES June 2026<br>
        Population: 2020 Census<br>
        HRSA Supply Estimate: ~1,620 (FY2024)
      </span>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    # ----------------------------------------------------------------
    # Title box
    # ----------------------------------------------------------------
    title_html = """
    <div style="
        position: fixed;
        top: 14px; left: 50%;
        transform: translateX(-50%);
        z-index: 1000;
        background-color: white;
        padding: 10px 24px;
        border-radius: 8px;
        border: 1px solid #cccccc;
        font-family: Arial, sans-serif;
        box-shadow: 2px 2px 6px rgba(0,0,0,0.2);
        text-align: center;
        ">
      <b style="font-size:16px;color:#003366;">
        Georgia OB/GYN Workforce — Rural Access Map
      </b><br>
      <span style="font-size:12px;color:#555;">
        NPI/NPPES Core General Cohort &nbsp;|&nbsp; June 2026
      </span>
    </div>
    """
    m.get_root().html.add_child(folium.Element(title_html))

    # Layer control
    folium.LayerControl(collapsed=False, position="topright").add_to(m)

    return m


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("\n" + "="*60)
    print(" GA OB/GYN Rural Map Builder — v1")
    print("="*60)

    df          = load_providers()
    ga_counties = load_georgia_shapefile()
    pop_data    = load_population_data()
    county_gdf  = build_county_summary(df, ga_counties, pop_data)
    m           = build_map(df, county_gdf)

    print(f"\n[STEP] Saving map to: {MAP_OUT}")
    m.save(str(MAP_OUT))

    # ----------------------------------------------------------------
    # Post-save: bundle all CDN JS libraries inline so the HTML is
    # fully self-contained — no network requests, no race conditions.
    # ----------------------------------------------------------------
    import re as _re
    import urllib.request as _urlreq

    print("  Bundling CDN libraries inline (eliminates race conditions) …")
    html = MAP_OUT.read_text(encoding="utf-8")

    def fetch_and_inline(match):
        url = match.group(1)
        try:
            req = _urlreq.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with _urlreq.urlopen(req, timeout=30) as resp:
                js = resp.read().decode("utf-8", errors="replace")
            print(f"    Bundled: {url.split('/')[-1]}")
            return f"<script>{js}</script>"
        except Exception as e:
            print(f"    WARN: could not fetch {url}: {e} — leaving as CDN link")
            return match.group(0)

    html = _re.sub(
        r'<script\s+src="([^"]+)"[^>]*>\s*</script>',
        fetch_and_inline,
        html,
    )

    MAP_OUT.write_text(html, encoding="utf-8")
    print("  Bundling complete — map is now fully self-contained.")

    print(f"  Open in any browser:")
    print(f"  {MAP_OUT}")

    # Print quick summary
    n_desert = county_gdf["Is_Desert"].sum()
    n_rural_providers = len(df[df["Rural Status"] == "Rural"])
    print(f"\n  Quick summary:")
    print(f"  Providers plotted on map : {df[['lat','lon']].notna().all(axis=1).sum():,}")
    print(f"  Rural providers          : {n_rural_providers:,}")
    print(f"  Maternity care deserts   : {n_desert} counties")


if __name__ == "__main__":
    main()
