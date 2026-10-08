"""
generate_dashboard.py
Part of the WeMap data pipeline.

Generates a single standalone HTML dashboard (index.html) containing:
  - 100% self-contained single-file HTML & embedded dataset
  - Okabe-Ito colorblind-friendly palette
  - Interactive cross-filtering (Market Segments, Top Countries, Company Scale)
  - Dynamic majority-colored cluster bubbles
  - Non-intrusive Map Legend
  - Slide-Over Company Profile Drawer (About, Products, Metadata, Social Links)
  - Click table row / popup to zoom & locate pin on map
  - Active filter chips bar
"""

import json
import argparse
from pathlib import Path
import openpyxl


def load_data(xlsx_path: Path, sheet_name: str = "WeMap_Live_Data", logos_dir: Path = None) -> list:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(c).strip() if c is not None else "" for c in next(rows_iter)]

    def col_idx(name: str):
        try:
            return headers.index(name)
        except ValueError:
            return None

    idx_org_id = col_idx("org_id")
    idx_name = col_idx("public_name")
    idx_cc = col_idx("city_country")
    idx_country = col_idx("country")
    idx_city = col_idx("city")
    idx_sug_city = col_idx("suggested_city")
    idx_lat = col_idx("lat")
    idx_lon = col_idx("lon")
    idx_segments = col_idx("market_segments")
    idx_biz = col_idx("business_model")
    idx_size = col_idx("employees_number_category")
    idx_emp_num = col_idx("employees_number")
    idx_website = col_idx("website")
    idx_desc = col_idx("org_desc")
    idx_member = col_idx("member_association")
    idx_lang = col_idx("product_language")
    idx_target_aud = col_idx("target_audience")
    idx_p1_name = col_idx("product_1_name")
    idx_p1_desc = col_idx("product_1_desc")
    idx_p2_name = col_idx("product_2_name")
    idx_p2_desc = col_idx("product_2_desc")
    idx_verified = col_idx("verified")
    idx_linkedin = col_idx("socials_linkedin")
    idx_x = col_idx("socials_x")
    idx_insta = col_idx("socials_insta")
    idx_fb = col_idx("facebook")

    valid_biz_models = {
        "Business to Business", "Business to Schools",
        "Business to Consumer", "Business to Government"
    }

    records = []
    for r in rows_iter:
        name = r[idx_name] if idx_name is not None else None
        if not name:
            continue

        lat_val = r[idx_lat] if idx_lat is not None else None
        lon_val = r[idx_lon] if idx_lon is not None else None

        lat, lon = None, None
        try:
            if lat_val is not None and str(lat_val).strip():
                lat = float(lat_val)
            if lon_val is not None and str(lon_val).strip():
                lon = float(lon_val)
        except (ValueError, TypeError):
            lat, lon = None, None

        cc_val = str(r[idx_cc]).strip() if idx_cc is not None and r[idx_cc] is not None else ""
        country_val = str(r[idx_country]).strip() if idx_country is not None and r[idx_country] is not None else ""
        city_val = str(r[idx_city]).strip() if idx_city is not None and r[idx_city] is not None else ""
        sug_city_val = str(r[idx_sug_city]).strip() if idx_sug_city is not None and r[idx_sug_city] is not None else ""
        web_val = str(r[idx_website]).strip() if idx_website is not None and r[idx_website] is not None else ""

        seg_list = []
        if idx_segments is not None and r[idx_segments]:
            for s in str(r[idx_segments]).split(","):
                s_clean = s.strip()
                if s_clean and not s_clean.startswith("http"):
                    seg_list.append(s_clean)

        biz_list = []
        if idx_biz is not None and r[idx_biz]:
            for b in str(r[idx_biz]).split(","):
                b_clean = b.strip()
                if b_clean in valid_biz_models:
                    biz_list.append(b_clean)

        size_val = ""
        if idx_size is not None and r[idx_size]:
            size_clean = str(r[idx_size]).strip()
            if not size_clean.startswith("http"):
                size_val = size_clean

        raw_org_id = str(r[idx_org_id]).strip() if idx_org_id is not None and r[idx_org_id] else ""
        org_id = raw_org_id if raw_org_id else f"ORG_{len(records)+1}"

        # Resolve logo: check logos_dir for {org_id}.png / .jpg / .gif / .webp
        logo_path = ""
        if logos_dir and logos_dir.is_dir():
            for ext in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
                candidate = logos_dir / f"{org_id}{ext}"
                if candidate.exists():
                    logo_path = f"logos/{org_id}{ext}"
                    break

        records.append({
            "id": org_id,
            "name": str(name).strip(),
            "logo": logo_path,
            "cc": cc_val,
            "country": country_val,
            "city": city_val,
            "suggested_city": sug_city_val,
            "lat": lat,
            "lon": lon,
            "segments": seg_list,
            "biz_models": biz_list,
            "size": size_val,
            "employees_num": str(r[idx_emp_num]).strip() if idx_emp_num is not None and r[idx_emp_num] else "",
            "website": web_val,
            "desc": str(r[idx_desc]).strip() if idx_desc is not None and r[idx_desc] else "",
            "member": str(r[idx_member]).strip() if idx_member is not None and r[idx_member] else "",
            "product_lang": str(r[idx_lang]).strip() if idx_lang is not None and r[idx_lang] else "",
            "target_audience": str(r[idx_target_aud]).strip() if idx_target_aud is not None and r[idx_target_aud] else "",
            "p1_name": str(r[idx_p1_name]).strip() if idx_p1_name is not None and r[idx_p1_name] else "",
            "p1_desc": str(r[idx_p1_desc]).strip() if idx_p1_desc is not None and r[idx_p1_desc] else "",
            "p2_name": str(r[idx_p2_name]).strip() if idx_p2_name is not None and r[idx_p2_name] else "",
            "p2_desc": str(r[idx_p2_desc]).strip() if idx_p2_desc is not None and r[idx_p2_desc] else "",
            "verified": str(r[idx_verified]).strip() if idx_verified is not None and r[idx_verified] else "",
            "social_linkedin": str(r[idx_linkedin]).strip() if idx_linkedin is not None and r[idx_linkedin] else "",
            "social_x": str(r[idx_x]).strip() if idx_x is not None and r[idx_x] else "",
            "social_insta": str(r[idx_insta]).strip() if idx_insta is not None and r[idx_insta] else "",
            "social_facebook": str(r[idx_fb]).strip() if idx_fb is not None and r[idx_fb] else ""
        })

    return records


def build_html(records: list, title: str = "WeMap European EdTech Explorer") -> str:
    records_json = json.dumps(records, ensure_ascii=False)

    geojson_path = Path(__file__).parent / "world_countries.geojson"
    if geojson_path.exists():
        countries_geojson_str = geojson_path.read_text(encoding="utf-8")
    else:
        countries_geojson_str = '{"type":"FeatureCollection","features":[]}'

    compiled_css_path = Path(__file__).parent / "compiled_tailwind.css"
    if compiled_css_path.exists():
        tailwind_css = compiled_css_path.read_text(encoding="utf-8")
    else:
        tailwind_css = "/* Tailwind fallback */"

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
  <meta http-equiv="Pragma" content="no-cache" />
  <meta http-equiv="Expires" content="0" />
  <title>{title}</title>
  
  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  
  <!-- Compiled Tailwind CSS (No CDN warnings, production ready) -->
  <style>
{tailwind_css}
  </style>

  <!-- Leaflet & MarkerCluster CSS/JS -->
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
  <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
  <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

  <!-- Chart.js CDN -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

  <style>
    body {{
      background-color: #f8fafc;
      color: #0f172a;
    }}
    #map {{
      height: 520px;
      width: 100%;
      border-radius: 0.75rem;
      z-index: 1;
    }}
    .leaflet-interactive, path.leaflet-interactive {{
      outline: none !important;
      -webkit-tap-highlight-color: transparent !important;
    }}
    path.leaflet-interactive:focus {{
      outline: none !important;
    }}
    .custom-cluster {{
      border: 2.5px solid #ffffff;
      color: #ffffff;
      font-weight: 700;
      font-size: 13px;
      border-radius: 9999px;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2), 0 2px 4px -2px rgba(0, 0, 0, 0.1);
      transition: transform 0.15s ease-in-out;
    }}
    .custom-cluster:hover {{
      transform: scale(1.1);
    }}
    .leaflet-popup-content-wrapper {{
      border-radius: 0.5rem;
      padding: 2px;
      box-shadow: 0 10px 15px -3px rgba(0,0,0,0.15);
    }}
    .leaflet-popup-content {{
      margin: 10px 14px;
      line-height: 1.4;
    }}
    .map-legend-box {{
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(4px);
      border: 1px solid #e2e8f0;
      border-radius: 0.5rem;
      padding: 8px 12px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }}
    #drawerContainer {{
      transition: visibility 0.3s ease-in-out;
    }}
    #drawerContainer.drawer-hidden {{
      visibility: hidden;
      pointer-events: none;
    }}
    #drawerContainer.drawer-visible {{
      visibility: visible;
      pointer-events: auto;
    }}
    .drawer-backdrop {{
      transition: opacity 0.3s ease-in-out;
    }}
    .drawer-panel {{
      transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }}
    .drawer-hidden .drawer-backdrop {{
      opacity: 0;
      pointer-events: none;
    }}
    .drawer-hidden .drawer-panel {{
      transform: translateX(100%);
    }}
    .drawer-visible .drawer-backdrop {{
      opacity: 1;
      pointer-events: auto;
    }}
    .drawer-visible .drawer-panel {{
      transform: translateX(0);
    }}
    .custom-scrollbar::-webkit-scrollbar {{
      width: 6px;
    }}
    .custom-scrollbar::-webkit-scrollbar-track {{
      background: #f1f5f9;
    }}
    .custom-scrollbar::-webkit-scrollbar-thumb {{
      background: #cbd5e1;
      border-radius: 3px;
    }}
  </style>
</head>
<body class="min-h-screen flex flex-col font-sans antialiased selection:bg-brand-500 selection:text-white relative overflow-x-hidden">

  <!-- Header -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-20 shadow-sm">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
      <div>
        <div class="flex items-center gap-3">
          <h1 class="text-xl sm:text-2xl font-bold text-slate-900 tracking-tight">{title}</h1>
          <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-brand-50 text-brand-700 border border-brand-100">
            Ecosystem Directory
          </span>
        </div>
        <p class="text-xs sm:text-sm text-slate-500 mt-0.5">Interactive geospatial map and cross-filtered analytics</p>
      </div>

      <!-- Market Segments Filter Dropdown -->
      <div class="flex items-center gap-2">
        <label for="segmentFilter" class="text-xs sm:text-sm font-semibold text-slate-700 whitespace-nowrap">Market Segment:</label>
        <select id="segmentFilter" class="bg-slate-50 border border-slate-300 text-slate-900 text-xs sm:text-sm rounded-lg focus:ring-brand-500 focus:border-brand-500 block p-2 pr-8 font-medium">
          <option value="ALL">All Segments</option>
        </select>
      </div>
    </div>
  </header>

  <!-- Main Content -->
  <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

    <!-- Active Filters Bar -->
    <div id="filterChipsBar" class="hidden bg-slate-100/80 border border-slate-200 p-3 rounded-xl flex items-center justify-between flex-wrap gap-2 text-xs">
      <div class="flex items-center gap-2 flex-wrap" id="chipsContainer">
        <!-- Chips inserted via JS -->
      </div>
      <button id="clearAllBtn" class="text-slate-500 hover:text-slate-800 font-semibold underline underline-offset-2 transition-colors">Clear All Filters</button>
    </div>

    <!-- Geospatial Explorer Card -->
    <div class="bg-white p-4 sm:p-5 rounded-xl border border-slate-200 shadow-sm relative">
      <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3">
        <div>
          <h2 class="text-base sm:text-lg font-bold text-slate-900">Geospatial Explorer</h2>
          <p class="text-xs text-slate-500">Clusters dynamically take the color of the <b>majority company scale</b> in that area. Click table rows to view full profile.</p>
        </div>
        
        <!-- View Controls & Count -->
        <div class="flex items-center gap-4 sm:gap-6 flex-wrap">
          <!-- Mode Toggle: Pin vs Choropleth (FIRST) -->
          <div class="inline-flex rounded-lg border border-slate-200 bg-slate-100/80 p-1 text-xs font-medium text-slate-600 shadow-2xs">
            <button id="btnPinMode" class="px-3.5 py-1.5 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all">📍 Pin View</button>
            <button id="btnChoroplethMode" class="px-3.5 py-1.5 rounded-md text-slate-600 hover:text-slate-900 transition-all">🗺️ Choropleth View</button>
          </div>

          <!-- Viewport Toggle: Focus vs Global (SECOND) -->
          <div class="inline-flex rounded-lg border border-slate-200 bg-slate-100/80 p-1 text-xs font-medium text-slate-600 shadow-2xs">
            <button id="btnFocusView" class="px-3.5 py-1.5 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all">🎯 Focus View</button>
            <button id="btnGlobalView" class="px-3.5 py-1.5 rounded-md text-slate-600 hover:text-slate-900 transition-all">🌐 Global View</button>
          </div>

          <span class="text-xs text-slate-500 whitespace-nowrap font-medium pl-1"><span id="mapCount" class="font-bold text-slate-900">0</span> mapped</span>
        </div>
      </div>

      <!-- Sleek Country Intelligence Card (shown on country selection) -->
      <div id="countryStatsCard" class="hidden mb-3 bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden transition-all">
        <!-- Top bar: Header & Action buttons -->
        <div class="flex items-center justify-between px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex-wrap gap-2 sm:gap-3">
          <!-- Left: Flag + Country Name + Org Count Badge -->
          <div class="flex items-center gap-2.5 sm:gap-3 min-w-0">
            <span id="cscFlag" class="shrink-0 flex items-center justify-center"></span>
            <h3 id="cscName" class="font-bold text-slate-900 text-base sm:text-lg truncate"></h3>
            <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold bg-brand-50 text-brand-700 border border-brand-200 shadow-2xs">
              <span id="cscCount" class="font-bold mr-1">0</span> Orgs
            </span>
          </div>

          <!-- Right: Filter Button & Close Button -->
          <div class="flex items-center gap-2 shrink-0 ml-auto sm:ml-0">
            <button id="cscFilterBtn" onclick="applyCountryStatsFilter()" class="px-3 py-1.5 text-xs font-semibold bg-brand-600 hover:bg-brand-700 text-white rounded-lg transition-colors shadow-2xs">
              Filter to this country →
            </button>
            <button onclick="closeCountryStatsCard()" class="text-slate-400 hover:text-slate-700 text-lg leading-none font-bold p-1.5 rounded-lg transition-colors" title="Close stats card">&times;</button>
          </div>
        </div>

        <!-- Responsive 4-Card Intelligence Grid -->
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 p-3.5 bg-slate-50/80 border-t border-slate-200">
          <div id="cscTopCities" class="bg-white rounded-xl border border-slate-200 p-3.5 shadow-xs hover:border-slate-300 transition-all flex flex-col justify-between"></div>
          <div id="cscBizMix" class="bg-white rounded-xl border border-slate-200 p-3.5 shadow-xs hover:border-slate-300 transition-all flex flex-col justify-between"></div>
          <div id="cscSegmentSynergy" class="bg-white rounded-xl border border-slate-200 p-3.5 shadow-xs hover:border-slate-300 transition-all flex flex-col justify-between"></div>
          <div id="cscProductExpansion" class="bg-white rounded-xl border border-slate-200 p-3.5 shadow-xs hover:border-slate-300 transition-all flex flex-col justify-between"></div>
        </div>
      </div>

      <!-- Map Container -->
      <div class="relative">
        <div id="map" class="shadow-inner border border-slate-100"></div>

        <!-- Map Legend (placed below the map) -->
        <div id="mapLegendBar" class="mt-3 bg-slate-50/90 border border-slate-200 rounded-lg px-4 py-2.5 text-xs text-slate-700 min-h-[44px] flex items-center justify-between flex-wrap gap-4">
          <!-- Pin Legend -->
          <div id="pinLegend" class="flex items-center gap-4 flex-wrap">
            <span class="font-bold text-slate-900 uppercase text-[10px] tracking-wider shrink-0">Company Scale Pins:</span>
            <div class="flex items-center gap-3.5 flex-wrap font-medium text-slate-700">
              <div class="flex items-center gap-1.5"><span class="w-3 h-3 min-w-[12px] min-h-[12px] rounded-full inline-block shrink-0 border border-black/10" style="background-color: #0072B2;"></span> Micro (1-9)</div>
              <div class="flex items-center gap-1.5"><span class="w-3 h-3 min-w-[12px] min-h-[12px] rounded-full inline-block shrink-0 border border-black/10" style="background-color: #E69F00;"></span> Small (10-49)</div>
              <div class="flex items-center gap-1.5"><span class="w-3 h-3 min-w-[12px] min-h-[12px] rounded-full inline-block shrink-0 border border-black/10" style="background-color: #009E73;"></span> Medium (50-249)</div>
              <div class="flex items-center gap-1.5"><span class="w-3 h-3 min-w-[12px] min-h-[12px] rounded-full inline-block shrink-0 border border-black/10" style="background-color: #CC79A7;"></span> Large (250+)</div>
              <div class="flex items-center gap-1.5"><span class="w-3 h-3 min-w-[12px] min-h-[12px] rounded-full inline-block shrink-0 border border-black/10" style="background-color: #D55E00;"></span> Solo Founder</div>
              <div class="flex items-center gap-1.5"><span class="w-3 h-3 min-w-[12px] min-h-[12px] rounded-full inline-block shrink-0 border border-black/10" style="background-color: #56B4E9;"></span> Med/Large</div>
            </div>
          </div>

          <!-- Choropleth Legend -->
          <div id="choroplethLegend" class="hidden flex items-center gap-4 flex-wrap">
            <span class="font-bold text-slate-900 uppercase text-[10px] tracking-wider shrink-0">Organization Density:</span>
            <div class="flex items-center gap-3.5 flex-wrap font-medium text-slate-700">
              <div class="flex items-center gap-1.5"><span class="w-4 h-3.5 rounded-xs border border-slate-300 inline-block shrink-0" style="background-color: #ffffcc;"></span> 1–5</div>
              <div class="flex items-center gap-1.5"><span class="w-4 h-3.5 rounded-xs border border-slate-300 inline-block shrink-0" style="background-color: #c7e9b4;"></span> 6–15</div>
              <div class="flex items-center gap-1.5"><span class="w-4 h-3.5 rounded-xs border border-slate-300 inline-block shrink-0" style="background-color: #7fcdbb;"></span> 16–35</div>
              <div class="flex items-center gap-1.5"><span class="w-4 h-3.5 rounded-xs border border-slate-300 inline-block shrink-0" style="background-color: #1d91c0;"></span> 36–75</div>
              <div class="flex items-center gap-1.5"><span class="w-4 h-3.5 rounded-xs border border-slate-300 inline-block shrink-0" style="background-color: #0c2c84;"></span> 75+</div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Charts Section -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
      
      <!-- Top Countries -->
      <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col">
        <div class="flex items-center justify-between mb-4">
          <div>
            <h2 class="text-base font-bold text-slate-900">Top Countries</h2>
            <p class="text-xs text-slate-500">Click a bar to filter & zoom to that country</p>
          </div>
          <span id="countryFilterStatus" class="text-xs font-semibold text-brand-700 bg-brand-50 px-2 py-0.5 rounded-md border border-brand-100 hidden">Filtered</span>
        </div>
        <div class="flex-1 min-h-[260px]">
          <canvas id="countryChart"></canvas>
        </div>
      </div>

      <!-- Company Scale / Employee Size Distribution -->
      <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col">
        <div class="flex items-center justify-between mb-4">
          <div>
            <h2 class="text-base font-bold text-slate-900">Company Scale (Employee Size)</h2>
            <p class="text-xs text-slate-500">Click a slice to filter by team size</p>
          </div>
          <span id="sizeFilterStatus" class="text-xs font-semibold text-brand-700 bg-brand-50 px-2 py-0.5 rounded-md border border-brand-100 hidden">Filtered</span>
        </div>
        <div class="flex-1 min-h-[260px] flex items-center justify-center">
          <canvas id="sizeChart"></canvas>
        </div>
      </div>

    </div>

    <!-- Directory Table Section -->
    <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div class="p-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h2 class="text-base font-bold text-slate-900">Organization Directory</h2>
          <p class="text-xs text-slate-500">Click any row to open full company profile, products & map location</p>
        </div>
        <div class="relative w-full sm:w-72">
          <input type="text" id="searchInput" placeholder="Search by name, city, country..." class="w-full bg-slate-50 border border-slate-300 text-slate-900 text-xs sm:text-sm rounded-lg focus:ring-brand-500 focus:border-brand-500 pl-3 pr-8 py-2">
          <svg class="w-4 h-4 text-slate-400 absolute right-2.5 top-2.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path>
          </svg>
        </div>
      </div>

      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs sm:text-sm text-slate-700">
          <thead class="bg-slate-50 text-slate-700 uppercase font-semibold text-[11px] tracking-wider border-b border-slate-200">
            <tr>
              <th scope="col" class="px-5 py-3">Organization</th>
              <th scope="col" class="px-5 py-3">Location</th>
              <th scope="col" class="px-5 py-3">Country</th>
              <th scope="col" class="px-5 py-3">Market Segments</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-100">
            <!-- Populated via JS -->
          </tbody>
        </table>
      </div>

      <!-- Pagination -->
      <div class="p-4 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
        <span id="pageInfo">Showing 0 to 0 of 0 entries</span>
        <div class="flex gap-1.5">
          <button id="prevBtn" class="px-3 py-1 bg-white border border-slate-300 rounded hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed font-medium">Previous</button>
          <button id="nextBtn" class="px-3 py-1 bg-white border border-slate-300 rounded hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed font-medium">Next</button>
        </div>
      </div>
      <div class="px-5 py-2 border-t border-slate-100 bg-slate-50/60 text-[11px] text-slate-500">
        <span class="text-amber-600 font-bold">*</span> Indicates location derived from raw dataset entry (suggested city unverified).
      </div>
    </div>

  </main>

  <!-- Slide-Over Company Profile Drawer Container -->
  <div id="drawerContainer" class="drawer-hidden fixed inset-0 z-50 overflow-hidden" role="dialog" aria-modal="true">
    <div onclick="closeDrawer()" class="drawer-backdrop fixed inset-0 bg-slate-900/40 backdrop-blur-xs"></div>
    <div class="fixed inset-y-0 right-0 max-w-full flex pl-10">
      <div class="drawer-panel w-screen max-w-md bg-white shadow-2xl border-l border-slate-200 flex flex-col">
        <div class="px-5 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50/80">
          <span class="text-xs font-bold uppercase tracking-wider text-slate-500">Organization Profile</span>
          <button onclick="closeDrawer()" class="p-1 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 transition-colors font-bold text-lg leading-none" title="Close Profile (Esc)">&times;</button>
        </div>
        <div id="drawerBody" class="p-6 overflow-y-auto custom-scrollbar flex-1 space-y-6"></div>
      </div>
    </div>
  </div>

  <!-- Footer -->
  <footer class="bg-white border-t border-slate-200 py-4 mt-8">
    <div class="max-w-7xl mx-auto px-4 text-center text-xs text-slate-400">
      WeMap European EdTech Intelligence Platform &bull; Open-source GIS Architecture
    </div>
  </footer>

  <!-- Application Script -->
  <script>
    const RAW_DATA = {records_json};

    // Colorblind-Friendly Okabe-Ito Palette
    const OKABE_ITO = {{
      blue: '#0072B2',
      orange: '#E69F00',
      green: '#009E73',
      pink: '#CC79A7',
      vermillion: '#D55E00',
      skyBlue: '#56B4E9',
      gray: '#999999',
      lightGray: '#cbd5e1'
    }};

    const SIZE_COLOR_MAP = {{
      'Micro': OKABE_ITO.blue,
      'Small': OKABE_ITO.orange,
      'Medium': OKABE_ITO.green,
      'Large': OKABE_ITO.pink,
      'Solo entrepreneur': OKABE_ITO.vermillion,
      'Medium or Large (unspecified)': OKABE_ITO.skyBlue
    }};

    function getCityOnly(d, includeAsterisk = false) {{
      const sug = d.suggested_city ? String(d.suggested_city).trim() : '';
      const rawCity = d.city ? String(d.city).trim() : '';
      const cCountry = d.country ? String(d.country).trim().toLowerCase() : '';

      if (sug && sug.toLowerCase() !== 'none' && sug.toLowerCase() !== 'null') {{
        if (cCountry && sug.toLowerCase() === cCountry) return '-';
        return sug;
      }}

      if (rawCity && rawCity.toLowerCase() !== 'none' && rawCity.toLowerCase() !== 'null') {{
        if (cCountry && rawCity.toLowerCase() === cCountry) return '-';
        if (includeAsterisk) {{
          return `${{rawCity}} <span class="text-amber-600 font-bold text-xs inline-block ml-0.5" title="Location derived from raw dataset entry; unverified suggested city">*</span>`;
        }}
        return rawCity;
      }}

      if (d.cc && d.cc.includes(',')) {{
        const parts = d.cc.split(',');
        const cityPart = parts[0].trim();
        if (cCountry && cityPart.toLowerCase() === cCountry) return '-';
        if (includeAsterisk) {{
          return `${{cityPart}} <span class="text-amber-600 font-bold text-xs inline-block ml-0.5" title="Location derived from raw dataset entry; unverified suggested city">*</span>`;
        }}
        return cityPart;
      }}

      return '-';
    }}

    const COUNTRY_GEOJSON = {countries_geojson_str};

    const COUNTRY_ALIASES = {{
      'the netherlands': 'netherlands',
      'turkiye': 'turkey',
      'scotland': 'united kingdom',
      'hong kong sar': 'china',
      'hong kong': 'china',
      'united states of america': 'united states',
      'usa': 'united states',
      'uk': 'united kingdom',
      'czech republic': 'czechia',
      'cz': 'czechia',
      'republic of czechia': 'czechia',
      'republic of serbia': 'serbia'
    }};

    function getCanonicalCountry(name) {{
      if (!name) return '';
      const clean = String(name).trim().toLowerCase();
      return COUNTRY_ALIASES[clean] || clean;
    }}

    const COUNTRY_ISO_MAP = {{
      'australia': 'au', 'au': 'au',
      'austria': 'at', 'at': 'at',
      'belarus': 'by', 'by': 'by',
      'belgium': 'be', 'be': 'be',
      'bulgaria': 'bg', 'bg': 'bg',
      'canada': 'ca', 'ca': 'ca',
      'croatia': 'hr', 'hr': 'hr',
      'cyprus': 'cy', 'cy': 'cy',
      'czech republic': 'cz', 'czechia': 'cz', 'cz': 'cz',
      'denmark': 'dk', 'dk': 'dk',
      'estonia': 'ee', 'ee': 'ee',
      'finland': 'fi', 'fi': 'fi',
      'france': 'fr', 'fr': 'fr',
      'germany': 'de', 'de': 'de',
      'greece': 'gr', 'gr': 'gr',
      'hong kong sar': 'hk', 'hong kong': 'hk', 'hk': 'hk',
      'hungary': 'hu', 'hu': 'hu',
      'iceland': 'is', 'is': 'is',
      'india': 'in', 'in': 'in',
      'ireland': 'ie', 'ie': 'ie',
      'israel': 'il', 'il': 'il',
      'italy': 'it', 'it': 'it',
      'kenya': 'ke', 'ke': 'ke',
      'latvia': 'lv', 'lv': 'lv',
      'lithuania': 'lt', 'lt': 'lt',
      'luxembourg': 'lu', 'lu': 'lu',
      'malta': 'mt', 'mt': 'mt',
      'moldova': 'md', 'md': 'md',
      'monaco': 'mc', 'mc': 'mc',
      'montenegro': 'me', 'me': 'me',
      'netherlands': 'nl', 'the netherlands': 'nl', 'nl': 'nl',
      'norway': 'no', 'no': 'no',
      'poland': 'pl', 'pl': 'pl',
      'portugal': 'pt', 'pt': 'pt',
      'romania': 'ro', 'ro': 'ro',
      'russia': 'ru', 'ru': 'ru',
      'scotland': 'gb-sct',
      'serbia': 'rs', 'rs': 'rs',
      'singapore': 'sg', 'sg': 'sg',
      'slovakia': 'sk', 'sk': 'sk',
      'slovenia': 'si', 'si': 'si',
      'south africa': 'za', 'za': 'za',
      'spain': 'es', 'es': 'es',
      'sweden': 'se', 'se': 'se',
      'switzerland': 'ch', 'ch': 'ch',
      'turkey': 'tr', 'turkiye': 'tr', 'tr': 'tr',
      'ukraine': 'ua', 'ua': 'ua',
      'united kingdom': 'gb', 'uk': 'gb', 'gb': 'gb',
      'united states': 'us', 'united states of america': 'us', 'usa': 'us', 'us': 'us'
    }};

    function countryFlagImg(nameOrIso) {{
      if (!nameOrIso) return '<span class="inline-flex items-center justify-center px-1.5 py-0.5 text-[10px] font-bold rounded bg-slate-200 text-slate-800 border border-slate-300">GLOBAL</span>';
      const clean = String(nameOrIso).trim().toLowerCase();
      const canon = getCanonicalCountry(clean);
      const iso = COUNTRY_ISO_MAP[clean] || COUNTRY_ISO_MAP[canon] || (clean.length === 2 ? clean : '');
      const displayTag = (iso || nameOrIso.slice(0, 2)).toUpperCase();

      if (!iso) {{
        return `<span class="inline-flex items-center justify-center px-1.5 py-0.5 text-[10px] font-bold rounded bg-slate-200 text-slate-800 border border-slate-300 shrink-0">${{displayTag}}</span>`;
      }}

      const flagIso = (iso === 'gb-sct' ? 'gb' : iso).toLowerCase();
      return `<img src="https://flagcdn.com/w40/${{flagIso}}.png" 
                   srcset="https://flagcdn.com/w80/${{flagIso}}.png 2x" 
                   width="28" height="20" 
                   alt="${{nameOrIso}} flag" 
                   class="rounded-xs shadow-2xs border border-slate-200 inline-block align-middle shrink-0 object-cover" 
                   onerror="this.onerror=null; this.replaceWith(Object.assign(document.createElement('span'), {{className: 'inline-flex items-center justify-center px-1.5 py-0.5 text-[10px] font-bold rounded bg-slate-200 text-slate-800 border border-slate-300 shrink-0', textContent: '${{displayTag}}'}}))">`;
    }}

    // ── Country Stats Card ─────────────────────────────────────────────────────
    let _cscPinnedCountry = null; // {{canon, displayName, iso2}}

    function openCountryStatsCard(canon, displayName, iso2, countsByCanon) {{
      _cscPinnedCountry = {{ canon, displayName }};

      const card = document.getElementById('countryStatsCard');
      if (!card) return;

      document.getElementById('cscFlag').innerHTML = countryFlagImg(displayName || iso2);
      document.getElementById('cscName').textContent = displayName;

      // Org count for this country from RAW_DATA
      const orgsHere = RAW_DATA.filter(d => getCanonicalCountry(d.country) === canon);
      const totalOrgs = orgsHere.length || 1;
      document.getElementById('cscCount').textContent = orgsHere.length.toLocaleString();

      // 1. Top Hubs
      const citiesEl = document.getElementById('cscTopCities');
      if (citiesEl) {{
        const cityCounts = {{}};
        orgsHere.forEach(d => {{
          const c = getCityOnly(d, false);
          if (c && c !== '-' && c !== 'Location unspecified') {{
            cityCounts[c] = (cityCounts[c] || 0) + 1;
          }}
        }});
        const topCities = Object.entries(cityCounts).sort((a, b) => b[1] - a[1]).slice(0, 3);
        if (topCities.length > 0) {{
          const cityTags = topCities.map(([cityName, cnt]) => `<span class="inline-flex items-center px-2 py-0.5 rounded-md bg-slate-100 font-semibold text-slate-800 text-xs">${{cityName}} <b class="text-brand-700 ml-1">${{cnt}}</b></span>`).join(' ');
          citiesEl.innerHTML = `
            <div>
              <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                  <span>🏙️</span> Top Innovation Hubs
                </span>
                <span class="text-xs font-bold text-brand-700 bg-brand-50 border border-brand-200 px-2 py-0.5 rounded-full">${{topCities.length}} Hubs</span>
              </div>
              <div class="flex flex-wrap gap-1.5 my-1">
                ${{cityTags}}
              </div>
              <p class="text-xs text-slate-500 mt-2 leading-tight">Key ecosystem hubs in ${{displayName}}</p>
            </div>`;
        }} else {{
          citiesEl.innerHTML = `
            <div>
              <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                  <span>🏙️</span> Top Innovation Hubs
                </span>
              </div>
              <div class="text-xs text-slate-500 italic my-1">Nationwide hubs</div>
              <p class="text-xs text-slate-500 mt-2 leading-tight">City distribution across ${{displayName}}</p>
            </div>`;
        }}
      }}

      // 2. Business Model Mix
      const bizEl = document.getElementById('cscBizMix');
      if (bizEl) {{
        const bmCounts = {{ 'B2B': 0, 'B2S': 0, 'B2C': 0, 'B2G': 0 }};
        orgsHere.forEach(d => {{
          if (d.biz_models) {{
            d.biz_models.forEach(b => {{
              if (b.includes('Business to Business')) bmCounts['B2B']++;
              if (b.includes('Business to Schools')) bmCounts['B2S']++;
              if (b.includes('Business to Consumer')) bmCounts['B2C']++;
              if (b.includes('Business to Government')) bmCounts['B2G']++;
            }});
          }}
        }});

        const sortedModels = Object.entries(bmCounts).sort((a, b) => b[1] - a[1]);
        const topModel = sortedModels[0];
        const topPct = topModel && totalOrgs > 0 ? Math.round((topModel[1] / totalOrgs) * 100) : 0;
        const b2bPct = Math.round((bmCounts['B2B'] / totalOrgs) * 100);
        const b2cPct = Math.round((bmCounts['B2C'] / totalOrgs) * 100);
        const b2sPct = Math.round((bmCounts['B2S'] / totalOrgs) * 100);
        const b2gPct = Math.round((bmCounts['B2G'] / totalOrgs) * 100);

        if (topModel && topModel[1] > 0) {{
          bizEl.innerHTML = `
            <div>
              <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                  <span>💼</span> Business Model Mix
                </span>
                <span class="text-xs font-bold text-brand-700 bg-brand-50 border border-brand-200 px-2 py-0.5 rounded-full">Top: ${{topModel[0]}} (${{topPct}}%)</span>
              </div>
              <div class="space-y-1.5 my-1">
                <div class="w-full bg-slate-100 h-2 rounded-full overflow-hidden flex shadow-inner">
                  <div class="h-full" style="width: ${{b2bPct}}%; background-color: #0072B2;" title="B2B ${{b2bPct}}%"></div>
                  <div class="h-full" style="width: ${{b2cPct}}%; background-color: #E69F00;" title="B2C ${{b2cPct}}%"></div>
                  <div class="h-full" style="width: ${{b2sPct}}%; background-color: #009E73;" title="B2S ${{b2sPct}}%"></div>
                  <div class="h-full" style="width: ${{b2gPct}}%; background-color: #CC79A7;" title="B2G ${{b2gPct}}%"></div>
                </div>
                <div class="flex items-center justify-between text-xs font-semibold text-slate-600 gap-1 flex-wrap">
                  <span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full inline-block" style="background-color: #0072B2;"></span> B2B: ${{b2bPct}}%</span>
                  <span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full inline-block" style="background-color: #E69F00;"></span> B2C: ${{b2cPct}}%</span>
                  <span class="inline-flex items-center gap-1"><span class="w-2 h-2 rounded-full inline-block" style="background-color: #009E73;"></span> B2S: ${{b2sPct}}%</span>
                </div>
              </div>
              <p class="text-xs text-slate-500 mt-2 leading-tight">Commercial model share across orgs</p>
            </div>`;
        }} else {{
          bizEl.innerHTML = `
            <div>
              <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                  <span>💼</span> Business Model Mix
                </span>
              </div>
              <div class="text-xs text-slate-400 italic my-1">Data pending</div>
              <p class="text-xs text-slate-500 mt-2 leading-tight">Commercial model breakdown</p>
            </div>`;
        }}
      }}

      // 3. Multi-Segment Reach
      const synergyEl = document.getElementById('cscSegmentSynergy');
      if (synergyEl) {{
        const multiSegOrgs = orgsHere.filter(d => d.segments && Array.isArray(d.segments) && d.segments.length > 1).length;
        const multiSegPct = Math.round((multiSegOrgs / totalOrgs) * 100);
        synergyEl.innerHTML = `
          <div>
            <div class="flex items-center justify-between mb-2">
              <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                <span>🔀</span> Multi-Segment Reach
              </span>
              <span class="text-xs font-extrabold text-brand-700 bg-brand-50 border border-brand-200 px-2 py-0.5 rounded-full">${{multiSegPct}}%</span>
            </div>
            <div class="space-y-1.5 my-1">
              <div class="w-full bg-slate-100 h-2 rounded-full overflow-hidden shadow-inner">
                <div class="h-full rounded-full" style="width: ${{multiSegPct}}%; background-color: #0072B2;"></div>
              </div>
              <div class="text-xs font-semibold text-slate-800">
                <span class="font-bold text-slate-900">${{multiSegOrgs}}</span> of ${{totalOrgs}} orgs serve 2+ sectors
              </div>
            </div>
            <p class="text-xs text-slate-500 mt-2 leading-tight">Companies active across multiple sectors</p>
          </div>`;
      }}

      // 4. Multi-Product Portfolio
      const productEl = document.getElementById('cscProductExpansion');
      if (productEl) {{
        const multiProductOrgs = orgsHere.filter(d => d.p1_name && d.p2_name).length;
        const multiProductPct = Math.round((multiProductOrgs / totalOrgs) * 100);
        productEl.innerHTML = `
          <div>
            <div class="flex items-center justify-between mb-2">
              <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                <span>📦</span> Multi-Product Portfolio
              </span>
              <span class="text-xs font-extrabold text-amber-700 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-full">${{multiProductPct}}%</span>
            </div>
            <div class="space-y-1.5 my-1">
              <div class="w-full bg-slate-100 h-2 rounded-full overflow-hidden shadow-inner">
                <div class="h-full rounded-full" style="width: ${{multiProductPct}}%; background-color: #E69F00;"></div>
              </div>
              <div class="text-xs font-semibold text-slate-800">
                <span class="font-bold text-slate-900">${{multiProductOrgs}}</span> of ${{totalOrgs}} orgs offer 2+ products
              </div>
            </div>
            <p class="text-xs text-slate-500 mt-2 leading-tight">Companies with multi-product catalog</p>
          </div>`;
      }}

      // Update filter button label
      const isAlreadyFiltered = (selectedCountry !== 'ALL' && getCanonicalCountry(selectedCountry) === canon);
      const btn = document.getElementById('cscFilterBtn');
      if (btn) {{
        btn.textContent = isAlreadyFiltered ? '✓ Filtered — click to clear' : 'Filter to this country →';
        btn.className = isAlreadyFiltered
          ? 'px-3 py-1.5 text-xs font-semibold bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors'
          : 'px-3 py-1.5 text-xs font-semibold bg-brand-600 hover:bg-brand-700 text-white rounded-lg transition-colors';
      }}

      card.classList.remove('hidden');
    }}

    function closeCountryStatsCard() {{
      const card = document.getElementById('countryStatsCard');
      if (card) card.classList.add('hidden');
      _cscPinnedCountry = null;
    }}

    function applyCountryStatsFilter() {{
      if (!_cscPinnedCountry) return;
      const {{ canon, displayName }} = _cscPinnedCountry;
      if (selectedCountry !== 'ALL' && getCanonicalCountry(selectedCountry) === canon) {{
        selectedCountry = 'ALL';
      }} else {{
        selectedCountry = displayName;
      }}
      applyFilters();
    }}
    // ── End Country Stats Card ─────────────────────────────────────────────────

    let map, markerClusterGroup, choroplethLayer = null;
    let currentMapMode = 'pin'; // 'pin' or 'choropleth'
    let countryChartInstance = null;
    let sizeChartInstance = null;

    let selectedSegment = 'ALL';
    let selectedCountry = 'ALL';
    let selectedSize = 'ALL';

    let filteredData = [...RAW_DATA];
    let currentPage = 1;
    const pageSize = 15;
    let currentBoundsMode = 'focus';
    let activeDrawerRecord = null;

    function initMap() {{
      map = L.map('map', {{
        zoomControl: true,
        scrollWheelZoom: true
      }});

      L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
        attribution: 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ, TomTom, USGS, FAO, NPS, NRCAN, GeoBase, Kadaster NL, Ordnance Survey',
        maxZoom: 18
      }}).addTo(map);

      markerClusterGroup = L.markerClusterGroup({{
        showCoverageOnHover: false,
        zoomToBoundsOnClick: true,
        spiderfyOnMaxZoom: true,
        maxClusterRadius: function(zoom) {{
          if (zoom >= 13) return 20;
          if (zoom >= 10) return 35;
          if (zoom >= 7) return 50;
          return 65;
        }},
        iconCreateFunction: function(cluster) {{
          const children = cluster.getAllChildMarkers();
          const count = children.length;

          const sizeCounts = {{}};
          children.forEach(m => {{
            const sz = m.options.companySize || 'Unspecified';
            sizeCounts[sz] = (sizeCounts[sz] || 0) + 1;
          }});

          let majoritySize = 'Unspecified';
          let maxFreq = -1;
          Object.entries(sizeCounts).forEach(([sz, freq]) => {{
            if (freq > maxFreq) {{
              maxFreq = freq;
              majoritySize = sz;
            }}
          }});

          const clusterBgColor = SIZE_COLOR_MAP[majoritySize] || OKABE_ITO.blue;

          let sizeClass = 'w-8 h-8 text-xs';
          if (count > 50) sizeClass = 'w-10 h-10 text-sm';
          if (count > 200) sizeClass = 'w-12 h-12 text-sm';

          return L.divIcon({{
            html: `<div class="custom-cluster ${{sizeClass}}" style="background-color: ${{clusterBgColor}};">${{count}}</div>`,
            className: '',
            iconSize: L.point(40, 40)
          }});
        }}
      }});

      map.addLayer(markerClusterGroup);

      document.getElementById('btnFocusView').addEventListener('click', () => {{
        currentBoundsMode = 'focus';
        updateToggleButtons();
        fitMapBounds();
      }});

      document.getElementById('btnGlobalView').addEventListener('click', () => {{
        currentBoundsMode = 'global';
        updateToggleButtons();
        fitMapBounds();
      }});

      document.getElementById('btnPinMode').addEventListener('click', () => {{
        currentMapMode = 'pin';
        updateToggleButtons();
        updateMap();
      }});

      document.getElementById('btnChoroplethMode').addEventListener('click', () => {{
        currentMapMode = 'choropleth';
        currentBoundsMode = 'focus';
        updateToggleButtons();
        updateMap(true);

        if (selectedCountry !== 'ALL') {{
          const canon = getCanonicalCountry(selectedCountry);
          const cRecs = RAW_DATA.filter(d => getCanonicalCountry(d.country) === canon && d.lat !== null && d.lon !== null);
          if (cRecs.length > 0) {{
            const cBounds = cRecs.map(d => [d.lat, d.lon]);
            map.fitBounds(cBounds, {{ padding: [40, 40], maxZoom: 6 }});
          }} else {{
            map.setView([51.0, 10.0], 4);
          }}
        }} else {{
          map.setView([51.0, 10.0], 4);
        }}
      }});
    }}

    function updateToggleButtons() {{
      const btnFocus = document.getElementById('btnFocusView');
      const btnGlobal = document.getElementById('btnGlobalView');

      if (currentBoundsMode === 'focus') {{
        btnFocus.className = 'px-3.5 py-1.5 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all';
        btnGlobal.className = 'px-3.5 py-1.5 rounded-md text-slate-600 hover:text-slate-900 transition-all';
      }} else {{
        btnGlobal.className = 'px-3.5 py-1.5 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all';
        btnFocus.className = 'px-3.5 py-1.5 rounded-md text-slate-600 hover:text-slate-900 transition-all';
      }}

      const btnPin = document.getElementById('btnPinMode');
      const btnChoropleth = document.getElementById('btnChoroplethMode');

      if (currentMapMode === 'pin') {{
        btnPin.className = 'px-3.5 py-1.5 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all';
        btnChoropleth.className = 'px-3.5 py-1.5 rounded-md text-slate-600 hover:text-slate-900 transition-all';
      }} else {{
        btnChoropleth.className = 'px-3.5 py-1.5 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all';
        btnPin.className = 'px-3.5 py-1.5 rounded-md text-slate-600 hover:text-slate-900 transition-all';
      }}
    }}

    function populateSegmentFilter() {{
      const segmentSet = new Set();
      RAW_DATA.forEach(d => {{
        if (d.segments && Array.isArray(d.segments)) {{
          d.segments.forEach(s => segmentSet.add(s));
        }}
      }});

      const sorted = Array.from(segmentSet).sort();
      const select = document.getElementById('segmentFilter');
      sorted.forEach(s => {{
        const opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        select.appendChild(opt);
      }});

      select.addEventListener('change', () => {{
        selectedSegment = select.value;
        applyFilters();
      }});

      document.getElementById('searchInput').addEventListener('input', () => {{
        currentPage = 1;
        renderTable();
      }});

      document.getElementById('clearAllBtn').addEventListener('click', () => {{
        selectedSegment = 'ALL';
        selectedCountry = 'ALL';
        selectedSize = 'ALL';
        document.getElementById('segmentFilter').value = 'ALL';
        closeCountryStatsCard();
        applyFilters();
      }});
    }}

    function applyFilters() {{
      filteredData = RAW_DATA.filter(d => {{
        const matchSeg = (selectedSegment === 'ALL') || (d.segments && d.segments.includes(selectedSegment));
        const matchCountry = (selectedCountry === 'ALL') || (d.country === selectedCountry);
        const matchSize = (selectedSize === 'ALL') || (d.size === selectedSize);
        return matchSeg && matchCountry && matchSize;
      }});

      const isOnlyCountryFiltered = (selectedCountry !== 'ALL' && selectedSegment === 'ALL' && selectedSize === 'ALL');
      if (isOnlyCountryFiltered) {{
        const canon = getCanonicalCountry(selectedCountry);
        const matchRec = RAW_DATA.find(d => getCanonicalCountry(d.country) === canon);
        const iso2 = matchRec ? matchRec.country : selectedCountry;
        openCountryStatsCard(canon, selectedCountry, iso2, {{}});
      }} else {{
        closeCountryStatsCard();
      }}

      currentPage = 1;
      updateFilterChips();
      updateMap();
      updateCharts();
      renderTable();
    }}

    function updateFilterChips() {{
      const bar = document.getElementById('filterChipsBar');
      const container = document.getElementById('chipsContainer');
      container.innerHTML = '';

      let activeCount = 0;

      if (selectedSegment !== 'ALL') {{
        activeCount++;
        container.appendChild(createChip(`Segment: ${{selectedSegment}}`, () => {{
          selectedSegment = 'ALL';
          document.getElementById('segmentFilter').value = 'ALL';
          applyFilters();
        }}));
      }}

      if (selectedCountry !== 'ALL') {{
        activeCount++;
        container.appendChild(createChip(`Country: ${{selectedCountry}}`, () => {{
          selectedCountry = 'ALL';
          closeCountryStatsCard();
          applyFilters();
        }}));
      }}

      if (selectedSize !== 'ALL') {{
        activeCount++;
        container.appendChild(createChip(`Scale: ${{selectedSize}}`, () => {{
          selectedSize = 'ALL';
          applyFilters();
        }}));
      }}

      if (activeCount > 0) {{
        bar.classList.remove('hidden');
      }} else {{
        bar.classList.add('hidden');
      }}

      const countryStatus = document.getElementById('countryFilterStatus');
      if (selectedCountry !== 'ALL') {{
        countryStatus.textContent = `Filtered: ${{selectedCountry}}`;
        countryStatus.classList.remove('hidden');
      }} else {{
        countryStatus.classList.add('hidden');
      }}

      const sizeStatus = document.getElementById('sizeFilterStatus');
      if (selectedSize !== 'ALL') {{
        sizeStatus.textContent = `Filtered: ${{selectedSize}}`;
        sizeStatus.classList.remove('hidden');
      }} else {{
        sizeStatus.classList.add('hidden');
      }}
    }}

    function createChip(text, onRemove) {{
      const chip = document.createElement('span');
      chip.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold bg-white border border-slate-300 text-slate-800 shadow-2xs';
      chip.innerHTML = `<span>${{text}}</span><button class="hover:text-red-600 transition-colors font-bold text-sm leading-none ml-1">&times;</button>`;
      chip.querySelector('button').addEventListener('click', onRemove);
      return chip;
    }}

    function fitMapBounds(forceRefocus = false) {{
      if (currentMapMode === 'choropleth' && !forceRefocus) {{
        return;
      }}

      const validPoints = filteredData.filter(d => d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon));
      if (validPoints.length === 0) {{
        map.setView([50.0, 10.0], 4);
        return;
      }}

      if (selectedCountry !== 'ALL' && !forceRefocus) {{
        const countryPoints = validPoints.map(d => [d.lat, d.lon]);
        map.fitBounds(countryPoints, {{ padding: [40, 40], maxZoom: 12 }});
        return;
      }}

      if (currentBoundsMode === 'global' && !forceRefocus) {{
        const allBounds = validPoints.map(d => [d.lat, d.lon]);
        map.fitBounds(allBounds, {{ padding: [30, 30], maxZoom: 13 }});
        return;
      }}

      const lats = validPoints.map(d => d.lat).sort((a, b) => a - b);
      const lons = validPoints.map(d => d.lon).sort((a, b) => a - b);

      const pLow = 0.02;
      const pHigh = 0.98;

      const minLat = lats[Math.floor(lats.length * pLow)];
      const maxLat = lats[Math.min(lats.length - 1, Math.floor(lats.length * pHigh))];
      const minLon = lons[Math.floor(lons.length * pLow)];
      const maxLon = lons[Math.min(lons.length - 1, Math.floor(lons.length * pHigh))];

      map.fitBounds([[minLat, minLon], [maxLat, maxLon]], {{ padding: [30, 30], maxZoom: 13 }});
    }}

    function getChoroplethColor(cCount) {{
      if (!cCount || cCount === 0) return '#f8fafc';
      if (cCount <= 5) return '#ffffcc';
      if (cCount <= 15) return '#c7e9b4';
      if (cCount <= 35) return '#7fcdbb';
      if (cCount <= 75) return '#1d91c0';
      return '#0c2c84';
    }}

    function updateMap(skipFitBounds = false) {{
      const pinLegend = document.getElementById('pinLegend');
      const choroplethLegend = document.getElementById('choroplethLegend');

      if (currentMapMode === 'pin') {{
        if (pinLegend) pinLegend.classList.remove('hidden');
        if (choroplethLegend) choroplethLegend.classList.add('hidden');

        if (choroplethLayer && map.hasLayer(choroplethLayer)) {{
          map.removeLayer(choroplethLayer);
        }}
        if (!map.hasLayer(markerClusterGroup)) {{
          map.addLayer(markerClusterGroup);
        }}

        markerClusterGroup.clearLayers();
        let count = 0;

        const coordGroups = {{}};
        filteredData.forEach(d => {{
          if (d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon)) {{
            const key = `${{d.lat.toFixed(5)}},${{d.lon.toFixed(5)}}`;
            if (!coordGroups[key]) coordGroups[key] = [];
            coordGroups[key].push(d);
          }}
        }});

        Object.values(coordGroups).forEach(group => {{
          const total = group.length;
          group.forEach((d, idx) => {{
            count++;
            let mLat = d.lat;
            let mLon = d.lon;

            if (total > 1 && idx > 0) {{
              const goldenAngle = 2.3999632;
              const r = 0.0016 * Math.sqrt(idx);
              const theta = idx * goldenAngle;
              const cosLat = Math.cos(d.lat * Math.PI / 180) || 1;
              mLat = d.lat + r * Math.cos(theta);
              mLon = d.lon + (r * Math.sin(theta)) / cosLat;
            }}

            const markerColor = SIZE_COLOR_MAP[d.size] || OKABE_ITO.gray;

            const marker = L.circleMarker([mLat, mLon], {{
              radius: 6,
              fillColor: markerColor,
              color: '#ffffff',
              weight: 2,
              opacity: 1,
              fillOpacity: 0.9,
              companySize: d.size,
              record: d
            }});

            const sizeBadge = d.size ? `<span class="inline-block mt-1 px-1.5 py-0.5 text-[10px] font-semibold rounded text-white" style="background-color: ${{markerColor}}">${{d.size}}</span>` : '';

            const logoHtml = d.logo
              ? `<img src="${{d.logo}}" alt="${{d.name}} logo" class="h-8 w-auto max-w-[80px] object-contain rounded mb-1" onerror="this.style.display='none'">`
              : '';

            const popupContent = `
              <div class="font-sans">
                ${{logoHtml}}
                <div class="font-bold text-slate-900 text-sm">${{d.name}}</div>
                <div class="text-xs text-slate-500 mt-0.5">${{d.cc || d.country || 'Location unspecified'}}</div>
                ${{sizeBadge}}
                <div class="mt-2 pt-1.5 border-t border-slate-100 flex items-center justify-between gap-2">
                  <button onclick="openDrawerById('${{d.id}}')" class="text-[11px] font-bold text-brand-600 hover:text-brand-800 transition-colors flex items-center gap-1">
                    <span>View Full Profile</span> &rarr;
                  </button>
                </div>
              </div>
            `;
            marker.bindPopup(popupContent);
            markerClusterGroup.addLayer(marker);
          }});
        }});

        document.getElementById('mapCount').textContent = count.toLocaleString();
        if (!skipFitBounds) fitMapBounds();
      }} else {{
        // Choropleth Mode
        if (choroplethLegend) choroplethLegend.classList.remove('hidden');
        if (pinLegend) pinLegend.classList.add('hidden');

        if (map.hasLayer(markerClusterGroup)) {{
          map.removeLayer(markerClusterGroup);
        }}
        if (choroplethLayer && map.hasLayer(choroplethLayer)) {{
          map.removeLayer(choroplethLayer);
        }}

        // Calculate counts based on RAW_DATA filtered by segment and size (so ALL countries remain clickable to switch)
        const baseCountsByCanon = {{}};
        const baseForChoropleth = RAW_DATA.filter(d => {{
          const matchSeg = (selectedSegment === 'ALL') || (d.segments && d.segments.includes(selectedSegment));
          const matchSize = (selectedSize === 'ALL') || (d.size === selectedSize);
          return matchSeg && matchSize;
        }});

        baseForChoropleth.forEach(d => {{
          const c = d.country ? getCanonicalCountry(d.country) : '';
          if (c) {{
            baseCountsByCanon[c] = (baseCountsByCanon[c] || 0) + 1;
          }}
        }});

        choroplethLayer = L.geoJSON(COUNTRY_GEOJSON, {{
          style: function(feature) {{
            const p = feature.properties;
            const name = p.name || p.name_long || '';
            const canon = getCanonicalCountry(name);
            const cCount = baseCountsByCanon[canon] || 0;
            const isSelected = (selectedCountry !== 'ALL' && getCanonicalCountry(selectedCountry) === canon);

            return {{
              fillColor: isSelected ? '#d97706' : getChoroplethColor(cCount),
              weight: isSelected ? 3.0 : (cCount > 0 ? 1.0 : 0.5),
              opacity: 1.0,
              color: isSelected ? '#78350f' : (cCount > 0 ? '#1d91c0' : '#cbd5e1'),
              fillOpacity: isSelected ? 0.9 : (cCount > 0 ? 0.8 : 0.15)
            }};
          }},
          onEachFeature: function(feature, layer) {{
            const p = feature.properties;
            const displayName = p.name || p.name_long || 'Unknown';
            const canon = getCanonicalCountry(displayName);
            const cCount = baseCountsByCanon[canon] || 0;

            if (cCount > 0) {{
              layer.bindTooltip(`
                <div class="px-2 py-1 font-sans text-xs">
                  <div class="font-bold text-slate-900">${{displayName}}</div>
                  <div class="text-slate-600 mt-0.5">${{cCount}} ${{cCount === 1 ? 'organization' : 'organizations'}}</div>
                </div>
              `, {{ sticky: true }});
            }}

            layer.on({{
              mouseover: function(e) {{
                if (cCount > 0) {{
                  const l = e.target;
                  const isSel = (selectedCountry !== 'ALL' && getCanonicalCountry(selectedCountry) === canon);
                  l.setStyle({{
                    weight: isSel ? 3.5 : 2.0,
                    color: isSel ? '#451a03' : '#0f172a',
                    fillOpacity: 0.95
                  }});
                }}
              }},
              mouseout: function(e) {{
                if (choroplethLayer) choroplethLayer.resetStyle(e.target);
              }},
              click: function() {{
                if (cCount > 0) {{
                  const iso2 = p.iso2 || displayName;
                  const matchRecord = RAW_DATA.find(d => getCanonicalCountry(d.country) === canon);
                  const targetCountryName = matchRecord ? matchRecord.country : displayName;

                  if (selectedCountry !== 'ALL' && getCanonicalCountry(selectedCountry) === canon) {{
                    selectedCountry = 'ALL';
                  }} else {{
                    selectedCountry = targetCountryName;
                  }}
                  applyFilters();
                }}
              }}
            }});
          }}
        }}).addTo(map);

        document.getElementById('mapCount').textContent = filteredData.length.toLocaleString();
        if (!skipFitBounds) fitMapBounds();
      }}
    }}

    function locateRecordOnMap(d) {{
      if (d && d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon)) {{
        if (currentMapMode !== 'pin') {{
          currentMapMode = 'pin';
          updateToggleButtons();
          updateMap(true);
        }}

        setTimeout(() => {{
          const layers = markerClusterGroup.getLayers();
          const targetMarker = layers.find(m => m.options.record && m.options.record.id === d.id);

          if (targetMarker) {{
            const latLng = targetMarker.getLatLng();
            map.flyTo(latLng, 14, {{ duration: 1.0 }});
            markerClusterGroup.zoomToShowLayer(targetMarker, () => {{
              targetMarker.openPopup();
            }});
          }} else {{
            map.flyTo([d.lat, d.lon], 14, {{ duration: 1.0 }});
          }}
        }}, 50);

        document.getElementById('map').scrollIntoView({{ behavior: 'smooth', block: 'center' }});
      }}
    }}

    function updateCharts() {{
      const countryCounts = {{}};
      const baseForCountries = RAW_DATA.filter(d => {{
        const matchSeg = (selectedSegment === 'ALL') || (d.segments && d.segments.includes(selectedSegment));
        const matchSize = (selectedSize === 'ALL') || (d.size === selectedSize);
        return matchSeg && matchSize;
      }});

      baseForCountries.forEach(d => {{
        const c = d.country || 'Unknown';
        countryCounts[c] = (countryCounts[c] || 0) + 1;
      }});

      const sortedCountries = Object.entries(countryCounts)
        .sort((a, b) => b[1] - a[1]);

      let topList = sortedCountries.slice(0, 8);

      if (selectedCountry !== 'ALL') {{
        const selCanon = getCanonicalCountry(selectedCountry);
        const inTop = topList.some(e => getCanonicalCountry(e[0]) === selCanon);
        if (!inTop) {{
          topList = sortedCountries.slice(0, 7);
          const selVal = countryCounts[selectedCountry] || baseForCountries.filter(d => getCanonicalCountry(d.country) === selCanon).length;
          topList.push([selectedCountry, selVal]);
        }}
      }}

      const rawCountryNames = topList.map(e => e[0]);
      const countryLabels = topList.map(e => {{
        const name = e[0];
        if (selectedCountry !== 'ALL' && getCanonicalCountry(name) === getCanonicalCountry(selectedCountry)) {{
          return `${{name}} ★`;
        }}
        return name;
      }});
      const countryValues = topList.map(e => e[1]);

      const barColors = topList.map(e => {{
        const isSel = (selectedCountry !== 'ALL' && getCanonicalCountry(e[0]) === getCanonicalCountry(selectedCountry));
        if (isSel) {{
          return '#d97706';
        }}
        return (selectedCountry !== 'ALL') ? '#cbd5e1' : OKABE_ITO.blue;
      }});

      const hoverBarColors = topList.map(e => {{
        const isSel = (selectedCountry !== 'ALL' && getCanonicalCountry(e[0]) === getCanonicalCountry(selectedCountry));
        if (isSel) {{
          return '#b45309';
        }}
        return OKABE_ITO.vermillion;
      }});

      if (countryChartInstance) countryChartInstance.destroy();
      const ctxC = document.getElementById('countryChart').getContext('2d');
      countryChartInstance = new Chart(ctxC, {{
        type: 'bar',
        data: {{
          labels: countryLabels,
          datasets: [{{
            label: 'Organizations',
            data: countryValues,
            backgroundColor: barColors,
            hoverBackgroundColor: hoverBarColors,
            borderRadius: 6
          }}]
        }},
        options: {{
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ display: false }},
            tooltip: {{ padding: 10, cornerRadius: 8 }}
          }},
          scales: {{
            x: {{ grid: {{ display: false }} }},
            y: {{ grid: {{ display: false }} }}
          }},
          onHover: (evt, elements) => {{
            evt.native.target.style.cursor = elements.length ? 'pointer' : 'default';
          }},
          onClick: (evt, elements) => {{
            if (elements.length > 0) {{
              const idx = elements[0].index;
              const clickedCountry = rawCountryNames[idx];
              if (selectedCountry !== 'ALL' && getCanonicalCountry(selectedCountry) === getCanonicalCountry(clickedCountry)) {{
                selectedCountry = 'ALL';
              }} else {{
                selectedCountry = clickedCountry;
              }}
              applyFilters();
            }}
          }}
        }}
      }});

      const sizeCounts = {{}};
      const baseForSizes = RAW_DATA.filter(d => {{
        const matchSeg = (selectedSegment === 'ALL') || (d.segments && d.segments.includes(selectedSegment));
        const matchCountry = (selectedCountry === 'ALL') || (d.country === selectedCountry);
        return matchSeg && matchCountry;
      }});

      baseForSizes.forEach(d => {{
        if (d.size) {{
          sizeCounts[d.size] = (sizeCounts[d.size] || 0) + 1;
        }}
      }});

      const preferredOrder = ['Micro', 'Small', 'Medium', 'Large', 'Solo entrepreneur', 'Medium or Large (unspecified)'];
      const sizeLabels = [];
      const sizeValues = [];

      preferredOrder.forEach(key => {{
        if (sizeCounts[key]) {{
          sizeLabels.push(key);
          sizeValues.push(sizeCounts[key]);
        }}
      }});

      Object.entries(sizeCounts).forEach(([k, v]) => {{
        if (!preferredOrder.includes(k)) {{
          sizeLabels.push(k);
          sizeValues.push(v);
        }}
      }});

      const donutColors = sizeLabels.map(lbl => {{
        if (selectedSize !== 'ALL' && selectedSize !== lbl) {{
          return OKABE_ITO.lightGray;
        }}
        return SIZE_COLOR_MAP[lbl] || OKABE_ITO.gray;
      }});

      if (sizeChartInstance) sizeChartInstance.destroy();
      const ctxS = document.getElementById('sizeChart').getContext('2d');
      sizeChartInstance = new Chart(ctxS, {{
        type: 'doughnut',
        data: {{
          labels: sizeLabels.length ? sizeLabels : ['Unspecified'],
          datasets: [{{
            data: sizeValues.length ? sizeValues : [1],
            backgroundColor: sizeValues.length ? donutColors : [OKABE_ITO.lightGray],
            borderWidth: 2,
            borderColor: '#ffffff'
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{
              position: 'bottom',
              labels: {{ boxWidth: 12, font: {{ size: 11 }} }}
            }},
            tooltip: {{ padding: 10, cornerRadius: 8 }}
          }},
          cutout: '65%',
          onHover: (evt, elements) => {{
            evt.native.target.style.cursor = elements.length ? 'pointer' : 'default';
          }},
          onClick: (evt, elements) => {{
            if (elements.length > 0) {{
              const idx = elements[0].index;
              const clickedSize = sizeLabels[idx];
              selectedSize = (selectedSize === clickedSize) ? 'ALL' : clickedSize;
              applyFilters();
            }}
          }}
        }}
      }});
    }}

    function renderTable() {{
      const query = document.getElementById('searchInput').value.toLowerCase().trim();
      let tableData = filteredData;

      if (query) {{
        tableData = filteredData.filter(d => {{
          return (d.name && d.name.toLowerCase().includes(query)) ||
                 (d.cc && d.cc.toLowerCase().includes(query)) ||
                 (d.country && d.country.toLowerCase().includes(query));
        }});
      }}

      const total = tableData.length;
      const totalPages = Math.ceil(total / pageSize) || 1;
      if (currentPage > totalPages) currentPage = totalPages;

      const start = (currentPage - 1) * pageSize;
      const end = Math.min(start + pageSize, total);
      const pageSlice = tableData.slice(start, end);

      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';

      if (pageSlice.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="4" class="px-5 py-6 text-center text-slate-400">No organizations found matching current filters</td></tr>`;
      }} else {{
        pageSlice.forEach(d => {{
          const tr = document.createElement('tr');
          const hasCoords = d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon);
          tr.className = `hover:bg-slate-100/80 transition-colors cursor-pointer ${{hasCoords ? '' : 'opacity-75'}}`;
          tr.title = 'Click to view full company profile & locate on map';

          let displaySegs = d.segments || [];
          if (selectedSegment !== 'ALL' && displaySegs.includes(selectedSegment)) {{
            displaySegs = [selectedSegment, ...displaySegs.filter(s => s !== selectedSegment)];
          }}

          let segBadges = '';
          if (displaySegs.length > 0) {{
            const visible = displaySegs.slice(0, 3);
            const extraCount = displaySegs.length - visible.length;
            segBadges = visible.map(s => {{
              const isMatch = (selectedSegment !== 'ALL' && s === selectedSegment);
              const badgeStyle = isMatch 
                ? 'bg-brand-50 text-brand-700 border-brand-200 font-semibold shadow-2xs' 
                : 'bg-slate-100 text-slate-700 border-slate-200';
              return `<span class="inline-block px-2 py-0.5 mr-1 mb-1 text-[11px] ${{badgeStyle}} rounded-md border">${{s}}</span>`;
            }}).join('');
            
            if (extraCount > 0) {{
              const extraTooltip = displaySegs.slice(3).join(', ');
              segBadges += `<span class="inline-block px-1.5 py-0.5 text-[11px] font-medium bg-slate-50 text-slate-500 rounded-md border border-slate-200 cursor-help" title="${{extraTooltip}}">+${{extraCount}} more</span>`;
            }}
          }} else {{
            segBadges = '<span class="text-slate-400 text-xs">-</span>';
          }}

          tr.innerHTML = `
            <td class="px-5 py-3 font-semibold text-slate-900 group-hover:text-brand-600 flex items-center gap-1.5">
              <span>${{d.name}}</span>
              ${{hasCoords ? '<svg class="w-3.5 h-3.5 text-slate-400 opacity-60" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"/></svg>' : ''}}
            </td>
            <td class="px-5 py-3 text-slate-600">${{getCityOnly(d, true)}}</td>
            <td class="px-5 py-3 text-slate-600">${{d.country || '-'}}</td>
            <td class="px-5 py-3">${{segBadges}}</td>
          `;

          tr.addEventListener('click', () => {{
            openDrawer(d);
          }});

          tbody.appendChild(tr);
        }});
      }}

      document.getElementById('pageInfo').textContent = `Showing ${{total === 0 ? 0 : start + 1}} to ${{end}} of ${{total.toLocaleString()}} entries`;
      document.getElementById('prevBtn').disabled = currentPage <= 1;
      document.getElementById('nextBtn').disabled = currentPage >= totalPages;
    }}

    function sanitizeSocialUrl(raw, fallbackPlat) {{
      if (!raw) return null;
      let val = String(raw).trim();
      const lower = val.toLowerCase();
      if (!val || ['none', 'n/a', '-', 'no', 'null', 'undefined', 'in progress'].includes(lower)) {{
        return null;
      }}
      if (val.startsWith('@')) {{
        const handle = val.replace(/^@+/, '').trim();
        if (!handle) return null;
        if (fallbackPlat === 'x') return `https://x.com/${{handle}}`;
        if (fallbackPlat === 'insta') return `https://www.instagram.com/${{handle}}`;
        if (fallbackPlat === 'facebook') return `https://www.facebook.com/${{handle}}`;
        if (fallbackPlat === 'linkedin') return `https://www.linkedin.com/company/${{handle}}`;
      }}
      if (val.includes(' ') && !val.includes('/') && !val.includes('.')) {{
        return null;
      }}
      if (!val.startsWith('http://') && !val.startsWith('https://')) {{
        if (val.includes('.') && !val.includes(' ')) {{
          val = 'https://' + val.replace(/^\\/+/, '');
        }} else {{
          return null;
        }}
      }}
      if (!/^https?:\\/\\/[a-zA-Z0-9.-]+\\.[a-zA-Z]{{2,}}/i.test(val)) {{
        return null;
      }}
      return val;
    }}

    function detectSocialPlatform(url) {{
      const u = url.toLowerCase();
      if (u.includes('linkedin.com')) return 'linkedin';
      if (u.includes('twitter.com') || u.includes('x.com')) return 'x';
      if (u.includes('instagram.com')) return 'insta';
      if (u.includes('facebook.com') || u.includes('fb.me') || u.includes('fb.com')) return 'facebook';
      if (u.includes('youtube.com') || u.includes('youtu.be')) return 'youtube';
      return 'other';
    }}

    function buildSocialLinksHtml(record) {{
      const rawEntries = [
        {{ raw: record.social_linkedin, plat: 'linkedin' }},
        {{ raw: record.social_x, plat: 'x' }},
        {{ raw: record.social_insta, plat: 'insta' }},
        {{ raw: record.social_facebook, plat: 'facebook' }}
      ];

      const platformConfig = {{
        linkedin: {{ label: 'LinkedIn', style: 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100' }},
        x: {{ label: 'X / Twitter', style: 'bg-slate-100 text-slate-800 border-slate-300 hover:bg-slate-200' }},
        insta: {{ label: 'Instagram', style: 'bg-pink-50 text-pink-700 border-pink-200 hover:bg-pink-100' }},
        facebook: {{ label: 'Facebook', style: 'bg-indigo-50 text-indigo-700 border-indigo-200 hover:bg-indigo-100' }},
        youtube: {{ label: 'YouTube', style: 'bg-red-50 text-red-700 border-red-200 hover:bg-red-100' }},
        other: {{ label: 'Social Profile', style: 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100' }}
      }};

      const seenUrls = new Set();
      const seenPlatforms = new Set();
      const badges = [];

      rawEntries.forEach(entry => {{
        const cleanUrl = sanitizeSocialUrl(entry.raw, entry.plat);
        if (!cleanUrl) return;

        const normKey = cleanUrl.toLowerCase().replace(/\\/+$/, '');
        if (seenUrls.has(normKey)) return;
        seenUrls.add(normKey);

        const realPlat = detectSocialPlatform(cleanUrl);
        if (seenPlatforms.has(realPlat) && realPlat !== 'other') return;
        seenPlatforms.add(realPlat);

        const cfg = platformConfig[realPlat] || platformConfig.other;
        badges.push(
          `<a href="${{cleanUrl}}" target="_blank" rel="noopener noreferrer" class="inline-flex items-center gap-1.5 px-3 py-1.5 ${{cfg.style}} rounded-lg text-xs font-semibold border transition-colors shadow-2xs"><span>${{cfg.label}}</span> &nearr;</a>`
        );
      }});

      if (badges.length === 0) {{
        return '<span class="text-xs text-slate-400 italic">No verified social profiles linked.</span>';
      }}
      return badges.join('');
    }}

    function renderProductBlock(pName, pDesc, defaultTitle) {{
      if (!pName && !pDesc) return '';

      const isNameUrl = pName && String(pName).trim().startsWith('http');
      const isDescUrl = pDesc && String(pDesc).trim().startsWith('http');

      const title = (!isNameUrl && pName) ? pName : (isNameUrl ? 'Product Resource' : defaultTitle);

      let contentHtml = '';
      if (isNameUrl && isDescUrl && pName.trim() === pDesc.trim()) {{
        contentHtml = `<a href="${{pName.trim()}}" target="_blank" rel="noopener noreferrer" class="inline-flex items-center gap-1 text-xs font-semibold text-brand-600 hover:text-brand-800 underline break-all"><span>View Linked Document / Resource</span> &nearr;</a>`;
      }} else if (isNameUrl) {{
        contentHtml = `<a href="${{pName.trim()}}" target="_blank" rel="noopener noreferrer" class="inline-flex items-center gap-1 text-xs font-semibold text-brand-600 hover:text-brand-800 underline break-all"><span>View Resource</span> &nearr;</a>`;
        if (pDesc) {{
          contentHtml += `<p class="text-xs text-slate-600 leading-relaxed mt-1">${{pDesc}}</p>`;
        }}
      }} else if (isDescUrl) {{
        contentHtml = `<a href="${{pDesc.trim()}}" target="_blank" rel="noopener noreferrer" class="inline-flex items-center gap-1 text-xs font-semibold text-brand-600 hover:text-brand-800 underline break-all"><span>View Product Documentation / Asset</span> &nearr;</a>`;
      }} else {{
        contentHtml = `<p class="text-xs text-slate-600 leading-relaxed">${{pDesc || 'No detailed description available.'}}</p>`;
      }}

      return `
        <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-200 space-y-1">
          <div class="font-bold text-slate-900 text-xs sm:text-sm flex items-center gap-2">
            <span>📦 ${{title}}</span>
          </div>
          ${{contentHtml}}
        </div>
      `;
    }}

    function openDrawer(record) {{
      if (!record) return;
      activeDrawerRecord = record;

      const container = document.getElementById('drawerContainer');
      const body = document.getElementById('drawerBody');

      let webUrl = record.website ? String(record.website).trim() : '';
      if (webUrl && !webUrl.startsWith('http://') && !webUrl.startsWith('https://')) {{
        if (webUrl.includes('.') && !webUrl.includes(' ')) {{
          webUrl = 'https://' + webUrl.replace(/^\\/+/, '');
        }} else {{
          webUrl = '';
        }}
      }}

      const isVerified = record.verified && record.verified.toLowerCase().includes('verified') && !record.verified.toLowerCase().includes('unverified');
      const verifiedBadge = isVerified 
        ? `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">✓ Verified</span>`
        : `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600 border border-slate-200">Unverified</span>`;

      const segBadges = (record.segments && record.segments.length > 0)
        ? record.segments.map(s => `<span class="inline-block px-2.5 py-1 text-xs font-semibold bg-brand-50 text-brand-700 rounded-md border border-brand-100 mr-1 mb-1">${{s}}</span>`).join('')
        : '<span class="text-slate-400 text-xs">Unspecified</span>';

      let productsHtml = '';
      const p1Html = renderProductBlock(record.p1_name, record.p1_desc, 'Product 1');
      const p2Html = renderProductBlock(record.p2_name, record.p2_desc, 'Product 2');
      if (p1Html) productsHtml += p1Html;
      if (p2Html) productsHtml += p2Html;
      if (!productsHtml) {{
        productsHtml = '<p class="text-xs text-slate-400 italic">No specific products documented for this entry.</p>';
      }}

      const socialsHtml = buildSocialLinksHtml(record);

      const hasCoords = record.lat !== null && record.lon !== null && !isNaN(record.lat) && !isNaN(record.lon);

      body.innerHTML = `
        <div class="space-y-3 border-b border-slate-200 pb-5">
          <div class="flex items-start justify-between gap-3">
            <div class="flex items-center gap-3">
              ${{record.logo ? `<img src="${{record.logo}}" alt="${{record.name}} logo" class="h-12 w-12 object-contain rounded-lg border border-slate-100 bg-white p-0.5 shrink-0" onerror="this.style.display='none'">` : ''}}
              <div>
                <h2 class="text-xl font-bold text-slate-900 tracking-tight leading-snug">${{record.name}}</h2>
                <div class="text-xs font-medium text-slate-500 mt-1 flex items-center gap-2">
                  <span>📍 ${{record.cc || record.country || 'Location Unspecified'}}</span>
                </div>
              </div>
            </div>
            ${{verifiedBadge}}
          </div>

          <div class="flex items-center gap-2 flex-wrap pt-1">
            ${{webUrl ? `<a href="${{webUrl}}" target="_blank" rel="noopener noreferrer" class="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-brand-500 hover:bg-brand-600 text-white rounded-lg text-xs font-bold shadow-xs transition-colors"><span>Visit Website</span> &nearr;</a>` : ''}}
            ${{hasCoords ? `<button onclick="closeDrawer(); locateRecordOnMap(activeDrawerRecord);" class="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-bold transition-colors shadow-2xs"><span>📍 Locate on Map</span></button>` : ''}}
          </div>
        </div>

        <div class="space-y-2">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">About Organization</h3>
          <p class="text-xs sm:text-sm text-slate-700 leading-relaxed bg-slate-50 p-4 rounded-xl border border-slate-200">
            ${{record.desc || 'No detailed organization description available in dataset.'}}
          </p>
        </div>

        <div class="space-y-3">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Products & Offerings</h3>
          ${{productsHtml}}
        </div>

        <div class="space-y-2">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Key Intelligence</h3>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
            
            <div class="bg-slate-50 p-3 rounded-lg border border-slate-200">
              <span class="text-slate-400 block text-[11px] font-semibold uppercase">Company Scale</span>
              <span class="font-bold text-slate-800">${{record.size || 'Unspecified'}} ${{record.employees_num ? `(${{record.employees_num}} employees)` : ''}}</span>
            </div>

            <div class="bg-slate-50 p-3 rounded-lg border border-slate-200">
              <span class="text-slate-400 block text-[11px] font-semibold uppercase">Business Model</span>
              <span class="font-bold text-slate-800">${{record.biz_models && record.biz_models.length ? record.biz_models.join(', ') : 'Unspecified'}}</span>
            </div>

            <div class="bg-slate-50 p-3 rounded-lg border border-slate-200">
              <span class="text-slate-400 block text-[11px] font-semibold uppercase">Target Audience</span>
              <span class="font-bold text-slate-800">${{record.target_audience || 'Unspecified'}}</span>
            </div>

            <div class="bg-slate-50 p-3 rounded-lg border border-slate-200">
              <span class="text-slate-400 block text-[11px] font-semibold uppercase">Product Languages</span>
              <span class="font-bold text-slate-800">${{record.product_lang || 'Unspecified'}}</span>
            </div>

            ${{record.member ? `
            <div class="bg-slate-50 p-3 rounded-lg border border-slate-200 sm:col-span-2">
              <span class="text-slate-400 block text-[11px] font-semibold uppercase">Member Association</span>
              <span class="font-bold text-slate-800">${{record.member}}</span>
            </div>` : ''}}

          </div>
        </div>

        <div class="space-y-2">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Market Segments</h3>
          <div>${{segBadges}}</div>
        </div>

        <div class="space-y-2 pt-2 border-t border-slate-200">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Social Profiles & Presence</h3>
          <div class="flex items-center gap-2 flex-wrap">${{socialsHtml}}</div>
        </div>
      `;

      container.classList.remove('drawer-hidden');
      container.classList.add('drawer-visible');
    }}

    function openDrawerById(id) {{
      const record = RAW_DATA.find(d => d.id === id);
      if (record) {{
        openDrawer(record);
      }}
    }}

    function closeDrawer() {{
      const container = document.getElementById('drawerContainer');
      if (container) {{
        container.classList.remove('drawer-visible');
        container.classList.add('drawer-hidden');
      }}
    }}

    window.addEventListener('keydown', (e) => {{
      if (e.key === 'Escape') closeDrawer();
    }});

    document.getElementById('prevBtn').addEventListener('click', () => {{
      if (currentPage > 1) {{
        currentPage--;
        renderTable();
      }}
    }});

    document.getElementById('nextBtn').addEventListener('click', () => {{
      currentPage++;
      renderTable();
    }});

    function boot() {{
      initMap();
      populateSegmentFilter();
      applyFilters();
    }}

    if (document.readyState === 'loading') {{
      document.addEventListener('DOMContentLoaded', boot);
    }} else {{
      boot();
    }}
  </script>
</body>
</html>"""
    return html_template


def _find_xlsx(description: str, pattern_geocoded: str = "*_geocoded.xlsx", pattern_any: str = "*.xlsx") -> Path:
    """Auto-detect an xlsx file in the current directory, preferring geocoded variants."""
    candidates = sorted(Path(".").glob(pattern_geocoded))
    if candidates:
        return candidates[0]
    candidates = sorted(Path(".").glob(pattern_any))
    if candidates:
        print(f"  Note: no *_geocoded.xlsx found; falling back to {candidates[0].name}")
        return candidates[0]
    print(
        f"Error: could not find a {description} xlsx file in the current directory.\n"
        f"  Pass the path explicitly with --data path/to/file.xlsx\n"
        f"  Or run geocode_cities.py first to produce a geocoded xlsx."
    )
    return None


def main():
    parser = argparse.ArgumentParser(
        description="Generate a standalone HTML dashboard from a geocoded WeMap xlsx dataset.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python generate_dashboard.py                         # auto-detects *_geocoded.xlsx in CWD\n"
            "  python generate_dashboard.py --data my_data.xlsx\n"
            "  python generate_dashboard.py --data my_data.xlsx --title 'My EdTech Map' --out dashboard.html\n"
        ),
    )
    parser.add_argument(
        "--data",
        default=None,
        help=(
            "Path to geocoded xlsx file. "
            "If omitted, auto-detects *_geocoded.xlsx (then any *.xlsx) in the current directory."
        ),
    )
    parser.add_argument("--sheet", default="WeMap_Live_Data",
                        help="Sheet name to read (default: WeMap_Live_Data; falls back to first sheet if not found)")
    parser.add_argument("--out", default="index.html",
                        help="Output HTML file path (default: index.html)")
    parser.add_argument("--title", default="WeMap European EdTech Explorer",
                        help="Dashboard title shown in the header (default: WeMap European EdTech Explorer)")
    args = parser.parse_args()

    # --- Resolve data file ---
    if args.data:
        data_path = Path(args.data)
        if not data_path.exists():
            print(
                f"Error: file not found: {data_path}\n"
                f"  Check the path, or omit --data to auto-detect an xlsx in the current directory."
            )
            return
    else:
        data_path = _find_xlsx("geocoded")
        if data_path is None:
            return
        print(f"Auto-detected dataset: {data_path}")

    print(f"Reading {data_path} (sheet: {args.sheet})...")
    logos_dir = data_path.parent / "logos"
    if logos_dir.is_dir():
        print(f"Logo directory found: {logos_dir} ({len(list(logos_dir.iterdir()))} files)")
    else:
        print("No logos/ directory found next to data file — logos will be omitted (run download_logos.py to add them).")
    records = load_data(data_path, args.sheet, logos_dir=logos_dir)
    print(f"Extracted {len(records)} organization records.")

    print(f"Building standalone dashboard → {args.out} ...")
    html_content = build_html(records, title=args.title)

    out_path = Path(args.out)
    out_path.write_text(html_content, encoding="utf-8")
    print(f"Done. Dashboard written to: {out_path.resolve()}")


if __name__ == "__main__":
    main()
