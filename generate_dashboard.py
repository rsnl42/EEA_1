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


def load_data(xlsx_path: Path, sheet_name: str = "WeMap_Live_Data") -> list:
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

    idx_name = col_idx("public_name")
    idx_cc = col_idx("city_country")
    idx_country = col_idx("country")
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

        records.append({
            "id": f"ORG_{len(records)+1}",
            "name": str(name).strip(),
            "cc": cc_val,
            "country": country_val,
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
        <div class="flex items-center gap-3">
          <div class="inline-flex rounded-lg border border-slate-200 bg-slate-50 p-0.5 text-xs font-medium text-slate-600">
            <button id="btnFocusView" class="px-2.5 py-1 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all">🎯 Focus View</button>
            <button id="btnGlobalView" class="px-2.5 py-1 rounded-md text-slate-600 hover:text-slate-900 transition-all">🌐 Global View</button>
          </div>
          <span class="text-xs text-slate-500 whitespace-nowrap"><span id="mapCount" class="font-bold text-slate-800">0</span> mapped</span>
        </div>
      </div>

      <!-- Map Container -->
      <div class="relative">
        <div id="map" class="shadow-inner border border-slate-100"></div>

        <!-- Non-intrusive Map Legend -->
        <div class="absolute bottom-3 right-3 z-10 pointer-events-none">
          <div class="map-legend-box pointer-events-auto text-[11px] text-slate-700 font-sans shadow-sm">
            <div class="font-bold text-slate-900 uppercase text-[10px] tracking-wider mb-1">Company Scale Pins</div>
            <div class="grid grid-cols-2 gap-x-3 gap-y-1">
              <div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full" style="background:#0072B2"></span> Micro</div>
              <div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full" style="background:#E69F00"></span> Small</div>
              <div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full" style="background:#009E73"></span> Medium</div>
              <div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full" style="background:#CC79A7"></span> Large</div>
              <div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full" style="background:#D55E00"></span> Solo</div>
              <div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full" style="background:#56B4E9"></span> Med/Large</div>
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

    let map, markerClusterGroup;
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
    }}

    function updateToggleButtons() {{
      const btnFocus = document.getElementById('btnFocusView');
      const btnGlobal = document.getElementById('btnGlobalView');

      if (currentBoundsMode === 'focus') {{
        btnFocus.className = 'px-2.5 py-1 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all';
        btnGlobal.className = 'px-2.5 py-1 rounded-md text-slate-600 hover:text-slate-900 transition-all';
      }} else {{
        btnGlobal.className = 'px-2.5 py-1 rounded-md bg-white text-slate-900 shadow-sm font-semibold transition-all';
        btnFocus.className = 'px-2.5 py-1 rounded-md text-slate-600 hover:text-slate-900 transition-all';
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

    function fitMapBounds() {{
      const validPoints = filteredData.filter(d => d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon));
      if (validPoints.length === 0) {{
        map.setView([50.0, 10.0], 4);
        return;
      }}

      if (selectedCountry !== 'ALL') {{
        const countryPoints = validPoints.map(d => [d.lat, d.lon]);
        map.fitBounds(countryPoints, {{ padding: [40, 40], maxZoom: 12 }});
        return;
      }}

      if (currentBoundsMode === 'global') {{
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

    function updateMap() {{
      markerClusterGroup.clearLayers();
      let count = 0;

      // Group records with identical coordinates so co-located pins uncluster when zooming
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
            // Golden spiral dispersal: at high zoom, points group into the city bubble;
            // as user zooms in, they separate into individual regional/city pins
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

          const popupContent = `
            <div class="font-sans">
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
      fitMapBounds();
    }}

    function locateRecordOnMap(d) {{
      if (d && d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon)) {{
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
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);

      const countryLabels = sortedCountries.map(e => e[0]);
      const countryValues = sortedCountries.map(e => e[1]);

      const barColors = countryLabels.map(c => (selectedCountry === c) ? OKABE_ITO.vermillion : OKABE_ITO.blue);

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
            borderRadius: 6,
            hoverBackgroundColor: OKABE_ITO.vermillion
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
              const clickedCountry = countryLabels[idx];
              selectedCountry = (selectedCountry === clickedCountry) ? 'ALL' : clickedCountry;
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
            <td class="px-5 py-3 text-slate-600">${{d.cc || '-'}}</td>
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
            <div>
              <h2 class="text-xl font-bold text-slate-900 tracking-tight leading-snug">${{record.name}}</h2>
              <div class="text-xs font-medium text-slate-500 mt-1 flex items-center gap-2">
                <span>📍 ${{record.cc || record.country || 'Location Unspecified'}}</span>
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

        <div class="space-y-3">
          <h3 class="text-xs font-bold uppercase tracking-wider text-slate-400">Products & Offerings</h3>
          ${{productsHtml}}
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


def main():
    parser = argparse.ArgumentParser(description="Generate single standalone HTML dashboard from geocoded WeMap dataset.")
    parser.add_argument("--data", default="202507_dataset_WeMap_LIVE_202607281022_CLEAN_geocoded.xlsx",
                        help="Path to geocoded xlsx file")
    parser.add_argument("--sheet", default="WeMap_Live_Data", help="Sheet name (default: WeMap_Live_Data)")
    parser.add_argument("--out", default="index.html", help="Output HTML file path (default: index.html)")
    parser.add_argument("--title", default="WeMap European EdTech Explorer", help="Dashboard title")
    args = parser.parse_args()

    data_path = Path(args.data)
    if not data_path.exists():
        print(f"Error: file not found: {data_path}")
        return

    print(f"Reading {data_path} (sheet: {args.sheet})...")
    records = load_data(data_path, args.sheet)
    print(f"Extracted {len(records)} organization records.")

    print("Building standalone index.html...")
    html_content = build_html(records, title=args.title)

    out_path = Path(args.out)
    out_path.write_text(html_content, encoding="utf-8")
    print(f"Successfully generated standalone dashboard: {out_path.resolve()}")


if __name__ == "__main__":
    main()
