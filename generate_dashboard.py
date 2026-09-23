"""
generate_dashboard.py
Part of the WeMap data pipeline.

Generates a standalone, responsive, aesthetic single-file HTML dashboard
from a geocoded WeMap dataset.

Features:
  - 100% Free, reliable tile layer (ArcGIS Topo / OSM without 403 blocks)
  - Smart outlier-filtered default zoom (focuses on primary geographic cluster)
  - Focus View / Global View toggle buttons
  - Marker clustering with clean count badges
  - Minimalist popups (public_name and city_country only)
  - Statistically sound charts (Top Countries & Company Scale / Employee Distribution)
  - Dynamic Market Segment filter and searchable directory table
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
    idx_website = col_idx("website")

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
            "name": str(name).strip(),
            "cc": cc_val,
            "country": country_val,
            "lat": lat,
            "lon": lon,
            "segments": seg_list,
            "biz_models": biz_list,
            "size": size_val,
            "website": web_val
        })

    return records


def build_html(records: list, title: str = "WeMap European EdTech Explorer") -> str:
    records_json = json.dumps(records, ensure_ascii=False)

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  
  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {{
      theme: {{
        extend: {{
          fontFamily: {{
            sans: ['"Plus Jakarta Sans"', 'sans-serif'],
          }},
          colors: {{
            brand: {{
              50: '#eef2ff',
              100: '#e0e7ff',
              500: '#6366f1',
              600: '#4f46e5',
              700: '#4338ca',
              900: '#312e81',
            }}
          }}
        }}
      }}
    }}
  </script>

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
      height: 540px;
      width: 100%;
      border-radius: 0.75rem;
      z-index: 1;
    }}
    .custom-cluster {{
      background: rgba(79, 70, 229, 0.9);
      border: 2.5px solid #ffffff;
      color: #ffffff;
      font-weight: 700;
      font-size: 13px;
      border-radius: 9999px;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.15), 0 2px 4px -2px rgba(0, 0, 0, 0.1);
    }}
    .leaflet-popup-content-wrapper {{
      border-radius: 0.5rem;
      padding: 2px;
      box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    }}
    .leaflet-popup-content {{
      margin: 10px 14px;
      line-height: 1.4;
    }}
  </style>
</head>
<body class="min-h-screen flex flex-col font-sans antialiased selection:bg-brand-500 selection:text-white">

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
        <p class="text-xs sm:text-sm text-slate-500 mt-0.5">Interactive geospatial map and ecosystem insights</p>
      </div>

      <!-- Market Segments Filter -->
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

    <!-- Geospatial Explorer Card -->
    <div class="bg-white p-4 sm:p-5 rounded-xl border border-slate-200 shadow-sm">
      <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3">
        <div>
          <h2 class="text-base sm:text-lg font-bold text-slate-900">Geospatial Explorer</h2>
          <p class="text-xs text-slate-500">Clustered by location. Click any pin or cluster to explore.</p>
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

      <div id="map" class="shadow-inner border border-slate-100"></div>
    </div>

    <!-- Charts Section -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
      
      <!-- Top Countries -->
      <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col">
        <div class="mb-4">
          <h2 class="text-base font-bold text-slate-900">Top Countries</h2>
          <p class="text-xs text-slate-500">Leading ecosystems in the active selection</p>
        </div>
        <div class="flex-1 min-h-[260px]">
          <canvas id="countryChart"></canvas>
        </div>
      </div>

      <!-- Company Scale / Employee Size Distribution -->
      <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col">
        <div class="mb-4">
          <h2 class="text-base font-bold text-slate-900">Company Scale (Employee Size)</h2>
          <p class="text-xs text-slate-500">Distribution of company maturity and team size</p>
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
          <p class="text-xs text-slate-500">Searchable list of organizations matching current selection</p>
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

  <!-- Footer -->
  <footer class="bg-white border-t border-slate-200 py-4 mt-8">
    <div class="max-w-7xl mx-auto px-4 text-center text-xs text-slate-400">
      WeMap European EdTech Intelligence Platform &bull; Open-source GIS Architecture
    </div>
  </footer>

  <!-- Application Script -->
  <script>
    const RAW_DATA = {records_json};

    let map, markerClusterGroup;
    let countryChartInstance = null;
    let sizeChartInstance = null;

    let filteredData = [...RAW_DATA];
    let currentPage = 1;
    const pageSize = 15;
    let currentBoundsMode = 'focus'; // 'focus' or 'global'

    // Initialize Map with 100% Free & Open High-Availability Tiles
    function initMap() {{
      map = L.map('map', {{
        zoomControl: true,
        scrollWheelZoom: false
      }});

      // High-reliability, open, high-resolution tile service
      L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
        attribution: 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ, TomTom, USGS, FAO, NPS, NRCAN, GeoBase, Kadaster NL, Ordnance Survey',
        maxZoom: 18
      }}).addTo(map);

      // Custom marker cluster group
      markerClusterGroup = L.markerClusterGroup({{
        showCoverageOnHover: false,
        zoomToBoundsOnClick: true,
        spiderfyOnMaxZoom: true,
        iconCreateFunction: function(cluster) {{
          const count = cluster.getChildCount();
          let sizeClass = 'w-8 h-8 text-xs';
          if (count > 50) sizeClass = 'w-10 h-10 text-sm';
          if (count > 200) sizeClass = 'w-12 h-12 text-sm';
          return L.divIcon({{
            html: `<div class="custom-cluster ${{sizeClass}}">${{count}}</div>`,
            className: '',
            iconSize: L.point(40, 40)
          }});
        }}
      }});

      map.addLayer(markerClusterGroup);

      // View Toggle Buttons
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

    // Populate Segment Filter
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
        applyFilters();
      }});

      document.getElementById('searchInput').addEventListener('input', () => {{
        currentPage = 1;
        renderTable();
      }});
    }}

    // Filter Logic
    function applyFilters() {{
      const selectedSegment = document.getElementById('segmentFilter').value;

      if (selectedSegment === 'ALL') {{
        filteredData = [...RAW_DATA];
      }} else {{
        filteredData = RAW_DATA.filter(d => d.segments && d.segments.includes(selectedSegment));
      }}

      currentPage = 1;
      updateMap();
      updateCharts();
      renderTable();
    }}

    // Compute Smart Bounds (Focus on core cluster / percentile vs. Global)
    function fitMapBounds() {{
      const validPoints = filteredData.filter(d => d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon));
      if (validPoints.length === 0) {{
        map.setView([50.0, 10.0], 4);
        return;
      }}

      if (currentBoundsMode === 'global') {{
        const allBounds = validPoints.map(d => [d.lat, d.lon]);
        map.fitBounds(allBounds, {{ padding: [30, 30], maxZoom: 13 }});
        return;
      }}

      // Focus Mode: Exclude isolated outliers (e.g., Australia, Singapore) using 3rd-97th percentile
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

    // Update Map Markers
    function updateMap() {{
      markerClusterGroup.clearLayers();
      let count = 0;

      filteredData.forEach(d => {{
        if (d.lat !== null && d.lon !== null && !isNaN(d.lat) && !isNaN(d.lon)) {{
          count++;

          const marker = L.circleMarker([d.lat, d.lon], {{
            radius: 6,
            fillColor: '#4f46e5',
            color: '#ffffff',
            weight: 2,
            opacity: 1,
            fillOpacity: 0.85
          }});

          // Minimalist popup: public_name & city_country only
          const popupContent = `
            <div class="font-sans">
              <div class="font-bold text-slate-900 text-sm">${{d.name}}</div>
              <div class="text-xs text-slate-500 mt-0.5">${{d.cc || d.country || 'Location unspecified'}}</div>
            </div>
          `;
          marker.bindPopup(popupContent);
          markerClusterGroup.addLayer(marker);
        }}
      }});

      document.getElementById('mapCount').textContent = count.toLocaleString();
      fitMapBounds();
    }}

    // Update Charts (Top Countries & Company Scale)
    function updateCharts() {{
      // 1. Top Countries (Horizontal Bar)
      const countryCounts = {{}};
      filteredData.forEach(d => {{
        const c = d.country || 'Unknown';
        countryCounts[c] = (countryCounts[c] || 0) + 1;
      }});

      const sortedCountries = Object.entries(countryCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);

      const countryLabels = sortedCountries.map(e => e[0]);
      const countryValues = sortedCountries.map(e => e[1]);

      if (countryChartInstance) countryChartInstance.destroy();
      const ctxC = document.getElementById('countryChart').getContext('2d');
      countryChartInstance = new Chart(ctxC, {{
        type: 'bar',
        data: {{
          labels: countryLabels,
          datasets: [{{
            label: 'Organizations',
            data: countryValues,
            backgroundColor: '#6366f1',
            borderRadius: 6,
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
          }}
        }}
      }});

      // 2. Company Scale (Mutually exclusive single-category distribution)
      const sizeCounts = {{}};
      filteredData.forEach(d => {{
        if (d.size) {{
          sizeCounts[d.size] = (sizeCounts[d.size] || 0) + 1;
        }}
      }});

      // Order by company scale
      const preferredOrder = ['Micro', 'Small', 'Medium', 'Large', 'Solo entrepreneur', 'Medium or Large (unspecified)'];
      const sizeLabels = [];
      const sizeValues = [];

      preferredOrder.forEach(key => {{
        if (sizeCounts[key]) {{
          sizeLabels.push(key);
          sizeValues.push(sizeCounts[key]);
        }}
      }});

      // Catch any other keys
      Object.entries(sizeCounts).forEach(([k, v]) => {{
        if (!preferredOrder.includes(k)) {{
          sizeLabels.push(k);
          sizeValues.push(v);
        }}
      }});

      const colors = ['#6366f1', '#3b82f6', '#06b6d4', '#10b981', '#f59e0b', '#94a3b8'];

      if (sizeChartInstance) sizeChartInstance.destroy();
      const ctxS = document.getElementById('sizeChart').getContext('2d');
      sizeChartInstance = new Chart(ctxS, {{
        type: 'doughnut',
        data: {{
          labels: sizeLabels.length ? sizeLabels : ['Unspecified'],
          datasets: [{{
            data: sizeValues.length ? sizeValues : [1],
            backgroundColor: sizeValues.length ? colors.slice(0, sizeLabels.length) : ['#cbd5e1'],
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
          cutout: '65%'
        }}
      }});
    }}

    // Render Table & Pagination
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
        tbody.innerHTML = `<tr><td colspan="4" class="px-5 py-6 text-center text-slate-400">No organizations found</td></tr>`;
      }} else {{
        pageSlice.forEach(d => {{
          const tr = document.createElement('tr');
          tr.className = 'hover:bg-slate-50 transition-colors';
          
          const segBadges = d.segments && d.segments.length > 0 
            ? d.segments.slice(0, 3).map(s => `<span class="inline-block px-2 py-0.5 mr-1 mb-1 text-[11px] font-medium bg-slate-100 text-slate-700 rounded-md border border-slate-200">${{s}}</span>`).join('')
            : '<span class="text-slate-400 text-xs">-</span>';

          tr.innerHTML = `
            <td class="px-5 py-3 font-semibold text-slate-900">${{d.name}}</td>
            <td class="px-5 py-3 text-slate-600">${{d.cc || '-'}}</td>
            <td class="px-5 py-3 text-slate-600">${{d.country || '-'}}</td>
            <td class="px-5 py-3">${{segBadges}}</td>
          `;
          tbody.appendChild(tr);
        }});
      }}

      document.getElementById('pageInfo').textContent = `Showing ${{total === 0 ? 0 : start + 1}} to ${{end}} of ${{total.toLocaleString()}} entries`;
      document.getElementById('prevBtn').disabled = currentPage <= 1;
      document.getElementById('nextBtn').disabled = currentPage >= totalPages;
    }}

    // Pagination Event Listeners
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

    // Boot
    window.addEventListener('DOMContentLoaded', () => {{
      initMap();
      populateSegmentFilter();
      applyFilters();
    }});
  </script>
</body>
</html>"""
    return html_template


def main():
    parser = argparse.ArgumentParser(description="Generate standalone HTML dashboard from geocoded WeMap dataset.")
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

    print("Building dashboard HTML...")
    html_content = build_html(records, title=args.title)

    out_path = Path(args.out)
    out_path.write_text(html_content, encoding="utf-8")
    print(f"Successfully generated dashboard: {out_path.resolve()}")


if __name__ == "__main__":
    main()
