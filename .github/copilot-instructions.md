# Repository instructions

## Build, test, and lint

- Build the embedded Tailwind stylesheet from the repository root:
  `cd tailwind_build && npm install && npx tailwindcss -i input.css -o compiled_tailwind.css --minify`
- Generate the dashboard with the default geocoded workbook:
  `python generate_dashboard.py`
  Use `--data`, `--sheet`, `--out`, and `--title` to override its input, worksheet, output path, and page title.
- Serve the generated page locally without browser caching:
  `python serve_dashboard.py 8080`
- To geocode a workbook before generating the dashboard:
  `python geocode_cities.py path/to/clean.xlsx --provider photon --out path/to/geocoded.xlsx`
  Geocoding contacts a shared public service and is rate-limited; cached results are stored in `geocode_cache.json`.
- There are no configured automated tests or lint commands. The root `npm test` script is a placeholder that exits with an error; there is no single-test command. For a Python syntax check, run `python -m py_compile generate_dashboard.py geocode_cities.py serve_dashboard.py`.

## Architecture

- `geocode_cities.py` optionally enriches the Excel source with latitude and longitude. It geocodes each unique `city_country` value, caches results, and writes a separate workbook.
- `generate_dashboard.py` reads the `WeMap_Live_Data` worksheet (falling back to the first worksheet if that name is absent), maps available spreadsheet columns into dashboard records, and renders the page. `public_name` is required to include a row; most other fields are optional.
- The generated `index.html` embeds the record data, compiled Tailwind CSS, and `world_countries.geojson`. Its inline JavaScript powers the Leaflet map, charts, filters, searchable table, and organization details. Leaflet, MarkerCluster, Chart.js, and fonts are loaded from external CDNs, so those features require network access in the browser.
- `serve_dashboard.py` is a small static HTTP server that disables caching; it does not generate or transform dashboard content.
- Treat `generate_dashboard.py` as the source for dashboard behavior and `index.html` as its generated output. Update both only when a change specifically requires editing the generated artifact.

## Code conventions

- The HTML/JavaScript page is an f-string in `generate_dashboard.py`. Escape JavaScript braces as `{{` and `}}`, and escape braces in JavaScript template interpolations as `${{...}}`, or Python will interpret them while generating the page.
- Keep spreadsheet header names and the record keys returned by `load_data` aligned with the consumers in the inline JavaScript. Preserve graceful handling of absent optional columns and missing or invalid coordinates.
- Keep Tailwind utility classes discoverable in the configured content sources (`generate_dashboard.py` and `index.html` in `tailwind_build/tailwind.config.js`). Dynamically assembled class names are not automatically detected by Tailwind.
- Dashboard filters share state across the map, charts, and table. Route filter interactions through the existing filter/update flow so each view remains synchronized.
- The geocoder uses Photon by default and supports Nominatim as an alternative. Keep requests polite to these shared services and retain the cache/resume behavior.
