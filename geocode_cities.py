"""
Part of the WeMap data pipeline.

RUN THIS LOCALLY, NOT IN A SANDBOXED/NETWORK-RESTRICTED ENVIRONMENT.

Geocodes via one of two free, open, no-API-key providers:
  - photon  (default): https://photon.komoot.io -- built on OpenStreetMap
    data, public demo instance, no signup/key/email required at all.
  - nominatim: https://nominatim.openstreetmap.org -- also free/no-key, but
    its usage policy asks for a descriptive User-Agent identifying the app
    (NOT your personal email -- an email is only suggested for heavy use,
    so it's contactable if something goes wrong; a generic app name like
    "WeMapGeocoder/1.0" below is fine for a one-off local run).

Both are shared community services -- be polite (this script rate-limits
itself to ~1 request/second either way) and don't hammer them. If you need
heavier/repeated use, consider self-hosting Nominatim or Photon instead.
Policies: https://operations.osmfoundation.org/policies/nominatim/
          https://photon.komoot.io  (no formal usage policy page, but same
          "don't abuse a free shared service" etiquette applies)

What it does:
  1. Reads the cleaned WeMap xlsx (WeMap_Live_Data sheet).
  2. Geocodes each *unique* city_country value once (not once per row) --
     with ~1,715 rows but far fewer unique city/country combos, this saves
     a lot of time and stays polite to the API.
  3. Caches results to geocode_cache.json as it goes, so if the script is
     interrupted (rate-limited, network drop, etc.) re-running it resumes
     instead of re-querying everything from scratch.
  4. Writes 'lat' and 'lon' columns into the workbook, right after
     city_country, and saves as a new file (does not overwrite your input).

Usage:
    python geocode_cities.py path/to/clean.xlsx
    python geocode_cities.py path/to/clean.xlsx --provider nominatim
    python geocode_cities.py path/to/clean.xlsx --out clean_geocoded.xlsx
"""

import sys
import json
import time
import argparse
from pathlib import Path

import openpyxl
import requests

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
PHOTON_URL = "https://photon.komoot.io/api/"
USER_AGENT = "WeMapGeocoder/1.0"  # generic app name -- no personal info needed
REQUEST_DELAY_SECONDS = 1.1  # be polite to these free shared services
CACHE_PATH = Path("geocode_cache.json")


def load_cache() -> dict:
    if CACHE_PATH.exists():
        with open(CACHE_PATH) as f:
            return json.load(f)
    return {}


def save_cache(cache: dict):
    with open(CACHE_PATH, "w") as f:
        json.dump(cache, f, indent=2, ensure_ascii=False)


def geocode_one_nominatim(query: str) -> tuple:
    """Returns (lat, lon) as floats, or (None, None) if not found/errored."""
    try:
        resp = requests.get(
            NOMINATIM_URL,
            params={"q": query, "format": "json", "limit": 1},
            headers={"User-Agent": USER_AGENT},
            timeout=10,
        )
        resp.raise_for_status()
        results = resp.json()
        if results:
            return float(results[0]["lat"]), float(results[0]["lon"])
        return None, None
    except Exception as e:
        print(f"  ERROR geocoding '{query}': {e}")
        return None, None


def geocode_one_photon(query: str) -> tuple:
    """Returns (lat, lon) as floats, or (None, None) if not found/errored."""
    try:
        resp = requests.get(
            PHOTON_URL,
            params={"q": query, "limit": 1},
            headers={"User-Agent": USER_AGENT},
            timeout=10,
        )
        resp.raise_for_status()
        results = resp.json().get("features", [])
        if results:
            lon, lat = results[0]["geometry"]["coordinates"]  # GeoJSON order is [lon, lat]
            return float(lat), float(lon)
        return None, None
    except Exception as e:
        print(f"  ERROR geocoding '{query}': {e}")
        return None, None


def geocode_all(unique_queries: list, cache: dict, provider: str) -> dict:
    geocode_fn = geocode_one_photon if provider == "photon" else geocode_one_nominatim
    to_fetch = [q for q in unique_queries if q not in cache]
    print(f"{len(unique_queries)} unique city_country values, "
          f"{len(to_fetch)} not yet cached -> fetching now via {provider}.")

    for i, query in enumerate(to_fetch, 1):
        lat, lon = geocode_fn(query)
        cache[query] = {"lat": lat, "lon": lon}
        status = "OK" if lat is not None else "NOT FOUND"
        print(f"  [{i}/{len(to_fetch)}] {query} -> {status}")

        if i % 20 == 0:
            save_cache(cache)

        time.sleep(REQUEST_DELAY_SECONDS)

    save_cache(cache)
    return cache


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx_path")
    parser.add_argument("--out", default=None, help="Output path (default: <input>_geocoded.xlsx)")
    parser.add_argument("--sheet", default="WeMap_Live_Data")
    parser.add_argument("--provider", default="photon", choices=["photon", "nominatim"],
                         help="Geocoding provider (default: photon -- no key/email needed)")
    args = parser.parse_args()

    in_path = Path(args.xlsx_path)
    out_path = Path(args.out) if args.out else in_path.with_name(in_path.stem + "_geocoded.xlsx")

    wb = openpyxl.load_workbook(in_path)
    ws = wb[args.sheet]
    headers = [c.value for c in ws[1]]
    cc_col = headers.index("city_country") + 1

    # Collect unique, non-empty city_country values
    values = set()
    for row in range(2, ws.max_row + 1):
        v = ws.cell(row=row, column=cc_col).value
        if v:
            values.add(str(v))

    cache = load_cache()
    cache = geocode_all(sorted(values), cache, args.provider)

    # Insert lat/lon columns right after city_country
    new_col_lat = cc_col + 1
    new_col_lon = cc_col + 2
    ws.insert_cols(new_col_lat, amount=2)
    ws.cell(row=1, column=new_col_lat, value="lat")
    ws.cell(row=1, column=new_col_lon, value="lon")

    n_filled, n_missing = 0, 0
    for row in range(2, ws.max_row + 1):
        v = ws.cell(row=row, column=cc_col).value
        entry = cache.get(str(v)) if v else None
        if entry and entry.get("lat") is not None:
            ws.cell(row=row, column=new_col_lat, value=entry["lat"])
            ws.cell(row=row, column=new_col_lon, value=entry["lon"])
            n_filled += 1
        else:
            n_missing += 1

    wb.save(out_path)
    print(f"\nDone. lat/lon filled for {n_filled} rows, missing for {n_missing} rows.")
    print(f"Saved to {out_path}")
    print(f"Cache saved to {CACHE_PATH} (re-run anytime to pick up where you left off, "
          f"or to re-run against an updated xlsx without re-geocoding known cities).")


if __name__ == "__main__":
    main()
