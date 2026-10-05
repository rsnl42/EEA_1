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


def load_cache(cache_path: Path) -> dict:
    if cache_path.exists():
        with open(cache_path) as f:
            return json.load(f)
    return {}


def save_cache(cache: dict, cache_path: Path):
    with open(cache_path, "w") as f:
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


def geocode_all(unique_queries: list, cache: dict, provider: str, cache_path: Path) -> dict:
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
            save_cache(cache, cache_path)

        time.sleep(REQUEST_DELAY_SECONDS)

    save_cache(cache, cache_path)
    return cache


def main():
    parser = argparse.ArgumentParser(
        description="Geocode city_country values in a WeMap xlsx and write lat/lon columns.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python geocode_cities.py                             # auto-detects a clean xlsx in CWD\n"
            "  python geocode_cities.py my_dataset.xlsx\n"
            "  python geocode_cities.py my_dataset.xlsx --provider nominatim\n"
            "  python geocode_cities.py my_dataset.xlsx --out my_dataset_geocoded.xlsx\n"
        ),
    )
    parser.add_argument(
        "xlsx_path",
        nargs="?",
        default=None,
        help=(
            "Path to the cleaned (non-geocoded) xlsx file. "
            "If omitted, auto-detects a suitable *.xlsx in the current directory "
            "(skips any *_geocoded.xlsx files)."
        ),
    )
    parser.add_argument("--out", default=None,
                        help="Output path (default: <input>_geocoded.xlsx next to the input file)")
    parser.add_argument("--sheet", default="WeMap_Live_Data",
                        help="Sheet name to read (default: WeMap_Live_Data)")
    parser.add_argument("--cache", default=None,
                        help="Path to geocode cache JSON (default: geocode_cache.json next to input file)")
    parser.add_argument("--provider", default="photon", choices=["photon", "nominatim"],
                        help="Geocoding provider (default: photon — no API key needed)")
    args = parser.parse_args()

    # --- Resolve input xlsx ---
    if args.xlsx_path:
        in_path = Path(args.xlsx_path)
        if not in_path.exists():
            print(
                f"Error: file not found: {in_path}\n"
                f"  Check the path, or omit the argument to auto-detect an xlsx in the current directory."
            )
            sys.exit(1)
    else:
        # Prefer non-geocoded files; skip *_geocoded.xlsx
        candidates = [p for p in sorted(Path(".").glob("*.xlsx")) if "_geocoded" not in p.name]
        if not candidates:
            print(
                "Error: no suitable xlsx file found in the current directory.\n"
                "  Pass the path as the first argument: python geocode_cities.py path/to/clean.xlsx"
            )
            sys.exit(1)
        in_path = candidates[0]
        print(f"Auto-detected input: {in_path}")

    out_path = Path(args.out) if args.out else in_path.with_name(in_path.stem + "_geocoded.xlsx")
    cache_path = Path(args.cache) if args.cache else in_path.parent / "geocode_cache.json"

    # --- Load workbook ---
    wb = openpyxl.load_workbook(in_path)
    if args.sheet not in wb.sheetnames:
        print(f"  Warning: sheet '{args.sheet}' not found; using first sheet '{wb.sheetnames[0]}' instead.")
        ws = wb[wb.sheetnames[0]]
    else:
        ws = wb[args.sheet]

    headers = [c.value for c in ws[1]]
    if "city_country" not in headers:
        print(
            f"Error: column 'city_country' not found in {in_path} (sheet: {ws.title}).\n"
            f"  Found columns: {[h for h in headers if h][:15]}"
        )
        sys.exit(1)
    cc_col = headers.index("city_country") + 1

    # Collect unique, non-empty city_country values
    values = set()
    for row in range(2, ws.max_row + 1):
        v = ws.cell(row=row, column=cc_col).value
        if v:
            values.add(str(v))

    cache = load_cache(cache_path)
    cache = geocode_all(sorted(values), cache, args.provider, cache_path)

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
    print(f"Saved to: {out_path}")
    print(f"Cache saved to: {cache_path}  (re-run anytime to resume or update without re-geocoding known cities)")


if __name__ == "__main__":
    main()
