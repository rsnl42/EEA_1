#!/usr/bin/env python3
"""
optimize_logos.py
Part of the WeMap data pipeline.

Optimizes, compresses, and resizes company logo images in the logos/ directory
so that they load fast, consume minimal bandwidth, and can be tracked in version
control (e.g., GitHub Pages) without bloating the repository.

Usage:
    python optimize_logos.py
    python optimize_logos.py --logos-dir logos --max-size 256 --quality 85 --workers 8
"""

import os
import sys
import argparse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from PIL import Image, ImageOps
    Image.MAX_IMAGE_PIXELS = None  # Allow processing unusually large source scans safely
except ImportError:
    print("Error: Pillow is required. Install it with: pip install Pillow")
    sys.exit(1)


def optimize_single_image(img_path: Path, max_size: int = 256, quality: int = 85) -> tuple[int, int, str]:
    """
    Optimizes a single image in-place.
    Returns (orig_size, new_size, status_msg).
    """
    orig_size = img_path.stat().st_size
    suffix = img_path.suffix.lower()

    # Skip vector SVGs
    if suffix == ".svg":
        return orig_size, orig_size, "Skipped (SVG)"

    try:
        with Image.open(img_path) as img:
            # Handle EXIF orientation if present
            try:
                img = ImageOps.exif_transpose(img)
            except Exception:
                pass

            orig_w, orig_h = img.size
            needs_resize = max(orig_w, orig_h) > max_size

            # If the image is already small in dimensions and bytes, keep it
            if not needs_resize and orig_size <= 40 * 1024:
                return orig_size, orig_size, "Already optimal"

            if needs_resize:
                img.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)

            temp_path = img_path.with_suffix(f".tmp{suffix}")

            if suffix in (".jpg", ".jpeg"):
                # Convert CMYK/RGBA to RGB for standard JPEG
                if img.mode != "RGB":
                    img = img.convert("RGB")
                img.save(temp_path, format="JPEG", quality=quality, optimize=True)
            elif suffix == ".webp":
                img.save(temp_path, format="WEBP", quality=quality, method=6)
            elif suffix == ".png":
                # Preserve transparency if RGBA, or convert to P/RGBA
                img.save(temp_path, format="PNG", optimize=True)
            else:
                img.save(temp_path, optimize=True)

            new_size = temp_path.stat().st_size

            # Only overwrite if the new file is smaller or we downscaled dimensions
            if new_size < orig_size or needs_resize:
                temp_path.replace(img_path)
                return orig_size, new_size, f"Optimized ({orig_w}x{orig_h} -> {img.size[0]}x{img.size[1]})"
            else:
                temp_path.unlink(missing_ok=True)
                return orig_size, orig_size, "Kept original"

    except Exception as e:
        # If Pillow cannot open (e.g. PDF masquerading as PNG), leave intact
        return orig_size, orig_size, f"Error: {e}"


def optimize_directory(logos_dir: Path, max_size: int = 256, quality: int = 85, workers: int = 8):
    if not logos_dir.is_dir():
        print(f"Error: Directory '{logos_dir}' not found.")
        return

    image_files = [
        f for f in logos_dir.iterdir()
        if f.is_file() and not f.name.startswith(".") and not f.name.endswith(".tmp")
    ]

    total_files = len(image_files)
    if total_files == 0:
        print(f"No image files found in '{logos_dir}'.")
        return

    orig_total_bytes = sum(f.stat().st_size for f in image_files)
    print(f"Starting optimization of {total_files} logos in '{logos_dir}'...")
    print(f"Initial total size: {orig_total_bytes / (1024 * 1024):.2f} MB")
    print(f"Parameters: max_dimension={max_size}px, jpeg_quality={quality}, workers={workers}")

    completed_count = 0
    new_total_bytes = 0

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(optimize_single_image, img_path, max_size, quality): img_path
            for img_path in image_files
        }

        for future in as_completed(futures):
            orig_s, new_s, status = future.result()
            new_total_bytes += new_s
            completed_count += 1

            if completed_count % 200 == 0 or completed_count == total_files:
                pct = (completed_count / total_files) * 100
                print(f"[{completed_count:4d}/{total_files}] ({pct:5.1f}%) processed...")

    saved_mb = (orig_total_bytes - new_total_bytes) / (1024 * 1024)
    ratio = (1 - (new_total_bytes / orig_total_bytes)) * 100 if orig_total_bytes else 0

    print("\n" + "=" * 50)
    print("Optimization Complete!")
    print(f"Original size:  {orig_total_bytes / (1024 * 1024):.2f} MB")
    print(f"Optimized size: {new_total_bytes / (1024 * 1024):.2f} MB")
    print(f"Saved:          {saved_mb:.2f} MB ({ratio:.1f}% reduction)")
    print("=" * 50 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Batch compress and resize logo images for web delivery and git tracking."
    )
    parser.add_argument(
        "--logos-dir",
        default="logos",
        help="Path to directory containing logo images (default: logos/)",
    )
    parser.add_argument(
        "--max-size",
        type=int,
        default=256,
        help="Max width or height in pixels (default: 256)",
    )
    parser.add_argument(
        "--quality",
        type=int,
        default=85,
        help="JPEG/WEBP quality (1-100, default: 85)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=8,
        help="Number of concurrent worker threads (default: 8)",
    )

    args = parser.parse_args()
    logos_path = Path(args.logos_dir)
    optimize_directory(logos_path, max_size=args.max_size, quality=args.quality, workers=args.workers)


if __name__ == "__main__":
    main()
