"""
image_source_finder.py

Two-step source finding:
  Step 1 — Check manual config file (image_sources_manual.json) for known sources
  Step 2 — Fall back to Google Lens reverse image search for unknown ones

Manual config lives in: src/image_sources_manual.json
"""

import os
import json
import time
import requests
import base64
from pathlib import Path

SERPAPI_ENDPOINT  = "https://serpapi.com/search"
IMGBB_UPLOAD_URL  = "https://api.imgbb.com/1/upload"
MAX_IMAGES        = 30
SEARCH_DELAY      = 1.5

MANUAL_CONFIG     = Path(__file__).parent / "image_sources_manual.json"

SKIP_DOMAINS = [
    "youtube.com", "facebook.com", "twitter.com", "instagram.com",
    "pinterest.com", "reddit.com", "tiktok.com", "linkedin.com",
    "amazon.com", "ebay.com", "x-raw-image",
]

PREFERRED_DOMAINS = [
    "geeksforgeeks.org", "javatpoint.com", "tutorialspoint.com",
    "w3schools.com", "realpython.com", "docs.python.org",
    "programiz.com", "pythonguides.com", "stackoverflow.com",
    "medium.com", "towardsdatascience.com",
]


def _load_manual_sources() -> dict:
    """Load manually specified sources from JSON config."""
    try:
        if MANUAL_CONFIG.exists():
            data = json.loads(MANUAL_CONFIG.read_text(encoding="utf-8"))
            # Remove the instructions key
            return {k: v for k, v in data.items()
                    if not k.startswith("_") and v}
    except Exception:
        pass
    return {}


def _get_mime(image_path: str) -> str:
    ext = Path(image_path).suffix.lower()
    return {
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".webp": "image/webp", ".gif": "image/gif", ".bmp": "image/bmp",
    }.get(ext, "image/png")


def _upload_to_imgbb(image_path: str, api_key: str):
    try:
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        r = requests.post(
            IMGBB_UPLOAD_URL,
            params={"key": api_key},
            data={"image": b64, "expiration": "600"},
            timeout=30,
        )
        data = r.json()
        if data.get("success"):
            return data["data"]["url"]
    except Exception:
        pass
    return None


def _google_lens_search(image_url: str, api_key: str):
    try:
        r = requests.get(
            SERPAPI_ENDPOINT,
            params={"engine": "google_lens", "url": image_url, "api_key": api_key},
            timeout=30,
        )
        data = r.json()
        if data.get("error"):
            return None

        visual_matches = data.get("visual_matches", [])
        if not visual_matches:
            return None

        # Rule 1: preferred domain page link
        for match in visual_matches:
            link = match.get("link", "")
            if link and any(d in link for d in PREFERRED_DOMAINS):
                return link

        # Rule 2: any credible site page link
        for match in visual_matches:
            link = match.get("link", "")
            if (link
                and link.startswith("http")
                and not any(s in link for s in SKIP_DOMAINS)):
                return link

    except Exception:
        pass
    return None


def find_sources_for_images(state) -> dict:
    """
    Find source URLs for images missing source links.
    Checks manual config first, then uses Google Lens reverse search.
    Returns: dict of {filename: "Source: ..."} strings
    """
    # Load manually specified sources
    manual = _load_manual_sources()
    if manual:
        print(f"[SOURCE FINDER] Loaded {len(manual)} manual sources from image_sources_manual.json")

    serpapi_key = os.environ.get("SERPAPI_KEY")
    imgbb_key   = os.environ.get("IMGBB_API_KEY")

    images_to_process = [
        img for img in state.images
        if not img.has_source_link
        and img.output_path
        and Path(img.output_path).exists()
    ][:MAX_IMAGES]

    if not images_to_process:
        return {}

    # Separate into manual vs needs-search
    manual_hits   = [(img, manual[img.filename]) for img in images_to_process if img.filename in manual]
    needs_search  = [img for img in images_to_process if img.filename not in manual]

    print(f"[SOURCE FINDER] {len(manual_hits)} from manual config, {len(needs_search)} need reverse search")

    image_sources = {}
    found_count   = 0
    unknown_count = 0

    # Apply manual sources
    for img, url in manual_hits:
        source_str = f"Source: {url}"
        image_sources[img.filename] = source_str
        found_count += 1
        print(f"  [MANUAL] {img.filename} -> {url}")

    # Reverse image search for the rest
    if needs_search:
        if not serpapi_key or not imgbb_key:
            if not serpapi_key:
                print("[SOURCE FINDER] SERPAPI_KEY not set — set $env:SERPAPI_KEY='key' for auto search")
            if not imgbb_key:
                print("[SOURCE FINDER] IMGBB_API_KEY not set — set $env:IMGBB_API_KEY='key' for auto search")
            for img in needs_search:
                image_sources[img.filename] = "Source: [Add source URL/reference here]"
                unknown_count += 1
        else:
            print(f"[SOURCE FINDER] Reverse searching {len(needs_search)} images via Google Lens...")
            for i, img in enumerate(needs_search, 1):
                print(f"  [{i}/{len(needs_search)}] {img.filename}: ", end="", flush=True)

                public_url = _upload_to_imgbb(img.output_path, imgbb_key)
                if not public_url:
                    image_sources[img.filename] = "Source: [Add source URL/reference here]"
                    unknown_count += 1
                    print("upload failed")
                    continue

                time.sleep(0.5)
                url = _google_lens_search(public_url, serpapi_key)

                if url:
                    image_sources[img.filename] = f"Source: {url}"
                    found_count += 1
                    print(f"found -> {url}")
                else:
                    image_sources[img.filename] = "Source: [Add source URL/reference here]"
                    unknown_count += 1
                    print("not found")

                if i < len(needs_search):
                    time.sleep(SEARCH_DELAY)

    print(f"[SOURCE FINDER] Done — {found_count} found, {unknown_count} not found.\n")
    return image_sources