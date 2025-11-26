#!/usr/bin/env python3
"""
Download Microsoft Ignite 2025 slide decks.

This script fetches the 2025 session catalogue and downloads any session
that provides a PowerPoint deck.  Files are saved into an "ignite_2025_slides"
directory.  It requires the 'requests' package.
"""

import json
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Tuple

import requests

API_URL = "https://api-v2.ignite.microsoft.com/api/session/all/en-US"
DEST_DIR = Path("ignite_2025_slides")
MAX_WORKERS = 10  # Number of concurrent downloads

def sanitize_filename(value: str) -> str:
    # Replace spaces with underscores and remove non-safe characters
    value = value.strip().replace(" ", "_")
    return re.sub(r"[^A-Za-z0-9_.-]", "", value)

def fetch_sessions() -> List[Dict]:
    # Retrieve all session data from the Ignite v2 API
    resp = requests.get(API_URL)
    resp.raise_for_status()
    return resp.json()

def download_file(url: str, dest_path: Path, session_code: str, chunk_size: int = 1024 * 1024) -> Tuple[bool, str]:
    # Send headers to mimic a browser; otherwise Medius may return HTTP 403
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://ignite.microsoft.com/",
    }
    try:
        with requests.get(url, headers=headers, stream=True, timeout=30) as r:
            if r.status_code != 200:
                return False, f"HTTP {r.status_code}"
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            with open(dest_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=chunk_size):
                    if chunk:
                        f.write(chunk)
        return True, session_code
    except Exception as e:
        return False, str(e)

def download_session(session: Dict) -> Tuple[str, int]:
    """
    Download a single session's slides.
    Returns: (status, code) where status is 'downloaded', 'no_deck', 'existing', or 'failed'
            and code is 1 for success, 0 for skip
    """
    slide_url = session.get("slideDeck") or ""
    if not slide_url:
        return "no_deck", 0

    session_code = sanitize_filename(session.get("sessionCode", ""))
    title = sanitize_filename(session.get("title", "untitled"))
    dest = DEST_DIR / f"{session_code}_{title}.pptx"

    if dest.exists() and dest.stat().st_size > 0:
        return "existing", 0

    print(f"Downloading {session_code} – {session.get('title')}")
    success, result = download_file(slide_url, dest, session_code)

    if success:
        return "downloaded", 1
    else:
        print(f"Failed to download {session_code}: {result}")
        if dest.exists():
            dest.unlink()
        return "failed", 0


def main() -> None:
    print("Fetching session catalogue...")
    sessions = fetch_sessions()
    print(f"Found {len(sessions)} sessions total\n")

    # Thread-safe counters
    lock = threading.Lock()
    counters = {"downloaded": 0, "no_deck": 0, "existing": 0, "failed": 0}

    # Use ThreadPoolExecutor for concurrent downloads
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Submit all download tasks
        futures = {executor.submit(download_session, session): session for session in sessions}

        # Process results as they complete
        for future in as_completed(futures):
            status, count = future.result()
            with lock:
                counters[status] += 1
                if status == "downloaded":
                    # Print progress
                    total_processed = sum(counters.values())
                    print(f"Progress: {total_processed}/{len(sessions)} sessions processed")

    print("\n" + "="*50)
    print("Summary:")
    print(f"Downloaded: {counters['downloaded']}")
    print(f"Skipped (no deck): {counters['no_deck']}")
    print(f"Skipped (already existed): {counters['existing']}")
    print(f"Failed: {counters['failed']}")
    print("="*50)

if __name__ == "__main__":
    main()
