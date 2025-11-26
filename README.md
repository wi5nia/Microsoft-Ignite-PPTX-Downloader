# Microsoft Ignite 2025 Slide Deck Downloader

A simple Python script to bulk download PowerPoint slide decks from Microsoft Ignite 2025 conference sessions.

> **⚠️ Disclaimer:** This script was vibe coded. No guarantees that it works perfectly or will continue to work if Microsoft changes their API. Use at your own risk!

## What It Does

This script:
- 🔍 Fetches all sessions from the Microsoft Ignite 2025 API
- 📥 Downloads PowerPoint slides for sessions that have them available
- ⚡ **Uses concurrent downloads (10 parallel downloads)** for much faster performance
- 💾 Saves files to `ignite_2025_slides/` directory with format: `{sessionCode}_{title}.pptx`
- ⏭️ Skips sessions without slides
- ✅ Skips files that already exist locally
- 📊 Shows a summary of downloaded, skipped, and existing files

## How to Run

### Prerequisites

- **Python 3.7+**

### Installation

1. Clone this repository:
   ```bash
   git clone <your-repo-url>
   cd Microsoft-Ignite-PPTX-Downloader
   ```

2. Install dependencies:
   ```bash
   python3 -m pip install --user -r requirements.txt
   ```

   This installs:
   - `requests` - for HTTP requests
   - `urllib3<2` - version 1.x for compatibility with LibreSSL on macOS

### Running the Script

Simply run:
```bash
python3 download_ignite2025_slides.py
```

The script will:
1. Connect to the Microsoft Ignite API
2. Fetch all session metadata
3. Download slides for each session that has them
4. Print progress as it downloads
5. Show a summary when complete

Example output:
```
Fetching session catalogue...
Found 334 sessions total

Downloading BRK123 – Building Modern Applications with Azure
Downloading BRK456 – AI and Machine Learning Best Practices
Downloading LAB789 – Hands-on with Azure Functions
Progress: 50/334 sessions processed
Progress: 100/334 sessions processed
...

==================================================
Summary:
Downloaded: 87
Skipped (no deck): 245
Skipped (already existed): 12
Failed: 0
==================================================
```

## Background

Microsoft publishes detailed metadata for Ignite sessions via a public API endpoint:
```
https://api-v2.ignite.microsoft.com/api/session/all/en-US
```

Each session record can include a `slideDeck` property containing a URL to the corresponding PowerPoint deck hosted on `medius.microsoft.com`. The script:
- Queries this API to get all sessions
- Filters for sessions with slides available
- Downloads the `.pptx` files with proper browser headers (to avoid 403 errors)
- Sanitizes filenames by converting spaces to underscores and removing special characters

## Troubleshooting

**HTTP 403 Errors**: The script includes browser-like User-Agent and Referer headers to avoid being blocked by the CDN. If you still get 403 errors, Microsoft may have changed their security policies.

**Download Failures**: Failed downloads are automatically removed (0-byte files won't be left behind). You can simply re-run the script to retry.

**API Changes**: If the API structure changes, the script may break. This is expected for a vibe-coded utility!

## License

Use freely. This is a simple utility script with no warranties or support.
