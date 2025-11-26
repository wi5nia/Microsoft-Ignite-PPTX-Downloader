# Microsoft Ignite 2025 Slide Deck Downloader

Bulk download PowerPoint slide decks from Microsoft Ignite 2025 conference sessions. Available in both Python and PowerShell Core versions.

> **⚠️ Disclaimer:** These scripts were vibe coded. No guarantees that they work perfectly or will continue to work if Microsoft changes their API. Use at your own risk!

## What It Does

Both scripts:
- 🔍 Fetch all sessions from the Microsoft Ignite 2025 API
- 📥 Download PowerPoint slides for sessions that have them available
- ⚡ **Use concurrent downloads (10 parallel downloads by default)** for much faster performance
- 💾 Save files to `ignite_2025_slides/` directory with format: `{sessionCode}_{title}.pptx`
- ⏭️ Skip sessions without slides
- ✅ Skip files that already exist locally
- 📊 Show a summary of downloaded, skipped, and existing files

## Choose Your Version

### 🐍 Python Version (`download_ignite2025_slides.py`)
**Best for:** Users comfortable with Python, Linux/macOS environments

**Prerequisites:**
- Python 3.7+
- `requests` library (see requirements.txt)

**Installation:**
```bash
git clone https://github.com/wi5nia/Microsoft-Ignite-PPTX-Downloader.git
cd Microsoft-Ignite-PPTX-Downloader
python3 -m pip install --user -r requirements.txt
```

**Usage:**
```bash
python3 download_ignite2025_slides.py
```

### 💻 PowerShell Core Version (`Download-Ignite2025Slides.ps1`)
**Best for:** Windows users, cross-platform environments, no external dependencies

**Prerequisites:**
- PowerShell Core 7.0+ (cross-platform PowerShell)

**Installation:**

**macOS:**
```bash
brew install powershell
```

**Windows:**
```powershell
winget install Microsoft.PowerShell
```

**Linux (Ubuntu/Debian):**
```bash
# Install via Snap
sudo snap install powershell --classic

# Or via package manager
wget -q https://packages.microsoft.com/config/ubuntu/20.04/packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y powershell
```

**Usage:**
```powershell
# Make executable (Linux/macOS)
chmod +x Download-Ignite2025Slides.ps1

# Basic usage
./Download-Ignite2025Slides.ps1

# Advanced usage with custom parameters
./Download-Ignite2025Slides.ps1 -MaxConcurrency 5 -DestinationPath "my_slides"

# Get help
Get-Help ./Download-Ignite2025Slides.ps1 -Full
```

## PowerShell Parameters

- **`-MaxConcurrency`** (optional): Number of concurrent downloads (default: 10)
- **`-DestinationPath`** (optional): Directory to save slides (default: "ignite_2025_slides")

## Sample Output

```
Fetching session catalogue...
Found 847 sessions total

Downloading BRK123 – Building Modern Applications with Azure
Downloading INT456 – Building Modern Applications with Containers
Progress: 15/847 sessions processed
Progress: 32/847 sessions processed
...

==================================================
Summary:
Downloaded: 234
Skipped (no deck): 456
Skipped (already existed): 12
Failed: 3
==================================================
```

## Features Comparison

| Feature | Python Version | PowerShell Version |
|---------|----------------|-------------------|
| **Cross-platform** | ✅ | ✅ |
| **External dependencies** | ❌ (requires `requests`) | ✅ (none needed) |
| **Concurrent downloads** | ✅ | ✅ |
| **Progress tracking** | ✅ | ✅ |
| **Resume capability** | ✅ | ✅ |
| **Parameter validation** | Basic | ✅ Advanced |
| **Built-in help system** | ❌ | ✅ |
| **Windows integration** | Basic | ✅ Native |

## Background

Microsoft publishes detailed metadata for Ignite sessions via a public API endpoint:

```text
https://api-v2.ignite.microsoft.com/api/session/all/en-US
```

Each session record can include a `slideDeck` property containing a URL to the corresponding PowerPoint deck hosted on `medius.microsoft.com`. Both scripts:

- Query this API to get all sessions
- Filter for sessions with slides available
- Download the `.pptx` files with proper browser headers (to avoid 403 errors)
- Sanitize filenames by converting spaces to underscores and removing special characters

## Troubleshooting

### Common Issues

**HTTP 403 Errors**: Both scripts include browser-like User-Agent and Referer headers to avoid being blocked by the CDN. If you still get 403 errors, Microsoft may have changed their security policies.

**Download Failures**: Failed downloads are automatically removed (0-byte files won't be left behind). You can simply re-run either script to retry.

**API Changes**: If the API structure changes, the scripts may break. This is expected for vibe-coded utilities!

### Python-Specific Issues

**Import Errors**: Make sure you've installed the requirements: `python3 -m pip install -r requirements.txt`

**SSL Issues on macOS**: The requirements.txt specifies `urllib3<2.0.0` for compatibility with LibreSSL on macOS.

### PowerShell-Specific Issues

**Permission Issues (Linux/macOS)**:
```bash
chmod +x Download-Ignite2025Slides.ps1
```

**Execution Policy (Windows)**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**PowerShell Core Not Found**: Make sure you have PowerShell Core 7.0+ installed, not just Windows PowerShell 5.1.

## License

Use freely. These are simple utility scripts with no warranties or support.
