# SheetHappens (Google Sheets Automated Synchronizer)

I wanted to format google sheets neatly in the web while still quickly viewing their contents inside my code editor. So I built SheetHappens, a tiny, zero-dependency tool that pulls public Google Sheets into local files in real time.

Edit your spreadsheets in your browser. Run SheetHappens. Get fresh local files (`.csv`, `.tsv`, `.xlsx`, `.pdf`) immediately in the `Sheets/` folder without touching the Google Cloud Console or managing OAuth tokens.

---

## Quick Start

### First Run
1. Double-click `sync_now.bat` (or `watch_sync.bat`).
2. If no spreadsheet ID is configured, the script will prompt you:
   ```text
   No Google Spreadsheet ID configured yet.
   Please paste your public Google Sheets link (or ID):
   > https://docs.google.com/spreadsheets/d/1zWOJfO08DNXUgOX_R___Dj7SfW5hP3EAHeWCPT9G0Nc/...
   ```
3. SheetHappens extracts the ID, saves it into `sync_config.json`, and downloads all your tabs into `Sheets/`.

---

## Usage

- **`sync_now.bat`**: Runs a single parallel sync and exits.
- **`watch_sync.bat`**: Keeps running in the background. It checks for edits every 10 seconds and downloads updates automatically. It also periodically scans for newly created or renamed tabs.

You can also run it directly with Python:
```bash
# Single sync
python sync_sheets.py

# Live watcher mode
python sync_sheets.py --watch
```

---

## Configuration (`sync_config.json`)

```json
{
  "spreadsheet_id": "1zWOJfO08DNXUgOX_R___Dj7SfW5hP3EAHeWCPT9G0Nc",
  "format": "csv",
  "watch_interval_seconds": 10
}
```

### Options

| Key | Default | Description |
| :--- | :--- | :--- |
| `spreadsheet_id` | `""` | The Google Sheet ID or full share link. |
| `format` | `"csv"` | Output format: `csv`, `tsv`, `xlsx`, `pdf`, `ods`, or `html`. |
| `watch_interval_seconds` | `10` | How often the watcher checks for cell changes (in seconds). |

#### Formats
- **`csv` / `tsv`**: Downloads every tab as an individual file (e.g. `Sheets/Weapons.csv`, `Sheets/Stats.csv`).
- **`xlsx` / `pdf` / `ods` / `html`**: Downloads the entire workbook as a single file (`Sheets/spreadsheet.xlsx`).

---

## How It Works

1. **Auto Tab Discovery**: Inspects the workbook structure in memory to detect all tab names without downloading bloated files to disk.
2. **Parallel Fetching**: Downloads all tabs concurrently using a thread pool.
3. **MD5 Change Detection**: Computes a checksum hash of new data in RAM and compares it against your local files. If nothing changed, **zero disk writes happen**.
4. **Pure Standard Library**: Built strictly on Python's built-in modules (`urllib`, `zipfile`, `xml.etree`, `pathlib`, `hashlib`). No `pip install` required.
