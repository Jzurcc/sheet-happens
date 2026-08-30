# SheetHappens (Google Sheets Automated Synchronizer)

I wanted to format Google Sheets neatly in the web while still quickly viewing/editing their contents inside my code editor so I built SheetHappens. It's a tiny, zero-dependency tool that pulls public Google Sheets into local files in real time.

Edit your spreadsheets in your browser. Run SheetHappens. Get fresh local files (`.csv`, `.tsv`, `.xlsx`, `.pdf`) immediately in the `Sheets/` folder without touching the Google Cloud Console or managing OAuth tokens.

I'll make another version for Google Drive folders soon, supporting all file types.

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

- **`sync_now.bat`**: Pulls the latest spreadsheet data into `Sheets/` and exits.
- **`watch_sync.bat`**: Keeps running in the background. It checks for edits every 10 seconds and downloads updates automatically. It also scans periodically for newly created or renamed tabs.
- **`push_now.bat`**: *(Optional)* Pushes your local CSV files in `Sheets/` back up to Google Sheets.

You can also run commands directly with Python:
```bash
# Single pull
python sync_sheets.py

# Live watcher mode
python sync_sheets.py --watch

# Push local CSVs to Google Sheets
python sync_sheets.py --push
```

---

## Configuration (`sync_config.json`)

```json
{
  "spreadsheet_id": "1zWOJfO08DNXUgOX_R___Dj7SfW5hP3EAHeWCPT9G0Nc",
  "format": "csv",
  "watch_interval_seconds": 10,
  "webhook_url": ""
}
```

### Options

| Key | Default | Description |
| :--- | :--- | :--- |
| `spreadsheet_id` | `""` | The Google Sheet ID or full share link. |
| `format` | `"csv"` | Output format: `csv`, `tsv`, `xlsx`, `pdf`, `ods`, or `html`. |
| `watch_interval_seconds` | `10` | How often the watcher checks for cell changes (in seconds). |
| `webhook_url` | `""` | *(Optional)* Google Apps Script Web App URL for pushing changes. |

#### Formats
- **`csv` / `tsv`**: Downloads every tab as an individual file (e.g. `Sheets/Weapons.csv`, `Sheets/Stats.csv`).
- **`xlsx` / `pdf` / `ods` / `html`**: Downloads the entire workbook as a single file (`Sheets/spreadsheet.xlsx`).

---

## Optional: Two-Way Sync (Pushing Local Edits to Google Sheets)

By default, SheetHappens is read-only (pulls from Google Sheets). If you want to edit local CSV files and push them back up to Google Sheets without overwriting formatting or formulas:

### 1. Add Apps Script in Google Sheets
1. Open your Google Sheet in your browser.
2. Click **Extensions > Apps Script**.
3. Replace any code in the editor with this script:

```javascript
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheetName = data.sheet_name;
    var csvText = data.csv_data;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    // Parse CSV rows
    var newRows = Utilities.parseCsv(csvText);
    if (!newRows || newRows.length === 0 || newRows[0].length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "empty" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // FORMULA SHIELDING: Read existing formulas so calculated summary columns are never destroyed
    var existingFormulas = sheet.getDataRange().getFormulas();
    for (var r = 0; r < newRows.length; r++) {
      for (var c = 0; c < newRows[r].length; c++) {
        // If Google Sheet already has a formula here and CSV does not provide a new one, keep it
        if (existingFormulas[r] && existingFormulas[r][c] && existingFormulas[r][c] !== "") {
          if (!String(newRows[r][c]).startsWith("=")) {
            newRows[r][c] = existingFormulas[r][c];
          }
        }
      }
    }
    
    // Clear only text/values (preserves colors, fonts, borders, widths, and validations)
    sheet.clearContents();
    sheet.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

### 2. Deploy as Web App
1. Click **Deploy > New deployment** (top-right).
2. Click the gear icon ⚙️ on the left $\rightarrow$ select **Web app**.
3. Set:
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. Click **Deploy**, authorize access, and copy the **Web app URL**.

### 3. Add to `sync_config.json`
Paste the URL into `sync_config.json` under `"webhook_url"`:
```json
{
  "spreadsheet_id": "YOUR_SHEET_ID",
  "webhook_url": "https://script.google.com/macros/s/AKfycb.../exec"
}
```

Now, double-clicking `push_now.bat` (or running `python sync_sheets.py --push`) will update your Google Sheet online while preserving all visual formatting and formulas!

---

## How It Works

1. **Auto Tab Discovery**: Inspects the workbook structure in memory to detect all tab names without downloading bloated files to disk.
2. **Parallel Fetching & Pushing**: Transfers all tabs concurrently using a thread pool.
3. **MD5 Change Detection**: Computes a checksum hash of new data in RAM and compares it against your local files. If nothing changed, **zero disk writes happen**.
4. **Formula Shielding**: Apps Script protects calculated formulas and visual styling (colors, fonts, borders) during push operations.
5. **Pure Standard Library**: Built strictly on Python's built-in modules (`urllib`, `zipfile`, `xml.etree`, `pathlib`, `hashlib`). No `pip install` required.
