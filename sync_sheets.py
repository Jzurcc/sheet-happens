import csv
import hashlib
import io
import json
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# config & constants
BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "sync_config.json"
SHEETS_DIR = BASE_DIR / "Sheets"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
SUPPORTED_FORMATS = {"csv", "tsv", "xlsx", "pdf", "ods", "html"}
DEFAULT_FORMAT = "csv"


# helper functions
def extract_spreadsheet_id(value: str) -> str:
    """Extracts raw spreadsheet ID from a string or full Google Sheets URL."""
    if not value:
        return ""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", value)
    return match.group(1) if match else value.strip()


def sanitize_filename(name: str) -> str:
    """Sanitizes invalid filesystem characters while keeping tab names readable."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def col2num(col_str: str) -> int:
    """Converts Excel column string (e.g. 'A', 'Z', 'AA') to 1-based integer."""
    num = 0
    for char in col_str:
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num


def fetch_full_workbook(spreadsheet_id: str, fmt: str = "xlsx") -> bytes:
    """Fetches the entire workbook in binary export format (xlsx, pdf, ods, html)."""
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.read()


def parse_workbook_to_tabs(xlsx_bytes: bytes, fmt: str = "csv") -> dict[str, bytes]:
    """Parses Google Sheets XLSX export into pure CSV/TSV per tab in memory without losing mixed-type cells."""
    delimiter = "\t" if fmt == "tsv" else ","
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
        # 1. Parse shared strings
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sst.iter(f"{ns}si"):
                text_parts = [t.text for t in si.iter(f"{ns}t") if t.text]
                strings.append("".join(text_parts))

        # 2. Parse workbook relationship mapping
        sheets_info = []
        wb_tree = ET.fromstring(z.read("xl/workbook.xml"))
        for sheet in wb_tree.iter(f"{ns}sheet"):
            sheet_name = sheet.attrib.get("name")
            r_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if sheet_name and r_id:
                sheets_info.append((sheet_name, r_id))

        rel_map = {}
        if "xl/_rels/workbook.xml.rels" in z.namelist():
            rels_tree = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            for rel in rels_tree.iter(f"{rel_ns}Relationship"):
                rel_map[rel.attrib.get("Id")] = rel.attrib.get("Target")

        tabs_data = {}
        for sheet_name, r_id in sheets_info:
            target_rel = rel_map.get(r_id, "")
            target_path = "xl/" + target_rel.lstrip("/")
            if target_path not in z.namelist():
                continue

            ws = ET.fromstring(z.read(target_path))
            rows_data = {}
            max_c = 0
            max_r = 0

            for c in ws.iter(f"{ns}c"):
                cell_ref = c.attrib.get("r", "")
                if not cell_ref:
                    continue
                m = re.match(r"([A-Z]+)(\d+)", cell_ref)
                if not m:
                    continue
                col_idx = col2num(m.group(1)) - 1
                row_idx = int(m.group(2)) - 1
                max_c = max(max_c, col_idx)
                max_r = max(max_r, row_idx)

                t_type = c.attrib.get("t")
                v_elem = c.find(f"{ns}v")
                val = ""
                if v_elem is not None and v_elem.text is not None:
                    if t_type == "s":
                        s_idx = int(v_elem.text)
                        val = strings[s_idx] if s_idx < len(strings) else ""
                    else:
                        val = v_elem.text
                        # Clean trailing .0 from pure integers
                        if val.endswith(".0"):
                            val = val[:-2]

                rows_data.setdefault(row_idx, {})[col_idx] = val

            out = io.StringIO()
            writer = csv.writer(out, delimiter=delimiter, lineterminator="\n")
            for r in range(max_r + 1):
                row_cells = [rows_data.get(r, {}).get(c, "") for c in range(max_c + 1)]
                writer.writerow(row_cells)

            tabs_data[sheet_name] = out.getvalue().encode("utf-8")

        return tabs_data


# core sync engine
def sync_file_content(target_path: Path, new_data: bytes, display_name: str, verbose: bool = True) -> bool:
    """Writes binary data to disk if the MD5 hash differs. Returns True if updated."""
    timestamp = time.strftime("%H:%M:%S")
    current_hash = hashlib.md5(target_path.read_bytes()).hexdigest() if target_path.exists() else ""
    new_hash = hashlib.md5(new_data).hexdigest()

    if current_hash != new_hash:
        target_path.write_bytes(new_data)
        print(f"[{timestamp}] [UPDATED] {display_name} -> Sheets/{target_path.name}")
        return True
    elif verbose:
        print(f"[{timestamp}] [UP-TO-DATE] {display_name} -> Sheets/{target_path.name}")

    return False


def sync_all(config: dict, verbose: bool = True) -> tuple[int, list[str]]:
    """Syncs sheets based on the configured format (per-tab or full workbook)."""
    fmt = config.get("format", DEFAULT_FORMAT).lower().strip().lstrip(".")
    if fmt not in SUPPORTED_FORMATS:
        print(f"[WARNING] Unknown format '{fmt}'. Falling back to '{DEFAULT_FORMAT}'.")
        fmt = DEFAULT_FORMAT

    spreadsheet_id = extract_spreadsheet_id(config.get("spreadsheet_id", ""))
    if not spreadsheet_id:
        print("[ERROR] No spreadsheet ID configured.")
        return 0, []

    SHEETS_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Full workbook exports (xlsx, pdf, ods, html)
    if fmt in {"xlsx", "pdf", "ods", "html"}:
        target_filename = f"spreadsheet.{fmt}"
        try:
            new_data = fetch_full_workbook(spreadsheet_id, fmt)
            is_updated = sync_file_content(SHEETS_DIR / target_filename, new_data, f"Full Workbook ({fmt.upper()})", verbose)
            return (1 if is_updated else 0), ["workbook"]
        except Exception as e:
            print(f"[ERROR] Failed to export workbook as {fmt}: {e}")
            return 0, []

    # 2. Per-tab exports (csv, tsv) - Fast single atomic workbook download + parse
    try:
        xlsx_bytes = fetch_full_workbook(spreadsheet_id, "xlsx")
        tabs = parse_workbook_to_tabs(xlsx_bytes, fmt)
    except Exception as e:
        print(f"[ERROR] Could not fetch sheets: {e}")
        return 0, []

    updated_count = 0
    for sheet_name, data in tabs.items():
        filename = f"{sanitize_filename(sheet_name)}.{fmt}"
        if sync_file_content(SHEETS_DIR / filename, data, sheet_name, verbose):
            updated_count += 1

    return updated_count, list(tabs.keys())


# push & two-way sync engine
def push_single_sheet(webhook_url: str, file_path: Path) -> bool:
    """Pushes a single local CSV file to the Google Sheets Apps Script webhook."""
    timestamp = time.strftime("%H:%M:%S")
    sheet_name = file_path.stem
    payload = {
        "sheet_name": sheet_name,
        "csv_data": file_path.read_text(encoding="utf-8")
    }

    try:
        if HAS_REQUESTS:
            resp = requests.post(webhook_url, json=payload, timeout=30)
            res = resp.json()
        else:
            data_bytes = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(webhook_url, data=data_bytes, headers={"Content-Type": "application/json", "User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=30) as response:
                res = json.loads(response.read().decode("utf-8"))

        if res.get("status") == "success":
            print(f"[{timestamp}] [PUSHED] Sheets/{file_path.name} -> Google Sheets ('{sheet_name}')")
            return True
        else:
            print(f"[{timestamp}] [ERROR] Google Sheets error for '{sheet_name}': {res.get('message')}")
            return False
    except Exception as e:
        print(f"[{timestamp}] [ERROR] Failed to push '{file_path.name}': {e}")
        return False


def push_all(config: dict) -> int:
    """Pushes all local CSV files in Sheets/ to Google Sheets via Webhook."""
    webhook_url = config.get("webhook_url", "").strip()

    if not webhook_url:
        print("=" * 60)
        print("  SheetHappens - Push to Google Sheets")
        print("=" * 60)
        print("No Webhook URL configured in sync_config.json yet.")
        print("\nTo enable pushing local changes back to Google Sheets:")
        print("1. Follow the 'Two-Way Sync' guide in README.md to deploy the Apps Script.")
        print("2. Paste the Web App URL below (or enter it into sync_config.json).")
        print("=" * 60)
        try:
            user_input = input("\nEnter Web App URL (or press Enter to cancel):\n> ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nCancelled.")
            return 0

        if not user_input or not user_input.startswith("http"):
            print("No URL provided. Push cancelled.")
            return 0

        webhook_url = user_input
        config["webhook_url"] = webhook_url
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2)
            print(f"[SUCCESS] Saved webhook_url to {CONFIG_FILE.name}\n")
        except Exception as e:
            print(f"[WARNING] Could not save config: {e}")

    csv_files = sorted(list(SHEETS_DIR.glob("*.csv")) + list(SHEETS_DIR.glob("*.tsv")))
    if not csv_files:
        print(f"[ERROR] No CSV or TSV files found in {SHEETS_DIR}")
        return 0

    print(f"Pushing {len(csv_files)} local file(s) to Google Sheets...")
    with ThreadPoolExecutor(max_workers=min(len(csv_files), 4)) as executor:
        futures = [executor.submit(push_single_sheet, webhook_url, f) for f in csv_files]
        pushed_count = sum(1 for f in futures if f.result())

    print(f"\nDone! {pushed_count} sheet(s) updated online.")
    return pushed_count


def two_way_sync_cycle(config: dict, known_hashes: dict[str, str]) -> dict[str, str]:
    """Performs a smart two-way check: pushes local edits, pulls cloud edits."""
    fmt = config.get("format", DEFAULT_FORMAT).lower().strip().lstrip(".")
    spreadsheet_id = extract_spreadsheet_id(config.get("spreadsheet_id", ""))
    webhook_url = config.get("webhook_url", "").strip()

    if not spreadsheet_id:
        return known_hashes

    SHEETS_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Check local files for user edits first
    for f in list(SHEETS_DIR.glob(f"*.{fmt}")):
        sheet_name = f.stem
        last_hash = known_hashes.get(sheet_name)
        local_hash = hashlib.md5(f.read_bytes()).hexdigest()

        if last_hash is not None and local_hash != last_hash:
            if webhook_url:
                if push_single_sheet(webhook_url, f):
                    known_hashes[sheet_name] = local_hash

    # 2. Fetch fresh cloud state in 1 fast atomic request
    try:
        xlsx_bytes = fetch_full_workbook(spreadsheet_id, "xlsx")
        tabs = parse_workbook_to_tabs(xlsx_bytes, fmt)

        for sheet_name, data in tabs.items():
            filename = f"{sanitize_filename(sheet_name)}.{fmt}"
            target_path = SHEETS_DIR / filename
            remote_hash = hashlib.md5(data).hexdigest()
            local_hash = hashlib.md5(target_path.read_bytes()).hexdigest() if target_path.exists() else ""

            if local_hash != remote_hash:
                target_path.write_bytes(data)
                timestamp = time.strftime("%H:%M:%S")
                print(f"[{timestamp}] [PULLED] {sheet_name} -> Sheets/{target_path.name}")

            known_hashes[sheet_name] = remote_hash
    except Exception:
        pass

    return known_hashes


# CLI & watcher lifecycle
def load_or_init_config() -> dict:
    """Loads existing config or interactively prompts user for their sheet link/ID."""
    config = {}
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            print(f"[WARNING] Could not parse {CONFIG_FILE.name}: {e}")

    spreadsheet_id = extract_spreadsheet_id(config.get("spreadsheet_id", ""))
    if not spreadsheet_id:
        print("=" * 60)
        print("  SheetHappens - Initial Setup")
        print("=" * 60)
        print("No Google Spreadsheet ID configured yet.")
        print("Please paste your public Google Sheets link (or ID):")
        try:
            user_input = input("> ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nSetup cancelled.")
            sys.exit(0)

        spreadsheet_id = extract_spreadsheet_id(user_input)
        if not spreadsheet_id:
            print("[ERROR] No valid Spreadsheet ID provided.")
            sys.exit(1)

        config["spreadsheet_id"] = spreadsheet_id
        config.setdefault("format", "csv")
        config.setdefault("watch_interval_seconds", 10)

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)

        print(f"\n[SUCCESS] Stored ID: {spreadsheet_id}")
        print(f"[INFO] Saved to {CONFIG_FILE.name} (changeable anytime)")
        print("=" * 60 + "\n")

    return config


def main():
    config = load_or_init_config()
    is_push_mode = "--push" in sys.argv or "-p" in sys.argv
    is_watch_mode = "--watch" in sys.argv or "-w" in sys.argv
    interval = config.get("watch_interval_seconds", 10)

    if is_push_mode:
        push_all(config)
    elif is_watch_mode:
        has_webhook = bool(config.get("webhook_url", "").strip())
        mode_label = "Two-Way Live Sync" if has_webhook else "Live Watcher (Pull-Only)"
        print("=" * 60)
        print(f"  SheetHappens - {mode_label}")
        print(f"  Watching every {interval}s... (Press Ctrl+C to stop)")
        print("=" * 60)

        _, _ = sync_all(config, verbose=True)
        known_hashes = {}
        for f in SHEETS_DIR.glob("*.*"):
            known_hashes[f.stem] = hashlib.md5(f.read_bytes()).hexdigest()

        try:
            while True:
                time.sleep(interval)
                known_hashes = two_way_sync_cycle(config, known_hashes)
        except KeyboardInterrupt:
            print("\nWatcher stopped.")
    else:
        print("Syncing all Google Sheets tabs to CSV files...")
        count, _ = sync_all(config, verbose=True)
        print(f"\nDone! {count} file(s) updated.")


if __name__ == "__main__":
    main()
