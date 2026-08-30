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


def get_all_sheet_names(spreadsheet_id: str) -> list[str]:
    """Discovers all sheet tab names from Google's workbook XML in memory."""
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

    with urllib.request.urlopen(req, timeout=15) as resp:
        content = resp.read()

    with zipfile.ZipFile(io.BytesIO(content)) as z:
        tree = ET.fromstring(z.read("xl/workbook.xml"))
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        return [elem.attrib["name"] for elem in tree.iter(f"{ns}sheet") if "name" in elem.attrib]


def fetch_sheet_tab(spreadsheet_id: str, sheet_name: str, fmt: str = "csv") -> bytes:
    """Fetches formatted data (csv/tsv) for a single sheet tab."""
    encoded_name = urllib.parse.quote(sheet_name)
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:{fmt}&sheet={encoded_name}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.read()


def fetch_full_workbook(spreadsheet_id: str, fmt: str = "xlsx") -> bytes:
    """Fetches the entire workbook in binary/export format (xlsx, pdf, ods, html)."""
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.read()


# core engine
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


def sync_single_tab(spreadsheet_id: str, sheet_name: str, target_file: str, fmt: str, verbose: bool = True) -> bool:
    """Syncs a single tab to disk if changed."""
    timestamp = time.strftime("%H:%M:%S")
    try:
        new_data = fetch_sheet_tab(spreadsheet_id, sheet_name, fmt)
        return sync_file_content(SHEETS_DIR / target_file, new_data, sheet_name, verbose)
    except Exception as e:
        print(f"[{timestamp}] [ERROR] Failed to sync tab '{sheet_name}': {e}")
        return False


def resolve_sheet_targets(config: dict, cached_names: list[str] = None, fmt: str = "csv") -> tuple[str, list[tuple[str, str]], list[str]]:
    """Resolves the spreadsheet ID and list of (sheet_name, target_filename) pairs."""
    spreadsheet_id = extract_spreadsheet_id(config.get("spreadsheet_id", ""))
    if not spreadsheet_id:
        return "", [], []

    if config.get("mappings"):
        targets = [(m["sheet_name"], m["target_file"]) for m in config["mappings"]]
        names = [t[0] for t in targets]
        return spreadsheet_id, targets, names

    names = cached_names or get_all_sheet_names(spreadsheet_id)
    targets = [(name, f"{sanitize_filename(name)}.{fmt}") for name in names]
    return spreadsheet_id, targets, names


def sync_all(config: dict, cached_names: list[str] = None, verbose: bool = True) -> tuple[int, list[str]]:
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

    if fmt in {"xlsx", "pdf", "ods", "html"}:
        target_filename = f"spreadsheet.{fmt}"
        try:
            new_data = fetch_full_workbook(spreadsheet_id, fmt)
            is_updated = sync_file_content(SHEETS_DIR / target_filename, new_data, f"Full Workbook ({fmt.upper()})", verbose)
            return (1 if is_updated else 0), ["workbook"]
        except Exception as e:
            print(f"[ERROR] Failed to export workbook as {fmt}: {e}")
            return 0, []

    try:
        spreadsheet_id, targets, sheet_names = resolve_sheet_targets(config, cached_names, fmt)
    except Exception as e:
        print(f"[ERROR] Could not discover sheets: {e}")
        return 0, []

    if not targets:
        print("[ERROR] No sheet tabs found.")
        return 0, []

    with ThreadPoolExecutor(max_workers=min(len(targets), 8)) as executor:
        futures = [
            executor.submit(sync_single_tab, spreadsheet_id, sheet_name, filename, fmt, verbose)
            for sheet_name, filename in targets
        ]
        updated_count = sum(1 for f in futures if f.result())

    return updated_count, sheet_names


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
            print(f"[{timestamp}] [ERROR] Google Sheets returned error for '{sheet_name}': {res.get('message')}")
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


def two_way_sync_cycle(config: dict, sheet_names: list[str], known_hashes: dict[str, str]) -> tuple[list[str], dict[str, str]]:
    """Performs a smart two-way check: pushes local edits, pulls cloud edits."""
    fmt = config.get("format", DEFAULT_FORMAT).lower().strip().lstrip(".")
    spreadsheet_id = extract_spreadsheet_id(config.get("spreadsheet_id", ""))
    webhook_url = config.get("webhook_url", "").strip()

    if not spreadsheet_id:
        return sheet_names, known_hashes

    try:
        _, targets, sheet_names = resolve_sheet_targets(config, sheet_names, fmt)
    except Exception:
        return sheet_names, known_hashes

    SHEETS_DIR.mkdir(parents=True, exist_ok=True)

    for sheet_name, target_name in targets:
        target_path = SHEETS_DIR / target_name
        last_hash = known_hashes.get(sheet_name)
        local_hash = hashlib.md5(target_path.read_bytes()).hexdigest() if target_path.exists() else None

        # 1. Check if user edited the local file (local hash changed since last sync)
        if last_hash is not None and local_hash is not None and local_hash != last_hash:
            if webhook_url:
                if push_single_sheet(webhook_url, target_path):
                    known_hashes[sheet_name] = local_hash
                continue

        # 2. Check if Google Sheets changed online
        try:
            new_data = fetch_sheet_tab(spreadsheet_id, sheet_name, fmt)
            remote_hash = hashlib.md5(new_data).hexdigest()

            if local_hash != remote_hash:
                target_path.write_bytes(new_data)
                timestamp = time.strftime("%H:%M:%S")
                print(f"[{timestamp}] [PULLED] {sheet_name} -> Sheets/{target_path.name}")

            known_hashes[sheet_name] = remote_hash
        except Exception:
            pass

    return sheet_names, known_hashes


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

        _, sheet_names = sync_all(config, verbose=True)
        known_hashes = {}
        for f in SHEETS_DIR.glob("*.*"):
            known_hashes[f.stem] = hashlib.md5(f.read_bytes()).hexdigest()

        last_tab_check = time.monotonic()
        tab_discovery_interval = 60

        try:
            while True:
                time.sleep(interval)
                if time.monotonic() - last_tab_check >= tab_discovery_interval:
                    sheet_names = None
                    last_tab_check = time.monotonic()

                sheet_names, known_hashes = two_way_sync_cycle(config, sheet_names, known_hashes)
        except KeyboardInterrupt:
            print("\nWatcher stopped.")
    else:
        print("Syncing all Google Sheets tabs to CSV files...")
        count, _ = sync_all(config, verbose=True)
        print(f"\nDone! {count} file(s) updated.")


if __name__ == "__main__":
    main()
