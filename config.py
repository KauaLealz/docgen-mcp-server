import os

def get_scan_max_matches() -> int:
    raw = os.environ.get("DOCGEN_SCAN_MAX_MATCHES")
    if not raw:
        return 500
    try:
        n = int(raw)
        return min(max(1, n), 50000)
    except ValueError:
        return 500

def get_read_sheet_max_rows() -> int:
    raw = os.environ.get("DOCGEN_READ_SHEET_MAX_ROWS")
    if not raw:
        return 10000
    try:
        n = int(raw)
        return min(max(1, n), 100000)
    except ValueError:
        return 10000
