"""
Provenance Pulse Daily Scraper
================================
Fetches Provenance Blockchain metrics directly from the Explorer API.
- 6 static metrics (single value, no time range variation)
- 6 time-series metrics captured across 24h, 1w, 1m, 3m

SETUP (run once):
    pip install requests openpyxl

SCHEDULE (run daily):
  Mac/Linux — crontab -e, add:
    0 13 * * * /usr/bin/python3 /path/to/scrape_provenance_pulse.py
  Windows — Task Scheduler, Daily at 8:00 AM ET
"""

import sys
import requests
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE_URL = "https://service-explorer.provenance.io/api/pulse/metric/type"
EXCEL_PATH = Path(__file__).parent / "provenance_pulse_tracker.xlsx"
# ──────────────────────────────────────────────────────────────────────────────

HEADERS_HTTP = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://provenance.io/",
    "Origin": "https://provenance.io",
}

# Static metrics — fetched once with 3m range (range doesn't affect their value)
STATIC_METRICS = {
    "TVL":                   "PULSE_TVL_METRIC",
    "Trading TVL":           "PULSE_TRADING_TVL_METRIC",
    "Total Participants":    "PULSE_PARTICIPANTS_METRIC",
    "Total Committed Value": "PULSE_COMMITTED_ASSETS_VALUE_METRIC",
    "Total Loan Balance":    "LOAN_LEDGER_TOTAL_BALANCE_METRIC",
    "Total Loans":           "LOAN_LEDGER_TOTAL_COUNT_METRIC",
}

# Time-series metrics — fetched for each of 24h, 1w, 1m, 3m
TIME_SERIES_METRICS = {
    "Chain Transactions":  "PULSE_TRANSACTION_VOLUME_METRIC",
    "Chain Fees":          "PULSE_CHAIN_FEES_VALUE_METRIC",
    "Loan Amount Funded":  "LOAN_LEDGER_DISBURSEMENTS_METRIC",
    "Loans Funded":        "LOAN_LEDGER_DISBURSEMENT_COUNT_METRIC",
    "Loan Amount Paid":    "LOAN_LEDGER_PAYMENTS_METRIC",
    "Loans Paid":          "LOAN_LEDGER_TOTAL_PAYMENTS_METRIC",
}

RANGES = ["24h", "1w", "1m", "3m"]

# Build full ordered header list
EXCEL_HEADERS = ["Date"] + list(STATIC_METRICS.keys())
for metric in TIME_SERIES_METRICS.keys():
    for r in RANGES:
        EXCEL_HEADERS.append(f"{metric} ({r})")


def fetch_metric(metric_name: str, range_val: str) -> str:
    url = f"{BASE_URL}/{metric_name}?range={range_val}"
    try:
        resp = requests.get(url, headers=HEADERS_HTTP, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        if isinstance(data, dict):
            value = data.get("amount") or data.get("quoteAmount") or data.get("base")
            if value is not None:
                return str(value)
        print(f"  Unexpected structure for {metric_name} ({range_val}): {list(data.keys()) if isinstance(data, dict) else type(data)}")
        return "Parse error"
    except requests.HTTPError as e:
        return f"HTTP {resp.status_code}"
    except Exception as e:
        print(f"  Error for {metric_name} ({range_val}): {e}")
        return "Error"


def scrape_all_metrics() -> dict:
    results = {}

    # Static metrics (3m range)
    for label, api_name in STATIC_METRICS.items():
        value = fetch_metric(api_name, "3m")
        results[label] = value
        print(f"  {label}: {value}")

    # Time-series metrics across all ranges
    for label, api_name in TIME_SERIES_METRICS.items():
        for r in RANGES:
            col = f"{label} ({r})"
            value = fetch_metric(api_name, r)
            results[col] = value
            print(f"  {col}: {value}")

    return results


def get_or_create_workbook():
    if EXCEL_PATH.exists():
        return load_workbook(EXCEL_PATH)

    wb = Workbook()
    ws = wb.active
    ws.title = "Pulse Metrics"
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", start_color="1F4E79")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    for i, h in enumerate(EXCEL_HEADERS, start=1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin
    ws.column_dimensions["A"].width = 14
    for col in range(2, len(EXCEL_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    ws.row_dimensions[1].height = 36
    wb.save(EXCEL_PATH)
    return wb


def append_row(wb, today: date, metrics: dict):
    ws = wb.active
    next_row = ws.max_row + 1
    row_data = [today.isoformat()] + [metrics.get(h, "N/A") for h in EXCEL_HEADERS[1:]]
    for col, value in enumerate(row_data, start=1):
        cell = ws.cell(row=next_row, column=col, value=value)
        cell.alignment = Alignment(horizontal="center")
    wb.save(EXCEL_PATH)
    print(f"\n✅ Saved {len(metrics)} metrics for {today.isoformat()}")


def main():
    today = date.today()
    print(f"Fetching Provenance Pulse metrics for {today}...")

    wb = get_or_create_workbook()
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] == today.isoformat():
            print(f"⚠️  Entry for {today} already exists. Skipping.")
            return

    try:
        metrics = scrape_all_metrics()
    except Exception as e:
        print(f"❌ Fatal error: {e}", file=sys.stderr)
        metrics = {h: "FAILED" for h in EXCEL_HEADERS[1:]}

    append_row(wb, today, metrics)


if __name__ == "__main__":
    main()
