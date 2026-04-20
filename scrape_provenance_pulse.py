"""
Provenance Pulse Daily Scraper - Direct API Version
=====================================================
Calls the Provenance Explorer API directly to fetch all 12 metrics
and appends them with today's date to an Excel spreadsheet.

SETUP (run once):
    pip install requests openpyxl

SCHEDULE (run daily):
  Mac/Linux — crontab -e, add:
    0 13 * * * /usr/bin/python3 /path/to/scrape_provenance_pulse.py
  Windows — Task Scheduler, Daily at 8:00 AM ET
"""

import sys
import json
import requests
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE_URL = "https://service-explorer.provenance.io/api/pulse/metric/type"
RANGE = "3m"
EXCEL_PATH = Path(__file__).parent / "provenance_pulse_tracker.xlsx"
# ──────────────────────────────────────────────────────────────────────────────

EXCEL_HEADERS = [
    "Date",
    "TVL",
    "Trading TVL",
    "3M Chain Transactions",
    "3M Chain Fees",
    "Total Participants",
    "Total Committed Value",
    "Total Loan Balance",
    "Total Loans",
    "3M Loan Amount Funded",
    "3M Loans Funded",
    "3M Loan Amount Paid",
    "3M Loans Paid",
]

# Maps Excel header -> API metric type name
METRICS = {
    "TVL":                    "PULSE_TVL_METRIC",
    "Trading TVL":            "PULSE_TRADING_TVL_METRIC",
    "3M Chain Transactions":  "PULSE_TRANSACTION_VOLUME_METRIC",
    "3M Chain Fees":          "PULSE_CHAIN_FEES_VALUE_METRIC",
    "Total Participants":     "PULSE_PARTICIPANTS_METRIC",
    "Total Committed Value":  "PULSE_COMMITTED_ASSETS_VALUE_METRIC",
    "Total Loan Balance":     "LOAN_LEDGER_TOTAL_BALANCE_METRIC",
    "Total Loans":            "LOAN_LEDGER_TOTAL_COUNT_METRIC",
    "3M Loan Amount Funded":  "LOAN_LEDGER_DISBURSEMENTS_METRIC",
    "3M Loans Funded":        "LOAN_LEDGER_DISBURSEMENT_COUNT_METRIC",
    "3M Loan Amount Paid":    "LOAN_LEDGER_PAYMENTS_METRIC",
    "3M Loans Paid":          "LOAN_LEDGER_TOTAL_PAYMENTS_METRIC",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://provenance.io/",
    "Origin": "https://provenance.io",
}


def fetch_metric(metric_name: str) -> str:
    """Fetch a single metric from the API and return its current value as a string."""
    url = f"{BASE_URL}/{metric_name}?range={RANGE}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        data = resp.json()

        # The API returns a list of data points; the last one is the most recent
        # Structure is typically: { "data": [{"value": ..., "date": ...}, ...] }
        # or { "metricData": [...] } — handle both
        points = (
            data.get("data") or
            data.get("metricData") or
            data.get("metrics") or
            (data if isinstance(data, list) else None)
        )

        # API returns a single object with keys: id, base, amount, quote, quoteAmount, trend, progress, series
        if isinstance(data, dict):
            value = data.get("amount") or data.get("quoteAmount") or data.get("base")
            if value is not None:
                return str(value)

        # Fallback: dump keys for debugging
        print(f"  Unexpected structure for {metric_name}: {list(data.keys()) if isinstance(data, dict) else type(data)}")
        return "Parse error"

    except requests.HTTPError as e:
        print(f"  HTTP error for {metric_name}: {e}")
        return f"HTTP {resp.status_code}"
    except Exception as e:
        print(f"  Error for {metric_name}: {e}")
        return "Error"


def scrape_all_metrics() -> dict:
    results = {}
    for header in EXCEL_HEADERS[1:]:
        metric_name = METRICS[header]
        value = fetch_metric(metric_name)
        results[header] = value
        print(f"  {header}: {value}")
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
    row_data = [today.isoformat()] + [metrics[h] for h in EXCEL_HEADERS[1:]]
    for col, value in enumerate(row_data, start=1):
        cell = ws.cell(row=next_row, column=col, value=value)
        cell.alignment = Alignment(horizontal="center")
    wb.save(EXCEL_PATH)
    print(f"\n✅ Saved {len(metrics)} metrics for {today.isoformat()}")


def main():
    today = date.today()
    print(f"Fetching Provenance Pulse metrics for {today} (range={RANGE})...")

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
