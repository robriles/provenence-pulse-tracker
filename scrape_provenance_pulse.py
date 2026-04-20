"""
Provenance Pulse Daily Scraper
================================
Scrapes all 12 metrics from https://provenance.io/pulse (3m time period)
and appends them with today's date to an Excel spreadsheet.

SETUP (run once):
    pip install playwright openpyxl
    playwright install chromium

SCHEDULE (run daily):
  Mac/Linux — add to crontab:
    crontab -e
    0 8 * * * /usr/bin/python3 /path/to/scrape_provenance_pulse.py

  Windows — use Task Scheduler to run:
    python C:\path\to\scrape_provenance_pulse.py
    (set trigger: Daily, at 8:00 AM)
"""

import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ── CONFIG ────────────────────────────────────────────────────────────────────
URL = "https://provenance.io/pulse"
EXCEL_PATH = Path(__file__).parent / "provenance_pulse_tracker.xlsx"
PAGE_TIMEOUT = 45_000
# ──────────────────────────────────────────────────────────────────────────────

METRIC_LABELS = [
    "TVL",
    "Trading TVL",
    "3 Months Chain Transactions",
    "3 Months Chain Fees",
    "Total Participants",
    "Total Committed Value",
    "Total Loan Balance",
    "Total Loans",
    "3 Months Loan Amount Funded",
    "3 Months Loans Funded",
    "3 Months Loan Amount Paid",
    "3 Months Loans Paid",
]

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


def scrape_metrics() -> dict:
    results = {label: "N/A" for label in METRIC_LABELS}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        try:
            page.goto(URL, timeout=PAGE_TIMEOUT, wait_until="networkidle")
        except PlaywrightTimeout:
            page.goto(URL, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(6000)

        # Click the "3m" time period button
        try:
            btn = page.locator("button, span, div").filter(has_text=re.compile(r"^3m$", re.IGNORECASE)).first
            btn.click(timeout=5000)
            page.wait_for_timeout(3000)
        except Exception:
            pass

        for label in METRIC_LABELS:
            try:
                label_el = page.locator(f"text={label}").first
                label_el.wait_for(timeout=8000, state="visible")
                card = label_el.locator("xpath=../..").first
                card_text = card.inner_text()
                value_match = re.search(
                    r"([\$]?\s*[\d,]+(?:\.\d+)?\s*[BMKTbmkt]?)",
                    card_text
                )
                if value_match:
                    results[label] = value_match.group(1).strip()
                else:
                    results[label] = "Parse error"
            except Exception as e:
                results[label] = f"Error: {str(e)[:40]}"

        browser.close()

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
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

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
    row_data = [today.isoformat()] + [metrics[label] for label in METRIC_LABELS]

    for col, value in enumerate(row_data, start=1):
        cell = ws.cell(row=next_row, column=col, value=value)
        cell.alignment = Alignment(horizontal="center")

    wb.save(EXCEL_PATH)
    print(f"✅ Saved {len(METRIC_LABELS)} metrics for {today.isoformat()}")
    for label, value in metrics.items():
        print(f"   {label}: {value}")


def main():
    today = date.today()
    print(f"Scraping {URL} for {today} (3m view)...")

    wb = get_or_create_workbook()
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] == today.isoformat():
            print(f"⚠️  Entry for {today} already exists. Skipping.")
            return

    try:
        metrics = scrape_metrics()
    except Exception as e:
        print(f"❌ Fatal scrape error: {e}", file=sys.stderr)
        metrics = {label: "SCRAPE FAILED" for label in METRIC_LABELS}

    append_row(wb, today, metrics)


if __name__ == "__main__":
    main()
