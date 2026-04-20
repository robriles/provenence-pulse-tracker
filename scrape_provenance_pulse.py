"""
Provenance Pulse Daily Scraper
================================
Scrapes all 12 Provenance Blockchain Metrics from https://provenance.io/pulse
on the 3m time period, and appends them with today's date to an Excel spreadsheet.

SETUP (run once):
    pip install playwright openpyxl
    playwright install chromium

SCHEDULE (run daily):
  Mac/Linux — crontab -e, add:
    0 13 * * * /usr/bin/python3 /path/to/scrape_provenance_pulse.py
  Windows — Task Scheduler, Daily at 8:00 AM ET
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
PAGE_TIMEOUT = 60_000
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

# These are the exact label texts on the page after clicking 3m.
# The page changes "Week's" to "3 Months" when 3m is selected.
# We match by partial text to handle both cases robustly.
METRIC_LABEL_PATTERNS = [
    "TVL",
    "Trading TVL",
    "Chain Transactions",
    "Chain Fees",
    "Total Participants",
    "Committed Value",
    "Loan Balance",
    "Total Loans",
    "Loan Amount Funded",
    "Loans Funded",
    "Loan Amount Paid",
    "Loans Paid",
]


def click_3m_button(page):
    """Click the 3m button in the Provenance Blockchain Metrics section."""
    # There are two sets of time buttons (Hash Metrics + Blockchain Metrics).
    # We want the SECOND set. Get all 3m buttons and click the last one.
    try:
        buttons = page.locator("button", has_text=re.compile(r"^3m$", re.IGNORECASE)).all()
        if buttons:
            buttons[-1].click()
            page.wait_for_timeout(4000)
            print(f"  Clicked 3m button (found {len(buttons)} total)")
        else:
            print("  No 3m button found, proceeding with default")
    except Exception as e:
        print(f"  Could not click 3m: {e}")


def scrape_metrics() -> dict:
    results = {h: "N/A" for h in EXCEL_HEADERS[1:]}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print("Loading page...")
        try:
            page.goto(URL, timeout=PAGE_TIMEOUT, wait_until="networkidle")
        except PlaywrightTimeout:
            page.goto(URL, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(8000)

        print("Selecting 3m time period...")
        click_3m_button(page)

        # Extract all SPAN text nodes with their positions
        print("Extracting text nodes...")
        nodes = page.evaluate("""
            () => {
                const spans = document.querySelectorAll('span');
                const results = [];
                for (const span of spans) {
                    const text = span.textContent.trim();
                    if (text.length === 0) continue;
                    const rect = span.getBoundingClientRect();
                    if (rect.width === 0) continue;
                    results.push({ text: text, y: rect.top, x: rect.left });
                }
                return results;
            }
        """)

        print(f"Found {len(nodes)} span nodes")

        # For each metric label pattern, find the matching span,
        # then find the next span that looks like a value (has digits)
        for i, pattern in enumerate(METRIC_LABEL_PATTERNS):
            header = EXCEL_HEADERS[i + 1]
            found = False

            for j, node in enumerate(nodes):
                if pattern.lower() in node['text'].lower() and len(node['text']) < 60:
                    # Look at the next few spans for a numeric value
                    for k in range(j + 1, min(j + 5, len(nodes))):
                        candidate = nodes[k]['text']
                        # Must contain digits and look like a real value
                        if re.search(r'\d', candidate) and candidate not in ['i', '%']:
                            # Skip pure percentages or tiny numbers that are change indicators
                            # A main value is usually on a similar x position and close y
                            if re.search(r'[\$\d,\.]+[BMKTbmkt]?', candidate):
                                clean = candidate.strip()
                                # Filter out change % values like "(1.27%)"
                                if not re.match(r'^\(.*%\)$', clean):
                                    results[header] = clean
                                    print(f"  {header}: {clean}  (label: '{node['text']}')")
                                    found = True
                                    break
                    if found:
                        break

            if not found:
                print(f"  {header}: NOT FOUND (pattern: '{pattern}')")

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
    row_data = [today.isoformat()] + [metrics[h] for h in EXCEL_HEADERS[1:]]

    for col, value in enumerate(row_data, start=1):
        cell = ws.cell(row=next_row, column=col, value=value)
        cell.alignment = Alignment(horizontal="center")

    wb.save(EXCEL_PATH)
    print(f"\n✅ Saved {len(metrics)} metrics for {today.isoformat()}")


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
        metrics = {h: "SCRAPE FAILED" for h in EXCEL_HEADERS[1:]}

    append_row(wb, today, metrics)


if __name__ == "__main__":
    main()
