"""
Provenance Pulse Daily Scraper - API Intercept Version
=======================================================
Intercepts the network API calls made by provenance.io/pulse
to find and directly call the underlying data endpoints.
"""

import re
import sys
import json
import requests
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
    "Date", "TVL", "Trading TVL", "3M Chain Transactions", "3M Chain Fees",
    "Total Participants", "Total Committed Value", "Total Loan Balance",
    "Total Loans", "3M Loan Amount Funded", "3M Loans Funded",
    "3M Loan Amount Paid", "3M Loans Paid",
]


def intercept_and_scrape():
    """Launch browser, intercept API calls, and extract data directly from responses."""
    api_calls = []
    api_responses = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
        )

        # Intercept all API responses
        def handle_response(response):
            url = response.url
            # Capture any JSON API calls that look relevant
            if any(kw in url.lower() for kw in ['pulse', 'metric', 'tvl', 'loan', 'chain', 'stats']):
                try:
                    body = response.body()
                    text = body.decode('utf-8', errors='ignore')
                    if text.strip().startswith('{') or text.strip().startswith('['):
                        api_calls.append(url)
                        api_responses[url] = text
                        print(f"  Captured API: {url}")
                except Exception:
                    pass

        # Also capture ALL json responses to find the right one
        def handle_any_response(response):
            url = response.url
            content_type = response.headers.get('content-type', '')
            if 'json' in content_type and 'provenance' in url:
                try:
                    body = response.body()
                    text = body.decode('utf-8', errors='ignore')
                    if len(text) > 50:
                        api_calls.append(url)
                        api_responses[url] = text[:2000]
                        print(f"  Captured JSON: {url[:120]}")
                except Exception:
                    pass

        page = context.new_page()
        page.on("response", handle_any_response)

        print("Loading page and intercepting API calls...")
        try:
            page.goto(URL, timeout=PAGE_TIMEOUT, wait_until="networkidle")
        except PlaywrightTimeout:
            page.goto(URL, timeout=PAGE_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(15000)

        # Wait extra to catch lazy-loaded data
        page.wait_for_timeout(8000)

        browser.close()

    # Save all captured API calls to debug file
    debug_path = Path(__file__).parent / "debug_output.txt"
    with open(debug_path, "w") as f:
        f.write(f"Total API calls captured: {len(api_calls)}\n\n")
        for url in api_calls:
            f.write(f"=== URL: {url} ===\n")
            f.write(api_responses.get(url, 'no body')[:3000])
            f.write("\n\n")

    print(f"Captured {len(api_calls)} API calls, saved to debug_output.txt")
    return api_responses


def parse_metrics_from_responses(api_responses):
    """Try to extract the 12 metrics from captured API responses."""
    results = {h: "N/A" for h in EXCEL_HEADERS[1:]}

    # Look through all responses for our metrics
    for url, body in api_responses.items():
        try:
            data = json.loads(body)
            text = json.dumps(data).lower()

            # Check if this response contains relevant data
            if any(kw in text for kw in ['tvl', 'loanamount', 'loan_amount', 'chainTransactions']):
                print(f"  Promising response from: {url}")
                print(f"  Keys: {list(data.keys()) if isinstance(data, dict) else 'list'}")

                # Try to extract known fields
                flat = {}
                def flatten(obj, prefix=''):
                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            flatten(v, f"{prefix}{k}.")
                    elif isinstance(obj, list):
                        for i, v in enumerate(obj):
                            flatten(v, f"{prefix}{i}.")
                    else:
                        flat[prefix.rstrip('.')] = obj
                flatten(data)

                print(f"  Flattened keys sample: {list(flat.keys())[:20]}")
        except Exception as e:
            pass

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


def main():
    today = date.today()
    print(f"Scraping {URL} for {today}...")

    api_responses = intercept_and_scrape()
    metrics = parse_metrics_from_responses(api_responses)

    wb = get_or_create_workbook()
    ws = wb.active

    # Remove existing entry for today if any
    rows_to_delete = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == today.isoformat():
            rows_to_delete.append(i)
    for i in reversed(rows_to_delete):
        ws.delete_rows(i)

    next_row = ws.max_row + 1
    row_data = [today.isoformat()] + [metrics[h] for h in EXCEL_HEADERS[1:]]
    for col, value in enumerate(row_data, start=1):
        cell = ws.cell(row=next_row, column=col, value=value)
        cell.alignment = Alignment(horizontal="center")
    wb.save(EXCEL_PATH)
    print(f"✅ Row written for {today.isoformat()}")


if __name__ == "__main__":
    main()
