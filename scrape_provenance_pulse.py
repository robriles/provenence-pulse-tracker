"""
Provenance Pulse Daily Scraper
================================
Scrapes all 12 metrics from https://provenance.io/pulse (3m time period)
and appends them with today's date to an Excel spreadsheet.

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
import json
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

# Partial label strings to search for (case-insensitive, flexible matching)
METRIC_SEARCH = [
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


def extract_value(text: str) -> str:
    """Pull the first recognizable numeric/dollar value from a block of text."""
    # Match things like $23.94B, $641.69M, 1,246,088, 74,559, $990,884,746
    match = re.search(
        r"(\$\s*[\d,]+(?:\.\d+)?\s*[BMKTbmkt]?|[\d,]+(?:\.\d+)?\s*[BMKTbmkt]?)",
        text
    )
    if match:
        val = match.group(1).strip()
        # Filter out bare single/double digit numbers (likely UI noise)
        digits_only = re.sub(r"[,$\s]", "", val)
        if len(digits_only) >= 3:
            return val
    return "Parse error"


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

        # Click the 3m button — try multiple selector strategies
        print("Selecting 3m time period...")
        clicked = False
        for selector in [
            "text=3m", "text=3M",
            "button:has-text('3m')", "button:has-text('3M')",
            "[class*='tab']:has-text('3m')", "[class*='button']:has-text('3m')",
        ]:
            try:
                el = page.locator(selector).first
                el.click(timeout=3000)
                page.wait_for_timeout(4000)
                clicked = True
                print(f"  Clicked 3m via: {selector}")
                break
            except Exception:
                continue
        if not clicked:
            print("  Could not click 3m button — proceeding with default view")

        # Dump full page text for debugging
        page.wait_for_timeout(3000)
        full_text = page.inner_text("body")
        print("\n--- PAGE TEXT DUMP (first 3000 chars) ---")
        print(full_text[:3000])
        print("--- END DUMP ---\n")

        # Strategy: get ALL card-like elements and match by label proximity
        # Try to find metric cards using common dashboard patterns
        cards_found = 0

        # Try approach 1: find all elements that contain a dollar sign or large number
        # paired with a nearby label
        all_elements = page.locator("body *").all()

        # Build a map of label -> value by scanning all text nodes
        # Look for patterns where a short label is near a large value
        # Use page.evaluate to extract structured data via JS
        card_data = page.evaluate("""
            () => {
                const results = [];
                // Get all visible text elements
                const walker = document.createTreeWalker(
                    document.body,
                    NodeFilter.SHOW_TEXT,
                    null
                );
                const texts = [];
                let node;
                while (node = walker.nextNode()) {
                    const text = node.textContent.trim();
                    if (text.length > 0 && text.length < 200) {
                        const rect = node.parentElement 
                            ? node.parentElement.getBoundingClientRect() 
                            : null;
                        if (rect && rect.width > 0) {
                            texts.push({
                                text: text,
                                y: rect.top,
                                x: rect.left,
                                tag: node.parentElement ? node.parentElement.tagName : ''
                            });
                        }
                    }
                }
                return texts;
            }
        """)

        print(f"Found {len(card_data)} text nodes on page")

        # Print all text nodes for debugging
        for item in card_data:
            print(f"  [{item['tag']}] y={item['y']:.0f} x={item['x']:.0f}: {item['text']!r}")

        browser.close()

    return results, card_data, full_text


def main():
    print(f"DEBUG MODE: Scraping {URL} to understand page structure...")
    metrics, card_data, full_text = scrape_metrics()

    # Save debug info
    debug_path = Path(__file__).parent / "debug_output.txt"
    with open(debug_path, "w") as f:
        f.write("=== FULL PAGE TEXT ===\n")
        f.write(full_text)
        f.write("\n\n=== TEXT NODES ===\n")
        for item in card_data:
            f.write(f"[{item['tag']}] y={item['y']:.0f} x={item['x']:.0f}: {item['text']!r}\n")

    print(f"\nDebug info saved to {debug_path}")
    print("Please share debug_output.txt so we can fix the scraper!")


if __name__ == "__main__":
    main()
