#!/usr/bin/env python3
"""
UGC Bangladesh University List Extractor
========================================
Scrapes university listings from the University Grants Commission (UGC) of Bangladesh
and exports them into a cleanly formatted Excel spreadsheet.

Supported Categories:
  - Public Universities
  - Private Universities
  - International Universities
  - CBHE (Cross-Border Higher Education)

Target Columns:
  - Uni Name
  - Website
  - Uni Type
"""

import argparse
import html
import json
import logging
import re
import sys
from typing import Dict, List, Optional
import urllib3

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Suppress insecure request warnings if government portal SSL certs have issues
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# Standard headers to mimic browser request
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,bn;q=0.8",
    "Connection": "keep-alive",
}

# Official modern portal URLs (UGC Bangladesh)
NEW_PORTAL_URLS = [
    {
        "type": "Public University",
        "url": "https://www.ugc.gov.bd/universities/public-universities",
    },
    {
        "type": "Private University",
        "url": "https://www.ugc.gov.bd/universities/private-universities",
    },
    {
        "type": "International University",
        "url": "https://www.ugc.gov.bd/universities/international-universities",
    },
    {
        "type": "CBHE",
        "url": "https://www.ugc.gov.bd/universities/cbhe",
    },
]

# Legacy portal URLs (ugc-universities.gov.bd)
OLD_PORTAL_URLS = [
    {
        "type": "Public University",
        "url": "http://www.ugc-universities.gov.bd/public-universities",
    },
    {
        "type": "Private University",
        "url": "http://www.ugc-universities.gov.bd/private-universities",
    },
    {
        "type": "International University",
        "url": "http://www.ugc-universities.gov.bd/international-universities",
    },
    {
        "type": "CBHE",
        "url": "http://www.ugc-universities.gov.bd/cbhe",
    },
]


def clean_text(text: Optional[str]) -> str:
    """Normalize whitespace and strip leading/trailing spaces."""
    if not text:
        return ""
    # Replace multiple whitespaces/newlines with a single space
    cleaned = re.sub(r"\s+", " ", str(text)).strip()
    return cleaned


def clean_url(url: Optional[str]) -> str:
    """Clean and standardize website URL."""
    if not url:
        return ""
    cleaned = str(url).strip()
    # Remove any internal whitespaces
    cleaned = re.sub(r"\s+", "", cleaned)
    # Strip any trailing slashes or junk
    return cleaned


def scrape_new_portal(session: requests.Session) -> List[Dict[str, str]]:
    """
    Scrapes universities from the modern ugc.gov.bd portal.
    Uses the Inertia.js data-page payload embedded in the HTML.
    """
    results: List[Dict[str, str]] = []

    for entry in NEW_PORTAL_URLS:
        type_name = entry["type"]
        url = entry["url"]
        logger.info(f"Fetching {type_name} from: {url}")

        try:
            resp = session.get(url, headers=HEADERS, timeout=20, verify=False)
            resp.raise_for_status()

            # Locate the Inertia data-page attribute
            match = re.search(r'data-page=([\"\'])(.*?)\1', resp.text)
            if not match:
                logger.warning(f"Could not find data-page attribute for {type_name} at {url}")
                continue

            raw_json = html.unescape(match.group(2))
            page_data = json.loads(raw_json)
            universities = page_data.get("props", {}).get("universities", [])

            logger.info(f"  -> Extracted {len(universities)} universities for {type_name}")

            for uni in universities:
                name = clean_text(uni.get("name_en"))
                website = clean_url(uni.get("website"))
                if name:
                    results.append({
                        "Uni Name": name,
                        "Website": website,
                        "Uni Type": type_name,
                    })

        except Exception as e:
            logger.error(f"Error scraping {url}: {e}")

    return results


def scrape_old_portal(session: requests.Session) -> List[Dict[str, str]]:
    """
    Scrapes universities from the legacy ugc-universities.gov.bd portal.
    Parses standard HTML tables.
    """
    results: List[Dict[str, str]] = []

    for entry in OLD_PORTAL_URLS:
        type_name = entry["type"]
        url = entry["url"]
        logger.info(f"Fetching {type_name} from legacy portal: {url}")

        try:
            resp = session.get(url, headers=HEADERS, timeout=20, verify=False)
            resp.raise_for_status()

            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.find("table")
            if not table:
                logger.warning(f"No table found for {type_name} at {url}")
                continue

            rows = table.find_all("tr")[1:]  # Skip header row
            count = 0
            for tr in rows:
                tds = tr.find_all("td")
                if len(tds) >= 3:
                    name_td = tds[1]
                    # Decompose note popups and star markings
                    for elem in name_td.find_all(class_=["uni-note", "red-star"]):
                        elem.decompose()
                    name = clean_text(name_td.get_text())

                    web_td = tds[2]
                    web_a = web_td.find("a")
                    if web_a and web_a.get("href"):
                        website = clean_url(web_a["href"])
                    else:
                        website = clean_url(web_td.get_text())

                    if name:
                        results.append({
                            "Uni Name": name,
                            "Website": website,
                            "Uni Type": type_name,
                        })
                        count += 1

            logger.info(f"  -> Extracted {count} universities for {type_name}")

        except Exception as e:
            logger.error(f"Error scraping legacy portal {url}: {e}")

    return results


def style_excel_sheet(excel_path: str):
    """
    Applies professional styling to the generated Excel workbook:
    - Navy header with bold white text
    - Auto-adjusted column widths
    - Alternating row zebra striping
    - Light borders and padding
    - Frozen top row and auto-filter
    """
    from openpyxl import load_workbook

    wb = load_workbook(excel_path)
    ws = wb.active
    ws.title = "UGC Universities"

    # Colors & Styles
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="000000")
    stripe_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border_side = Side(style="thin", color="E2E8F0")
    cell_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )

    # Style Header
    ws.row_dimensions[1].height = 28
    for col_idx in range(1, 4):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = cell_border

    # Style Data Rows
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 22
        fill = stripe_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, 4):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        # Add buffer padding
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    # Enable AutoFilter and Freeze Panes
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(excel_path)
    logger.info(f"Applied styling to Excel file: {excel_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape university lists from UGC Bangladesh into an Excel spreadsheet."
    )
    parser.add_argument(
        "-o",
        "--output",
        default="ugc_universities.xlsx",
        help="Output Excel filename (default: ugc_universities.xlsx)",
    )
    parser.add_argument(
        "--source",
        choices=["new", "old", "both"],
        default="new",
        help="Data source: 'new' (ugc.gov.bd), 'old' (ugc-universities.gov.bd), or 'both' (default: new)",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help="Also export a copy to CSV format",
    )
    args = parser.parse_args()

    session = requests.Session()
    universities: List[Dict[str, str]] = []

    logger.info("Starting UGC University scraper...")
    if args.source == "new":
        universities = scrape_new_portal(session)
    elif args.source == "old":
        universities = scrape_old_portal(session)
    elif args.source == "both":
        new_data = scrape_new_portal(session)
        old_data = scrape_old_portal(session)
        # Deduplicate based on Uni Name (case-insensitive)
        seen = set()
        for u in new_data + old_data:
            key = u["Uni Name"].lower()
            if key not in seen:
                seen.add(key)
                universities.append(u)

    if not universities:
        logger.error("No universities were scraped. Please check network connectivity.")
        sys.exit(1)

    df = pd.DataFrame(universities)

    # Ensure correct column order
    df = df[["Uni Name", "Website", "Uni Type"]]

    # Save to Excel
    output_excel = args.output
    df.to_excel(output_excel, index=False, engine="openpyxl")
    logger.info(f"Raw data saved to {output_excel}")

    # Apply professional styling
    style_excel_sheet(output_excel)

    # Optional CSV export
    if args.csv:
        csv_path = output_excel.rsplit(".", 1)[0] + ".csv"
        df.to_csv(csv_path, index=False)
        logger.info(f"CSV copy saved to {csv_path}")

    # Summary
    print("\n" + "=" * 55)
    print("Scraping Summary:")
    print("-" * 55)
    print(df["Uni Type"].value_counts().to_string())
    print("-" * 55)
    print(f"Total Universities Scraped: {len(df)}")
    print(f"Excel file created: {output_excel}")
    if args.csv:
        print(f"CSV file created:   {csv_path}")
    print("=" * 55 + "\n")


if __name__ == "__main__":
    main()
