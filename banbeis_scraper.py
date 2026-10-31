#!/usr/bin/env python3
"""
BANBEIS Educational Institution Scraper
=======================================
Scrapes all registered educational institutions in Bangladesh from the official
BANBEIS (Bangladesh Bureau of Educational Information and Statistics) IEIMS portal:
https://institute-search.ieims.banbeis.gov.bd/search-result

API Endpoint:
https://apigw.ieims.banbeis.gov.bd/eatl-config-service/api/v1/institute/public/list

Outputs:
  - banbeis_institutions.xlsx (Styled Excel workbook)
  - banbeis_institutions.csv (UTF-8 with BOM for Excel compatibility)
"""

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

BASE_API_URL = "https://apigw.ieims.banbeis.gov.bd/eatl-config-service/api/v1/institute/public/list"
PORTAL_DETAILS_URL = "https://institute-search.ieims.banbeis.gov.bd/details"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://institute-search.ieims.banbeis.gov.bd",
    "Referer": "https://institute-search.ieims.banbeis.gov.bd/search-result",
}

PAGE_SIZE = 50


def create_session() -> requests.Session:
    """Creates a configured requests session with retries and pooling."""
    session = requests.Session()
    session.headers.update(HEADERS)
    session.verify = False

    retries = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(
        max_retries=retries,
        pool_connections=25,
        pool_maxsize=25,
    )
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def fetch_page(session: requests.Session, page: int) -> List[Dict[str, Any]]:
    """Fetches a single page of institutes from the BANBEIS public API."""
    params = {"page": page, "size": PAGE_SIZE}
    try:
        resp = session.get(BASE_API_URL, params=params, timeout=20)
        resp.raise_for_status()
        json_data = resp.json()
        if json_data.get("status") == "success" and "data" in json_data:
            return json_data["data"]
        return []
    except Exception as e:
        logger.error(f"Error fetching page {page}: {e}")
        return []


def scrape_all_institutions(max_pages: Optional[int] = None, max_workers: int = 12) -> List[Dict[str, Any]]:
    """
    Scrapes all institutions across all pages concurrently.
    """
    session = create_session()
    
    # 1. Fetch page 1 to get metadata and total count
    logger.info("Connecting to BANBEIS portal API...")
    resp = session.get(BASE_API_URL, params={"page": 1, "size": PAGE_SIZE}, timeout=20)
    resp.raise_for_status()
    first_data = resp.json()
    
    meta = first_data.get("meta", {})
    total_institutions = meta.get("total", 0)
    total_pages = (total_institutions + PAGE_SIZE - 1) // PAGE_SIZE
    
    logger.info(f"BANBEIS Database: {total_institutions:,} total institutions across {total_pages:,} pages.")
    
    if max_pages and max_pages < total_pages:
        total_pages = max_pages
        logger.info(f"Limiting scrape to {total_pages} pages as requested.")

    page_results: Dict[int, List[Dict[str, Any]]] = {}
    page_results[1] = first_data.get("data", [])

    pages_to_fetch = list(range(2, total_pages + 1))
    logger.info(f"Starting parallel download of {len(pages_to_fetch)} remaining pages using {max_workers} workers...")

    start_time = time.time()
    completed_count = 1  # Page 1 already fetched

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_page = {
            executor.submit(fetch_page, session, page_num): page_num
            for page_num in pages_to_fetch
        }

        for future in as_completed(future_to_page):
            page_num = future_to_page[future]
            try:
                data = future.result()
                page_results[page_num] = data
            except Exception as exc:
                logger.error(f"Page {page_num} generated an exception: {exc}")
                page_results[page_num] = []

            completed_count += 1
            if completed_count % 100 == 0 or completed_count == total_pages:
                elapsed = time.time() - start_time
                pct = (completed_count / total_pages) * 100
                logger.info(f"Progress: {completed_count}/{total_pages} pages ({pct:.1f}%) in {elapsed:.1f}s")

    # Combine in correct order
    raw_list: List[Dict[str, Any]] = []
    for p in range(1, total_pages + 1):
        raw_list.extend(page_results.get(p, []))

    logger.info(f"Fetched {len(raw_list):,} raw records. Processing and cleaning...")

    cleaned_records: List[Dict[str, str]] = []
    seen_eiin = set()

    for item in raw_list:
        eiin = str(item.get("eiinNo") or "").strip()
        if not eiin or eiin in seen_eiin:
            continue
        seen_eiin.add(eiin)

        name_en = re.sub(r"\s+", " ", str(item.get("instituteName") or "")).strip()
        name_bn = re.sub(r"\s+", " ", str(item.get("instituteNameBn") or "")).strip()
        type_name = str(item.get("instituteTypeName") or "").strip()
        division = str(item.get("divisionName") or "").strip()
        district = str(item.get("districtName") or "").strip()
        thana = str(item.get("thanaName") or "").strip()
        mobile = str(item.get("mobile") or "").strip()
        is_stopped = item.get("stopLegacy", False)
        status = "Inactive" if is_stopped else "Active"
        profile_url = f"{PORTAL_DETAILS_URL}/{eiin}"

        cleaned_records.append({
            "EIIN": eiin,
            "Institute Name": name_en,
            "Institute Name (Bangla)": name_bn,
            "Institute Type": type_name,
            "Division": division,
            "District": district,
            "Thana / Upazila": thana,
            "Mobile": mobile,
            "Status": status,
            "BANBEIS Profile URL": profile_url,
        })

    logger.info(f"Final cleaned institution count: {len(cleaned_records):,}")
    return cleaned_records


def export_to_excel(institutes: List[Dict[str, str]], output_path: str):
    """
    Exports institutions into a styled openpyxl Excel spreadsheet.
    """
    logger.info(f"Writing {len(institutes):,} institutions to Excel workbook: {output_path}...")
    wb = Workbook()
    ws = wb.active
    ws.title = "BANBEIS Institutions"

    columns = list(institutes[0].keys())
    ws.append(columns)

    for item in institutes:
        ws.append([item[col] for col in columns])

    # Corporate Navy styling
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10, color="000000")
    stripe_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    ws.row_dimensions[1].height = 28
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = thin_border

    # Sample maximum lengths to compute column widths quickly without scanning 40k rows
    sample_size = min(len(institutes), 2000)
    col_widths = {col: len(col) for col in columns}
    for item in institutes[:sample_size]:
        for col in columns:
            val_len = len(str(item[col]))
            if val_len > col_widths[col]:
                col_widths[col] = min(val_len, 50)

    for idx, col in enumerate(columns, 1):
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = max(col_widths[col] + 4, 12)

    # Enable AutoFilter and freeze header row
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(output_path)
    logger.info(f"Excel file successfully generated: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape all educational institutions in Bangladesh from BANBEIS IEIMS."
    )
    parser.add_argument(
        "-o",
        "--output",
        default="banbeis_institutions.xlsx",
        help="Output Excel filename (default: banbeis_institutions.xlsx)",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        default=True,
        help="Also export to CSV format (default: True)",
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=None,
        help="Maximum pages to scrape (omit to scrape all 805 pages)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=12,
        help="Number of concurrent threads (default: 12)",
    )
    args = parser.parse_args()

    t_start = time.time()
    institutes = scrape_all_institutions(max_pages=args.pages, max_workers=args.workers)

    if not institutes:
        logger.error("No institutions extracted. Exiting.")
        sys.exit(1)

    export_to_excel(institutes, args.output)

    if args.csv:
        csv_path = os.path.splitext(args.output)[0] + ".csv"
        logger.info(f"Exporting CSV to {csv_path} (utf-8-sig)...")
        df = pd.DataFrame(institutes)
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        logger.info(f"CSV file saved: {csv_path}")

    total_time = time.time() - t_start
    logger.info(f"Completed full scrape in {total_time:.1f} seconds. Total institutes: {len(institutes):,}")


if __name__ == "__main__":
    main()

