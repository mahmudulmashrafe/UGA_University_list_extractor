#!/usr/bin/env python3
"""
Sohopathi Bangladesh Institution Scraper
========================================
Scrapes educational institutions across Bangladesh from sohopathi.com:
  1. Primary Schools      (https://www.sohopathi.com/primary-schools/)
  2. Secondary Schools    (https://www.sohopathi.com/schools/)
  3. Colleges             (https://www.sohopathi.com/colleges/)
  4. Madrasahs            (https://www.sohopathi.com/madrasahs/)
  5. Medical Colleges     (https://www.sohopathi.com/medical-colleges/)

Exports the consolidated institution list into:
  - sohopathi_institutions.xlsx
  - sohopathi_institutions.csv
"""

import argparse
import concurrent.futures
import csv
import html
import logging
import os
import re
import sys
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,bn;q=0.8",
}

CATEGORIES = [
    {
        "slug": "medical-colleges",
        "type": "Medical College",
        "url": "https://www.sohopathi.com/medical-colleges/",
    },
    {
        "slug": "colleges",
        "type": "College",
        "url": "https://www.sohopathi.com/colleges/",
    },
    {
        "slug": "madrasahs",
        "type": "Madrasah",
        "url": "https://www.sohopathi.com/madrasahs/",
    },
    {
        "slug": "primary-schools",
        "type": "Primary School",
        "url": "https://www.sohopathi.com/primary-schools/",
    },
    {
        "slug": "schools",
        "type": "Secondary School",
        "url": "https://www.sohopathi.com/schools/",
    },
]


def create_session(workers: int) -> requests.Session:
    """Configures a thread-safe requests session with connection pooling."""
    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(
        pool_connections=workers * 2,
        pool_maxsize=workers * 2,
        max_retries=3,
    )
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(HEADERS)
    return session


def get_total_pages(session: requests.Session, base_url: str) -> int:
    """Detects total pages from the pagination element on page 1."""
    try:
        r = session.get(base_url, timeout=15)
        if r.status_code != 200:
            logger.warning(f"Failed to fetch {base_url} (HTTP {r.status_code})")
            return 1

        # Search for page numbers in links
        matches = re.findall(r'/page/(\d+)/', r.text)
        if matches:
            return max(int(m) for m in matches)
        return 1
    except Exception as e:
        logger.error(f"Error finding total pages for {base_url}: {e}")
        return 1


def parse_page_html(html_text: str, inst_type: str) -> List[Dict[str, str]]:
    """Fast regex extraction of institute cards from HTML."""
    chunks = html_text.split('class="sp-post-content"')[1:]
    results = []

    for c in chunks:
        title_m = re.search(r'class=[\"\']sp-post-title[\"\']>\s*(.*?)\s*</div>', c, re.DOTALL)
        if not title_m:
            continue

        name = html.unescape(title_m.group(1)).strip()
        # Clean extra whitespace
        name = re.sub(r"\s+", " ", name)

        dist_m = re.search(r'title=[\'\"]District[\'\"][^>]*>\s*<a[^>]*>\s*(.*?)\s*</a>', c, re.DOTALL)
        district = html.unescape(dist_m.group(1)).strip() if dist_m else ""

        eiin_m = re.search(r'title=[\'\"]EIIN[\'\"][^>]*>\s*(.*?)\s*</div>', c, re.DOTALL)
        eiin = html.unescape(eiin_m.group(1)).strip() if eiin_m else ""

        link_m = re.search(r'<div class=[\"\']sp-post-action-bar[\"\']>\s*<a href=[\"\']([^\"\']+)[\"\']', c)
        link = link_m.group(1).strip() if link_m else ""

        results.append({
            "Institute Name": name,
            "Institute Type": inst_type,
            "District": district,
            "EIIN": eiin,
            "URL": link,
        })

    return results


def scrape_category(
    session: requests.Session,
    slug: str,
    inst_type: str,
    base_url: str,
    max_pages_limit: Optional[int],
    workers: int,
    csv_writer,
    csv_file,
) -> int:
    """Scrapes all pages of a specific category using concurrent threads."""
    total_pages = get_total_pages(session, base_url)
    if max_pages_limit:
        total_pages = min(total_pages, max_pages_limit)

    logger.info(f"Starting {inst_type} ({slug}): {total_pages} total pages to scrape...")

    def fetch_single_page(p: int) -> List[Dict[str, str]]:
        url = f"{base_url}page/{p}/" if p > 1 else base_url
        for attempt in range(3):
            try:
                r = session.get(url, timeout=15)
                if r.status_code == 200:
                    return parse_page_html(r.text, inst_type)
                elif r.status_code == 404:
                    return []
            except Exception:
                time.sleep(1)
        return []

    collected_count = 0
    start_time = time.time()
    batch_size = 100

    # Scrape in page chunks for steady memory and regular CSV flushing
    for start_p in range(1, total_pages + 1, batch_size):
        end_p = min(start_p + batch_size - 1, total_pages)
        page_range = range(start_p, end_p + 1)

        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_page = {executor.submit(fetch_single_page, p): p for p in page_range}
            for future in concurrent.futures.as_completed(future_to_page):
                items = future.result()
                for item in items:
                    csv_writer.writerow(item)
                    collected_count += 1

        csv_file.flush()
        elapsed = time.time() - start_time
        speed = collected_count / elapsed if elapsed > 0 else 0
        pct = (end_p / total_pages) * 100
        logger.info(
            f"  [{inst_type}] Pages {start_p}-{end_p}/{total_pages} ({pct:.1f}%) "
            f"-> {collected_count} items total ({speed:.1f} items/sec)"
        )

    return collected_count


def style_excel(excel_path: str):
    """Styles the final Excel workbook with modern professional formatting."""
    from openpyxl import load_workbook

    wb = load_workbook(excel_path)
    ws = wb.active
    ws.title = "Sohopathi Institutions"

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

    ws.row_dimensions[1].height = 28
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = cell_border

    # Sample width inspection across first 500 rows to keep memory fast
    for col in ws.iter_cols(min_row=1, max_row=min(ws.max_row, 500)):
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            v_len = len(str(cell.value or ""))
            if v_len > max_len:
                max_len = v_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(excel_path)
    logger.info(f"Applied styling to {excel_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape schools, colleges, madrasahs, and medical colleges from sohopathi.com."
    )
    parser.add_argument(
        "--categories",
        nargs="+",
        choices=["medical-colleges", "colleges", "madrasahs", "primary-schools", "schools"],
        help="Specify which categories to scrape (default: all)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=18,
        help="Concurrent network worker threads (default: 18)",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Maximum pages to scrape per category (for testing)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="sohopathi_institutions.xlsx",
        help="Output Excel filename (default: sohopathi_institutions.xlsx)",
    )
    parser.add_argument(
        "--csv-output",
        default="sohopathi_institutions.csv",
        help="Output CSV filename (default: sohopathi_institutions.csv)",
    )
    args = parser.parse_args()

    session = create_session(args.workers)

    target_categories = CATEGORIES
    if args.categories:
        target_categories = [c for c in CATEGORIES if c["slug"] in args.categories]

    csv_path = args.csv_output
    excel_path = args.output

    fieldnames = ["Institute Name", "Institute Type", "District", "EIIN", "URL"]

    # Open CSV writer for continuous streaming
    csv_file = open(csv_path, "w", newline="", encoding="utf-8")
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()

    overall_start = time.time()
    category_summary = {}

    for cat in target_categories:
        count = scrape_category(
            session=session,
            slug=cat["slug"],
            inst_type=cat["type"],
            base_url=cat["url"],
            max_pages_limit=args.max_pages,
            workers=args.workers,
            csv_writer=writer,
            csv_file=csv_file,
        )
        category_summary[cat["type"]] = count

    csv_file.close()

    total_scraped = sum(category_summary.values())
    elapsed = time.time() - overall_start

    logger.info(f"Scraping completed! Total {total_scraped} institutions saved to {csv_path} in {elapsed/60:.2f} mins.")

    # Convert to Excel
    logger.info(f"Generating formatted Excel file: {excel_path}...")
    df = pd.read_csv(csv_path)
    df.to_excel(excel_path, index=False, engine="openpyxl")
    style_excel(excel_path)

    # Print summary
    print("\n" + "=" * 65)
    print("Sohopathi Scraping Summary:")
    print("-" * 65)
    for k, v in category_summary.items():
        print(f"  {k:20}: {v:,} institutions")
    print("-" * 65)
    print(f"  Total Institutions  : {total_scraped:,}")
    print(f"  Total Time Elapsed  : {elapsed/60:.2f} minutes")
    print(f"  Excel File Created  : {excel_path}")
    print(f"  CSV File Created    : {csv_path}")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    main()
