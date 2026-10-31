#!/usr/bin/env python3
"""
Bangladesh Medical Colleges & Universities Scraper
==================================================
Scrapes all medical colleges and universities in Bangladesh from Wikipedia:
https://en.wikipedia.org/wiki/List_of_medical_colleges_in_Bangladesh

Categories scraped:
  - Public Medical Colleges
  - Private Medical Colleges
  - Military Medical Colleges
  - Public Medical Universities

Generates professionally formatted Excel spreadsheet and CSV files.
"""

import argparse
import logging
import os
import re
from typing import Dict, List, Optional

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

WIKI_URL = "https://en.wikipedia.org/wiki/List_of_medical_colleges_in_Bangladesh"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Known supplemental websites for institutes where Wikipedia cell is blank
SUPPLEMENTAL_WEBSITES = {
    "Noakhali Medical College": "https://nmc.gov.bd/",
    "Abdul Malek Ukil Medical College": "https://nmc.gov.bd/",
}

CATEGORY_MAP = {
    0: "Public Medical College",
    1: "Private Medical College",
    2: "Military Medical College",
    3: "Public Medical University",
}


def clean_cell_text(cell) -> str:
    """Removes Wikipedia reference tags [1], [note 1], hidden spans, and trims whitespace."""
    if not cell:
        return ""
    c = BeautifulSoup(str(cell), "html.parser")
    for sup in c.find_all("sup"):
        sup.decompose()
    for span in c.find_all("span", style=lambda s: s and "display:none" in s):
        span.decompose()
    text = c.get_text(separator=" ")
    text = re.sub(r"\[\s*(?:\d+|note\s*\d+|citation needed)\s*\]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_url(cell, inst_name: str = "") -> str:
    """Extracts and normalizes the website URL."""
    if not cell:
        return SUPPLEMENTAL_WEBSITES.get(inst_name, "")
    
    a_tag = cell.find("a", href=True)
    if a_tag and a_tag["href"]:
        url = a_tag["href"].strip()
    else:
        url = cell.get_text().strip()

    url = re.sub(r"\s+", "", url)
    if url in ["-", "–", "—", "N/A", "n/a", "None", "none", ""]:
        return SUPPLEMENTAL_WEBSITES.get(inst_name, "")

    # Clean multiple trailing slashes (e.g., https://ssmcmh.gov.bd// -> https://ssmcmh.gov.bd/)
    url = re.sub(r"/+$", "/", url)

    # Add https:// if protocol is missing but domain is present
    if url and not url.startswith(("http://", "https://")):
        url = "https://" + url

    return url


def scrape_wikipedia(html_content: Optional[str] = None) -> List[Dict[str, str]]:
    """
    Fetches and parses the Wikipedia medical colleges page.
    """
    if html_content is None:
        logger.info(f"Fetching live Wikipedia page: {WIKI_URL}")
        try:
            resp = requests.get(WIKI_URL, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            html_content = resp.text
        except Exception as e:
            logger.warning(f"Failed to fetch live page ({e}), checking local cache...")
            if os.path.exists("wiki_medical.html"):
                with open("wiki_medical.html", "r", encoding="utf-8") as f:
                    html_content = f.read()
                logger.info("Loaded cached HTML from wiki_medical.html")
            else:
                raise

    soup = BeautifulSoup(html_content, "html.parser")
    tables = soup.find_all("table", class_=lambda c: c and "wikitable" in c)

    all_institutes: List[Dict[str, str]] = []

    for t_idx, table in enumerate(tables):
        if t_idx not in CATEGORY_MAP:
            continue
        
        type_name = CATEGORY_MAP[t_idx]
        rows = table.find_all("tr")
        if not rows:
            continue

        header_cells = rows[0].find_all(["th", "td"])
        headers = [clean_cell_text(th) for th in header_cells]

        # Determine column positions based on header text
        name_col = None
        web_col = None
        acronym_col = None
        location_col = None
        division_col = None
        univ_col = None
        est_col = None
        enroll_col = None

        for c_i, h in enumerate(headers):
            hl = h.lower()
            if ("college" in hl or "university" in hl) and name_col is None and "affiliated" not in hl:
                name_col = c_i
            elif "website" in hl:
                web_col = c_i
            elif "acronym" in hl:
                acronym_col = c_i
            elif "location" in hl:
                location_col = c_i
            elif "division" in hl:
                division_col = c_i
            elif "university" in hl or "affiliated" in hl:
                univ_col = c_i
            elif "established" in hl or "founded" in hl or "started" in hl:
                if est_col is None:
                    est_col = c_i
            elif "enroll" in hl:
                enroll_col = c_i

        count = 0
        for r in rows[1:]:
            cols = r.find_all(["td", "th"])
            if len(cols) < 3:
                continue

            name = clean_cell_text(cols[name_col]) if name_col is not None and name_col < len(cols) else ""
            if not name:
                continue

            website = clean_url(cols[web_col], name) if web_col is not None and web_col < len(cols) else clean_url(None, name)
            acronym = clean_cell_text(cols[acronym_col]) if acronym_col is not None and acronym_col < len(cols) else ""
            affil_univ = clean_cell_text(cols[univ_col]) if univ_col is not None and univ_col < len(cols) else ""
            location = clean_cell_text(cols[location_col]) if location_col is not None and location_col < len(cols) else ""
            division = clean_cell_text(cols[division_col]) if division_col is not None and division_col < len(cols) else ""
            est = clean_cell_text(cols[est_col]) if est_col is not None and est_col < len(cols) else ""
            enroll = clean_cell_text(cols[enroll_col]) if enroll_col is not None and enroll_col < len(cols) else ""

            record = {
                "Uni Name": name,
                "Website": website,
                "Uni Type": type_name,
                "Acronym": acronym,
                "Affiliated University": affil_univ,
                "Location": location,
                "Division": division,
                "Established": est,
                "Enrollment": enroll,
            }
            all_institutes.append(record)
            count += 1

        logger.info(f"Extracted {count} institutes for {type_name}")

    logger.info(f"Total medical institutes scraped: {len(all_institutes)}")
    return all_institutes


def format_worksheet(ws, title: str):
    """Applies professional corporate styling to an openpyxl worksheet."""
    ws.title = title

    # Styles
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

    max_row = ws.max_row
    max_col = ws.max_column

    # Format Header Row
    ws.row_dimensions[1].height = 28
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = cell_border

    # Format Data Rows
    for row_idx in range(2, max_row + 1):
        ws.row_dimensions[row_idx].height = 22
        fill = stripe_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Adjust Column Widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Freeze Header and Enable Filter
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def export_to_excel(institutes: List[Dict[str, str]], output_path: str):
    """
    Exports institutes to a formatted Excel file with:
    - Master sheet: 'All Medical Institutes'
    - Category sheets for each institute type
    """
    df_all = pd.DataFrame(institutes)

    wb = Workbook()
    # Master sheet
    ws_all = wb.active
    ws_all.title = "All Institutes"

    # Write master sheet headers
    columns = list(institutes[0].keys())
    ws_all.append(columns)
    for inst in institutes:
        ws_all.append([inst[col] for col in columns])
    format_worksheet(ws_all, "All Institutes")

    # Add category sheets
    categories = [
        ("Public Medical Colleges", "Public Medical College"),
        ("Private Medical Colleges", "Private Medical College"),
        ("Military Medical Colleges", "Military Medical College"),
        ("Medical Universities", "Public Medical University"),
    ]

    for sheet_title, cat_type in categories:
        cat_insts = [inst for inst in institutes if inst["Uni Type"] == cat_type]
        if not cat_insts:
            continue
        ws = wb.create_sheet(title=sheet_title)
        ws.append(columns)
        for inst in cat_insts:
            ws.append([inst[col] for col in columns])
        format_worksheet(ws, sheet_title)

    wb.save(output_path)
    logger.info(f"Saved styled Excel file to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Scrape Bangladesh medical colleges and universities into Excel."
    )
    parser.add_argument(
        "-o",
        "--output",
        default="medical_institutes.xlsx",
        help="Output Excel filename (default: medical_institutes.xlsx)",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        default=True,
        help="Also export a copy to CSV format (default: True)",
    )
    args = parser.parse_args()

    institutes = scrape_wikipedia()

    export_to_excel(institutes, args.output)

    # Also save as medical_colleges_bangladesh.xlsx
    export_to_excel(institutes, "medical_colleges_bangladesh.xlsx")

    if args.csv:
        csv_path = os.path.splitext(args.output)[0] + ".csv"
        df = pd.DataFrame(institutes)
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        df.to_csv("medical_colleges_bangladesh.csv", index=False, encoding="utf-8-sig")
        logger.info(f"Saved CSV file to: {csv_path} and medical_colleges_bangladesh.csv")


if __name__ == "__main__":
    main()

