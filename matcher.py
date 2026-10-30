#!/usr/bin/env python3
"""
University List Matcher
=======================
Matches institute names from `Institute list.xlsx` against the official UGC Bangladesh
university list (`ugc_universities.xlsx`) and populates:
  - Fixed Name (official UGC name)
  - Confidence (match confidence percentage)
  - Uni Type (Public, Private, International, CBHE)
  - Website (official website)
"""

import argparse
import logging
import re
from typing import Dict, Optional, Tuple

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# Non-university indicators. If an institution name matches these and is not a genuine
# UGC university, it must not be falsely classified as a university.
NEGATIVE_PATTERNS = [
    r"\bschool\b",
    r"\bcollege\b",
    r"\bcollegiate\b",
    r"\bmadrasha\b",
    r"\bmadrasah\b",
    r"\balim\b",
    r"\bdakhil\b",
    r"\bfazil\b",
    r"\bkamil\b",
    r"\bpolytechnic\b",
    r"\btechnical\b",
    r"\bvocational\b",
    r"\btraining\b",
    r"\bacademy\b",
    r"\bbiddalaya\b",
    r"\bhigh school\b",
    r"\bgirls school\b",
    r"\bjunior school\b",
    r"\bsecondary school\b",
    r"\bprimary school\b",
    r"\blaboratory school\b",
    r"\buniversity college\b",
    r"\buniversity school\b",
]

NEGATIVE_REGEX = re.compile("|".join(NEGATIVE_PATTERNS), re.IGNORECASE)


def normalize_text(text: str) -> str:
    """Normalize university names for matching."""
    if not text:
        return ""
    s = text.lower()
    # Normalize common spelling variations and typos
    s = re.sub(r"\bunivers(?:i|y|ty|city)\b", "university", s)
    s = re.sub(r"\bvarsity\b", "university", s)
    s = re.sub(r"\buni\.?\b", "university", s)
    s = re.sub(r"\buinv\.?\b", "university", s)
    s = re.sub(r"\bdaka\b", "dhaka", s)
    s = re.sub(r"\bchattogram\b", "chittagong", s)
    s = re.sub(r"\bjashore\b", "jessore", s)
    s = re.sub(r"\bcumilla\b", "comilla", s)
    s = re.sub(r"\bbogura\b", "bogra", s)
    s = re.sub(r"\bbarisal\b", "barishal", s)
    # Remove special punctuation but keep letters and numbers
    s = re.sub(r"[^\w\s]", " ", s)
    # Collapse multiple whitespaces
    s = re.sub(r"\s+", " ", s).strip()
    return s


class UniversityMatcher:
    def __init__(self, ugc_file_path: str):
        self.ugc_file_path = ugc_file_path
        self.lookup: Dict[str, Dict[str, str]] = {}
        self.canonical_names: Dict[str, str] = {}
        self._load_ugc_database()

    def _load_ugc_database(self):
        """Loads UGC universities and constructs comprehensive alias mappings."""
        logger.info(f"Loading UGC database from: {self.ugc_file_path}")
        df = pd.read_excel(self.ugc_file_path)

        for _, row in df.iterrows():
            official_name = str(row["Uni Name"]).strip()
            uni_type = str(row["Uni Type"]).strip()
            web = str(row["Website"]).strip() if pd.notna(row["Website"]) else ""
            if web.lower() == "nan":
                web = ""

            record = {
                "Fixed Name": official_name,
                "Uni Type": uni_type,
                "Website": web,
            }

            norm_official = normalize_text(official_name)
            self.lookup[norm_official] = record
            self.canonical_names[norm_official] = norm_official

            # Also register inverted form: "University of X" <-> "X University"
            if norm_official.startswith("university of "):
                inv = norm_official.replace("university of ", "").strip() + " university"
                self.lookup[inv] = record
                self.canonical_names[inv] = norm_official
            elif norm_official.endswith(" university"):
                inv = "university of " + norm_official[:-11].strip()
                self.lookup[inv] = record
                self.canonical_names[inv] = norm_official

        # Comprehensive aliases for abbreviations, former names, and common variations
        manual_aliases = {
            "University of Dhaka": [
                "dhaka university", "dhaka university du", "dhaka uni", "dhaka univ",
                "dhaka varsity", "daka university", "daka universty", "du",
                "iba dhaka university", "iba du", "iba of dhaka university"
            ],
            "Bangladesh University of Engineering & Technology": [
                "bangladesh university of engineering and technology", "buet"
            ],
            "University of Rajshahi": [
                "rajshahi university", "rajshahi uni", "rajshahi univ", "ru",
                "iba rajshahi university", "iba ru"
            ],
            "Bangladesh Agricultural University": [
                "bangladesh agricultural university mymensingh", "bau", "bau mymensingh"
            ],
            "University of Chittagong": [
                "chittagong university", "chattogram university", "cu", "cu chittagong"
            ],
            "Jahangirnagar University": [
                "jahangirnagar uni", "ju", "iba jahangirnagar university", "iba ju"
            ],
            "Islamic University, Bangladesh": [
                "islamic university", "iu", "iu kushtia"
            ],
            "Shahjalal University of Science & Technology": [
                "shahjalal university of science and technology", "sust", "sust sylhet"
            ],
            "Khulna University": ["ku"],
            "National University": ["nu"],
            "Bangladesh Open University": ["bou"],
            "Bangladesh Medical University": [
                "bangabandhu sheikh mujib medical university", "bsmmu", "bmu medical"
            ],
            "Gazipur Agricultural University": [
                "bangabandhu sheikh mujibur rahman agricultural university", "bsmrau"
            ],
            "Hajee Mohammad Danesh Science & Technology University": [
                "hajee mohammad danesh science and technology university", "hstu"
            ],
            "Mawlana Bhashani Science & Technology University": [
                "mawlana bhashani science and technology university", "mbstu"
            ],
            "Patuakhali Science And Technology University": [
                "patuakhali science and technology university", "pstu"
            ],
            "Sher-e-Bangla Agricultural University": [
                "sher e bangla nagar agricultural university",
                "sher e bangla agricultural university", "sau"
            ],
            "Chittagong University of Engineering & Technology (CUET)": [
                "chittagong university of engineering technology",
                "chittagong university of engineering and technology",
                "chattogram university of engineering and technology", "cuet"
            ],
            "Rajshahi University of Engineering & Technology (RUET)": [
                "rajshahi university of engineering technology",
                "rajshahi university of engineering and technology", "ruet"
            ],
            "Khulna University of Engineering & Technology (KUET)": [
                "khulna university of engineering technology",
                "khulna university of engineering and technology", "kuet"
            ],
            "Dhaka University of Engineering & Technology (DUET)": [
                "dhaka university of engineering technology",
                "dhaka university of engineering and technology", "duet"
            ],
            "Noakhali Science & Technology University": [
                "noakhali science and technology university", "nstu"
            ],
            "Jagannath University": ["jnu", "jnu dhaka"],
            "Comilla University": ["cumilla university", "cou"],
            "Jatiya Kabi Kazi Nazrul Islam University": ["jkkniu"],
            "Chittagong Veterinary and Animal Sciences University": [
                "chattogram veterinary and animal sciences university", "cvasu"
            ],
            "Sylhet Agricultural University": ["sau sylhet"],
            "Jessore University of Science & Technology": [
                "jashore science technology university",
                "jashore science and technology university",
                "jessore science and technology university", "just"
            ],
            "Pabna University of Science and Technology": [
                "pabna university of science technology", "pust"
            ],
            "Begum Rokeya University, Rangpur": [
                "begum rokeya university", "brur"
            ],
            "Bangladesh University of Professionals": [
                "bup", "military institute of science and technology mist", "mist"
            ],
            "Gopalganj Science & Technology University": [
                "bangabandhu sheikh mujibur rahaman science technology university",
                "bangabandhu sheikh mujibur rahman science and technology university",
                "bsmrstu"
            ],
            "Bangladesh University of Textiles": ["butex"],
            "University of Barishal": ["barishal university", "barisal university", "bu"],
            "Rangamati Science and Technology University": ["rmstu"],
            "Bangladesh Maritime University": [
                "bangabandhu sheikh mujibur rahman maritime university bangladesh",
                "bangabandhu sheikh mujibur rahman maritime university",
                "bsmrmu", "bmu"
            ],
            "Islamic Arabic University": ["iau"],
            "Chittagong Medical University": ["chattogram medical university", "cmu"],
            "Rajshahi Medical University": ["rmu"],
            "Rabindra University, Bangladesh": ["rabindra university", "rub"],
            "University of Frontier Technology, Bangladesh": [
                "bangabandhu sheikh mujibur rahman digital university bangladesh",
                "bangabandhu sheikh mujibur rahman digital university",
                "bsmrdu", "uftb"
            ],
            "Netrokona University": ["sheikh hasina university", "shu"],
            "Khulna Agricultural University": ["kau"],
            "Jamalpur Science and Technology University": [
                "bangamata sheikh fojilatunnesa mujib science and technology university",
                "bangamata sheikh fazilatunnesa mujib science and technology university",
                "bsfmstu"
            ],
            "Sylhet Medical University, Sylhet": [
                "sylhet medical university",
                "bangamata sheikh fazilatunnesa mujib medical university",
                "bsfmmu"
            ],
            "Aviation And Aerospace University, Bangladesh": [
                "bangabandhu sheikh mujibur rahman aviation and aerospace university bsmraau",
                "bangabandhu sheikh mujibur rahman aviation and aerospace university",
                "bsmraau", "aaub"
            ],
            "Chandpur Science and Technology University": ["cstu"],
            "Kishoreganj University": [
                "bangabandhu sheikh mujibur rahman university kishoreganj", "bsmru"
            ],
            "Hobiganj Agricultural University": ["habiganj agricultural university", "hau"],
            "Khulna Medical University, Khulna": [
                "khulna medical university",
                "sheikh hasina medical university khulna",
                "sheikh hasina medical university",
                "shmu"
            ],
            "Kurigram Agricultural University": ["kuriau"],
            "Sunamganj Science and Technology University": ["sstu"],
            "Pirojpur Science & Technology University": [
                "bangabandhu sheikh mujibur rahman science technology university pirojpur",
                "prstu"
            ],
            "Bogra Science and Technology University": ["bogura science and technology university", "bstu"],
            "Dhaka Central University": ["dcu"],
            "North South University": ["nsu"],
            "University of Science & Technology Chittagong": [
                "university of science and technology chattogram", "ustc"
            ],
            "Independent University, Bangladesh": ["independent university bangladesh", "iub"],
            "Central Women's University": ["central womens university", "cwu"],
            "International University of Business Agriculture & Technology (IUBAT)": [
                "international university of business agriculture and technology", "iubat"
            ],
            "International Islamic University Chittagong": ["iiuc"],
            "Ahsanullah University of Science and Technology": [
                "ahsanullah university of science and technology aust", "aust"
            ],
            "American International University-Bangladesh": [
                "american international university bangladesh aiub",
                "american international university bangladesh", "aiub"
            ],
            "East West University": ["ewu"],
            "University of Asia Pacific": ["uap"],
            "Gono Bishwabidyalay": ["gono university"],
            "Asian University of Bangladesh": ["aub"],
            "Dhaka International University": ["diu dhaka"],
            "Manarat International University": ["miu"],
            "BRAC University": ["brac", "bracu"],
            "Bangladesh University": ["bu bangladesh"],
            "Leading University": ["lu"],
            "BGC Trust University Bangladesh": ["bgc trust university"],
            "Sylhet International University": ["sylhet international university sylhet", "siu"],
            "University of Development Alternative": ["uoda", "uda"],
            "Premier University": ["premier university chittagong"],
            "Southeast University": ["seu"],
            "Daffodil International University": [
                "daffodil international university diu", "diu"
            ],
            "Stamford University Bangladesh": ["stamford university"],
            "State University of Bangladesh (SUB)": [
                "state university of bangladesh", "sub"
            ],
            "Northern University Bangladesh": ["northern university", "nub"],
            "Southern University Bangladesh": ["southern university"],
            "Green University of Bangladesh": ["green university", "gub"],
            "World University of Bangladesh": ["wub"],
            "Shanto-Mariam University of Creative Technology": ["smuct"],
            "The Millennium University": ["the millenium university", "tmu"],
            "Eastern University": ["eu"],
            "Metropolitan University": ["mu"],
            "Uttara University": ["uu"],
            "United International University": ["uiu"],
            "University of South Asia": ["unisa"],
            "Victoria University of Bangladesh": ["vub"],
            "Bangladesh University of Business & Technology (BUBT)": [
                "bangladesh university of business and technology", "bubt"
            ],
            "Presidency University": ["pu"],
            "University of Information Technology & Sciences": [
                "university of information technology science", "uits"
            ],
            "Royal University of Dhaka": ["rud"],
            "University of Liberal Arts Bangladesh": ["ulab"],
            "Atish Dipankar University of Science & Technology": [
                "atish dipankar university of science and technology", "adust"
            ],
            "Bangladesh Islami University": ["biu"],
            "ASA University Bangladesh": ["asaub"],
            "East Delta University": ["edu"],
            "European University of Bangladesh": ["eub"],
            "Varendra University": ["vu"],
            "Hamdard University Bangladesh": ["hamdard university"],
            "BGMEA University of Fashion & Technology(BUFT)": [
                "bgmea university of fashion technology buft",
                "bgmea university of fashion technology", "buft"
            ],
            "North East University Bangladesh": ["neub"],
            "First Capital University of Bangladesh": ["fcub"],
            "Ishakha International University, Bangladesh": ["ishakha international university"],
            "Z.H Sikder University of Science & Technology": [
                "z h sikder univfrsity of science technology",
                "z h sikder university of science technology", "zhsust"
            ],
            "Exim Bank Agricultural University, Bangladesh": ["ebaub"],
            "North Western University": ["nwu"],
            "Khwaja Yunus Ali University": ["kyau"],
            "Sonargaon University": ["su"],
            "Feni University": ["fu"],
            "Port City International University": ["pciu"],
            "Bangladesh University of Health Sciences": ["buhs"],
            "Chittagong Independent University": [
                "chattogram independent university", "ciu"
            ],
            "Notre Dame University Bangladesh": ["ndub"],
            "North Bengal International University": ["nbiu"],
            "Fareast International University": ["fiu"],
            "Rajshahi Science & Technology University (RSTU), Natore": ["rstu"],
            "Brahmaputra International University": [
                "sheikh fazilatunnesa mujib university", "sfmu"
            ],
            "Cox's Bazar International University": ["cbiu"],
            "R. P. Shaha University": [
                "ranada prasad shaha university", "rpsu"
            ],
            "German University Bangladesh": ["gub german"],
            "Global University Bangladesh": ["gub global"],
            "CCN University of Science & Technology": ["ccnust"],
            "Bangladesh Army University of Science and Technology (BAUST), Saidpur": [
                "bangladesh army university of science and technology baust saidpur", "baust"
            ],
            "Bangladesh Army University of Engineering and Technology (BAUET), Qadirabad": ["bauet"],
            "Bangladesh Army International University of Science & Technology(BAIUST) ,Comilla": [
                "bangladesh army international university of science technology baiust cumilla", "baiust"
            ],
            "University of Scholars": ["the international university of scholars", "ius"],
            "Canadian University of Bangladesh": ["cub"],
            "N.P.I University of Bangladesh": ["npiub"],
            "Northern University of Business & Technology, Khulna": ["nubtk"],
            "University of Creative Technology, Chittagong": [
                "university of creative technology chattogram", "uctc"
            ],
            "Central University of Science and Technology": ["cust"],
            "Tagore University of Creative Arts, Uttara, Dhaka, Bangladesh": ["tuca"],
            "University of Global Village": ["ugv"],
            "Anwer Khan Modern University": ["akmu"],
            "ZNRF University of Management Sciences": ["zums"],
            "Ahsania Mission University of Science and Technology": ["amust"],
            "Khulna Khan Bahadur Ahsanullah University": ["kkbau"],
            "International Standard University": ["isu"],
            "University of Brahmanbaria": ["uob"],
            "University of Skill Enrichment and Technology": ["uset"],
            "R.T.M Al-Kabir Technical University": ["rtm aktu"],
            "Dr. Momtaz Begum University of Science and Technology (MUST)": ["must"],
            "Chattogram BGMEA University of Fashion and Technology": ["cbuft"],
            "Teesta University, Rangpur": ["teesta university"],
            "International Islami University of Science and Technology Bangladesh (IIUSTB)": ["iiustb"],
            "Lalon University of Science and Arts": ["lusa"],
            "Islamic University of Technology (IUT), Gazipur": [
                "islamic university of technology", "iut"
            ],
            "Asian University for Women": ["auw"],
            "South Asian University": ["sau intl"],
            "UCSI University, Malaysia ( Bangladesh Campus)": [
                "ucsi university", "ucsi"
            ],
        }

        for target_name, alias_list in manual_aliases.items():
            norm_target = normalize_text(target_name)
            # Find the matching record from lookup
            matched_record = None
            if norm_target in self.lookup:
                matched_record = self.lookup[norm_target]
            else:
                for k in self.lookup:
                    if fuzz.ratio(k, norm_target) > 88:
                        matched_record = self.lookup[k]
                        break

            if matched_record:
                for alias in alias_list:
                    norm_alias = normalize_text(alias)
                    self.lookup[norm_alias] = matched_record
                    self.canonical_names[norm_alias] = norm_target

        logger.info(f"Lookup dictionary constructed with {len(self.lookup)} normalized entries.")

    def match(self, raw_name: Optional[str]) -> Optional[Tuple[str, str, str, str]]:
        """
        Matches a raw institute name to a UGC university.
        Returns Tuple of (Fixed Name, Confidence, Uni Type, Website) or None.
        """
        if not raw_name or pd.isna(raw_name):
            return None

        clean_raw = str(raw_name).strip()
        if not clean_raw:
            return None

        norm = normalize_text(clean_raw)

        # 1. Exact match in lookup (Handles exact official names & all known acronyms/aliases)
        if norm in self.lookup:
            rec = self.lookup[norm]
            return rec["Fixed Name"], "100%", rec["Uni Type"], rec["Website"]

        # 2. Check if this is explicitly a school, college, madrasah, or vocational institute
        if NEGATIVE_REGEX.search(clean_raw):
            # It's an affiliated school/college, not a university
            return None

        # 3. If it contains "university" or "varsity" or acronym-like string, run fuzzy match
        is_candidate = any(w in norm for w in ["university", "varsity", "uni"]) or len(norm.split()) <= 4

        if not is_candidate:
            return None

        # Fuzzy matching against all registered lookup keys
        best_match = process.extractOne(
            norm,
            list(self.lookup.keys()),
            scorer=fuzz.token_sort_ratio,
        )

        if best_match:
            match_key, score, _ = best_match
            if score >= 85:
                rec = self.lookup[match_key]
                confidence_str = f"{round(score)}%"
                return rec["Fixed Name"], confidence_str, rec["Uni Type"], rec["Website"]

        return None


def style_matched_excel(excel_path: str):
    """Styles the matched Excel sheet with clean headers, zebra striping, and auto-filters."""
    from openpyxl import load_workbook

    wb = load_workbook(excel_path)
    ws = wb.active
    ws.title = "Matched Institutes"

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

    # Format data rows
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 21
        fill = stripe_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Set column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    wb.save(excel_path)
    logger.info(f"Applied styling to {excel_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Match Institute list against UGC Universities list."
    )
    parser.add_argument(
        "--institutes",
        default="Institute list.xlsx",
        help="Input institute list Excel file (default: Institute list.xlsx)",
    )
    parser.add_argument(
        "--ugc",
        default="ugc_universities.xlsx",
        help="UGC universities list Excel file (default: ugc_universities.xlsx)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="Institute_list_matched.xlsx",
        help="Output Excel file (default: Institute_list_matched.xlsx)",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help="Also export a CSV copy",
    )
    args = parser.parse_args()

    matcher = UniversityMatcher(args.ugc)

    logger.info(f"Reading institute list from: {args.institutes}")
    inst_df = pd.read_excel(args.institutes)

    if "InstituteName" not in inst_df.columns:
        logger.error("Required column 'InstituteName' not found in institute file.")
        return

    logger.info(f"Processing {len(inst_df)} institute entries...")

    fixed_names = []
    confidences = []
    uni_types = []
    websites = []

    matched_count = 0
    for name in inst_df["InstituteName"]:
        res = matcher.match(name)
        if res:
            fixed_name, conf, u_type, web = res
            fixed_names.append(fixed_name)
            confidences.append(conf)
            uni_types.append(u_type)
            websites.append(web)
            matched_count += 1
        else:
            fixed_names.append("")
            confidences.append("")
            uni_types.append("")
            websites.append("")

    # Construct result DataFrame with required columns
    out_df = pd.DataFrame({
        "InstituteName": inst_df["InstituteName"],
        "Fixed Name": fixed_names,
        "Confidence": confidences,
        "Uni Type": uni_types,
        "Website": websites,
    })

    logger.info(f"Matched {matched_count} universities out of {len(inst_df)} rows.")

    output_path = args.output
    out_df.to_excel(output_path, index=False, engine="openpyxl")
    logger.info(f"Saved matched results to {output_path}")

    style_matched_excel(output_path)

    if args.csv:
        csv_path = output_path.rsplit(".", 1)[0] + ".csv"
        out_df.to_csv(csv_path, index=False)
        logger.info(f"Saved CSV copy to {csv_path}")

    # Display some matched samples
    matched_samples = out_df[out_df["Fixed Name"] != ""]
    print("\n" + "=" * 80)
    print("Sample Matched Universities:")
    print("-" * 80)
    for _, row in matched_samples.sample(min(15, len(matched_samples)), random_state=42).iterrows():
        print(f"[{row['Confidence']}] '{row['InstituteName']}' -> '{row['Fixed Name']}' ({row['Uni Type']})")
    print("-" * 80)
    print(f"Total rows processed: {len(out_df)}")
    print(f"Total universities identified: {matched_count}")
    print(f"Output saved to: {output_path}")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    main()
