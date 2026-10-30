# UGA_University_list_extractor & Sohopathi Scraper

A comprehensive educational institution dataset extractor and cleaner for Bangladesh.

---

## 📁 Generated Files & Data

### 1. Sohopathi Educational Institutions (167,396 records)
Scraped from [sohopathi.com](https://www.sohopathi.com/) across 5 institution categories:
- [sohopathi_institutions.xlsx](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/sohopathi_institutions.xlsx) (Excel workbook, 5.4 MB)
- [sohopathi_institutions.csv](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/sohopathi_institutions.csv) (CSV file, 19.6 MB)
- [sohopathi_scraper.py](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/sohopathi_scraper.py) (Multithreaded scraper script)

#### Category Breakdown:
| Category | URL Source | Count |
|---|---|---|
| **Secondary School** | `https://www.sohopathi.com/schools/` | 82,427 |
| **Primary School** | `https://www.sohopathi.com/primary-schools/` | 63,824 |
| **Madrasah** | `https://www.sohopathi.com/madrasahs/` | 16,417 |
| **College** | `https://www.sohopathi.com/colleges/` | 4,682 |
| **Medical College** | `https://www.sohopathi.com/medical-colleges/` | 46 |
| **Total** | | **167,396** |

#### Columns:
- `Institute Name`
- `Institute Type` (`Secondary School`, `Primary School`, `Madrasah`, `College`, `Medical College`)
- `District`
- `EIIN`
- `URL`

---

### 2. UGC Bangladesh Universities (181 records)
- [ugc_universities.xlsx](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/ugc_universities.xlsx)
- [ugc_universities.csv](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/ugc_universities.csv)
- [scraper.py](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/scraper.py)

---

### 3. Matched BDJobs Institute List (35,068 records)
- [Institute list.xlsx](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/Institute%20list.xlsx)
- [Institute_list_matched.csv](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/Institute_list_matched.csv)
- [matcher.py](file:///Users/mahmudulmashrafe/Programming/4.%20Bdjobs/UGA_University_list_extractor/matcher.py)

---

## 💻 How to Run the Sohopathi Scraper

```bash
# Scrape all 5 categories (creates sohopathi_institutions.xlsx & .csv)
python3 sohopathi_scraper.py --workers 20

# Scrape specific categories only
python3 sohopathi_scraper.py --categories medical-colleges colleges

# Quick test run (e.g. 2 pages per category)
python3 sohopathi_scraper.py --max-pages 2
```
