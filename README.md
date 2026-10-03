# B2B Data Scraper & ETL Pipeline (Demo Project)

> **Project Type:** Personal open-source & engineering demo by Tuan Anh Trinh ([@tuananh4865](https://github.com/tuananh4865))  
> **Tech Stack:** Python 3.11+, BeautifulSoup4, Requests, OpenPyXL, Pytest

---

## 1. Overview

This repository demonstrates a modular Python data extraction and ETL pipeline designed for catalog and directory scraping. It showcases robust error handling, session management, and automated multi-format export.

### Core Features
- **Session Pooling & Headers**: Uses `requests.Session` with realistic desktop User-Agent rotation.
- **Resilience**: Configurable exponential backoff retries on network failures or rate limits (HTTP 429/503).
- **Automated Pagination**: Crawls multi-page listings sequentially with configurable delays.
- **Data Deduplication**: In-memory URL tracking to eliminate duplicate entries during extraction.
- **Safe Multi-Format Export**:
  - **CSV**: UTF-8 BOM encoding for seamless display across Windows and macOS spreadsheet applications without encoding artifacts.
  - **Excel (.xlsx)**: Structured tables with header styling and protection against spreadsheet formula injection (`=`, `@`, `+`, `-`).

---

## 2. Project Structure

```text
b2b_data_scraper/
├── scraper.py                 # Core crawler and ETL pipeline script
├── requirements.txt           # Python dependencies
├── tests/
│   └── test_scraper.py        # Automated test suite (5 unit tests)
├── sample_output.csv          # Sample generated CSV output
└── sample_output.xlsx         # Sample generated Excel output
```

---

## 3. Quickstart & Verification

```bash
# 1. Setup virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run automated tests
python -m pytest tests/ -v

# 4. Run the scraper demo (crawls 2 sample pages)
python scraper.py --pages 2 --csv output.csv --excel output.xlsx
```

---

## 4. Automated Test Suite

The test suite in `tests/test_scraper.py` includes 5 automated tests verifying:
- HTML product card extraction and field parsing.
- URL deduplication behavior across batches.
- Exponential backoff retry logic on network timeouts.
- Formula-injection escaping during export.
- Integrity of generated CSV and Excel files.

---

## 5. License

MIT License &copy; 2026 Tuan Anh Trinh. Open for personal, learning, and reference use.
