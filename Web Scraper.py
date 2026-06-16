"""
Web Scraper - Universal Website Data Extractor
Author: M-R-Tech0007
Description: A professional Python web scraping tool that extracts
             titles, links, images, tables, and custom CSS selectors
             from any website and exports to CSV / JSON / Excel.
"""

import os
import csv
import json
import time
import logging
import requests
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin, urlparse

try:
    from bs4 import BeautifulSoup
except ImportError:
    raise SystemExit("❌  Missing dependency: run  pip install beautifulsoup4 requests openpyxl")

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# ──────────────────────────────────────────────
#  LOGGING
# ──────────────────────────────────────────────
def setup_logging() -> logging.Logger:
    logger = logging.getLogger("WebScraper")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    log_file = f"scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


# ──────────────────────────────────────────────
#  HTTP HELPERS
# ──────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,/;q=0.8",
}


def fetch_page(url: str, retries: int = 3, delay: float = 2.0,
               logger: logging.Logger = None) -> BeautifulSoup | None:
    """Fetch a URL and return a BeautifulSoup object, with retry logic."""
    for attempt in range(1, retries + 1):
        try:
            logger and logger.info("Fetching (attempt %d): %s", attempt, url)
            resp = requests.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding
            return BeautifulSoup(resp.text, "html.parser")
        except requests.exceptions.HTTPError as e:
            logger and logger.warning("HTTP error %s — %s", e.response.status_code, url)
        except requests.exceptions.ConnectionError:
            logger and logger.warning("Connection error — %s", url)
        except requests.exceptions.Timeout:
            logger and logger.warning("Timeout — %s", url)
        except requests.exceptions.RequestException as e:
            logger and logger.error("Request failed: %s", e)

        if attempt < retries:
            logger and logger.info("Retrying in %.1f seconds…", delay)
            time.sleep(delay)

    logger and logger.error("All %d attempts failed for: %s", retries, url)
    return None


# ──────────────────────────────────────────────
#  EXTRACTION FUNCTIONS
# ──────────────────────────────────────────────
def scrape_titles(soup: BeautifulSoup) -> list[dict]:
    """Extract all headings (h1–h3) from the page."""
    results = []
    for tag in soup.find_all(["h1", "h2", "h3"]):
        text = tag.get_text(strip=True)
        if text:
            results.append({"tag": tag.name.upper(), "text": text})
    return results


def scrape_links(soup: BeautifulSoup, base_url: str) -> list[dict]:
    """Extract all hyperlinks with their anchor text."""
    results = []
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        full_url = urljoin(base_url, href)
        if full_url not in seen and full_url.startswith("http"):
            seen.add(full_url)
            results.append({
                "text": a.get_text(strip=True) or "(no text)",
                "url": full_url,
            })
    return results


def scrape_images(soup: BeautifulSoup, base_url: str) -> list[dict]:
    """Extract all images with src and alt text."""
    results = []
    for img in soup.find_all("img"):
        src = img.get("src", "").strip()
        if src:
            full_src = urljoin(base_url, src)
            results.append({
                "alt": img.get("alt", "").strip() or "(no alt)",
                "src": full_src,
            })
    return results


def scrape_tables(soup: BeautifulSoup) -> list[list[list[str]]]:
    """Extract all HTML tables as a list of 2-D arrays."""
    all_tables = []
    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True)
                     for td in tr.find_all(["td", "th"])]
            if any(cells):
                rows.append(cells)
        if rows:
            all_tables.append(rows)
    return all_tables


def scrape_custom(soup: BeautifulSoup, css_selector: str) -> list[dict]:
    """Extract elements matching a custom CSS selector."""
    results = []
    for el in soup.select(css_selector):
        results.append({
            "tag":  el.name,
            "text": el.get_text(strip=True),
            "html": str(el)[:300],          # first 300 chars of raw HTML
        })
    return results


# ──────────────────────────────────────────────
#  EXPORT FUNCTIONS
# ──────────────────────────────────────────────
def export_csv(data: list[dict], filepath: Path) -> None:
    if not data:
        return
    filepath.parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)


def export_json(data: list | dict, filepath: Path) -> None:
    filepath.parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def export_excel(sheets: dict[str, list[dict]], filepath: Path) -> None:
    if not EXCEL_AVAILABLE:
        print("⚠️  openpyxl not installed — skipping Excel export.")
        return
    filepath.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)                        # remove default empty sheet

    for sheet_name, rows in sheets.items():
        if not rows:
            continue
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel limit: 31 chars
        ws.append(list(rows[0].keys()))              # header row
        for row in rows:
            ws.append(list(row.values()))

    wb.save(filepath)


# ──────────────────────────────────────────────
#  MAIN SCRAPER
# ──────────────────────────────────────────────
def run_scraper(
    url: str,
    mode: str,
    output_dir: str = "output",
    css_selector: str = "",
    logger: logging.Logger = None,
) -> dict:
    """
    Scrape a URL and export results.

    Parameters
    ----------
    url          : Target URL (must start with http/https)
    mode         : 'titles' | 'links' | 'images' | 'tables' | 'custom' | 'all'
    output_dir   : Folder where exports are saved
    css_selector : Required when mode='custom'
    logger       : Logger instance

    Returns
    -------
    dict  with all scraped data
    """
    if logger is None:
        logger = setup_logging()

    # Validate URL
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        logger.error("Invalid URL (must start with http:// or https://): %s", url)
        return {}

    soup = fetch_page(url, logger=logger)
    if soup is None:
        return {}

    base_url = f"{parsed.scheme}://{parsed.netloc}"
    out = Path(output_dir)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    domain = parsed.netloc.replace("www.", "").split(".")[0]
    prefix = f"{domain}_{timestamp}"

    results = {}

    # ── Titles ──
    if mode in ("titles", "all"):
        titles = scrape_titles(soup)
        logger.info("Titles found: %d", len(titles))
        results["titles"] = titles
        if titles:
            export_csv(titles,  out / f"{prefix}_titles.csv")
            export_json(titles, out / f"{prefix}_titles.json")

    # ── Links ──
    if mode in ("links", "all"):
        links = scrape_links(soup, base_url)
        logger.info("Links found: %d", len(links))
        results["links"] = links
        if links:
            export_csv(links,  out / f"{prefix}_links.csv")
            export_json(links, out / f"{prefix}_links.json")

    # ── Images ──
    if mode in ("images", "all"):
        images = scrape_images(soup, base_url)
        logger.info("Images found: %d", len(images))
        results["images"] = images
        if images:
            export_csv(images,  out / f"{prefix}_images.csv")
            export_json(images, out / f"{prefix}_images.json")

    # ── Tables ──
    if mode in ("tables", "all"):
        tables = scrape_tables(soup)
        logger.info("Tables found: %d", len(tables))
        results["tables"] = tables
        for i, table in enumerate(tables, 1):
            rows_as_dicts = [{"col_" + str(j): v for j, v in enumerate(row)}
                             for row in table]
            export_csv(rows_as_dicts, out / f"{prefix}table{i}.csv")

    # ── Custom selector ──
    if mode == "custom":
        if not css_selector:
            logger.error("Custom mode requires a CSS selector.")
        else:
            custom = scrape_custom(soup, css_selector)
            logger.info("Custom elements found: %d  (selector: %s)",
                        len(custom), css_selector)
            results["custom"] = custom
            if custom:
                export_csv(custom,  out / f"{prefix}_custom.csv")
                export_json(custom, out / f"{prefix}_custom.json")

    # ── All → Excel summary ──
    if mode == "all" and EXCEL_AVAILABLE:
        sheets = {}
        if results.get("titles"):  sheets["Titles"] = results["titles"]
        if results.get("links"):   sheets["Links"]  = results["links"]
        if results.get("images"):  sheets["Images"] = results["images"]
        if sheets:
            export_excel(sheets, out / f"{prefix}_summary.xlsx")
            logger.info("Excel summary saved.")

    logger.info("All exports saved to: %s/", out)
    return results


# ──────────────────────────────────────────────
#  CLI
# ──────────────────────────────────────────────
MODES = {
    "1": ("titles",  "Page headings (H1, H2, H3)"),
    "2": ("links",   "All hyperlinks"),
    "3": ("images",  "All images"),
    "4": ("tables",  "HTML tables"),
    "5": ("custom",  "Custom CSS selector"),
    "6": ("all",     "Everything (titles + links + images + tables)"),
}


def main():
    print("\n" + "=" * 60)
    print("      WEB SCRAPER  |  by M-R-Tech0007")
    print("=" * 60)

    url = input("\nEnter the URL to scrape\n> ").strip()
    if not url.startswith("http"):
        url = "https://" + url

    print("\nWhat do you want to scrape?")
    for key, (_, label) in MODES.items():
        print(f"  [{key}] {label}")

    choice = input("\nYour choice [6]: ").strip() or "6"
    mode, _ = MODES.get(choice, ("all", ""))

    css_selector = ""
    if mode == "custom":
        css_selector = input("Enter CSS selector (e.g. div.product-title): ").strip()

    output_dir = input("\nOutput folder [output]: ").strip() or "output"

    logger = setup_logging()
    print()
    run_scraper(url, mode, output_dir, css_selector, logger)
    print("\n✅  Done! Check the output folder and log file.")


if _name_ == "_main_":
    main()