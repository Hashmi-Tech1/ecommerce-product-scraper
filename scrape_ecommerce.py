"""
Production-grade scraper for https://scrapeme.live/shop/

Features:
- Two-tier scraping: listing pages -> product detail pages (for SKU + Stock)
- Anti-bot mitigation: rotating User-Agents, random delay (1-3s), retry logic
  for 429 / 5xx errors with exponential backoff
- Output: Excel (.xlsx) + SQLite database (products.db)
- Logging: execution.log with timings, record counts, and failed URLs
"""

import logging
import random
import re
import sqlite3
import time
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ----------------------------- CONFIG ----------------------------- #
BASE_URL = "https://scrapeme.live/shop/"
EXCEL_OUTPUT = "ecommerce_products.xlsx"
DB_OUTPUT = "products.db"
LOG_FILE = "execution.log"

MAX_RETRIES = 4
BACKOFF_BASE = 2          # seconds, doubles each retry
MIN_DELAY, MAX_DELAY = 1, 3   # polite random delay between requests

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
    "(KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 "
    "Firefox/125.0",
]

# ----------------------------- LOGGING ----------------------------- #
logging.basicConfig(
    filename=LOG_FILE,
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)
# also print to console
console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(console)

session = requests.Session()
failed_urls = []


def get_headers():
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }


def fetch(url):
    """GET a URL with retry logic for 429 / 5xx errors, exponential backoff."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, headers=get_headers(), timeout=15)
        except requests.RequestException as e:
            logger.warning(f"Attempt {attempt} network error on {url}: {e}")
            time.sleep(BACKOFF_BASE * attempt)
            continue

        if resp.status_code == 200:
            return resp
        if resp.status_code == 429 or resp.status_code >= 500:
            wait = BACKOFF_BASE * attempt
            logger.warning(
                f"Attempt {attempt} got HTTP {resp.status_code} on {url}. "
                f"Retrying in {wait}s..."
            )
            time.sleep(wait)
            continue
        if resp.status_code == 404:
            return None  # no more pages
        logger.warning(f"Unexpected status {resp.status_code} on {url}")
        return None

    logger.error(f"Giving up on {url} after {MAX_RETRIES} attempts")
    failed_urls.append(url)
    return None


def clean_price(text):
    if not text:
        return None
    match = re.search(r"[\d,]+\.?\d*", text.replace(",", ""))
    return float(match.group()) if match else None


# ----------------------------- TIER 1: LISTING PAGES ----------------------------- #
def scrape_listing_pages():
    """Walk paginated shop pages, return list of dicts with basic info + product URL."""
    products = []
    page = 1

    while True:
        url = BASE_URL if page == 1 else f"{BASE_URL}page/{page}/"
        resp = fetch(url)
        if resp is None:
            logger.info(f"No more listing pages after page {page - 1}.")
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        items = soup.select("li.product")
        if not items:
            logger.info(f"No products found on page {page}, stopping.")
            break

        for item in items:
            try:
                link_tag = item.select_one("a.woocommerce-LoopProduct-link, a.product_type_simple, h2 ~ a, a")
                name_tag = item.select_one("h2, h3")
                price_tag = item.select_one("span.woocommerce-Price-amount")
                img_tag = item.select_one("img")

                product_url = link_tag["href"] if link_tag and link_tag.has_attr("href") else None
                name = name_tag.get_text(strip=True) if name_tag else None
                price = clean_price(price_tag.get_text()) if price_tag else None
                image_url = None
                if img_tag:
                    image_url = (
                        img_tag.get("data-src")
                        or img_tag.get("data-lazy-src")
                        or img_tag.get("src")
                    )
                    if image_url:
                        image_url = urljoin(url, image_url)

                if not product_url:
                    continue

                products.append(
                    {
                        "Product_Name": name,
                        "Price": price,
                        "Image_URL": image_url,
                        "Product_URL": urljoin(url, product_url),
                    }
                )
            except Exception as e:
                logger.warning(f"Skipped one item on page {page}: {e}")

        logger.info(f"Listing page {page} done ({len(products)} products so far)")
        page += 1
        time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))

    return products


# ----------------------------- TIER 2: PRODUCT DETAIL PAGES ----------------------------- #
def scrape_detail(product):
    """Visit a single product page and enrich the dict with SKU, Stock, Categories."""
    resp = fetch(product["Product_URL"])
    if resp is None:
        product["SKU"] = None
        product["Stock_Status"] = None
        product["Categories"] = None
        return product

    soup = BeautifulSoup(resp.text, "html.parser")

    sku_tag = soup.select_one("span.sku")
    product["SKU"] = sku_tag.get_text(strip=True) if sku_tag else None

    stock_tag = soup.select_one("p.stock, .stock")
    product["Stock_Status"] = stock_tag.get_text(strip=True) if stock_tag else "Unknown"

    cat_tags = soup.select("span.posted_in a")
    product["Categories"] = ", ".join(c.get_text(strip=True) for c in cat_tags) if cat_tags else None

    # fall back to a higher-res image from the detail page if available
    detail_img = soup.select_one("div.woocommerce-product-gallery img")
    if detail_img:
        src = detail_img.get("data-large_image") or detail_img.get("src")
        if src:
            product["Image_URL"] = urljoin(product["Product_URL"], src)

    return product


def scrape_all_details(products):
    enriched = []
    for i, product in enumerate(products, start=1):
        enriched.append(scrape_detail(product))
        if i % 10 == 0 or i == len(products):
            logger.info(f"Detail pages processed: {i}/{len(products)}")
        time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))
    return enriched


# ----------------------------- STORAGE ----------------------------- #
def save_to_excel(df):
    columns = [
        "Product_Name", "Price", "SKU", "Stock_Status",
        "Categories", "Image_URL", "Product_URL",
    ]
    df = df.reindex(columns=columns)
    df.to_excel(EXCEL_OUTPUT, index=False)


def save_to_sqlite(df):
    conn = sqlite3.connect(DB_OUTPUT)
    df.to_sql("products", conn, if_exists="replace", index=False)
    conn.close()


# ----------------------------- MAIN ----------------------------- #
def main():
    start_time = time.time()
    logger.info("===== SCRAPE STARTED =====")
    logger.info(f"Target: {BASE_URL}")

    listing_products = scrape_listing_pages()
    logger.info(f"Tier 1 complete: {len(listing_products)} products found in listing pages.")

    full_products = scrape_all_details(listing_products)

    df = pd.DataFrame(full_products).drop_duplicates(subset="Product_URL")

    save_to_excel(df)
    save_to_sqlite(df)

    elapsed = time.time() - start_time
    avg_speed = elapsed / len(df) if len(df) else 0

    logger.info("===== EXECUTION REPORT =====")
    logger.info(f"Total records scraped : {len(df)}")
    logger.info(f"Failed URLs           : {len(failed_urls)}")
    for u in failed_urls:
        logger.info(f"  - {u}")
    logger.info(f"Total execution time  : {elapsed:.2f} seconds")
    logger.info(f"Avg time per record   : {avg_speed:.2f} seconds")
    logger.info(f"Excel saved to        : {EXCEL_OUTPUT}")
    logger.info(f"SQLite saved to       : {DB_OUTPUT}")
    logger.info(f"Run finished at       : {datetime.now().isoformat()}")


if __name__ == "__main__":
    main()