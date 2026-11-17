# E-commerce Product Scraper

A Python scraper that collects every product from a WooCommerce shop ([scrapeme.live](https://scrapeme.live/shop/), a site made for scraping practice) and saves the data to Excel and a SQLite database.

**Result: 755 products** scraped with name, price, SKU, stock status, categories, image URL and product URL. See [`sample_output.csv`](sample_output.csv) for the first rows.

## How it works

The scraper works in two tiers:

1. **Listing pages:** it walks through every paginated shop page (`/shop/page/2/`, `/shop/page/3/` ...) and collects the name, price, image and link of each product. It stops when a page returns 404 or has no products.
2. **Product pages:** it opens each product's own page to get the details that are not on the listing: **SKU**, **stock status**, **categories** and a higher-resolution image.

Then it removes duplicates and saves everything to:

- `ecommerce_products.xlsx` (Excel)
- `products.db` (SQLite table `products`)
- `execution.log` (timings, record count and any failed URLs)

## Polite and reliable scraping

- **Random delay** of 1 to 3 seconds between requests, so the site is not overloaded.
- **Rotating User-Agents** from a small list of real browser strings.
- **Retry with backoff** on network errors, HTTP 429 (too many requests) and 5xx server errors, up to 4 attempts.
- **Failed URLs are logged**, and one broken product never stops the whole run.

## Setup

```
python -m venv venv
venv\Scripts\activate          (Windows)
pip install -r requirements.txt
python scrape_ecommerce.py
```

The run takes a while because of the polite delays. Progress is printed in the terminal and written to `execution.log`.

## Output columns

| Column | Example |
|---|---|
| Product_Name | Bulbasaur |
| Price | 63 |
| SKU | 4391 |
| Stock_Status | 45 in stock |
| Categories | Pokemon, Seed |
| Image_URL | https://scrapeme.live/wp-content/uploads/2018/08/001.png |
| Product_URL | https://scrapeme.live/shop/Bulbasaur/ |

## Tech

Python · Requests · BeautifulSoup · Pandas · openpyxl · SQLite

## Author

**Muhammad Hashmi** · Swat, Pakistan · hashmi.freelance@gmail.com
