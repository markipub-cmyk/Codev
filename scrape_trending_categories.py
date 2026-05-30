"""
Scraper: q84sale.com — Trending Categories (shortCuts section)

Fetches the homepage, parses the __NEXT_DATA__ JSON embedded in the HTML,
extracts items from the "الأقسام الأكثر بحثاً" (most-searched categories) block,
writes an Excel file, then uploads it to Cloudflare R2 under:

  <bucket>/4sale-data/trending_categories/year=YYYY/month=MM/day=DD/excel-files/
"""

import json
import os
from datetime import datetime, timezone

import boto3
import openpyxl
import requests
from bs4 import BeautifulSoup
from botocore.config import Config

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
TARGET_URL = "https://www.q84sale.com/ar"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ar-KW,ar;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# DataImpulse proxy — Account 1 (user already has __cr.XX suffix appended by workflow)
_PROXY_USER = os.environ["DATAIMPULSE_USER"]   # e.g. myuser__cr.kw
_PROXY_PASS = os.environ["DATAIMPULSE_PASS"]
_PROXY_HOST = os.environ.get("PROXY_HOST", "gw.dataimpulse.com")
_PROXY_PORT = os.environ.get("PROXY_PORT", "823")

PROXIES = {
    "http":  f"http://{_PROXY_USER}:{_PROXY_PASS}@{_PROXY_HOST}:{_PROXY_PORT}",
    "https": f"http://{_PROXY_USER}:{_PROXY_PASS}@{_PROXY_HOST}:{_PROXY_PORT}",
}

# Cloudflare R2 — loaded from environment (set by GitHub Actions secrets)
CF_ACCESS_KEY_ID     = os.environ["CF_R2_ACCESS_KEY_ID"]
CF_SECRET_ACCESS_KEY = os.environ["CF_R2_SECRET_ACCESS_KEY"]
CF_ENDPOINT_URL      = os.environ["CF_R2_ENDPOINT_URL"]
CF_BUCKET_NAME       = os.environ["CF_R2_BUCKET_NAME"]


# ---------------------------------------------------------------------------
# Step 1: Fetch page and extract __NEXT_DATA__
# ---------------------------------------------------------------------------
def fetch_next_data(url: str) -> dict:
    resp = requests.get(url, headers=HEADERS, proxies=PROXIES, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    tag = soup.find("script", {"id": "__NEXT_DATA__", "type": "application/json"})
    if not tag:
        raise RuntimeError("__NEXT_DATA__ script tag not found on page")

    return json.loads(tag.string)


# ---------------------------------------------------------------------------
# Step 2: Parse items from shortCuts
# ---------------------------------------------------------------------------
def extract_trending_items(next_data: dict) -> list[dict]:
    shortcuts = next_data.get("props", {}).get("pageProps", {}).get("shortCuts", [])

    category_blocks = [s for s in shortcuts if s.get("data_type") == "category"]
    if not category_blocks:
        raise RuntimeError("No shortCuts block with data_type='category' found")

    rows = []
    scrape_ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    for block in category_blocks:
        block_title = block.get("title", "")
        for item in block.get("items", []):
            filters = item.get("filters", {})
            rows.append(
                {
                    "scrape_date":    datetime.now(timezone.utc).strftime("%Y-%m-%d"),
                    "scrape_ts":      scrape_ts,
                    "block_title":    block_title,
                    "item_id":        item.get("item_id"),
                    "name":           item.get("name"),
                    "slug":           item.get("slug"),
                    "full_path_slug": filters.get("full_path_slug"),
                    "category_id":    filters.get("category_id"),
                    "display_order":  item.get("display_order"),
                    "icon":           item.get("icon"),
                }
            )

    return rows


# ---------------------------------------------------------------------------
# Step 3: Write Excel
# ---------------------------------------------------------------------------
COLUMNS = [
    "scrape_date",
    "scrape_ts",
    "block_title",
    "item_id",
    "name",
    "slug",
    "full_path_slug",
    "category_id",
    "display_order",
    "icon",
]


def write_excel(rows: list[dict], filepath: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "trending_categories"

    ws.append(COLUMNS)

    for row in rows:
        ws.append([row.get(col) for col in COLUMNS])

    # Auto-fit column widths (approximate)
    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 60)

    wb.save(filepath)
    print(f"Excel written → {filepath}  ({len(rows)} rows)")


# ---------------------------------------------------------------------------
# Step 4: Upload to Cloudflare R2
# ---------------------------------------------------------------------------
def r2_client():
    return boto3.client(
        "s3",
        endpoint_url=CF_ENDPOINT_URL,
        aws_access_key_id=CF_ACCESS_KEY_ID,
        aws_secret_access_key=CF_SECRET_ACCESS_KEY,
        config=Config(signature_version="s3v4"),
        region_name="auto",
    )


def upload_to_r2(local_path: str, s3_key: str) -> None:
    client = r2_client()
    client.upload_file(
        local_path,
        CF_BUCKET_NAME,
        s3_key,
        ExtraArgs={"ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
    )
    print(f"Uploaded → r2://{CF_BUCKET_NAME}/{s3_key}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    now = datetime.now(timezone.utc)
    year  = now.strftime("%Y")
    month = now.strftime("%m")
    day   = now.strftime("%d")
    date_str = now.strftime("%Y-%m-%d")

    filename = f"trending_categories_{date_str}.xlsx"
    local_path = f"/tmp/{filename}"

    s3_key = (
        f"4sale-data/trending_categories/"
        f"year={year}/month={month}/day={day}/"
        f"excel-files/{filename}"
    )

    print(f"Proxy user : {_PROXY_USER}  host: {_PROXY_HOST}:{_PROXY_PORT}")
    print(f"Fetching {TARGET_URL} ...")
    next_data = fetch_next_data(TARGET_URL)

    print("Extracting trending category items ...")
    rows = extract_trending_items(next_data)
    print(f"Found {len(rows)} items")

    write_excel(rows, local_path)

    print(f"Uploading to Cloudflare R2 ...")
    upload_to_r2(local_path, s3_key)

    print("Done.")


if __name__ == "__main__":
    main()
