"""Google/DuckDuckGo email scraper utility."""

from __future__ import annotations

import argparse
import re
import time
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    from selenium.webdriver import Chrome
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager
except Exception:  # pragma: no cover - selenium optional
    Chrome = None  # type: ignore
    Options = None  # type: ignore
    ChromeDriverManager = None  # type: ignore

EMAIL_RE = re.compile(r"[\w.%-]+@[\w.-]+\.[A-Za-z]{2,}")
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


def build_driver() -> Chrome:
    """Return a headless Chrome webdriver."""
    if Chrome is None or Options is None or ChromeDriverManager is None:
        raise RuntimeError("selenium dependencies are missing")
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    return Chrome(ChromeDriverManager().install(), options=opts)


@dataclass
class StoreInfo:
    name: str
    email: Optional[str]


def parse_title(title: str) -> str:
    """Extract store name from a search result title."""
    return title.split("-")[0].split("|")[0].strip()


def fetch_serp(keyword: str, use_selenium: bool = False) -> List[str]:
    """Return store names from the first two DuckDuckGo SERP pages."""
    stores: List[str] = []
    base_url = "https://duckduckgo.com/html/"
    driver: Optional[Chrome] = None

    if use_selenium:
        driver = build_driver()

    try:
        for page in range(2):
            params = {"q": keyword, "s": str(page * 50)}
            if use_selenium and driver:
                driver.get(f"{base_url}?q={keyword}&s={page * 50}")
                html = driver.page_source
            else:
                resp = requests.get(
                    base_url, params=params, headers=HEADERS, timeout=10
                )
                resp.raise_for_status()
                html = resp.text

            soup = BeautifulSoup(html, "html.parser")
            for res in soup.select("div.result__body"):
                a = res.select_one("a.result__a")
                if not a:
                    continue
                title = a.get_text(strip=True)
                store = parse_title(title)
                if store and store not in stores:
                    stores.append(store)
    finally:
        if driver:
            driver.quit()
    return stores


def fetch_page(url: str, use_selenium: bool = False) -> str:
    """Return HTML of the given URL."""
    if use_selenium:
        driver = build_driver()
        try:
            driver.get(url)
            time.sleep(1)
            return driver.page_source
        finally:
            driver.quit()
    resp = requests.get(url, headers=HEADERS, timeout=10)
    resp.raise_for_status()
    return resp.text


def find_email(store_name: str, use_selenium: bool = False) -> Optional[str]:
    """Find an email related to the given store name."""
    query = f"{store_name} 이메일"
    base_url = "https://duckduckgo.com/html/"
    resp = requests.get(
        base_url,
        params={"q": query},
        headers=HEADERS,
        timeout=10,
    )
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    for res in soup.select("div.result__snippet"):
        email_match = EMAIL_RE.search(res.get_text())
        if email_match:
            return email_match.group(0)

    first = soup.select_one("a.result__a")
    if first and first.has_attr("href"):
        page_html = fetch_page(first["href"], use_selenium)
        match = EMAIL_RE.search(page_html)
        if match:
            return match.group(0)
    return None


def create_dataframe(rows: List[StoreInfo], keyword: str) -> pd.DataFrame:
    """Return DataFrame from scraped store information."""
    data = [
        {
            "브랜드명": r.name,
            "키워드": keyword,
            "이메일": r.email or "",
            "수집시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        for r in rows
    ]
    return pd.DataFrame(data)


def main() -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(description="Email scraper")
    parser.add_argument("keyword", nargs="?", default="실버 목걸이")
    parser.add_argument("--selenium", action="store_true")
    parser.add_argument("--limit", type=int, default=None)
    args = parser.parse_args()

    stores = fetch_serp(args.keyword, use_selenium=args.selenium)
    if args.limit:
        stores = stores[: args.limit]

    rows: List[StoreInfo] = []
    for store in stores:
        email = find_email(store, use_selenium=args.selenium)
        rows.append(StoreInfo(store, email))

    df = create_dataframe(rows, args.keyword)
    out_path = "output.xlsx"
    df.to_excel(out_path, index=False, engine="openpyxl")
    print(out_path)


if __name__ == "__main__":
    main()
    # Demo run
    if False:  # pragma: no cover - example usage
        from pprint import pprint

        stores_demo = fetch_serp("실버 목걸이")[:5]
        emails_demo = [find_email(s) for s in stores_demo]
        pprint(list(zip(stores_demo, emails_demo)))
