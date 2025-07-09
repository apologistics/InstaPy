from __future__ import annotations

import argparse
import datetime as dt
import re
from typing import List, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

EMAIL_RE = re.compile(r"[\w.%-]+@[\w.-]+\.[A-Za-z]{2,}")
HEADERS = {"User-Agent": "Mozilla/5.0"}


def fetch_serp(keyword: str, use_selenium: bool = False) -> List[str]:
    """Return store names from DuckDuckGo SERP titles."""
    stores: List[str] = []
    if use_selenium:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from urllib.parse import quote
        from webdriver_manager.chrome import ChromeDriverManager

        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        for page in range(2):
            start = page * 30
            url = f"https://duckduckgo.com/html/?q={quote(keyword)}&s={start}"
            driver.get(url)
            soup = BeautifulSoup(driver.page_source, "html.parser")
            for result in soup.select("div.result"):
                if "result--ad" in result.get("class", []):
                    continue
                a = result.select_one("a.result__a")
                if not a:
                    continue
                title = a.get_text()
                brand = re.split(r"[|\-]", title)[0].strip()
                if brand and brand not in stores:
                    stores.append(brand)
        driver.quit()
    else:
        for page in range(2):
            start = page * 30
            resp = requests.get(
                "https://duckduckgo.com/html/",
                params={"q": keyword, "s": start},
                headers=HEADERS,
                timeout=10,
            )
            soup = BeautifulSoup(resp.text, "html.parser")
            for result in soup.select("div.result"):
                if "result--ad" in result.get("class", []):
                    continue
                a = result.select_one("a.result__a")
                if not a:
                    continue
                title = a.get_text()
                brand = re.split(r"[|\-]", title)[0].strip()
                if brand and brand not in stores:
                    stores.append(brand)
    return stores


def find_email(store_name: str, use_selenium: bool = False) -> Optional[str]:
    """Search for an email address related to the store."""
    query = f"{store_name} 이메일"
    if use_selenium:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from urllib.parse import quote
        from webdriver_manager.chrome import ChromeDriverManager

        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        url = f"https://duckduckgo.com/html/?q={quote(query)}"
        driver.get(url)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        driver.quit()
    else:
        resp = requests.get(
            "https://duckduckgo.com/html/",
            params={"q": query},
            headers=HEADERS,
            timeout=10,
        )
        soup = BeautifulSoup(resp.text, "html.parser")

    for result in soup.select("div.result"):
        text = result.get_text(" ", strip=True)
        match = EMAIL_RE.search(text)
        if match:
            return match.group()
        a = result.select_one("a.result__a")
        if a and a.has_attr("href"):
            try:
                page = requests.get(a["href"], headers=HEADERS, timeout=10)
                match = EMAIL_RE.search(page.text)
                if match:
                    return match.group()
            except requests.RequestException:
                continue
    return None


def main(argv: Optional[List[str]] = None) -> None:
    """Entry point for command-line execution."""
    parser = argparse.ArgumentParser()
    parser.add_argument("-k", "--keyword", default="실버 목걸이")
    parser.add_argument("--selenium", action="store_true")
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("-o", "--output", default="emails.xlsx")
    args = parser.parse_args(argv)

    stores = fetch_serp(args.keyword, args.selenium)
    if args.limit is not None:
        stores = stores[: args.limit]

    data = []
    for name in stores:
        email = find_email(name, args.selenium)
        data.append(
            {
                "브랜드명": name,
                "키워드": args.keyword,
                "이메일": email or "",
                "수집시각": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )

    df = pd.DataFrame(data, columns=["브랜드명", "키워드", "이메일", "수집시각"])
    df.to_excel(args.output, index=False, engine="openpyxl")
    print(args.output)


if __name__ == "__main__":
    main(["--keyword", "실버 목걸이", "--limit", "5"])
