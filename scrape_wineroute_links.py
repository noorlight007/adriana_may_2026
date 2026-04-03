#!/usr/bin/env python3
import json
import sys
from typing import Iterable, List
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

DEFAULT_URLS = [
    "https://www.wineroute.alsace/wineries/",
    # "https://www.wineroute.alsace/wineries/page/2/",
]

for i in range(2, 57):  # Adjust the range as needed to cover more pages
    DEFAULT_URLS.append(f"https://www.wineroute.alsace/wineries/page/{i}/")

OUTPUT_FILE = "wineries_links.json"


def fetch_html(url: str, timeout: int = 30) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0 Safari/537.36"
        )
    }
    response = requests.get(url, headers=headers, timeout=timeout)
    response.raise_for_status()
    return response.text


def extract_links(html: str, base_url: str) -> List[str]:
    soup = BeautifulSoup(html, "html.parser")

    # Primary target: the card listing container shown in the provided HTML snippet.
    container = soup.select_one("div.wrapper-cards.wrapper.js-listing-card-container")

    # Fallbacks in case the class list or page structure changes slightly.
    if container is None:
        container = soup.select_one("div.wrapper-cards")
    if container is None:
        container = soup

    links: List[str] = []
    seen = set()

    for a_tag in container.select("a.cards-v2.card-sit[href]"):
        href = a_tag.get("href", "").strip()
        if not href:
            continue
        absolute_url = urljoin(base_url, href)
        if absolute_url not in seen:
            seen.add(absolute_url)
            links.append(absolute_url)

    return links


def collect_links(urls: Iterable[str]) -> List[str]:
    all_links: List[str] = []
    seen = set()

    for url in urls:
        html = fetch_html(url)
        links = extract_links(html, url)
        for link in links:
            if link not in seen:
                seen.add(link)
                all_links.append(link)

    return all_links


def main() -> None:
    urls = sys.argv[1:] if len(sys.argv) > 1 else DEFAULT_URLS
    links = collect_links(urls)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(links, f, ensure_ascii=False, indent=2)

    print(f"Saved {len(links)} links to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
