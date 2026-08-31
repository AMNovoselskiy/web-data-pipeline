"""Scraper for books.toscrape.com — a public sandbox site built for scraping practice.

Crawls every category listed on the homepage sidebar, walks its listing pages
and yields one raw record per book. Stays polite: single session, custom
User-Agent, configurable delay between requests, retries with backoff.
"""

import time
from typing import Dict, Iterator, List, Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://books.toscrape.com/"
USER_AGENT = "web-data-pipeline-demo/1.0 (portfolio project; polite crawler)"


class Scraper:
    def __init__(self, delay: float = 0.3, retries: int = 3, timeout: float = 15.0):
        self.delay = delay
        self.retries = retries
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT
        self.requests_made = 0

    def _get(self, url: str) -> BeautifulSoup:
        last_error: Optional[Exception] = None
        for attempt in range(self.retries):
            try:
                time.sleep(self.delay if self.requests_made else 0)
                response = self.session.get(url, timeout=self.timeout)
                self.requests_made += 1
                response.raise_for_status()
                return BeautifulSoup(response.text, "html.parser")
            except requests.RequestException as error:
                last_error = error
                time.sleep(self.delay * (2 ** attempt))
        raise RuntimeError("Failed to fetch %s: %s" % (url, last_error))

    def categories(self) -> List[Dict[str, str]]:
        """Category names and URLs from the homepage sidebar."""
        soup = self._get(BASE_URL)
        links = soup.select("div.side_categories ul ul a")
        return [
            {"name": link.get_text(strip=True), "url": urljoin(BASE_URL, link["href"])}
            for link in links
        ]

    def category_books(self, category: Dict[str, str], max_pages: int) -> Iterator[Dict[str, str]]:
        """Raw book records from up to `max_pages` listing pages of one category."""
        url = category["url"]
        for _ in range(max_pages):
            soup = self._get(url)
            for pod in soup.select("article.product_pod"):
                title_link = pod.select_one("h3 a")
                rating_classes = pod.select_one("p.star-rating")["class"]
                yield {
                    "title": title_link["title"],
                    "url": urljoin(url, title_link["href"]),
                    "category": category["name"],
                    "price_raw": pod.select_one("p.price_color").get_text(strip=True),
                    "rating_raw": next(c for c in rating_classes if c != "star-rating"),
                    "availability_raw": pod.select_one("p.instock.availability").get_text(strip=True),
                }
            next_link = soup.select_one("li.next a")
            if next_link is None:
                break
            url = urljoin(url, next_link["href"])
