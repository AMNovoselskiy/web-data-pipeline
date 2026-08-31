"""End-to-end run: scrape -> clean -> Excel report.

Usage:
    python pipeline.py --out examples/report.xlsx --max-pages-per-category 1
"""

import argparse
import sys
import time

from cleaner import clean
from report import build_report
from scraper import Scraper


def main() -> int:
    parser = argparse.ArgumentParser(description="books.toscrape.com -> Excel report")
    parser.add_argument("--out", default="report.xlsx", help="output .xlsx path")
    parser.add_argument("--max-pages-per-category", type=int, default=1,
                        help="listing pages to crawl per category (20 books each)")
    parser.add_argument("--delay", type=float, default=0.3,
                        help="seconds between requests")
    args = parser.parse_args()

    started = time.time()
    scraper = Scraper(delay=args.delay)

    categories = scraper.categories()
    print("Categories found: %d" % len(categories))

    raw = []
    for index, category in enumerate(categories, start=1):
        books = list(scraper.category_books(category, args.max_pages_per_category))
        raw.extend(books)
        print("  [%2d/%d] %-24s %3d books" % (index, len(categories), category["name"], len(books)))

    rows, quality = clean(raw)
    build_report(rows, quality, args.out)

    print("Done in %.1fs, %d requests, %d clean rows -> %s"
          % (time.time() - started, scraper.requests_made, quality["clean"], args.out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
