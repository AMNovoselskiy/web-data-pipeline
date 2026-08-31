"""Turns raw scraped records into typed, validated rows.

Every transformation that can fail is counted instead of silently dropped,
so the pipeline can report data quality alongside the data itself.
"""

import re
from typing import Dict, List, Tuple

RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
PRICE_PATTERN = re.compile(r"(\d+\.\d+)")


def clean(raw_records: List[Dict[str, str]]) -> Tuple[List[Dict], Dict[str, int]]:
    """Returns (clean rows, quality counters)."""
    rows = []
    quality = {"total_raw": len(raw_records), "duplicates": 0, "bad_price": 0, "bad_rating": 0}
    seen_urls = set()

    for record in raw_records:
        if record["url"] in seen_urls:
            quality["duplicates"] += 1
            continue
        seen_urls.add(record["url"])

        price_match = PRICE_PATTERN.search(record["price_raw"])
        if price_match is None:
            quality["bad_price"] += 1
            continue

        rating = RATING_WORDS.get(record["rating_raw"])
        if rating is None:
            quality["bad_rating"] += 1
            continue

        rows.append(
            {
                "title": record["title"],
                "category": record["category"],
                "price": float(price_match.group(1)),
                "rating": rating,
                "in_stock": record["availability_raw"].lower().startswith("in stock"),
                "url": record["url"],
            }
        )

    quality["clean"] = len(rows)
    return rows, quality
