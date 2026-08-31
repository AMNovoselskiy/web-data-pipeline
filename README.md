# Web Data Pipeline

Turn-key data extraction: **scrape → clean → validated Excel report**, in one command.

This demo runs against [books.toscrape.com](https://books.toscrape.com/) — a public
sandbox site built specifically for scraping practice — and produces a formatted
Excel workbook: raw data, per-category summary with a chart, and a data-quality sheet.

The same pipeline structure adapts to product catalogs, directories, listings,
price monitoring — any structured public source.

## What the client gets

Not a script — **data ready to use**:

- `Data` sheet — typed, deduplicated rows (title, category, price, rating, stock, URL)
- `Summary` sheet — aggregates per category + chart
- `Quality` sheet — how many raw records came in, what was dropped and why

See [`examples/report.xlsx`](examples/report.xlsx) for real output.

## Run it

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python pipeline.py --out report.xlsx --max-pages-per-category 1
```

Typical run: ~50 requests, ~500 rows, 30 seconds.

## Design notes

- **Polite crawling**: single session, honest User-Agent, configurable delay,
  retries with exponential backoff. Scraping is only done where it is allowed.
- **Quality is reported, not hidden**: every transformation that can fail is
  counted (`duplicates`, `bad price`, `bad rating`) and shipped with the data.
- **Small and dependency-light**: `requests`, `beautifulsoup4`, `openpyxl`.
  No framework to fight with when requirements change.

## Structure

| File | Role |
|---|---|
| `scraper.py` | crawling and HTML parsing |
| `cleaner.py` | typing, validation, dedup + quality counters |
| `report.py` | Excel workbook: formatting, aggregates, chart |
| `pipeline.py` | CLI entry point |
