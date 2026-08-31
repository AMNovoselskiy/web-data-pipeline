"""Builds the Excel deliverable: a Data sheet, a Summary sheet and a chart."""

from collections import defaultdict
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def _style_header(sheet, columns_count: int) -> None:
    for col in range(1, columns_count + 1):
        cell = sheet.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    sheet.freeze_panes = "A2"


def _fit_columns(sheet, widths: List[int]) -> None:
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def build_report(rows: List[Dict], quality: Dict[str, int], out_path: str) -> None:
    workbook = Workbook()

    data = workbook.active
    data.title = "Data"
    data.append(["Title", "Category", "Price", "Rating", "In stock", "URL"])
    for row in sorted(rows, key=lambda r: (r["category"], r["title"])):
        data.append(
            [row["title"], row["category"], row["price"], row["rating"],
             "yes" if row["in_stock"] else "no", row["url"]]
        )
    _style_header(data, 6)
    _fit_columns(data, [52, 24, 10, 8, 9, 60])
    for row_cells in data.iter_rows(min_row=2, min_col=3, max_col=3):
        row_cells[0].number_format = "0.00"

    by_category = defaultdict(list)
    for row in rows:
        by_category[row["category"]].append(row)

    summary = workbook.create_sheet("Summary")
    summary.append(["Category", "Books", "Avg price", "Min price", "Max price", "4-5 star share"])
    for category in sorted(by_category, key=lambda c: -len(by_category[c])):
        items = by_category[category]
        prices = [item["price"] for item in items]
        top_rated = sum(1 for item in items if item["rating"] >= 4)
        summary.append(
            [category, len(items), sum(prices) / len(prices), min(prices), max(prices),
             top_rated / len(items)]
        )
    _style_header(summary, 6)
    _fit_columns(summary, [24, 8, 10, 10, 10, 14])
    last_row = summary.max_row
    for row_cells in summary.iter_rows(min_row=2, max_row=last_row, min_col=3, max_col=5):
        for cell in row_cells:
            cell.number_format = "0.00"
    for row_cells in summary.iter_rows(min_row=2, max_row=last_row, min_col=6, max_col=6):
        row_cells[0].number_format = "0%"

    chart = BarChart()
    chart.title = "Books per category (top 12)"
    chart.height = 9
    chart.width = 24
    chart.legend = None
    top = min(13, last_row)
    chart.add_data(Reference(summary, min_col=2, min_row=1, max_row=top), titles_from_data=True)
    chart.set_categories(Reference(summary, min_col=1, min_row=2, max_row=top))
    summary.add_chart(chart, "H2")

    notes = workbook.create_sheet("Quality")
    notes.append(["Metric", "Value"])
    for key in ("total_raw", "clean", "duplicates", "bad_price", "bad_rating"):
        notes.append([key.replace("_", " "), quality.get(key, 0)])
    _style_header(notes, 2)
    _fit_columns(notes, [18, 10])

    workbook.save(out_path)
