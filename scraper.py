"""
Book Scraper -> Excel report
----------------------------
Collects every book from https://books.toscrape.com (a website made for
scraping practice) and saves the results into a formatted Excel file with
a summary sheet and a chart.

Usage:
    python scraper.py
    python scraper.py --pages 3            # only the first 3 pages (quick test)
    python scraper.py --output my_file.xlsx
"""

import argparse
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

BASE_URL = "https://books.toscrape.com/"

# The site stores the star rating as a word in the CSS class, e.g. "star-rating Three"
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


# ---------------------------------------------------------------------------
# 1) DOWNLOADING
# ---------------------------------------------------------------------------
def get_page(session, url):
    """Download one page and return it as a BeautifulSoup object."""
    response = session.get(url, timeout=15)
    response.raise_for_status()          # stop with an error if the page failed (404, 500...)
    response.encoding = "utf-8"          # the site uses £, so make sure it's read correctly
    return BeautifulSoup(response.text, "html.parser")


# ---------------------------------------------------------------------------
# 2) EXTRACTING DATA FROM ONE PAGE
# ---------------------------------------------------------------------------
def parse_books(soup, page_url):
    """Find all books on one listing page and return them as a list of dicts."""
    books = []
    for card in soup.select("article.product_pod"):       # every book is one <article>
        link = card.h3.a
        title = link["title"]                              # full title (the visible text is shortened)
        url = urljoin(page_url, link["href"])              # turn relative link into full URL

        price_text = card.select_one("p.price_color").text  # e.g. "£51.77"
        price = float(price_text.replace("£", "").replace("Â", "").strip())

        rating_classes = card.select_one("p.star-rating")["class"]   # ["star-rating", "Three"]
        rating = next((RATING_WORDS[c] for c in rating_classes if c in RATING_WORDS), None)

        in_stock = "In stock" in card.select_one("p.availability").text

        books.append({
            "Title": title,
            "Price (£)": price,
            "Rating (1-5)": rating,
            "In stock": "Yes" if in_stock else "No",
            "URL": url,
        })
    return books


def find_next_page(soup, page_url):
    """Return the URL of the next page, or None if this is the last page."""
    next_link = soup.select_one("li.next a")
    return urljoin(page_url, next_link["href"]) if next_link else None


# ---------------------------------------------------------------------------
# 3) GOING THROUGH ALL PAGES
# ---------------------------------------------------------------------------
def scrape(start_url, max_pages=None, delay=0.5):
    all_books = []
    url = start_url
    page_no = 0

    with requests.Session() as session:
        session.headers["User-Agent"] = "Mozilla/5.0 (portfolio scraper demo)"
        while url and (max_pages is None or page_no < max_pages):
            page_no += 1
            soup = get_page(session, url)
            books = parse_books(soup, url)
            all_books.extend(books)
            print(f"Page {page_no}: {len(books)} books (total {len(all_books)})")

            url = find_next_page(soup, url)
            time.sleep(delay)            # be polite – don't overload the website

    return all_books


# ---------------------------------------------------------------------------
# 4) SAVING TO EXCEL
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_header(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"               # header stays visible when scrolling


def save_to_excel(books, filename):
    wb = Workbook()

    # --- Sheet 1: all books ---
    ws = wb.active
    ws.title = "Books"
    columns = list(books[0].keys())
    ws.append(columns)
    for book in books:
        ws.append([book[c] for c in columns])

    style_header(ws)
    ws.auto_filter.ref = ws.dimensions   # filter arrows on every column
    for col, width in zip("ABCDE", [60, 12, 13, 10, 70]):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "£#,##0.00"
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):   # clickable links
        row[0].hyperlink = row[0].value
        row[0].font = Font(color="0563C1", underline="single")

    # --- Sheet 2: summary ---
    summary = wb.create_sheet("Summary")
    prices = [b["Price (£)"] for b in books]
    summary.append(["Metric", "Value"])
    summary.append(["Total books", len(books)])
    summary.append(["Average price (£)", round(sum(prices) / len(prices), 2)])
    summary.append(["Cheapest (£)", min(prices)])
    summary.append(["Most expensive (£)", max(prices)])
    summary.append(["In stock", sum(b["In stock"] == "Yes" for b in books)])
    table_start = summary.max_row + 2    # leave one empty row between the two tables
    summary.cell(row=table_start, column=1, value="Rating")
    summary.cell(row=table_start, column=2, value="Number of books")
    summary.cell(row=table_start, column=3, value="Average price (£)")
    for stars in range(1, 6):
        group = [b["Price (£)"] for b in books if b["Rating (1-5)"] == stars]
        avg = round(sum(group) / len(group), 2) if group else 0
        summary.append([f"{stars} ★", len(group), avg])
    table_end = summary.max_row

    style_header(summary)
    for cell in summary[table_start]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 18
    summary.column_dimensions["C"].width = 18

    # --- Chart: number of books per rating ---
    chart = BarChart()
    chart.title = "Books by rating"
    chart.y_axis.title = "Number of books"
    chart.x_axis.title = "Rating"
    chart.legend = None
    chart.x_axis.delete = False          # make axes visible in newer Excel versions
    chart.y_axis.delete = False
    data = Reference(summary, min_col=2, min_row=table_start, max_row=table_end)
    labels = Reference(summary, min_col=1, min_row=table_start + 1, max_row=table_end)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.width, chart.height = 16, 8
    summary.add_chart(chart, "E2")

    wb.save(filename)


# ---------------------------------------------------------------------------
# 5) RUN
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com into Excel")
    parser.add_argument("--pages", type=int, default=None, help="max pages to scrape (default: all)")
    parser.add_argument("--output", default="books.xlsx", help="output Excel file name")
    parser.add_argument("--url", default=BASE_URL, help="start URL")
    args = parser.parse_args()

    books = scrape(args.url, max_pages=args.pages)
    if not books:
        print("No books found.")
        return
    save_to_excel(books, args.output)
    print(f"\nDone! Saved {len(books)} books to {args.output}")


if __name__ == "__main__":
    main()
