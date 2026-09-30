"""Demo: scrape all 1,000 books from books.toscrape.com (a public practice site) into Excel.

Collects title, price, rating, stock, category and URL from 50 listing pages
plus each detail page, then writes a formatted .xlsx and a .csv.
"""
import csv
import re
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "https://books.toscrape.com/catalogue/"
RATING = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
s = requests.Session()
s.headers["User-Agent"] = "Mozilla/5.0 (demo scraper)"


def get(url):
    for attempt in range(3):
        try:
            r = s.get(url, timeout=20)
            r.raise_for_status()
            r.encoding = "utf-8"
            return BeautifulSoup(r.text, "html.parser")
        except requests.RequestException:
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"failed: {url}")


def listing(page):
    soup = get(f"{BASE}page-{page}.html")
    return [urljoin(f"{BASE}page-{page}.html", a["href"]) for a in soup.select("article.product_pod h3 a")]


def detail(url):
    soup = get(url)
    table = {r.th.text: r.td.text for r in soup.select("table.table tr")}
    stock = int(re.search(r"\d+", table.get("Availability", "0")).group() if re.search(r"\d+", table.get("Availability", "")) else 0)
    return {
        "Title": soup.h1.text.strip(),
        "Category": soup.select("ul.breadcrumb li a")[-1].text.strip(),
        "Price (GBP)": float(table["Price (incl. tax)"].lstrip("£Â")),
        "Rating (1-5)": RATING[soup.select_one("p.star-rating")["class"][1]],
        "In stock": stock,
        "UPC": table["UPC"],
        "URL": url,
    }


def main():
    t0 = time.time()
    with ThreadPoolExecutor(8) as ex:
        urls = [u for page in ex.map(listing, range(1, 51)) for u in page]
        rows = list(ex.map(detail, urls))
    rows.sort(key=lambda r: (r["Category"], r["Title"]))
    cols = list(rows[0])
    with open("books.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, cols)
        w.writeheader()
        w.writerows(rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "Books"
    ws.append(cols)
    for r in rows:
        ws.append([r[c] for c in cols])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F6F43")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, width in enumerate([48, 18, 12, 12, 10, 18, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for c in ws["C"][1:]:
        c.number_format = "£#,##0.00"
    wb.save("books.xlsx")
    print(f"{len(rows)} books, {len({r['Category'] for r in rows})} categories, {time.time() - t0:.0f}s")


if __name__ == "__main__":
    main()
