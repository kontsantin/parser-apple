"""Script for counting products in categories on techno-buyer.ru.
Counts both total SKUs (with variants) and unique products (without variants)."""

import re
import requests
import sys
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup


BASE = "https://techno-buyer.ru"

CATEGORIES = [
    "/category/apple/apple-watch/",
    "/category/apple/apple-iphone/",
    "/category/apple/apple-ipad/",
    "/category/apple/apple-macbook/",
    "/category/apple/airpods/",
]

# Tokens that mark the start of variant attributes in a product URL slug.
# The base product is everything BEFORE the first match.
# Ordered so compound/longer patterns match first.
_VARIANT_TOKENS = [
    r"-bez-rustore",
    r"-bez-apple-intelligence",
    r"-sim-e-sim",
    r"-e-sim",
    r"-wi-fi-cellular",
    r"-wi-fi",
    r"-nanoteksturnoe-steklo",
    r"-matovoe-steklo",
    r"-ssd",
    r"-usb-c",
    r"-lightning",
    r"-magsafe",
    r"-kupit",
    r"-remeshok-[\w-]+",
    r"-tsveta-[\w-]+",
    r"-\d+tb",
    r"-\d+gb",
    # Compound colors (must be before simple colors)
    r"-chernyy-kosmos",
    r"-siyayushchaya-zvezda",
    r"-goluboe-nebo",
    r"-seryy-kosmos",
    r"-polunochnyi-chernyi",
    r"-polunochnyi",
    # Russian simple colors
    r"-chyornyy",
    r"-chernyy",
    r"-belyy",
    r"-siniy",
    r"-goluboy",
    r"-krasnyy",
    r"-rozovyy",
    r"-seryy",
    r"-zelenyy",
    r"-fioletovyy",
    r"-zheltyy",
    r"-oranzhevyy",
    r"-korichnevyy",
    r"-bezhevyy",
    r"-serebristyy",
    # English compound
    r"-silver",
    r"-gold",
    r"-space",
    r"-starlight",
    r"-midnight",
    # English simple colors (as standalone segments)
    r"-black\b",
    r"-green\b",
    r"-blue\b",
    r"-white\b",
    r"-red\b",
    r"-pink\b",
    r"-gray\b",
    r"-yellow\b",
    r"-purple\b",
    r"-orange\b",
    r"-brown\b",
    r"-natural\b",
    # Watch: case / body variants (appear before band info)
    r"-korpus",          # korpus-iz-titana, titanovyy-korpus, etc.
    r"-titanovyy",
    # Watch band specific
    r"-charcoal",
    r"-neon",
    r"-anchor",
    r"-bright",
    r"-light\b",
    r"-terra",
    r"-cotta",
    r"-milanese",
    r"-loop\b",
    r"-band\b",
    r"-alpine",
    r"-trail",
    r"-ocean",
    r"-titanium",
    r"-alyumini",
    # Trailing dedup number
    r"-\d+$",
]

VARIANT_RE = re.compile("|".join(_VARIANT_TOKENS), re.IGNORECASE)


def get_soup(url: str) -> BeautifulSoup:
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return BeautifulSoup(resp.text, "html.parser")


def get_product_urls(soup: BeautifulSoup) -> list[str]:
    urls = []
    for card in soup.select(".products__item-flying"):
        link = card.find("a", href=True)
        if link:
            urls.append(urljoin(BASE, link["href"]))
    return urls


def strip_variants(url: str) -> str:
    """Strip variant suffix from a product URL to get the base product URL."""
    path = urlparse(url).path.strip("/")
    segments = [s for s in path.split("/") if s]
    if not segments:
        return url

    slug = segments[-1]

    # Find first variant match and take everything before it as base
    match = VARIANT_RE.search(slug)
    if match:
        base = slug[: match.start()].rstrip("-")
    else:
        base = slug

    if base:
        segments[-1] = base
    else:
        segments.pop()

    return "/" + "/".join(segments) + "/"


def get_page_urls(soup: BeautifulSoup, base_url: str) -> list[str]:
    max_page = 1
    for link in soup.select(".pagin a"):
        href = link.get("href", "")
        m = re.search(r"[?&]page=(\d+)", href)
        if m:
            max_page = max(max_page, int(m.group(1)))
        text = link.text.strip()
        if text.isdigit():
            max_page = max(max_page, int(text))

    page_base = base_url if base_url.endswith("/") else base_url + "/"
    page_base = re.sub(r"([?&])page=\d+", "", page_base)

    return [f"{page_base}?page={p}" for p in range(2, max_page + 1)]


def get_subcategory_links(soup: BeautifulSoup) -> list[str]:
    urls, seen = [], set()
    for link in soup.select(".h-categ__link"):
        href = link.get("href")
        if href:
            full = urljoin(BASE, href)
            if full not in seen:
                seen.add(full)
                urls.append(full)
    return urls


def count_category(category_url: str) -> dict:
    print(f"\nCategory: {category_url}")

    soup = get_soup(category_url)

    all_urls = get_product_urls(soup)
    sku_page1 = len(all_urls)
    print(f"  Page 1: {sku_page1} SKUs", end="")

    for page_url in get_page_urls(soup, category_url):
        page_soup = get_soup(page_url)
        page_urls = get_product_urls(page_soup)
        all_urls.extend(page_urls)
        print(f" | p{page_url.split('page=')[-1]}: {len(page_urls)}", end="")

    unique_base = sorted(set(strip_variants(u) for u in all_urls))
    unique_count = len(unique_base)
    sku_count = len(all_urls)

    if unique_count != sku_count:
        print(f"  => {sku_count} SKUs, {unique_count} unique")

    result = {
        "url": category_url,
        "total": sku_count,
        "unique": unique_count,
        "subcategories": [],
    }

    for sub_url in get_subcategory_links(soup):
        result["subcategories"].append(count_category(sub_url))

    return result


def print_tree(result: dict, indent: int = 0, last: bool = True) -> None:
    prefix = "  " * indent
    connector = "+- " if last else "|- "
    name = result["url"].replace(BASE, "")
    u = result.get("unique", result["total"])
    t = result["total"]
    label = f"{u} unique ({t} SKUs)" if u != t else f"{u} items"
    print(f"{prefix}{connector}{name} - {label}")

    subs = result.get("subcategories", [])
    for i, sub in enumerate(subs):
        print_tree(sub, indent + 1, i == len(subs) - 1)


def main() -> None:
    if len(sys.argv) > 1:
        urls = sys.argv[1:]
    else:
        urls = CATEGORIES

    all_results = []
    grand_skus = 0
    grand_unique = 0

    for url in urls:
        full_url = url if url.startswith("http") else urljoin(BASE, url)
        result = count_category(full_url)
        all_results.append(result)
        grand_skus += result["total"]
        grand_unique += result.get("unique", result["total"])

    print("\n" + "=" * 60)
    print("RESULTS:")
    print("=" * 60)
    for result in all_results:
        print_tree(result)
    print("=" * 60)
    print(f"TOTAL SKUs (with variants): {grand_skus}")
    print(f"TOTAL UNIQUE PRODUCTS:     {grand_unique}")


if __name__ == "__main__":
    main()
    