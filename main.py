import requests
import json
import re
import sys
from collections import deque
from typing import Any
from urllib.parse import urljoin, urlparse

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


DEFAULT_URL = "https://techno-buyer.ru/iphone-17-pro-1tb-bez-rustore-oranzhevyy-2/"
OUTPUT_FILE = "techno_buyer_product_variants.json"
WAIT_SECONDS = 20


def clean_text(value: str) -> str:
	return re.sub(r"\s+", " ", value or "").strip()


def parse_price(value: str | None) -> int | None:
	if not value:
		return None
	digits = re.sub(r"[^\d]", "", value)
	return int(digits) if digits else None


def absolute_url(base_url: str, href: str | None) -> str | None:
	if not href:
		return None
	return urljoin(base_url, href)


def same_product_domain(url: str) -> bool:
	return urlparse(url).netloc == "techno-buyer.ru"


def first_existing_text(driver: webdriver.Chrome, selectors: list[str]) -> str:
	for selector in selectors:
		try:
			text = clean_text(driver.find_element(By.CSS_SELECTOR, selector).text)
		except NoSuchElementException:
			continue
		if text:
			return text
	return ""


def first_existing_attr(driver: webdriver.Chrome, selectors: list[str], attr_name: str) -> str:
	for selector in selectors:
		try:
			value = driver.find_element(By.CSS_SELECTOR, selector).get_attribute(attr_name)
		except NoSuchElementException:
			continue
		if value:
			return value
	return ""


class TechnoBuyerProductParser:
	def __init__(self, start_url: str, headless: bool = True) -> None:
		self.start_url = start_url
		self.driver = self._build_driver(headless=headless)
		self.wait = WebDriverWait(self.driver, WAIT_SECONDS)

	def _build_driver(self, headless: bool) -> webdriver.Chrome:
		options = ChromeOptions()
		options.add_argument("--window-size=1600,2200")
		options.add_argument("--disable-blink-features=AutomationControlled")
		options.add_argument("--no-sandbox")
		options.add_argument("--disable-dev-shm-usage")
		if headless:
			options.add_argument("--headless=new")
		return webdriver.Chrome(options=options)

	def close(self) -> None:
		self.driver.quit()

	def open_page(self, url: str) -> None:
		self.driver.get(url)
		self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "form#cart-form, .product__wrap")))
		self._dismiss_cookie_banner()
		self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".product-group, .product__prices")))

	def _dismiss_cookie_banner(self) -> None:
		buttons = self.driver.find_elements(By.XPATH, "//button[normalize-space()='OK'] | //a[normalize-space()='OK']")
		for button in buttons:
			if button.is_displayed() and button.is_enabled():
				try:
					button.click()
					return
				except Exception:
					return

	def collect_all_variant_urls(self) -> list[str]:
		queue: deque[str] = deque([self.start_url])
		seen: set[str] = set()

		while queue:
			url = queue.popleft()
			if url in seen:
				continue

			self.open_page(url)
			seen.add(self.driver.current_url)

			for candidate in self.extract_variant_links(self.driver.current_url):
				if candidate not in seen:
					queue.append(candidate)

		return sorted(seen)

	def extract_variant_links(self, base_url: str) -> list[str]:
		links: set[str] = {self.driver.current_url}
		elements = self.driver.find_elements(By.CSS_SELECTOR, ".product-group a.product-group__item[href]")
		for element in elements:
			href = absolute_url(base_url, element.get_attribute("href"))
			if href and same_product_domain(href):
				links.add(href)
		return sorted(links)

	def parse_current_product(self) -> dict[str, Any]:
		current_url = self.driver.current_url
		title = first_existing_text(
			self.driver,
			[
				"h1",
				".product__title",
				".product__wrap h1",
			],
		)
		article = first_existing_text(self.driver, [".product__code", ".art-1942", "[class*='art-']"])
		price_text = first_existing_text(self.driver, [".product__price", ".price.product__price"])
		old_price_text = first_existing_text(self.driver, [".product__price-old", ".compare-at-price"])
		stock_text = first_existing_text(self.driver, [".stocks__stock", ".stock__block", ".product__stocks"])
		badge = first_existing_text(self.driver, [".badges .badge", ".badge span"])

		sku_id = first_existing_attr(self.driver, ["input[name='sku_id']"], "value")
		product_id = first_existing_attr(self.driver, ["input[name='product_id']"], "value")

		return {
			"url": current_url,
			"title": title,
			"article": clean_text(article.replace("Артикул:", "")),
			"sku_id": sku_id or None,
			"product_id": product_id or None,
			"price": parse_price(price_text),
			"price_text": price_text,
			"old_price": parse_price(old_price_text),
			"old_price_text": old_price_text,
			"stock": stock_text,
			"badge": badge,
			"brand": first_existing_text(self.driver, [".p-images__brand-link", ".p-images__brand-name a"]),
			"variation_groups": self.parse_variation_groups(),
			"images": self.parse_images(current_url),
			"services": self.parse_services(),
			"features": self.parse_features(),
		}

	def parse_variation_groups(self) -> list[dict[str, Any]]:
		groups_data: list[dict[str, Any]] = []
		groups = self.driver.find_elements(By.CSS_SELECTOR, ".product-group")

		for group in groups:
			title = clean_text(self._safe_find_text(group, ".product-group__title"))
			items_data: list[dict[str, Any]] = []
			active_value = ""

			items = group.find_elements(By.CSS_SELECTOR, ".product-group__item")
			for item in items:
				classes = item.get_attribute("class") or ""
				is_active = "product-group__item--active" in classes
				href = item.get_attribute("href")
				style = clean_text(item.get_attribute("style") or "")
				text = clean_text(item.text)

				if not text and "background-color" in style:
					text = style

				item_data = {
					"value": text,
					"is_active": is_active,
					"url": absolute_url(self.driver.current_url, href),
					"style": style or None,
				}
				items_data.append(item_data)

				if is_active:
					active_value = text

			groups_data.append(
				{
					"title": title,
					"active": active_value,
					"items": items_data,
				}
			)

		return groups_data

	def parse_images(self, base_url: str) -> list[str]:
		images: list[str] = []
		seen: set[str] = set()

		anchors = self.driver.find_elements(By.CSS_SELECTOR, ".p-images__slider-item")
		for anchor in anchors:
			href = absolute_url(base_url, anchor.get_attribute("href"))
			if href and href not in seen:
				seen.add(href)
				images.append(href)

		thumbs = self.driver.find_elements(By.CSS_SELECTOR, ".p-images__dop-img")
		for image in thumbs:
			src = image.get_attribute("data-src") or image.get_attribute("src")
			src = absolute_url(base_url, src)
			if src and src not in seen:
				seen.add(src)
				images.append(src)

		return images

	def parse_services(self) -> list[dict[str, Any]]:
		services: list[dict[str, Any]] = []
		labels = self.driver.find_elements(By.CSS_SELECTOR, ".services__list label")
		for label in labels:
			try:
				checkbox = label.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
				price = checkbox.get_attribute("data-price")
			except NoSuchElementException:
				price = None

			services.append(
				{
					"name": clean_text(label.text),
					"price": int(price) if price and price.isdigit() else None,
				}
			)
		return services

	def parse_features(self) -> dict[str, str]:
		features: dict[str, str] = {}
		rows = self.driver.find_elements(By.CSS_SELECTOR, ".product__features-item")
		for row in rows:
			name = clean_text(self._safe_find_text(row, ".product__features-name")).rstrip(":")
			value = clean_text(self._safe_find_text(row, ".product__features-value"))
			if name and value:
				features[name] = value
		return features

	def _safe_find_text(self, root: WebElement, selector: str) -> str:
		try:
			return root.find_element(By.CSS_SELECTOR, selector).text
		except NoSuchElementException:
			return ""


def parse_product(start_url: str, headless: bool = True) -> dict[str, Any]:
	parser = TechnoBuyerProductParser(start_url=start_url, headless=headless)
	try:
		variant_urls = parser.collect_all_variant_urls()
		products: list[dict[str, Any]] = []

		for index, url in enumerate(variant_urls, start=1):
			parser.open_page(url)
			product_data = parser.parse_current_product()
			product_data["variant_index"] = index
			products.append(product_data)
			print(f"[{index}/{len(variant_urls)}] {product_data['title']} -> {product_data['price_text']}")

		return {
			"start_url": start_url,
			"variant_count": len(products),
			"variants": products,
		}
	finally:
		parser.close()


def main() -> None:
	start_url = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_URL
	headless = "--headed" not in sys.argv

	try:
		result = parse_product(start_url=start_url, headless=headless)
	except TimeoutException as error:
		raise SystemExit(f"Не удалось дождаться элементов страницы: {error}") from error

	with open(OUTPUT_FILE, "w", encoding="utf-8") as file:
		json.dump(result, file, ensure_ascii=False, indent=2)

	print()
	print(f"Сохранено вариантов: {result['variant_count']}")
	print(f"Файл: {OUTPUT_FILE}")


if __name__ == "__main__":
	main()
