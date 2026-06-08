# -*- coding: utf-8 -*-
import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
import json
import os
from urllib.parse import urlparse
from datetime import date

# ==========================================================
# CONFIG
# ==========================================================

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
}

DATA_DIR = "data"
REPORT_DIR = "reports"

MY_MENU_FILE = os.path.join(DATA_DIR, "my_menu.xlsx")
COMPETITORS_FILE = os.path.join(DATA_DIR, "competitors.xlsx")

MATCH_THRESHOLD = 0.25

CONCEPT_MAP = {
    "کافه": "لیبرو",
    "کباب ایرانی": "اسپیئو رویال",
    "تایلندی و ایتالیایی": "روماتای",
    "پیترا و سوخاری": "چیکن فکتوری",
    "برگر": "برگر فکتوری",
    "غذای ایرانی": "اسپیئو اکسپرس",
    "شیکبار": "لفو",
}

# ==========================================================
# UTILS
# ==========================================================

def safe_int(x):
    try:
        return int(float(x))
    except:
        return 0

def parse_price(text):
    text = re.sub(r"[^\d]", "", str(text))
    return safe_int(text) if text else None

def normalize(text):
    text = str(text).lower().strip()
    text = re.sub(r"[‌\s]+", " ", text)
    return text

# ==========================================================
# LOAD DATA
# ==========================================================

def load_my_menu():
    if not os.path.exists(MY_MENU_FILE):
        print("❌ my_menu.xlsx not found")
        return pd.DataFrame()

    df = pd.read_excel(MY_MENU_FILE, sheet_name=0)

    required_cols = ["نام فارسي", "کانسپت", "فی واحد  با ارزش افزوده - ریال"]
    for col in required_cols:
        if col not in df.columns:
            print(f"❌ Column missing in my_menu: {col}")
            return pd.DataFrame()

    df = df[df["فی واحد  با ارزش افزوده - ریال"] > 100]
    df = df.dropna(subset=["نام فارسي", "کانسپت"])

    return df


def load_competitors():
    if not os.path.exists(COMPETITORS_FILE):
        print("⚠ competitors.xlsx not found")
        return {}

    xl = pd.read_excel(COMPETITORS_FILE, sheet_name=None)
    result = {}

    for sheet_name, df in xl.items():

        if sheet_name not in CONCEPT_MAP:
            print(f"⚠ Skipping sheet not in concept map: {sheet_name}")
            continue

        if df.shape[1] < 2:
            print(f"⚠ Sheet has less than 2 columns: {sheet_name}")
            continue

        df = df.iloc[:, :2]
        df.columns = ["نام_برند", "URL"]

        df = df.dropna(subset=["URL"])
        df = df[df["URL"].astype(str).str.startswith("http")]

        if df.empty:
            print(f"⚠ No valid URLs in sheet: {sheet_name}")
            continue

        result[sheet_name] = df

    return result

# ==========================================================
# SCRAPER (Generic Only - Stable)
# ==========================================================

def scrape_generic(url, brand_name):
    items = []
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return items

        soup = BeautifulSoup(r.text, "html.parser")
        price_pattern = re.compile(r"\d{4,}")

        for el in soup.find_all(["div", "li", "article"]):
            text = el.get_text(" ", strip=True)
            prices = price_pattern.findall(text)

            if prices and len(text) < 200:
                price = safe_int(prices[0])
                if 10000 < price < 500000000:
                    name = re.sub(r"\d+", "", text).strip()
                    if 2 < len(name) < 80:
                        items.append({
                            "نام_رقیب": brand_name,
                            "نام_آیتم_رقیب": name[:80],
                            "قیمت_رقیب_تومان": price,
                        })

        # dedupe
        seen = set()
        unique = []
        for i in items:
            if i["نام_آیتم_رقیب"] not in seen:
                seen.add(i["نام_آیتم_رقیب"])
                unique.append(i)

        return unique[:300]

    except Exception as e:
        print(f"⚠ Scrape error {brand_name}: {e}")
        return []


def scrape_all(competitors):
    scraped = {}

    for sheet_name, df in competitors.items():
        print(f"\n📍 Scraping concept: {sheet_name}")
        all_items = []

        for _, row in df.iterrows():
            print(f"  → {row['نام_برند']}")
            items = scrape_generic(row["URL"], row["نام_برند"])
            print(f"     ✓ {len(items)} items")
            all_items.extend(items)
            time.sleep(2)

        scraped[sheet_name] = all_items
        print(f"✅ Total for {sheet_name}: {len(all_items)}")

    return scraped

# ==========================================================
# MATCHING
# ==========================================================

def match_score(a, b):
    a_words = set(normalize(a).split())
    b_words = set(normalize(b).split())
    if not a_words or not b_words:
        return 0
    return len(a_words & b_words) / len(a_words | b_words)


def find_best_match(my_item, comp_items):
    best = None
    best_score =0

    for item in comp_items:
        score = match_score(my_item, item["نام_آیتم_رقیب"])
        if score > best_score:
            best_score = score
            best = item

    if best_score >= MATCH_THRESHOLD:
        return best
    return None

# ==========================================================
# BUILD EXCEL
# ==========================================================

def build_excel(my_menu, competitors, scraped):
    from openpyxl import Workbook

    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"

    ws_summary.append(["Concept", "My Item", "My Price", "Competitor Item", "Comp Price", "Brand", "Diff"])

    any_data = False

    for concept_sheet, my_concept in CONCEPT_MAP.items():

        if concept_sheet not in competitors:
            continue

        my_items = my_menu[my_menu["کانسپت"] == my_con
