import pandas as pd
import requests
from bs4 import BeautifulSoup
from rapidfuzz import fuzz
from playwright.sync_api import sync_playwright
from datetime import datetime
import os

# -------------------------
# تنظیمات
# -------------------------

MATCH_THRESHOLD = 70

OUTPUT_COLUMNS = [
    "دسته‌بندی",
    "نام آیتم (من)",
    "قیمت من (تومان)",
    "نام آیتم رقیب",
    "قیمت رقیب (تومان)",
    "نام رقیب",
    "اختلاف قیمت (تومان)",
    "وضعیت",
]


# -------------------------
# ابزار کمکی
# -------------------------

def normalize_columns(df):
    df.columns = df.columns.str.strip().str.replace("\u200c", "")
    return df


def normalize_text(t):

    if not isinstance(t, str):
        return ""

    t = t.strip()
    t = t.replace("ي", "ی")
    t = t.replace("ك", "ک")

    return t.lower()


def safe_price(v):

    try:
        return int(float(v))
    except:
        return None


# -------------------------
# خواندن منوی خودت
# -------------------------

def load_my_menu():

    df = pd.read_excel("my_menu.xlsx")

    df = normalize_columns(df)

    df = df.rename(columns={
        "نام فارسي": "item",
        "فی واحد  با ارزش افزوده - ریال": "price",
        "کانسپت": "category"
    })

    df["price"] = df["price"].apply(safe_price)

    return df[["category", "item", "price"]]


# -------------------------
# خواندن رقبا
# -------------------------

def load_competitors():

    df = pd.read_excel("competitors.xlsx")

    df = normalize_columns(df)

    df = df.rename(columns={
        "نام برند": "brand",
        "منبع": "url"
    })

    return df


# -------------------------
# اسکرپ سایت معمولی
# -------------------------

def scrape_normal_site(url):

    items = []

    try:

        r = requests.get(url, timeout=20)

        soup = BeautifulSoup(r.text, "html.parser")

        for tag in soup.find_all(["h1","h2","h3","li","span","p","div"]):

            text = tag.get_text(strip=True)

            if 3 < len(text) < 80:
                items.append(text)

    except:
        pass

    return list(set(items))


# -------------------------
# اسکرپ Snappfood
# -------------------------

def scrape_snappfood(url):

    items = []

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=True)

        page = browser.new_page()

        page.goto(url, timeout=60000)

        page.wait_for_timeout(5000)

        for _ in range(8):
            page.mouse.wheel(0, 5000)
            page.wait_for_timeout(800)

        texts = page.locator("text=").all_inner_texts()

        for t in texts:

            t = t.strip()

            if 3 < len(t) < 80:
                items.append(t)

        browser.close()

    return list(set(items))


# -------------------------
# انتخاب نوع اسکرپ
# -------------------------

def scrape_menu(url):

    if "snappfood" in url.lower():
        return scrape_snappfood(url)

    return scrape_normal_site(url)


# -------------------------
# fuzzy match
# -------------------------

def find_match(my_item, competitor_items):

    best_score = 0
    best_item = None

    my_norm = normalize_text(my_item)

    for item in competitor_items:

        score = fuzz.token_sort_ratio(my_norm, normalize_text(item))

        if score > best_score:
            best_score = score
            best_item = item

    if best_score >= MATCH_THRESHOLD:
        return best_item

    return None


# -------------------------
# ساخت ردیف خروجی
# -------------------------

def build_row(category, my_item, my_price, comp_item, comp_price, brand):

    if comp_item is None:

        return {
            "دسته‌بندی": category,
            "نام آیتم (من)": my_item,
            "قیمت من (تومان)": my_price,
            "نام آیتم رقیب": "—",
            "قیمت رقیب (تومان)": "—",
            "نام رقیب": "—",
            "اختلاف قیمت (تومان)": "—",
            "وضعیت": "⚪ بدون تطابق",
        }

    diff = "—"

    if comp_price and my_price:
        diff = comp_price - my_price

    if diff == "—":
        status = "⚪ بدون تطابق"
    elif diff > 0:
        status = "🔴 گران‌تر"
    elif diff < 0:
        status = "🟢 ارزان‌تر"
    else:
        status = "⚪ برابر"

    return {
        "دسته‌بندی": category,
        "نام آیتم (من)": my_item,
        "قیمت من (تومان)": my_price,
        "نام آیتم رقیب": comp_item,
        "قیمت رقیب (تومان)": comp_price if comp_price else "—",
        "نام رقیب": brand,
        "اختلاف قیمت (تومان)": diff,
        "وضعیت": status,
    }


# -------------------------
# موتور اصلی
# -------------------------

def run():

    print("Loading menus...")

    my_menu = load_my_menu()

    competitors = load_competitors()

    competitor_data = {}

    for _, row in competitors.iterrows():

        brand = row["brand"]
        url = row["url"]

        print("Scraping:", brand)

        competitor_data[brand] = scrape_menu(url)

        print("Items found:", len(competitor_data[brand]))

    results = []

    for _, row in my_menu.iterrows():

        category = row["category"]
        my_item = row["item"]
        my_price = row["price"]

        matched = False

        for brand, items in competitor_data.items():

            match = find_match(my_item, items)

            if match:

                results.append(
                    build_row(category, my_item, my_price, match, None, brand)
                )

                matched = True
                break

        if not matched:

            results.append(
                build_row(category, my_item, my_price, None, None, None)
            )

    df = pd.DataFrame(results, columns=OUTPUT_COLUMNS)

    if not os.path.exists("outputs"):
        os.makedirs("outputs")

    filename = f"outputs/MenuRadar_{datetime.today().date()}.xlsx"

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:

        df.to_excel(writer, sheet_name="Report", startrow=1, index=False)

        ws = writer.sheets["Report"]

        ws["A1"] = "📊 مقایسه منو — MenuRadar"

    print("Report saved:", filename)


# -------------------------
# اجرا
# -------------------------

if __name__ == "__main__":
    run()
