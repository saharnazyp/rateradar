"""
MenuRadar Scraper v3
- مسیر صحیح فایل‌ها از data/
- mapping دقیق کانسپت
- scraping با requests + BS4
- خروجی کامل: هر کانسپت یک شیت + خلاصه کل
"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import json
import os
import time
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# تنظیمات
# ─────────────────────────────────────────────

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "fa,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

MATCH_THRESHOLD = 0.22

# mapping: شیت رقبا → کانسپت منوی من
CONCEPT_MAP = {
    "کافه ":          "لیبرو",
    "کباب ایرانی ":   "اسپیئو رویال",
    "تایلندی و ایتالیایی": "روماتای",
    "پیترا و سوخاری ": "چیکن فکتوری ",
    "برگر":           "برگر فکتوری",
    "غذای ایرانی ":   "اسپیئو اکسپرس",
    "شیکبار":         "لفو",
}

SYNONYMS = {
    "فرایز": ["سیب زمینی", "فرنچ فرایز", "فری", "fries"],
    "برگر": ["همبرگر", "burger"],
    "مرغ": ["چیکن", "chicken", "مرغی"],
    "سوخاری": ["فراید", "crispy", "fried"],
    "پاستا": ["ماکارونی", "پنه", "اسپاگتی", "pasta"],
    "پیتزا": ["پیزا", "pizza"],
    "قهوه": ["coffee", "کافه", "لاته", "latte", "اسپرسو", "espresso", "کاپوچینو"],
    "شیک": ["میلک شیک", "milkshake", "shake"],
    "کباب": ["کبابی", "kabab"],
    "سالاد": ["salad"],
}

# ─────────────────────────────────────────────
# خواندن فایل‌ها
# ─────────────────────────────────────────────

def load_my_menu():
    df = pd.read_excel("data/my_menu.xlsx", sheet_name="لیست نهایی")
    df.columns = df.columns.str.strip().str.replace("\u200c", "")
    # فقط آیتم‌های فعال و دارای قیمت
    df = df[df["فعال"].notna()]
    df = df[df["فی واحد  با ارزش افزوده - ریال"].notna()]
    df = df[df["فی واحد  با ارزش افزوده - ریال"] > 100000]
    df = df[df["نام فارسي"].notna()]
    # حذف ردیف‌های غیر غذایی
    if "نوع کالا" in df.columns:
        df = df[~df["نوع کالا"].isin(["ظروف", "بسته بندی", "متفرقه"])]
    print(f"  ✓ My menu: {len(df)} items, {df['کانسپت'].nunique()} concepts")
    print(f"  Concepts: {list(df['کانسپت'].dropna().unique())}")
    return df

def load_competitors():
    xl = pd.ExcelFile("data/competitors.xlsx")
    result = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel("data/competitors.xlsx", sheet_name=sheet)
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=["منبع"])
        df = df[df["منبع"].str.startswith("http", na=False)]
        result[sheet] = df
        print(f"  Sheet '{sheet}': {len(df)} competitors")
    return result

# ─────────────────────────────────────────────
# Scraping
# ─────────────────────────────────────────────

def clean_price(text):
    digits = re.sub(r"[^\d]", "", str(text))
    return int(digits) if digits else None

def toman(price_rial):
    """تبدیل ریال به تومان"""
    if not price_rial:
        return None
    return price_rial // 10

def scrape_snappfood_api(url, brand):
    """اسنپ‌فود: استفاده از API مستقیم"""
    items = []
    try:
        # استخراج slug از URL
        slug_match = re.search(r"menu/([^/?#]+)", url)
        if not slug_match:
            return items
        slug = slug_match.group(1)

        # API endpoint اسنپ‌فود
        api_urls = [
            f"https://snappfood.ir/restaurant/api/v2/restaurant/details?restaurantId={slug}",
            f"https://snappfood.ir/restaurant/api/v1/restaurant/details?restaurantId={slug}",
        ]
        for api_url in api_urls:
            try:
                r = requests.get(api_url, headers=HEADERS, timeout=15)
                if r.status_code == 200:
                    data = r.json()
                    products = data.get("data", {}).get("products", [])
                    if not products:
                        # nested structure
                        for cat in data.get("data", {}).get("menu", {}).get("menuCategories", []):
                            products.extend(cat.get("products", []))
                    for p in products:
                        name = p.get("title") or p.get("name", "")
                        price = p.get("price") or p.get("priceAfterDiscount") or p.get("basePrice", 0)
                        if name and price:
                            p_int = int(price)
                            items.append({
                                "brand": brand, "item": name,
                                "price_toman": p_int if p_int < 5000000 else p_int // 10
                            })
                    if items:
                        break
            except:
                continue
    except Exception as e:
        print(f"    snappfood api error: {e}")
    return items

def scrape_generic(url, brand):
    """scraper عمومی برای سایت‌های معمولی"""
    items = []
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            print(f"    HTTP {r.status_code}")
            return items

        soup = BeautifulSoup(r.text, "html.parser")

        # ۱. JSON-LD
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                raw = script.string or ""
                data = json.loads(raw)
                if not isinstance(data, dict):
                    continue
                # schema.org Restaurant
                menu = data.get("hasMenu") or data.get("menu")
                if menu and isinstance(menu, dict):
                    for section in menu.get("hasMenuSection", []):
                        for mi in section.get("hasMenuItem", []):
                            n = mi.get("name", "")
                            p = mi.get("offers", {}).get("price", 0)
                            if n and p:
                                pv = int(float(p))
                                items.append({"brand": brand, "item": n,
                                              "price_toman": pv if pv < 5000000 else pv // 10})
            except:
                pass

        # ۲. JSON داخل script tag
        if not items:
            for script in soup.find_all("script"):
                txt = script.string or ""
                if len(txt) < 100:
                    continue
                # دنبال آرایه‌های محصولات
                for pat in [
                    r'"products"\s*:\s*(\[.{20,5000}?\])',
                    r'"items"\s*:\s*(\[.{20,5000}?\])',
                    r'"menuItems"\s*:\s*(\[.{20,5000}?\])',
                    r'"foods"\s*:\s*(\[.{20,5000}?\])',
                ]:
                    m = re.search(pat, txt, re.DOTALL)
                    if m:
                        try:
                            arr = json.loads(m.group(1))
                            for obj in arr[:100]:
                                if not isinstance(obj, dict):
                                    continue
                                n = obj.get("title") or obj.get("name") or obj.get("fa_name", "")
                                p = obj.get("price") or obj.get("basePrice") or obj.get("cost", 0)
                                if n and p:
                                    pv = int(float(p))
                                    items.append({"brand": brand, "item": str(n),
                                                  "price_toman": pv if pv < 5000000 else pv // 10})
                        except:
                            pass
                if items:
                    break

        # ۳. HTML selectors
        if not items:
            selectors = [
                (".menu-item", ".menu-item-title,.item-title,.food-title", ".price,.menu-price,.item-price"),
                (".product-card", "h3,h4,.name,.title", ".price"),
                (".food-card", ".food-name,.name", ".food-price,.price"),
                ("li.item", ".name,.title", ".price"),
                (".menu-card", ".card-title", ".card-price"),
            ]
            for container_sel, name_sel, price_sel in selectors:
                cards = soup.select(container_sel)
                if not cards:
                    continue
                for card in cards[:200]:
                    ne = card.select_one(name_sel)
                    pe = card.select_one(price_sel)
                    if ne and pe:
                        pv = clean_price(pe.get_text())
                        if pv and 10000 < pv < 999999999:
                            items.append({"brand": brand, "item": ne.get_text(strip=True)[:80],
                                          "price_toman": pv if pv < 5000000 else pv // 10})
                if items:
                    break

        # ۴. Generic fallback — هر عنصری که عدد قیمت داره
        if not items:
            price_re = re.compile(r"\b(\d[\d,]{2,})\b")
            seen_names = set()
            for tag in soup.find_all(["div","li","tr","article"], limit=2000):
                if tag.find(["div","li","tr","article"]):
                    continue  # فقط leaf nodes
                text = tag.get_text(separator=" ", strip=True)
                if len(text) > 300 or len(text) < 5:
                    continue
                prices = [int(m.group(1).replace(",","")) for m in price_re.finditer(text)]
                valid = [p for p in prices if 10000 < p < 999999999]
                if not valid:
                    continue
                name = price_re.sub("", text).strip()
                name = re.sub(r"[,،٬؛:\-/\\|]", " ", name)
                name = re.sub(r"\s+", " ", name).strip()[:80]
                if len(name) < 3 or name in seen_names:
                    continue
                seen_names.add(name)
                pv = valid[0]
                items.append({"brand": brand, "item": name,
                              "price_toman": pv if pv < 5000000 else pv // 10})

        # deduplicate
        seen = set()
        unique = []
        for it in items:
            k = it["item"].strip()
            if k and k not in seen:
                seen.add(k)
                unique.append(it)
        items = unique

    except Exception as e:
        print(f"    Error: {e}")

    return items

def scrape_brand(url, brand):
    print(f"    → {brand} ...")
    url = url.strip()
    if not url.startswith("http"):
        return []

    if "snappfood.ir" in url or "snapp-store.com" in url or "snapp.ir" in url:
        items = scrape_snappfood_api(url, brand)
        if not items:
            items = scrape_generic(url, brand)
    else:
        items = scrape_generic(url, brand)

    print(f"      ✓ {len(items)} items")
    time.sleep(1.5)
    return items

# ─────────────────────────────────────────────
# Matching
# ─────────────────────────────────────────────

def normalize_text(text):
    text = str(text).lower().strip()
    text = text.replace("ي", "ی").replace("ك", "ک").replace("‌", " ")
    text = re.sub(r"\s+", " ", text)
    for canonical, variants in SYNONYMS.items():
        for v in variants:
            text = text.replace(v.lower(), canonical)
    return text

def match_score(a, b):
    wa = set(normalize_text(a).split())
    wb = set(normalize_text(b).split())
    if not wa or not wb:
        return 0.0
    inter = wa & wb
    if not inter:
        return 0.0
    jaccard = len(inter) / len(wa | wb)
    overlap = len(inter) / min(len(wa), len(wb))
    return (jaccard + overlap) / 2.0

def find_best_match(my_item, comp_items):
    best_score, best = 0.0, None
    for ci in comp_items:
        s = match_score(my_item, ci["item"])
        if s > best_score:
            best_score = s
            best = ci
    return (best, best_score) if best_score >= MATCH_THRESHOLD else (None, 0.0)

# ─────────────────────────────────────────────
# Excel Builder
# ─────────────────────────────────────────────

def build_excel(my_df, comp_by_concept):
    wb = Workbook()
    wb.remove(wb.active)

    # استایل‌ها
    HDR = PatternFill("solid", start_color="1F3864")
    SUB = PatternFill("solid", start_color="2E75B6")
    ALT = PatternFill("solid", start_color="EBF3FB")
    WHT = PatternFill("solid", start_color="FFFFFF")
    GRN = PatternFill("solid", start_color="E2EFDA")
    RED = PatternFill("solid", start_color="FFDAD9")
    YLW = PatternFill("solid", start_color="FFF2CC")

    hdr_font = Font(bold=True, color="FFFFFF", name="Tahoma", size=11)
    sub_font = Font(bold=True, color="FFFFFF", name="Tahoma", size=9)
    nrm_font = Font(name="Tahoma", size=9)
    thin = Side(style="thin", color="BFBFBF")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    rgt = Alignment(horizontal="right", vertical="center", wrap_text=True)

    all_rows = []
    COLS = ["دسته‌بندی", "نام آیتم (من)", "قیمت من (تومان)",
            "نام آیتم رقیب", "قیمت رقیب (تومان)", "نام رقیب",
            "اختلاف (تومان)", "وضعیت"]
    COL_W = [20, 38, 18, 38, 18, 22, 18, 14]

    for sheet_name, my_concept in CONCEPT_MAP.items():
        comp_items = comp_by_concept.get(sheet_name, [])
        subset = my_df[my_df["کانسپت"].str.strip() == my_concept.strip()].copy()

        if subset.empty:
            print(f"  ⚠ No items for concept '{my_concept}'")
            continue

        ws = wb.create_sheet(title=my_concept.strip()[:31])
        ws.sheet_view.rightToLeft = True

        # عنوان
        ws.merge_cells(f"A1:{get_column_letter(len(COLS))}1")
        ws["A1"] = f"📊 مقایسه منو — {my_concept.strip()}  |  رقبا: {len(comp_by_concept.get(sheet_name,[]))} آیتم"
        ws["A1"].font = Font(bold=True, color="FFFFFF", name="Tahoma", size=12)
        ws["A1"].fill = HDR
        ws["A1"].alignment = ctr
        ws.row_dimensions[1].height = 30

        # هدر ستون‌ها
        for c, h in enumerate(COLS, 1):
            cell = ws.cell(row=2, column=c, value=h)
            cell.font = sub_font; cell.fill = SUB
            cell.alignment = ctr; cell.border = brd
        ws.row_dimensions[2].height = 24

        row = 3
        matched = 0
        for _, r in subset.iterrows():
            my_name = str(r["نام فارسي"]).strip()
            price_rial = r.get("فی واحد  با ارزش افزوده - ریال", 0)
            try:
                my_price_t = int(float(price_rial)) // 10
            except:
                my_price_t = 0
            cat = str(r.get("نوع کالا", "—")).strip()

            best, score = find_best_match(my_name, comp_items)

            if best:
                matched += 1
                cp = best["price_toman"]
                diff = my_price_t - cp
                if diff < -5000:
                    status = "🟢 ارزان‌تر"
                    row_fill = GRN
                elif diff > 5000:
                    status = "🔴 گران‌تر"
                    row_fill = RED
                else:
                    status = "🟡 مشابه"
                    row_fill = YLW
                vals = [cat, my_name, my_price_t, best["item"], cp, best["brand"], diff, status]
                all_rows.append({
                    "کانسپت": my_concept.strip(),
                    "دسته": cat, "آیتم_من": my_name, "قیمت_من": my_price_t,
                    "آیتم_رقیب": best["item"], "قیمت_رقیب": cp,
                    "رقیب": best["brand"], "اختلاف": diff, "وضعیت": status
                })
            else:
                row_fill = ALT if row % 2 == 0 else WHT
                vals = [cat, my_name, my_price_t, "—", "—", "—", "—", "⚪ بدون تطابق"]

            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=c, value=v)
                cell.font = nrm_font
                cell.fill = row_fill
                cell.alignment = rgt
                cell.border = brd
            row += 1

        # عرض ستون‌ها
        for i, w in enumerate(COL_W, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        print(f"  ✓ {my_concept.strip()}: {len(subset)} my items, {matched} matched")

    # ─── شیت خلاصه کل ───────────────────────────
    if all_rows:
        ws0 = wb.create_sheet("📋 خلاصه کل", 0)
        ws0.sheet_view.rightToLeft = True

        COLS0 = ["کانسپت", "دسته‌بندی", "آیتم من", "قیمت من (تومان)",
                 "آیتم رقیب", "قیمت رقیب (تومان)", "رقیب", "اختلاف (تومان)", "وضعیت"]
        COL_W0 = [18, 18, 38, 18, 38, 18, 22, 18, 14]

        ws0.merge_cells(f"A1:{get_column_letter(len(COLS0))}1")
        ws0["A1"] = f"📋 خلاصه کل MenuRadar — {datetime.date.today()}"
        ws0["A1"].font = Font(bold=True, color="FFFFFF", name="Tahoma", size=13)
        ws0["A1"].fill = HDR
        ws0["A1"].alignment = ctr
        ws0.row_dimensions[1].height = 32

        for c, h in enumerate(COLS0, 1):
            cell = ws0.cell(row=2, column=c, value=h)
            cell.font = sub_font; cell.fill = SUB
            cell.alignment = ctr; cell.border = brd
        ws0.row_dimensions[2].height = 24

        for r, item in enumerate(all_rows, 3):
            if item["وضعیت"] == "🟢 ارزان‌تر":
                rfill = GRN
            elif item["وضعیت"] == "🔴 گران‌تر":
                rfill = RED
            elif item["وضعیت"] == "🟡 مشابه":
                rfill = YLW
            else:
                rfill = ALT if r % 2 == 0 else WHT

            vals = [item["کانسپت"], item["دسته"], item["آیتم_من"], item["قیمت_من"],
                    item["آیتم_رقیب"], item["قیمت_رقیب"], item["رقیب"],
                    item["اختلاف"], item["وضعیت"]]
            for c, v in enumerate(vals, 1):
                cell = ws0.cell(row=r, column=c, value=v)
                cell.font = nrm_font; cell.fill = rfill
                cell.alignment = rgt; cell.border = brd

        for i, w in enumerate(COL_W0, 1):
            ws0.column_dimensions[get_column_letter(i)].width = w

    return wb

# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    print("🚀 MenuRadar v3 starting...")

    print("\n📂 Loading my menu...")
    my_df = load_my_menu()

    print("\n📂 Loading competitors list...")
    comp_sheets = load_competitors()

    print("\n🌐 Scraping competitors...")
    comp_by_concept = {}
    for sheet_name, df in comp_sheets.items():
        print(f"\n  📍 {sheet_name}")
        all_items = []
        for _, row in df.iterrows():
            brand = str(row.get("نام برند", "")).strip()
            url = str(row.get("منبع", "")).strip()
            if url.startswith("http"):
                all_items.extend(scrape_brand(url, brand))
        comp_by_concept[sheet_name] = all_items
        print(f"  → Total: {len(all_items)} items")

    print("\n📊 Building Excel report...")
    wb = build_excel(my_df, comp_by_concept)

    os.makedirs("reports", exist_ok=True)
    date_str = datetime.date.today().strftime("%Y-%m-%d")
    out_path = f"reports/MenuRadar_{date_str}.xlsx"
    wb.save(out_path)
    print(f"\n✅ Report saved: {out_path}")
    return out_path

if __name__ == "__main__":
    main()
