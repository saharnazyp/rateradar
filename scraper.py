import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import json
import os
import time
import datetime

# ─────────────────────────────────────────────
# تنظیمات
# ─────────────────────────────────────────────

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
}

MATCH_THRESHOLD = 0.25

SYNONYMS = {
    "فرایز": ["سیب زمینی", "فرنچ فرایز", "فری"],
    "برگر": ["همبرگر"],
    "چیز": ["پنیر"],
    "چیکن": ["مرغ", "chicken"],
    "سوخاری": ["فراید", "crispy"],
    "پاستا": ["ماکارونی", "پنه", "اسپاگتی"],
    "پیتزا": ["پیزا"],
    "لاته": ["latte", "لته"],
    "اسپرسو": ["espresso"],
    "کاپوچینو": ["cappuccino"],
}

# ─────────────────────────────────────────────
# خواندن فایل‌ها
# ─────────────────────────────────────────────

def load_my_menu():
    path = "data/my_menu.xlsx"
    xl = pd.ExcelFile(path)
    print(f"  Sheets in my_menu: {xl.sheet_names}")
    # اولین شیت رو بخون
    df = pd.read_excel(path, sheet_name=0)
    df.columns = df.columns.str.strip().str.replace("\u200c", "")
    print(f"  Columns: {list(df.columns)}")
    print(f"  Rows: {len(df)}")
    return df

def load_competitors():
    path = "data/competitors.xlsx"
    xl = pd.ExcelFile(path)
    print(f"  Sheets in competitors: {xl.sheet_names}")
    result = {}
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        df.columns = df.columns.str.strip()
        result[sheet] = df
        print(f"  Sheet '{sheet}': {len(df)} rows, cols: {list(df.columns)}")
    return result

# ─────────────────────────────────────────────
# Scraping
# ─────────────────────────────────────────────

def parse_price(text):
    text = re.sub(r"[^\d]", "", str(text))
    return int(text) if text else None

def scrape_url(url, brand_name):
    items = []
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            print(f"    HTTP {r.status_code} for {url}")
            return items
        soup = BeautifulSoup(r.text, "html.parser")

        # JSON-LD
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                if isinstance(data, dict) and "hasMenu" in data:
                    for section in data["hasMenu"].get("hasMenuSection", []):
                        for item in section.get("hasMenuItem", []):
                            name = item.get("name", "")
                            price = item.get("offers", {}).get("price", 0)
                            if name and price:
                                p = int(float(price))
                                items.append({
                                    "brand": brand_name,
                                    "item": name,
                                    "price_toman": p if p < 1000000 else p // 10,
                                })
            except:
                pass

        # Generic fallback
        if not items:
            price_pat = re.compile(r"\b\d[\d,]{3,}\b")
            seen = set()
            for el in soup.find_all(["div", "li", "tr", "article", "section"]):
                text = el.get_text(separator=" ", strip=True)
                prices = [int(p.replace(",", "")) for p in price_pat.findall(text)]
                valid = [p for p in prices if 10000 < p < 500_000_000]
                if valid and len(text) < 400:
                    name = re.sub(r"[\d,،٬]+", "", text).strip()
                    name = re.sub(r"\s+", " ", name).strip()[:80]
                    if 2 < len(name) < 80 and name not in seen:
                        seen.add(name)
                        p = valid[0]
                        items.append({
                            "brand": brand_name,
                            "item": name,
                            "price_toman": p if p < 1000000 else p // 10,
                        })
    except Exception as e:
        print(f"    Error scraping {brand_name}: {e}")
    print(f"    {brand_name}: {len(items)} items")
    time.sleep(1)
    return items

# ─────────────────────────────────────────────
# Matching
# ─────────────────────────────────────────────

def normalize(text):
    text = str(text).lower().strip()
    text = text.replace("ي", "ی").replace("ك", "ک")
    text = re.sub(r"[‌\s]+", " ", text)
    for canonical, variants in SYNONYMS.items():
        for v in variants:
            text = text.replace(v.lower(), canonical.lower())
    return text

def match_score(a, b):
    wa = set(normalize(a).split())
    wb = set(normalize(b).split())
    if not wa or not wb:
        return 0
    inter = wa & wb
    union = wa | wb
    jaccard = len(inter) / len(union)
    ratio = len(inter) / min(len(wa), len(wb))
    return (jaccard + ratio) / 2

# ─────────────────────────────────────────────
# ساخت Excel خروجی
# ─────────────────────────────────────────────

def build_excel(my_df, comp_items_by_concept, my_col_map):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)

    HDR_FILL = PatternFill("solid", start_color="1F3864")
    HDR_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    SUB_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    ALT_FILL = PatternFill("solid", start_color="EBF3FB")
    NRM_FONT = Font(name="Arial", size=9)
    thin = Side(style="thin", color="BDD7EE")
    BRD = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
    RGT = Alignment(horizontal="right", vertical="center", wrap_text=True)

    all_rows = []

    name_col = my_col_map.get("name")
    price_col = my_col_map.get("price")
    concept_col = my_col_map.get("concept")
    cat_col = my_col_map.get("category")

    concepts = my_df[concept_col].dropna().unique() if concept_col else ["همه"]

    for concept in concepts:
        if concept_col:
            subset = my_df[my_df[concept_col] == concept].copy()
        else:
            subset = my_df.copy()

        comp_items = comp_items_by_concept.get(str(concept), [])

        ws = wb.create_sheet(title=str(concept)[:31])
        ws.merge_cells("A1:H1")
        ws["A1"] = f"📊 مقایسه منو — {concept}"
        ws["A1"].font = Font(bold=True, color="FFFFFF", name="Arial", size=12)
        ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CTR
        ws.row_dimensions[1].height = 28

        headers = ["دسته‌بندی", "نام آیتم (من)", "قیمت من (تومان)",
                   "نام آیتم رقیب", "قیمت رقیب (تومان)", "نام رقیب",
                   "اختلاف (تومان)", "وضعیت"]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=c, value=h)
            cell.font = SUB_FONT; cell.fill = SUB_FILL
            cell.alignment = CTR; cell.border = BRD
        ws.row_dimensions[2].height = 22

        row = 3
        for _, r in subset.iterrows():
            my_name = str(r[name_col]) if name_col else "—"
            my_price_raw = r[price_col] if price_col else None
            try:
                my_price = int(float(my_price_raw)) // 10 if my_price_raw else None
            except:
                my_price = None
            cat = str(r[cat_col]) if cat_col and cat_col in r else "—"

            best, best_score = None, 0
            for ci in comp_items:
                s = match_score(my_name, ci["item"])
                if s > best_score:
                    best_score = s
                    best = ci

            fill = ALT_FILL if row % 2 == 0 else PatternFill("solid", start_color="FFFFFF")

            if best and best_score >= MATCH_THRESHOLD:
                cp = best["price_toman"]
                diff = (my_price or 0) - cp
                status = "🟢 ارزان‌تر" if diff < -5000 else ("🔴 گران‌تر" if diff > 5000 else "🟡 مشابه")
                vals = [cat, my_name, my_price, best["item"], cp, best["brand"], diff, status]
                all_rows.append({"کانسپت": concept, "آیتم من": my_name, "قیمت من": my_price,
                                  "آیتم رقیب": best["item"], "قیمت رقیب": cp,
                                  "رقیب": best["brand"], "اختلاف": diff, "وضعیت": status})
            else:
                vals = [cat, my_name, my_price, "—", "—", "—", "—", "⚪ بدون تطابق"]

            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=c, value=v)
                cell.font = NRM_FONT; cell.fill = fill
                cell.alignment = RGT; cell.border = BRD
            row += 1

        for i, w in enumerate([18, 35, 18, 35, 18, 22, 20, 14], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.sheet_view.rightToLeft = True

    # شیت خلاصه
    if all_rows:
        ws0 = wb.create_sheet("خلاصه کل", 0)
        ws0.merge_cells("A1:H1")
        ws0["A1"] = "📋 خلاصه کل — MenuRadar"
        ws0["A1"].font = Font(bold=True, color="FFFFFF", name="Arial", size=13)
        ws0["A1"].fill = HDR_FILL
        ws0["A1"].alignment = CTR
        for c, h in enumerate(["کانسپت","آیتم من","قیمت من","آیتم رقیب","قیمت رقیب","رقیب","اختلاف","وضعیت"], 1):
            cell = ws0.cell(row=2, column=c, value=h)
            cell.font = SUB_FONT; cell.fill = SUB_FILL; cell.alignment = CTR; cell.border = BRD
        for r, item in enumerate(all_rows, 3):
            fill = ALT_FILL if r % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
            for c, v in enumerate([item["کانسپت"],item["آیتم من"],item["قیمت من"],
                                    item["آیتم رقیب"],item["قیمت رقیب"],item["رقیب"],
                                    item["اختلاف"],item["وضعیت"]], 1):
                cell = ws0.cell(row=r, column=c, value=v)
                cell.font = NRM_FONT; cell.fill = fill; cell.alignment = RGT; cell.border = BRD
        for i, w in enumerate([20,35,18,35,20,22,20,14], 1):
            ws0.column_dimensions[get_column_letter(i)].width = w
        ws0.sheet_view.rightToLeft = True

    return wb

# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    print("🚀 MenuRadar starting...")

    print("\n📂 Loading my menu...")
    my_df = load_my_menu()

    print("\n📂 Loading competitors...")
    comp_sheets = load_competitors()

    # پیدا کردن ستون‌های منو
    cols = list(my_df.columns)
    print(f"\nDetected columns: {cols}")

    # ستون‌ها رو شناسایی می‌کنیم
    name_col = next((c for c in cols if "نام" in c and "فارس" in c), None) or \
               next((c for c in cols if "نام" in c), None)
    price_col = next((c for c in cols if "ارزش" in c or "قیمت" in c or "فی" in c), None)
    concept_col = next((c for c in cols if "کانسپت" in c), None)
    cat_col = next((c for c in cols if "نوع" in c or "دسته" in c or "گروه" in c), None)

    print(f"  name_col={name_col}, price_col={price_col}, concept_col={concept_col}, cat_col={cat_col}")

    my_col_map = {"name": name_col, "price": price_col, "concept": concept_col, "category": cat_col}

    # Scrape رقبا
    print("\n🌐 Scraping competitors...")
    comp_items_by_concept = {}

    for sheet_name, df in comp_sheets.items():
        print(f"\n  Concept: {sheet_name}")
        url_col = next((c for c in df.columns if "url" in c.lower() or "سایت" in c or "لینک" in c or "آدرس" in c), None)
        brand_col = next((c for c in df.columns if "برند" in c or "نام" in c), None)

        if not url_col:
            print(f"    ⚠ No URL column found in sheet '{sheet_name}', cols: {list(df.columns)}")
            continue

        all_items = []
        for _, row in df.iterrows():
            url = str(row[url_col]).strip()
            brand = str(row[brand_col]).strip() if brand_col else sheet_name
            if url.startswith("http"):
                all_items.extend(scrape_url(url, brand))

        comp_items_by_concept[sheet_name] = all_items
        print(f"  Total for {sheet_name}: {len(all_items)} items")

    # ساخت Excel
    print("\n📊 Building Excel...")
    wb = build_excel(my_df, comp_items_by_concept, my_col_map)

    os.makedirs("reports", exist_ok=True)
    date_str = datetime.date.today().strftime("%Y-%m-%d")
    out = f"reports/MenuRadar_{date_str}.xlsx"
    wb.save(out)
    print(f"✅ Saved: {out}")
    return out

if __name__ == "__main__":
    main()
