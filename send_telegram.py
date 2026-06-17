"""
send_telegram.py — انتشار گزارش‌های هفتگی به کانال تلگرام
هم RawMenus_*.xlsx و هم Compare_*.xlsx رو می‌فرسته
"""
import os
import glob
import datetime
import requests

TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
CHANNEL = os.environ["TELEGRAM_CHANNEL_ID"]
BASE = f"https://api.telegram.org/bot{TOKEN}"

REPORTS_DIR = "reports"


def find_latest(pattern):
    """جدیدترین فایل با الگوی مشخص"""
    files = sorted(
        glob.glob(f"{REPORTS_DIR}/{pattern}"),
        key=os.path.getmtime,
        reverse=True,
    )
    return files[0] if files else None


def extract_date(filename):
    name = os.path.basename(filename)
    parts = name.replace(".xlsx", "").split("_")
    return parts[-1] if len(parts) >= 2 else str(datetime.date.today())


def send_document(file_path, caption):
    file_name = os.path.basename(file_path)
    print(f"↑ ارسال {file_name} ...")
    with open(file_path, "rb") as f:
        resp = requests.post(
            f"{BASE}/sendDocument",
            data={
                "chat_id": CHANNEL,
                "caption": caption,
                "parse_mode": "Markdown",
            },
            files={"document": (file_name, f)},
            timeout=120,
        )
    if resp.status_code == 200 and resp.json().get("ok"):
        print(f"  ✓ {file_name} فرستاده شد")
        return True
    else:
        print(f"  ✗ خطا: {resp.text[:200]}")
        return False


def send_text(text):
    requests.post(
        f"{BASE}/sendMessage",
        json={
            "chat_id": CHANNEL,
            "text": text,
            "parse_mode": "Markdown",
        },
        timeout=30,
    )


def main():
    raw = find_latest("RawMenus_*.xlsx")
    compare = find_latest("Compare_*.xlsx")

    if not raw and not compare:
        print("❌ هیچ فایلی برای ارسال پیدا نشد")
        return

    date_str = datetime.date.today().strftime("%Y/%m/%d")

    # پیام سرتیتر
    send_text(
        f"📊 *گزارش هفتگی MenuRadar*\n"
        f"📅 {date_str}\n\n"
        f"به‌روزرسانی خودکار از سرور — این هفته دو فایل خواهید دید 👇"
    )

    # ۱. RawMenus — منو خام رقبا
    if raw:
        caption_raw = (
            f"📊 *گزارش منوی رقبا*\n"
            f"📅 تاریخ: `{extract_date(raw)}`\n\n"
            f"منو و قیمت ۴۰+ برند رقیب در ۷ کانسپت"
        )
        send_document(raw, caption_raw)

    # ۲. Compare — مقایسه با منو SPO
    if compare:
        caption_compare = (
            f"🔍 *مقایسه منوی SPO با رقبا*\n"
            f"📅 تاریخ: `{extract_date(compare)}`\n\n"
            f"هر آیتم SPO + نزدیک‌ترین match رقبا\n"
            f"🟢 ارزون‌تر  🔴 گرون‌تر  ⚪ مساوی\n"
            f"به‌علاوه آیتم‌های انحصاری رقبا"
        )
        send_document(compare, caption_compare)

    print()
    print("✅ تمام")


if __name__ == "__main__":
    main()
