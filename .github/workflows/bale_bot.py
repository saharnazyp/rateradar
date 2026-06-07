"""
ربات بله — MenuRadar
دکمه گزارش رقبا → ارسال آخرین Excel از cache
"""

import os
import glob
import asyncio
import datetime
import requests

BALE_TOKEN = os.environ["BALE_BOT_TOKEN"]
BASE_URL = f"https://tapi.bale.ai/bot{BALE_TOKEN}"

# آیدی عددی کاربران مجاز (از @userinfobot در بله بگیر)
ALLOWED_USERS = [
    int(uid.strip())
    for uid in os.environ.get("ALLOWED_USER_IDS", "").split(",")
    if uid.strip()
]

REPORT_DIR = "reports"  # محل ذخیره Excel‌های تولیدشده توسط GitHub Actions

# ─────────────────────────────────────────────
# Telegram / Bale API helpers
# ─────────────────────────────────────────────

def api(method: str, **kwargs):
    resp = requests.post(f"{BASE_URL}/{method}", json=kwargs, timeout=30)
    return resp.json()

def send_message(chat_id: int, text: str, reply_markup=None):
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    return api("sendMessage", **payload)

def send_document(chat_id: int, file_path: str, caption: str = ""):
    url = f"{BASE_URL}/sendDocument"
    with open(file_path, "rb") as f:
        resp = requests.post(url, data={
            "chat_id": chat_id,
            "caption": caption,
            "parse_mode": "Markdown",
        }, files={"document": (os.path.basename(file_path), f)}, timeout=60)
    return resp.json()

def answer_callback(callback_id: str, text: str = ""):
    api("answerCallbackQuery", callback_query_id=callback_id, text=text)

# ─────────────────────────────────────────────
# Keyboard
# ─────────────────────────────────────────────

MAIN_KEYBOARD = {
    "inline_keyboard": [
        [{"text": "📊 گزارش رقبا", "callback_data": "report"}],
        [{"text": "📅 آخرین آپدیت", "callback_data": "last_update"}],
        [{"text": "ℹ️ راهنما", "callback_data": "help"}],
    ]
}

# ─────────────────────────────────────────────
# Handlers
# ─────────────────────────────────────────────

def get_latest_report():
    """آخرین فایل Excel رو پیدا می‌کنه"""
    os.makedirs(REPORT_DIR, exist_ok=True)
    files = sorted(glob.glob(f"{REPORT_DIR}/MenuRadar_*.xlsx"))
    return files[-1] if files else None

def handle_report(chat_id: int):
    report = get_latest_report()
    if not report:
        send_message(chat_id, "⚠️ هنوز هیچ گزارشی آماده نیست.\nگزارش بعدی توسط سیستم خودکار ارسال می‌شه.")
        return
    # date from filename: MenuRadar_2024-01-15.xlsx
    try:
        date_part = os.path.basename(report).replace("MenuRadar_", "").replace(".xlsx", "")
        caption = f"📊 *گزارش MenuRadar*\n📅 تاریخ تولید: `{date_part}`\n\nمقایسه قیمت منو SPO با رقبا\nهر کانسپت در یک شیت جداگانه"
    except Exception:
        caption = "📊 گزارش MenuRadar"
    send_document(chat_id, report, caption)

def handle_last_update(chat_id: int):
    report = get_latest_report()
    if not report:
        send_message(chat_id, "⚠️ هنوز هیچ گزارشی آماده نیست.")
        return
    try:
        date_part = os.path.basename(report).replace("MenuRadar_", "").replace(".xlsx", "")
        mtime = os.path.getmtime(report)
        dt = datetime.datetime.fromtimestamp(mtime).strftime("%Y/%m/%d — %H:%M")
        send_message(chat_id,
            f"📅 *آخرین آپدیت داده‌ها*\n\n"
            f"تاریخ گزارش: `{date_part}`\n"
            f"زمان دریافت فایل: `{dt}`\n\n"
            f"_گزارش بعدی ۱۰ روز دیگه به‌روز می‌شه_"
        )
    except Exception as e:
        send_message(chat_id, f"خطا: {e}")

def handle_help(chat_id: int):
    send_message(chat_id,
        "📌 *راهنمای ربات MenuRadar*\n\n"
        "📊 *گزارش رقبا* — آخرین فایل Excel مقایسه قیمت رو دریافت کن\n"
        "📅 *آخرین آپدیت* — ببین داده‌ها چه زمانی به‌روز شدن\n\n"
        "⏱ گزارش‌ها هر ۱۰ روز یک‌بار به‌صورت خودکار آپدیت می‌شن\n"
        "📢 گزارش جدید در کانال تلگرام هم ارسال می‌شه"
    )

def handle_start(chat_id: int, first_name: str):
    send_message(chat_id,
        f"سلام {first_name}! 👋\n\n"
        f"به ربات *MenuRadar* خوش اومدی 📊\n"
        f"از دکمه‌های زیر استفاده کن:",
        reply_markup=MAIN_KEYBOARD
    )

def handle_unauthorized(chat_id: int):
    send_message(chat_id, "⛔ شما دسترسی به این ربات ندارید.")

# ─────────────────────────────────────────────
# Polling loop
# ─────────────────────────────────────────────

def process_update(update: dict):
    # Callback query (دکمه‌ها)
    if "callback_query" in update:
        cq = update["callback_query"]
        user_id = cq["from"]["id"]
        chat_id = cq["message"]["chat"]["id"]
        data = cq.get("data", "")
        answer_callback(cq["id"])

        if ALLOWED_USERS and user_id not in ALLOWED_USERS:
            handle_unauthorized(chat_id)
            return

        if data == "report":
            handle_report(chat_id)
        elif data == "last_update":
            handle_last_update(chat_id)
        elif data == "help":
            handle_help(chat_id)

    # پیام متنی
    elif "message" in update:
        msg = update["message"]
        user_id = msg["from"]["id"]
        chat_id = msg["chat"]["id"]
        first_name = msg["from"].get("first_name", "")
        text = msg.get("text", "")

        if ALLOWED_USERS and user_id not in ALLOWED_USERS:
            handle_unauthorized(chat_id)
            return

        if text in ["/start", "start"]:
            handle_start(chat_id, first_name)
        else:
            send_message(chat_id,
                "از دکمه‌های زیر استفاده کن 👇",
                reply_markup=MAIN_KEYBOARD
            )

def run_polling():
    print("🤖 Bale bot polling started...")
    offset = 0
    while True:
        try:
            resp = requests.get(f"{BASE_URL}/getUpdates", params={
                "offset": offset,
                "timeout": 30,
                "limit": 100,
            }, timeout=40)
            data = resp.json()
            if data.get("ok"):
                for update in data.get("result", []):
                    try:
                        process_update(update)
                    except Exception as e:
                        print(f"Error processing update: {e}")
                    offset = update["update_id"] + 1
        except requests.exceptions.Timeout:
            pass
        except Exception as e:
            print(f"Polling error: {e}")
            import time; time.sleep(5)

if __name__ == "__main__":
    run_polling()
