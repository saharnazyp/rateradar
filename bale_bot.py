"""
ربات بله — MenuRadar
سیستم ورود با پسورد (بدون نیاز به User ID)
v2 — آپدیت برای الگوی RawMenus_*.xlsx
"""

import os
import glob
import datetime
import time
import requests

BALE_TOKEN = os.environ["BALE_BOT_TOKEN"]
BASE_URL = f"https://tapi.bale.ai/bot{BALE_TOKEN}"

BOT_PASSWORD = os.environ.get("BOT_PASSWORD", "spo1245")

REPORT_DIR = "reports"

authenticated_users = {}

# ─────────────────────────────────────────────
# API helpers
# ─────────────────────────────────────────────

def api(method, **kwargs):
    try:
        resp = requests.post(f"{BASE_URL}/{method}", json=kwargs, timeout=30)
        return resp.json()
    except Exception as e:
        print(f"API error ({method}): {e}")
        return {}

def send_message(chat_id, text, reply_markup=None):
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    return api("sendMessage", **payload)

def send_document(chat_id, file_path, caption=""):
    url = f"{BASE_URL}/sendDocument"
    try:
        with open(file_path, "rb") as f:
            resp = requests.post(url, data={
                "chat_id": chat_id,
                "caption": caption,
                "parse_mode": "Markdown",
            }, files={"document": (os.path.basename(file_path), f)}, timeout=60)
        return resp.json()
    except Exception as e:
        print(f"sendDocument error: {e}")
        return {}

def answer_callback(callback_id, text=""):
    api("answerCallbackQuery", callback_query_id=callback_id, text=text)

# ─────────────────────────────────────────────
# Keyboards
# ─────────────────────────────────────────────

MAIN_KEYBOARD = {
    "inline_keyboard": [
        [{"text": "📊 گزارش رقبا", "callback_data": "report"}],
        [{"text": "📅 آخرین آپدیت", "callback_data": "last_update"}],
        [{"text": "ℹ️ راهنما", "callback_data": "help"}],
    ]
}

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def get_latest_report():
    """جدیدترین گزارش RawMenus رو پیدا می‌کنه — اول بر اساس mtime، نه اسم"""
    os.makedirs(REPORT_DIR, exist_ok=True)
    # الگوی جدید: RawMenus_*.xlsx
    files = glob.glob(f"{REPORT_DIR}/RawMenus_*.xlsx")
    # اگه چیزی نبود، الگوی قدیمی رو هم چک کن (سازگاری عقب)
    if not files:
        files = glob.glob(f"{REPORT_DIR}/MenuRadar_*.xlsx")
    if not files:
        return None
    # جدیدترین بر اساس زمان فایل (mtime) — نه اسم
    return max(files, key=os.path.getmtime)

def is_authenticated(chat_id):
    return authenticated_users.get(chat_id, False)

# ─────────────────────────────────────────────
# Handlers
# ─────────────────────────────────────────────

def handle_start(chat_id, first_name):
    if is_authenticated(chat_id):
        send_message(chat_id,
            f"سلام {first_name}! 👋\nقبلاً وارد شدی. از دکمه‌های زیر استفاده کن:",
            reply_markup=MAIN_KEYBOARD
        )
    else:
        send_message(chat_id,
            f"سلام {first_name}! 👋\n\n"
            f"به ربات *MenuRadar* خوش اومدی 📊\n\n"
            f"🔐 برای ورود، رمز عبور رو وارد کن:"
        )

def handle_password(chat_id, first_name, text):
    if text.strip() == BOT_PASSWORD:
        authenticated_users[chat_id] = True
        send_message(chat_id,
            f"✅ *خوش اومدی {first_name}!*\n\nاز دکمه‌های زیر استفاده کن 👇",
            reply_markup=MAIN_KEYBOARD
        )
    else:
        send_message(chat_id, "❌ رمز اشتباهه. دوباره امتحان کن:")

def _extract_date(filename):
    """تاریخ رو از اسم فایل می‌گیره"""
    name = os.path.basename(filename)
    # RawMenus_2026-06-15.xlsx یا MenuRadar_2026-06-15.xlsx
    parts = name.replace(".xlsx", "").split("_")
    if len(parts) >= 2:
        return parts[-1]
    return "نامعلوم"

def handle_report(chat_id):
    report = get_latest_report()
    if not report:
        send_message(chat_id,
            "⚠️ هنوز هیچ گزارشی آماده نیست.\n"
            "گزارش بعدی توسط سیستم خودکار ارسال می‌شه."
        )
        return
    date_part = _extract_date(report)
    caption = (
        f"📊 *گزارش MenuRadar*\n"
        f"📅 تاریخ: `{date_part}`\n\n"
        f"مقایسه قیمت منو SPO با رقبا\n"
        f"هر کانسپت در یک شیت جداگانه"
    )
    send_message(chat_id, "⏳ در حال ارسال فایل...")
    send_document(chat_id, report, caption)

def handle_last_update(chat_id):
    report = get_latest_report()
    if not report:
        send_message(chat_id, "⚠️ هنوز هیچ گزارشی آماده نیست.")
        return
    try:
        date_part = _extract_date(report)
        mtime = os.path.getmtime(report)
        dt = datetime.datetime.fromtimestamp(mtime).strftime("%Y/%m/%d — %H:%M")
        send_message(chat_id,
            f"📅 *آخرین آپدیت*\n\n"
            f"تاریخ گزارش: `{date_part}`\n"
            f"زمان فایل: `{dt}`\n\n"
            f"_گزارش هر هفته یکشنبه به‌روز می‌شه_ ⏱"
        )
    except Exception as e:
        send_message(chat_id, f"خطا: {e}")

def handle_help(chat_id):
    send_message(chat_id,
        "📌 *راهنمای ربات MenuRadar*\n\n"
        "📊 *گزارش رقبا* — آخرین فایل Excel مقایسه قیمت\n"
        "📅 *آخرین آپدیت* — تاریخ به‌روزرسانی داده‌ها\n\n"
        "⏱ گزارش‌ها هر هفته یکشنبه خودکار آپدیت می‌شن\n"
        "📢 گزارش جدید در کانال تلگرام هم ارسال می‌شه"
    )

def handle_not_authenticated(chat_id):
    send_message(chat_id, "🔐 ابتدا رمز عبور رو وارد کن:")

# ─────────────────────────────────────────────
# Process updates
# ─────────────────────────────────────────────

def process_update(update):
    if "callback_query" in update:
        cq = update["callback_query"]
        chat_id = cq["message"]["chat"]["id"]
        data = cq.get("data", "")
        answer_callback(cq["id"])
        if not is_authenticated(chat_id):
            handle_not_authenticated(chat_id)
            return
        if data == "report":
            handle_report(chat_id)
        elif data == "last_update":
            handle_last_update(chat_id)
        elif data == "help":
            handle_help(chat_id)
    elif "message" in update:
        msg = update["message"]
        chat_id = msg["chat"]["id"]
        first_name = msg["from"].get("first_name", "کاربر")
        text = msg.get("text", "").strip()
        if text in ["/start", "start"]:
            handle_start(chat_id, first_name)
        elif is_authenticated(chat_id):
            send_message(chat_id,
                "از دکمه‌های زیر استفاده کن 👇",
                reply_markup=MAIN_KEYBOARD
            )
        else:
            handle_password(chat_id, first_name, text)

# ─────────────────────────────────────────────
# Polling
# ─────────────────────────────────────────────

def run_polling():
    print("🤖 Bale bot started (v2 — RawMenus pattern)...")
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
                        print(f"Update error: {e}")
                    offset = update["update_id"] + 1
        except requests.exceptions.Timeout:
            pass
        except Exception as e:
            print(f"Polling error: {e}")
            time.sleep(5)

if __name__ == "__main__":
    run_polling()
