"""
ربات بله — MenuRadar v3
دو نوع گزارش: RawMenus (رقبا) + Compare (مقایسه)
"""

import os
import glob
import datetime
import time
import requests

BALE_TOKEN = os.environ["BALE_BOT_TOKEN"]
BASE_URL = f"https://tapi.bale.ai/bot{BALE_TOKEN}"

BOT_PASSWORD = os.environ.get("BOT_PASSWORD", "SPO1403")

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
        [{"text": "📊 گزارش رقبا", "callback_data": "report_raw"}],
        [{"text": "🔍 مقایسه منو ما با رقبا", "callback_data": "report_compare"}],
        [{"text": "📅 آخرین آپدیت", "callback_data": "last_update"}],
        [{"text": "ℹ️ راهنما", "callback_data": "help"}],
    ]
}

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def get_latest_file(pattern):
    """جدیدترین فایل با الگوی مشخص — بر اساس mtime"""
    os.makedirs(REPORT_DIR, exist_ok=True)
    files = glob.glob(f"{REPORT_DIR}/{pattern}")
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def is_authenticated(chat_id):
    return authenticated_users.get(chat_id, False)

def _extract_date(filename):
    name = os.path.basename(filename)
    parts = name.replace(".xlsx", "").split("_")
    if len(parts) >= 2:
        return parts[-1]
    return "نامعلوم"

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

def handle_report_raw(chat_id):
    """گزارش رقبا — RawMenus_*.xlsx"""
    report = get_latest_file("RawMenus_*.xlsx")
    if not report:
        send_message(chat_id,
            "⚠️ هنوز گزارش رقبا آماده نیست.\n"
            "گزارش بعدی توسط سیستم خودکار ارسال می‌شه."
        )
        return
    date_part = _extract_date(report)
    caption = (
        f"📊 *گزارش منوی رقبا*\n"
        f"📅 تاریخ: `{date_part}`\n\n"
        f"منو و قیمت ۴۰+ برند رقیب در ۷ کانسپت"
    )
    send_message(chat_id, "⏳ در حال ارسال فایل...")
    send_document(chat_id, report, caption)

def handle_report_compare(chat_id):
    """مقایسه منو ما با رقبا — Compare_*.xlsx"""
    report = get_latest_file("Compare_*.xlsx")
    if not report:
        send_message(chat_id,
            "⚠️ هنوز گزارش مقایسه آماده نیست.\n"
            "گزارش بعدی توسط سیستم خودکار ساخته می‌شه."
        )
        return
    date_part = _extract_date(report)
    caption = (
        f"🔍 *مقایسه منوی SPO با رقبا*\n"
        f"📅 تاریخ: `{date_part}`\n\n"
        f"هر آیتم SPO + نزدیک‌ترین match رقبا\n"
        f"🟢 ارزون‌تر / 🔴 گرون‌تر / ⚪ مساوی\n"
        f"به‌علاوه آیتم‌های انحصاری رقبا"
    )
    send_message(chat_id, "⏳ در حال ارسال فایل...")
    send_document(chat_id, report, caption)

def handle_last_update(chat_id):
    raw = get_latest_file("RawMenus_*.xlsx")
    compare = get_latest_file("Compare_*.xlsx")
    if not raw and not compare:
        send_message(chat_id, "⚠️ هنوز هیچ گزارشی آماده نیست.")
        return
    parts = ["📅 *آخرین آپدیت*\n"]
    if raw:
        dt = datetime.datetime.fromtimestamp(os.path.getmtime(raw)).strftime("%Y/%m/%d — %H:%M")
        parts.append(f"📊 گزارش رقبا: `{_extract_date(raw)}` ({dt})")
    if compare:
        dt = datetime.datetime.fromtimestamp(os.path.getmtime(compare)).strftime("%Y/%m/%d — %H:%M")
        parts.append(f"🔍 مقایسه: `{_extract_date(compare)}` ({dt})")
    parts.append("\n_گزارش هر هفته یکشنبه به‌روز می‌شه_ ⏱")
    send_message(chat_id, "\n".join(parts))

def handle_help(chat_id):
    send_message(chat_id,
        "📌 *راهنمای ربات MenuRadar*\n\n"
        "📊 *گزارش رقبا* — منو و قیمت رقبا (فایل خام)\n"
        "🔍 *مقایسه منو ما با رقبا* — تطبیق هوشمند آیتم به آیتم\n"
        "📅 *آخرین آپدیت* — تاریخ به‌روزرسانی فایل‌ها\n\n"
        "⏱ گزارش‌ها هر هفته یکشنبه خودکار آپدیت می‌شن\n"
        "📢 گزارش جدید در کانال تلگرام @SPOMenuRadar هم منتشر می‌شه"
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
        if data == "report_raw":
            handle_report_raw(chat_id)
        elif data == "report_compare":
            handle_report_compare(chat_id)
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
    print("🤖 Bale bot v3 started — دو نوع گزارش")
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
