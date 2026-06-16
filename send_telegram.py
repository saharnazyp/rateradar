import os
import sys
import requests
import datetime

def send_to_telegram(file_path: str):
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    channel = os.environ["TELEGRAM_CHANNEL_ID"]  # e.g. @SPOMenuRadar or -100xxxxx

    date_str = datetime.date.today().strftime("%Y/%m/%d")

    caption = (
        f"📊 *گزارش MenuRadar*\n"
        f"📅 تاریخ: {date_str}\n\n"
        f" قیمت منو رستوران‌های SPO با رقبا\n"
        f"هر کانسپت در یک شیت جداگانه"
    )

    url = f"https://api.telegram.org/bot{token}/sendDocument"

    with open(file_path, "rb") as f:
        resp = requests.post(url, data={
            "chat_id": channel,
            "caption": caption,
            "parse_mode": "Markdown",
        }, files={"document": (os.path.basename(file_path), f)})

    if resp.status_code == 200:
        print(f"✅ File sent to Telegram channel {channel}")
    else:
        print(f"❌ Telegram error: {resp.status_code} — {resp.text}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python send_telegram.py <excel_file>")
        sys.exit(1)
    send_to_telegram(sys.argv[1])
