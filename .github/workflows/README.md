# 📊 MenuRadar — سیستم تحلیل قیمت رقبا

## ساختار پروژه

```
MenuRadar/
├── .github/
│   └── workflows/
│       └── menu_radar.yml     ← زمان‌بندی GitHub Actions
├── data/
│   ├── my_menu.xlsx           ← منوی رستوران‌های خودت (ثابت)
│   └── competitors.xlsx       ← لیست URL رقبا (ثابت)
├── scraper.py                 ← اسکریپر اصلی + مقایسه
├── send_telegram.py           ← ارسال به کانال تلگرام
├── requirements.txt
└── README.md
```

---

## راه‌اندازی اول (یک‌بار)

### ۱. ساخت ربات تلگرام
- به [@BotFather](https://t.me/BotFather) پیام بده
- `/newbot` بزن، اسم و username بده
- **Token** رو ذخیره کن

### ۲. اضافه کردن ربات به کانال
- ربات رو به کانال `@SPOMenuRadar` به عنوان **Admin** اضافه کن
- دسترسی **Post Messages** رو بده

### ۳. گرفتن Channel ID
- اگه کانال public هست: همون `@SPOMenuRadar` کافیه
- اگه private: با [@userinfobot](https://t.me/userinfobot) ID رو بگیر (شروع با `-100`)

### ۴. ساخت Repository در GitHub
```bash
git init
git add .
git commit -m "Initial MenuRadar setup"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/MenuRadar.git
git push -u origin main
```

### ۵. اضافه کردن Secrets در GitHub
برو به: **Settings → Secrets and variables → Actions → New repository secret**

| Secret Name | Value |
|---|---|
| `TELEGRAM_BOT_TOKEN` | توکن ربات از BotFather |
| `TELEGRAM_CHANNEL_ID` | مثلاً `@SPOMenuRadar` یا `-100xxxxxxxxx` |

---

## اجرای دستی

از GitHub بکش بالا → **Actions** → **MenuRadar** → **Run workflow**

---

## زمان‌بندی خودکار

هر ماه روزهای ۱، ۱۱، ۲۱ ساعت ۸ صبح ایران اجرا می‌شه.
برای تغییر زمان‌بندی، فایل `.github/workflows/menu_radar.yml` رو ویرایش کن.

---

## خروجی

یک فایل Excel با:
- **شیت «خلاصه کل»**: همه کانسپت‌ها کنار هم
- **شیت جداگانه** برای هر کانسپت (لیبرو، برگر فکتوری، ...)
- ستون‌ها: آیتم من | قیمت من | آیتم رقیب | قیمت رقیب | رقیب | اختلاف | وضعیت
