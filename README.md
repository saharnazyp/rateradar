<div align="center">

# 📊 MenuRadar

**Automated competitor menu price intelligence for SPO Restaurant Group**

🇬🇧 **English** &nbsp; | &nbsp; [🇮🇷 فارسی](./README.fa.md)

</div>

---

## Overview

MenuRadar is an automated weekly system that scrapes menus and prices from 40+ competitor restaurants in Tehran, intelligently matches them against SPO's own menu, and delivers actionable price-comparison reports to the team via a Telegram channel and a Bale bot.

The project covers all **7 SPO concepts**: Cafe Libro, Royal (Iranian kebab), Romatay (Thai/Italian), Chicken Factory, Burger Factory, SPO Express (Iranian food), and Lefoo (shake bar).

## Why this exists

SPO needed continuous visibility into competitor pricing across multiple concepts and brands, but manual price tracking was impractical. MenuRadar automates the entire pipeline — from scraping to delivery — and runs every Sunday morning without any human intervention.

A core constraint: most Iranian restaurant websites block requests from foreign IPs, which forced the scraping layer to live on a domestic VPS.

## What it does

- **Scrapes** 4000+ menu items weekly from 42 competitor restaurants across snappfood, menew.ir, delino, and direct restaurant sites
- **Compares** 532+ active SPO items against competitor pricing using a smart matching engine
- **Categorizes** each match as:
  - 🟢 **Strong match** — same item, all keywords align
  - 🟡 **Possible match** — same core item with a distinguishing qualifier (double, cheese, spicy, etc.)
  - ❌ **No match** — unique SPO item not found in competitor menus
- **Detects price gaps** with three states: 🔴 competitor more expensive, 🟢 competitor cheaper, ⚪ equivalent
- **Surfaces competitor-exclusive items** — dishes competitors offer that SPO doesn't, useful for menu R&D
- **Delivers** two Excel reports weekly: raw competitor menus + intelligent price comparison

## Architecture

```
┌─────────────────────────────────────────────────┐
│ Iranian VPS  (Sunday 9:00 AM Tehran, via cron)  │
│                                                 │
│  ┌──────────────┐  → reports/RawMenus_*.xlsx    │
│  │ scraper.py   │                               │
│  └──────────────┘                               │
│         │                                       │
│         ▼                                       │
│  ┌──────────────┐  → reports/Compare_*.xlsx     │
│  │ compare.py   │   (uses ai_matcher.py)        │
│  └──────────────┘                               │
│         │                                       │
│         ▼                                       │
│  ┌──────────────┐  → push to GitHub via REST    │
│  │  upload.py   │                               │
│  └──────────────┘                               │
└─────────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────┐
│ GitHub Actions  (triggered on push)             │
│                                                 │
│  ┌──────────────────┐                           │
│  │ send_telegram.py │ → @SPOMenuRadar channel   │
│  └──────────────────┘                           │
│                                                 │
│  ┌──────────────────┐                           │
│  │   bale_bot.py    │ → always-on Bale bot      │
│  └──────────────────┘                           │
└─────────────────────────────────────────────────┘
```

## The matching engine

The hardest problem was matching items across menus where the same dish appears under wildly different names — Persian, English, transliterated English, and various qualifiers. Examples:

| SPO item | Competitor variant |
|---|---|
| سیب زمینی فکتوری | Fries / French Fries / چیپس / Potato |
| کلد برو لیبرو | Cold Brew / Cold-Brew |
| دبل چیز برگر | Double Cheese Burger |
| میگو سوخاری | Fried Shrimp |

**`ai_matcher.py`** solves this with:
- A **synonym dictionary** of ~180 word families (Persian, English, transliterations)
- **Protein detection** — items with conflicting proteins (chicken vs. shrimp) never match
- **Qualifier awareness** — "double", "cheese", "spicy", "iced" surface in the report when they differ between SPO and the competitor
- **LRU caching** — canonicalization is memoized, dropping comparison time from ~50s to ~0.3s per concept

Result: 89-90% of SPO items find at least one comparable competitor item.

## Tech stack

- **Python 3.10+** — scraping, matching, report generation
- **openpyxl** — Excel reports with RTL support, color-coded rows, multiple sheets
- **Playwright + BeautifulSoup** — scraping (with Chromium on the VPS, since Playwright's installer is blocked from Iran)
- **GitHub REST API** — file transport between VPS and GitHub
- **python-telegram-bot** style HTTP polling — for the Bale bot
- **GitHub Actions** — orchestration for Telegram + Bale delivery
- **systemd-style cron** — weekly scheduling on the VPS

## Repository structure

```
rateradar/
├── menu_reader.py           # scraper
├── ai_matcher.py            # smart matching engine + synonym dictionary
├── compare_menus.py         # comparison report generator
├── upload_to_github.py      # VPS → GitHub uploader
├── bale_bot.py              # Bale messenger bot
├── send_telegram.py         # Telegram channel publisher
├── reports/                 # weekly Excel output
│   ├── RawMenus_*.xlsx      # competitor menus
│   └── Compare_*.xlsx       # price comparison
└── .github/workflows/
    ├── menuradar_send.yml   # Telegram publishing
    └── bale_bot.yml         # always-on bot runner
```

## Status

Currently in production for SPO Restaurant Group. Reports are delivered weekly without manual intervention.

---

<div align="center">

Built with care for SPO Restaurant Group · Tehran

</div>
