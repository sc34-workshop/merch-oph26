# Merch OPH 2026 × MWIT 36th Anniversary

A lightweight pre-order system for commemorative merchandise sold at MWIT Open House 2026, built and maintained by the **MWIT-SC 34 IT Recruitment team**.

**Live site →** https://sc34-workshop.github.io/merch-oph26/  
**Admin dashboard →** https://sc34-workshop.github.io/merch-oph26/admin.html

---

## What is this?

Mahidol Wittayanusorn School (MWIT) celebrates its **36th anniversary** at OPH 2026. This repo hosts the full pre-order flow — customers fill a form, generate a PromptPay QR, upload a payment slip, and an admin dashboard tracks everything in real time.

No server, no database subscription, no frameworks — just HTML/CSS/JS on GitHub Pages backed by **Google Apps Script + Google Sheets**.

---

## Merchandise

| Category | Items |
|---|---|
| 🔑 Keychains | มังกร MWIT · โลโก้ 36 ปี · มาสคอต |
| 📋 Stickers | Science Series · School Life Series |
| 👜 Tote Bag | Canvas Tote (1 design) |
| 👕 T-Shirts | S / M / L / XL × ขาว (white) / กรมท่า (navy) |
| 🎁 Set A | Keychain + Sticker bundle |
| 🎁 Set B | Keychain + Sticker + T-Shirt bundle |

Pickup at school (free) or postal delivery (+฿50).  
Payment via **PromptPay** — slip upload required to confirm the order.

---

## Architecture

```
Browser  ──────────────────────────────────────────────────►  GitHub Pages
                                                               (index.html / admin.html)
                                         │
                                    fetch (JSON)
                                         │
                                         ▼
                              Google Apps Script (doPost / doGet)
                                         │
                          ┌──────────────┴─────────────┐
                          ▼                             ▼
                   Google Sheets                  Google Drive
                  (order records)               (payment slip images)
```

| Layer | Technology |
|---|---|
| Hosting | GitHub Pages (static) |
| Frontend | Vanilla HTML · CSS variables · ES6 JS |
| Backend | Google Apps Script (serverless, zero cost) |
| Database | Google Sheets |
| File storage | Google Drive (payment slips) |
| Payment | PromptPay QR (generated client-side) |

---

## Project Structure

```
merch-oph26/
├── index.html          # Customer-facing pre-order form
├── admin.html          # Admin dashboard (orders + stats)
├── gas/
│   └── Code.gs         # Google Apps Script backend
└── SETUP.md            # Full deployment walkthrough
```

---

## Features

### Order Form (`index.html`)
- Product catalogue with live quantity selectors and running total
- Automatic PromptPay QR generation
- Payment slip upload (stored in Google Drive)
- Order reference number generated on submission
- Mobile-responsive · Light / dark mode

### Admin Dashboard (`admin.html`)
- Live stats: total orders, revenue, pending count, shipping split
- Donut chart — units sold per category
- Per-item quantity breakdown with bar chart
- Full order table with search + filter (by shipping method / status)
- **Pagination** — 100 orders per page with ← X–Y จาก Z → controls
- Status management (รอยืนยัน → ยืนยันแล้ว → จัดส่งแล้ว)
- CSV export (UTF-8 BOM for Thai characters in Excel)
- Light / dark mode

---

## Quick Setup

Full instructions are in [`SETUP.md`](SETUP.md). The short version:

1. **Google Sheet** — create a new sheet, copy its ID.
2. **Google Drive folder** — create a folder for slips, copy its ID.
3. **Apps Script** — paste `gas/Code.gs`, fill in the two IDs in `CONFIG`, deploy as a Web App, copy the URL.
4. **`index.html` + `admin.html`** — paste the Web App URL into `GAS_URL`. Set your PromptPay number in `index.html`.
5. **GitHub Pages** — push to `master`, enable Pages in repo settings.

> Run `setupSummarySheet()` from the GAS editor once to create the Summary tab in your Sheet.

---

## Development

No build step required. Open `index.html` directly in a browser, or serve locally:

```bash
npx serve .
# or
python -m http.server 8000
```

The `admin.html` page ships a **Sample Data** button (📊) that loads 150 generated orders — useful for previewing charts and pagination without connecting to a real Sheet.

---

## Tech Notes

- **Empty-row filtering** — `setupCheckboxes()` pre-fills data-validation checkboxes to row 10 000 for scalability. `doGet` filters on `OrderRef` so those blank rows never appear in the dashboard count.
- **Dark-mode CSS variables** — all inline `color: var(--navy)` table-cell styles inherit the overridden variable from `[data-theme="dark"]` on `<html>`, so no `!important` hacks are needed.
- **Thai text in CSV** — a UTF-8 BOM (`﻿`) is prepended so Excel on Windows opens the file correctly without a manual import wizard.

---

## Credits

Built by the **IT Recruitment team, MWIT Science Club Batch 34** for OPH 2026.  
School: [Mahidol Wittayanusorn School](https://www.mwit.ac.th) · ครบรอบ 36 ปี มหิดลวิทยานุสรณ์

**Developer:** Tidsanat Kaewlad (MWIT-SC 34)
