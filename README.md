# 📊 Stock Analyzer — One Glance Dashboard

A **zero-dependency, client-side** stock analysis tool. Upload a [Screener.in](https://screener.in) Excel export and get an instant visual dashboard with charts, traffic-light indicators, and an auto-generated investment report.

> **No server, no API keys, no setup** — just open `index.html` in your browser.

---

## ✨ Features

| Category | Feature |
|----------|---------|
| 🎯 **Core** | Upload `.xlsx` → instant dashboard with BUY / HOLD / AVOID verdict |
| 📊 **5 Charts** | Revenue & Profit trend, Quarterly Net Profit bars, Balance Sheet stacked, Cash Flow grouped, Quarterly OPM trend |
| 🚦 **6 Indicator Cards** | Revenue Growth, Net Profit, Quarterly Trend, OPM, Debt Health, Cash Flow — each with sparklines & color signals |
| 📐 **7 Ratios** | EPS, P/E, P/B, ROE, ROA, D/E, Dividend Yield |
| 🎯 **Score Gauge** | Animated SVG dial (0–10) with breakdown tags showing exactly why |
| ⚖️ **Compare Mode** | Upload multiple stocks and see side-by-side comparison table |
| 🏦 **Banking Detection** | Auto-detects banking stocks, computes NIM |
| 📄 **PDF Export** | One-click download full dashboard as PDF |
| 🌙 **Dark / Light Mode** | Toggle with persistence via localStorage |
| 📁 **Analysis History** | Last 10 analyses saved locally, shown on upload screen |
| 🖨️ **Print-Friendly** | Clean print layout via `@media print` |
| 💬 **Tooltips** | Hover any indicator card for plain-English explanation |
| 🎬 **Animations** | Gauge fill, number counters, sparklines, hover effects |

---

## 🚀 Quick Start

1. Download or clone this repo
2. Open `index.html` in any modern browser
3. Drag & drop a Screener.in `.xlsx` file (or click to browse)
4. Done — your analysis dashboard appears instantly

### Getting the Excel File

1. Go to [screener.in](https://www.screener.in)
2. Search for any company
3. Click **"Export to Excel"** (requires free login)
4. Upload the downloaded `.xlsx` here

---

## 📁 Project Structure

```
Stock/
├── index.html   # Main app shell (upload screen + dashboard layout)
├── index.css    # Premium dark/light theme with glassmorphism
├── app.js       # Excel parser, analysis engine, charts, all features
└── README.md    # You are here
```

**No build step.** No `npm install`. No framework. Just 3 files.

---

## 🧠 How the Analysis Works

The engine extracts data from the Excel file and scores **8 financial metrics**:

| Metric | Green | Amber | Red |
|--------|-------|-------|-----|
| Revenue CAGR | > 10% | 0–10% | Negative |
| Net Profit | Growing | Declining | Loss |
| Quarterly Trend | All profitable | 1 loss quarter | 3+ loss quarters |
| Operating Margin | > 15% | 0–15% | Negative |
| Debt-to-Equity | < 1x | 1–3x | > 3x |
| Cash Flow | Positive | — | Negative |
| Expense Spike | Normal | — | > 20% above revenue growth |
| Valuation (P/B) | < 1x | 1–5x | Value trap |

**Scoring:** Each metric adds or subtracts points → mapped to 0–10 gauge → verdict:
- **≥ 4** → Buy | **≥ 2** → Hold | **< 0** → Avoid

---

## 🛠️ Tech Stack

- **[SheetJS](https://sheetjs.com/)** — Excel parsing (CDN)
- **[Chart.js](https://www.chartjs.org/)** — Interactive charts (CDN)
- **[html2canvas](https://html2canvas.hertzen.com/)** + **[jsPDF](https://parall.ax/products/jspdf)** — PDF export (CDN)
- **Vanilla HTML/CSS/JS** — No frameworks

---

## 📸 Screenshots

> Upload your own screenshots here after running the app!
>
> Suggested: Upload Screen | Dashboard (Dark) | Dashboard (Light) | Compare Mode

---

## ⚠️ Disclaimer

This tool is for **educational and informational purposes only**. It does not constitute financial advice. Always consult a SEBI-registered advisor before making investment decisions.

---

## 📄 License

MIT — free to use, modify, and distribute.
