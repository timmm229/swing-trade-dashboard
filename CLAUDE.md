# CLAUDE.md — Swing Trade Dashboard

This file documents the codebase structure, conventions, and development workflows for AI assistants working on this project.

## Project Overview

Automated swing trading intelligence platform built in Python. Fetches real-time stock data, runs technical analysis, generates formatted Excel spreadsheets with embedded charts, delivers reports via email, and serves a live web dashboard.

**No database. No API keys required for stock data.** Everything runs from a single Python file.

## Technology Stack

- **Python 3.12+**
- **Flask 3.0+** — Web server and API
- **yfinance 0.2.36+** — Free real-time/historical stock data (no API key)
- **pandas 2.0+** — Data manipulation and Excel I/O
- **numpy 1.24+** — Numerical calculations
- **matplotlib 3.8+** — Technical chart generation (PNG)
- **openpyxl 3.1+** — Professional Excel (.xlsx) spreadsheet creation
- **APScheduler 3.10+** — Cron-style background job scheduling
- **SMTP (Gmail)** — Email delivery

## Repository Structure

```
swing-trade-dashboard/
├── app.py              # Entire application (~1200 lines, monolithic)
├── requirements.txt    # Python dependencies
├── .env.example        # Required environment variable template
├── Dockerfile          # Container deployment config
├── render.yaml         # Render.com deployment config
├── README.md           # User-facing documentation
└── output/             # Generated files (gitignored at runtime)
    ├── latest.xlsx     # Most recent dashboard spreadsheet
    └── charts/         # Generated technical chart PNGs
```

## Architecture: Single-File Monolith

All logic lives in `app.py`. Do not split it into modules unless explicitly asked — the simplicity is intentional.

### Code Layout in app.py

| Lines | Section |
|-------|---------|
| 1–76 | Imports, configuration constants, global state |
| 79–113 | Technical indicator calculations (RSI, MACD, Bollinger, SMA) |
| 119–374 | Data fetching (stocks, buy signals, futures, news) |
| 391–502 | Chart generation (matplotlib, dark theme, 4-subplot) |
| 509–852 | Spreadsheet building (5-sheet .xlsx with formatting) |
| 859–908 | Email module (SMTP/TLS, Gmail) |
| 915–944 | Main job runner (`run_job()`) |
| 951–1087 | Embedded HTML/CSS/JavaScript for web dashboard |
| 1089–1154 | Flask route definitions |
| 1161–1170 | APScheduler setup |
| 1173–1193 | CLI argument parsing and main entry point |

## Key Constants (top of app.py)

```python
TICKERS = ["NVDA", "MSTR", "PLTR", "TSLA", "AMD", "AVGO", "PANW", "CRWD", "MU", "NFLX"]
SECTORS = { ... }    # Ticker -> sector string mapping
OUTPUT_DIR = "output"
```

To change the tracked stocks, edit `TICKERS` and `SECTORS` only.

## Global State

A single `latest` dict holds in-memory results between requests:

```python
latest = {
    "file": str,              # Path to latest .xlsx
    "summary": {
        "top3": list,         # Top 3 tickers by swing score
        "futures_verdict": str, # "BULLISH" | "MIXED" | "BEARISH"
        "futures": list
    },
    "generated_at": str,      # ISO timestamp
    "error": str | None,
    "buy_signals": list,
    "chart_images": dict      # ticker -> base64 PNG string
}
```

## Flask API Routes

| Route | Method | Description |
|-------|--------|-------------|
| `/` | GET | Web dashboard (HTML, full page) |
| `/api/data` | GET | JSON — all analysis results |
| `/refresh` | GET | Trigger immediate job run |
| `/chart/<ticker>` | GET | PNG image for a ticker |
| `/download` | GET | Download latest .xlsx |

## Scheduler Schedule (CST timezone)

Jobs run automatically at:
- 7:00 AM — Pre-market
- 9:30 AM — Market open
- 12:00 PM — Midday
- 2:45 PM — Late session

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and configure environment
cp .env.example .env
# Edit .env with your Gmail credentials

# Run modes
python app.py                  # Full: scheduler + web dashboard + emails
python app.py --run-once       # Generate once, send email, exit
python app.py --no-email       # Run without sending email (testing)
python app.py --web-only       # Web dashboard only, no scheduled emails

# Docker
docker build -t swing-bot .
docker run -d -p 5000:5000 \
  -e SMTP_USER=you@gmail.com \
  -e SMTP_PASSWORD=apppassword16chars \
  -e EMAIL_TO=recipient@gmail.com \
  --name swing-bot swing-bot
```

## Environment Variables

All required. Defined in `.env` locally or platform env vars in production.

| Variable | Description |
|----------|-------------|
| `SMTP_USER` | Gmail address used to send reports |
| `SMTP_PASSWORD` | Gmail App Password (16-char, not account password) |
| `EMAIL_TO` | Recipient email address |
| `SMTP_SERVER` | Default: `smtp.gmail.com` |
| `SMTP_PORT` | Default: `587` |
| `PORT` | Flask port, default `5000` |
| `OUTPUT_DIR` | Output directory, default `output` |

## Technical Analysis Logic

### Swing Score (0–100)
Combines 3 factors per ticker:
- **Volatility** — ATR-based, favors moderate volatility
- **Momentum** — Rate of change over recent periods
- **Liquidity** — Average volume weighting

### Buy Score (0–100)
Combines 5 binary signals (20 pts each):
1. RSI(14) < 70 (not overbought)
2. MACD line above signal line
3. Price above lower Bollinger Band
4. Price above SMA(50)
5. SMA(50) above SMA(200) (golden cross)

**Ratings:** BUY (≥70) | WATCH (45–69) | AVOID (<45)

### Futures Verdict
- Count bullish vs bearish macro signals
- BULLISH: majority green | BEARISH: majority red | MIXED: otherwise

## Excel Spreadsheet Structure (5 sheets)

1. **Market Overview** — Futures, macro indicators, Fed/FOMC status
2. **Top 10 Swing Trades** — Ranked by swing score, top 3 highlighted in green
3. **Buy Signals** — Technical ratings with color-coded BUY/WATCH/AVOID
4. **Technical Charts** — 6-month price/volume/RSI/MACD charts embedded
5. **Trading Notes** — Disclaimer and usage guide

## Chart Specifications

- **Theme:** Dark (GitHub-style color scheme)
- **Duration:** 6-month historical data
- **4 Subplots:** Price + Bollinger Bands + SMAs | Volume | RSI(14) | MACD
- **Output:** PNG saved to `output/charts/<ticker>.png`, also base64 for web

## Development Conventions

- **No new files** unless truly necessary — keep the monolithic structure
- **No database** — file-based outputs are intentional
- **No external APIs** for stock data — yfinance is free and sufficient
- **Inline HTML/CSS/JS** — the web dashboard is embedded in `app.py` as a Python string; do not extract to separate files unless asked
- When adding a new ticker, add it to both `TICKERS` list and `SECTORS` dict
- Chart colors and spreadsheet styling are hardcoded inline near usage — no separate theme file
- Error handling follows the pattern: catch exceptions broadly in `run_job()`, store `error` in `latest`, surface via API

## Deployment

**Render.com** (recommended): Push to GitHub, connect repo, set env vars in dashboard.

**Docker**: Build image, run with `-e` flags for env vars. Exposes port 5000.

**Other platforms:** Railway, Fly.io, PythonAnywhere — all supported (see README.md).

## No Tests

There is no test suite. Validation is done manually via:
```bash
python app.py --run-once --no-email
curl http://localhost:5000/api/data
```
