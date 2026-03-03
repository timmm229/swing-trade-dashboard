# ⚡ Swing Trade Dashboard Bot

**Automated NASDAQ swing trade intelligence — live data, formatted spreadsheets, scheduled emails, and a web dashboard.**

![Schedule: 7am, 9:30am, 12pm, 2:45pm CST](https://img.shields.io/badge/Schedule-4x_daily_CST-blue)
![Python 3.12+](https://img.shields.io/badge/Python-3.12+-green)

---

## What It Does

| Feature | Details |
|---|---|
| **Live Stock Data** | Pulls real-time prices for 10 NASDAQ swing candidates via yfinance (free, no API key) |
| **Formatted .xlsx** | Professional spreadsheet with color-coded % changes, yellow top-3, swing scores |
| **Futures & Macro** | Pre-market futures, VIX, treasury yields, dollar index |
| **Fed/FOMC Status** | Current rate, next meeting, rate cut expectations |
| **Swing Score** | Composite score (0-100) from volatility + momentum + volume |
| **Email Delivery** | Sends .xlsx attachment to your inbox 4x/day |
| **Web Dashboard** | Live browser view at your cloud URL with 1-click refresh |

### Schedule (CST)
| Time | Purpose |
|------|---------|
| 7:00 AM | Pre-market overview before open |
| 9:30 AM | Market open snapshot |
| 12:00 PM | Midday update |
| 2:45 PM | Late session / before-close positioning |

---

## Quick Start (Local)

```bash
# 1. Clone / download this folder
cd swing_trade_bot

# 2. Install dependencies
pip install -r requirements.txt

# 3. Set up email credentials (see Email Setup below)
cp .env.example .env
# Edit .env with your Gmail + App Password

# 4. Source the .env file
export $(cat .env | grep -v '^#' | xargs)

# 5. Run it
python app.py                  # Full app: scheduler + web dashboard + email
python app.py --run-once       # Generate one spreadsheet and email it, then exit
python app.py --no-email       # Run everything but skip email (for testing)
python app.py --web-only       # Web dashboard only, no scheduled emails

# 6. Open your browser
#    → http://localhost:5000
```

---

## Email Setup (Gmail — 5 minutes)

The app sends emails via Gmail SMTP. You need a **Gmail App Password** (not your regular password).

### Step-by-step:
1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable **2-Step Verification** if not already on
3. Go to [App Passwords](https://myaccount.google.com/apppasswords)
4. Select **"Mail"** and **"Other (Custom name)"** → name it "Swing Bot"
5. Google gives you a 16-character password like `abcd efgh ijkl mnop`
6. Put that in your `.env` file as `SMTP_PASSWORD` (remove spaces)

```env
SMTP_USER=your-gmail@gmail.com
SMTP_PASSWORD=abcdefghijklmnop
EMAIL_TO=el.capitan.44@gmail.com
```

---

## ☁️ Cloud Deployment Options

### Option A: Render.com (Recommended — Free Tier)

1. Push this folder to a GitHub repo
2. Go to [render.com](https://render.com) → New → **Web Service**
3. Connect your GitHub repo
4. Render auto-detects `render.yaml`
5. Add environment variables in Render dashboard:
   - `SMTP_USER` = your Gmail
   - `SMTP_PASSWORD` = your App Password
6. Deploy → your dashboard is live at `https://your-app.onrender.com`

### Option B: Railway.app (Simple)

1. Push to GitHub
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Add env vars in Railway dashboard
4. Gets a public URL automatically

### Option C: Fly.io (Always-On)

```bash
# Install flyctl
brew install flyctl   # or see https://fly.io/docs/flyctl/install/

# Deploy
cd swing_trade_bot
fly launch             # follow prompts
fly secrets set SMTP_USER=your@gmail.com SMTP_PASSWORD=yourapppassword EMAIL_TO=el.capitan.44@gmail.com
fly deploy
```

### Option D: Docker (Any Server / VPS)

```bash
# Build
docker build -t swing-bot .

# Run
docker run -d -p 5000:5000 \
  -e SMTP_USER=your@gmail.com \
  -e SMTP_PASSWORD=yourapppassword \
  -e EMAIL_TO=el.capitan.44@gmail.com \
  --name swing-bot \
  swing-bot
```

### Option E: PythonAnywhere (Free)

1. Upload files to PythonAnywhere
2. Set up a **scheduled task** (free tier gets 1/day; paid gets unlimited):
   - `cd /home/yourusername/swing_trade_bot && python app.py --run-once`
3. Set up a **web app** pointing to the Flask app for the dashboard

---

## Web Dashboard Endpoints

| URL | Description |
|-----|-------------|
| `/` | Live dashboard with futures, top 10 stocks, scores |
| `/download` | Download latest .xlsx file |
| `/refresh` | Trigger an immediate data refresh |
| `/api/data` | JSON API for programmatic access |

---

## Customizing Tickers

Edit the `TICKERS` list near the top of `app.py`:

```python
TICKERS = ["NVDA", "MSTR", "PLTR", "TSLA", "AMD", "AVGO", "PANW", "CRWD", "MU", "NFLX"]
```

Change to any NASDAQ-listed stocks with market cap > $1B.

---

## File Structure

```
swing_trade_bot/
├── app.py              # Main application (everything in one file)
├── requirements.txt    # Python dependencies
├── .env.example        # Template for credentials
├── Dockerfile          # Docker deployment
├── render.yaml         # Render.com auto-deploy config
├── README.md           # This file
└── output/             # Generated spreadsheets (auto-created)
    └── latest.xlsx
```

---

## Disclaimer

This tool is for **informational purposes only**. It is not financial advice.
Swing trading involves significant risk of loss. Always do your own research
and use proper risk management (stop-losses, position sizing).
