# 🚚 Logistics Tracking Bot — Setup Guide

## What's in this folder

| File | Purpose |
|------|---------|
| `bot.py` | The Telegram bot (main script) |
| `orders.xlsx` | Excel file with order data |
| `requirements.txt` | Python dependencies |

---

## Step 1 — Create your Telegram Bot (2 minutes)

1. Open Telegram and search for **@BotFather**
2. Send `/newbot`
3. Choose a name, e.g. *MyLogistics Bot*
4. Choose a username ending in `bot`, e.g. *mylogistics_bot*
5. BotFather gives you a token like:
   ```
   7412638901:AAFxyz_example_token_here
   ```
6. Copy that token — you'll need it in Step 3

---

## Step 2 — Install Python & dependencies

Make sure Python 3.10+ is installed, then run:

```bash
pip install -r requirements.txt
```

---

## Step 3 — Add your bot token

Open `bot.py` and find this line near the top:

```python
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "PASTE_YOUR_BOT_TOKEN_HERE")
```

**Option A** — Edit the file directly:
```python
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "7412638901:AAFxyz_your_real_token")
```

**Option B** — Set an environment variable (more secure):
```bash
# On Windows (Command Prompt):
set TELEGRAM_BOT_TOKEN=7412638901:AAFxyz_your_real_token

# On Mac/Linux:
export TELEGRAM_BOT_TOKEN=7412638901:AAFxyz_your_real_token
```

---

## Step 4 — Set up your Excel file

The bot reads `orders.xlsx` which must have a sheet named **Orders** with these columns:

| Column | Description |
|--------|-------------|
| Tracking Code | Unique ID clients will send (e.g. TRK-10001) |
| Client Name | Customer's full name |
| Phone Number | Customer's phone |
| Package Description | What's in the package |
| Date Sent | Date the package was shipped |
| Estimated Delivery Date | Expected arrival date |
| Status | Processing / In Transit / Delivered / Cancelled |
| Price ($) | Order price |
| Weight (kg) | Package weight |
| Notes | Any extra notes |

> ✅ **You can update the Excel file at any time while the bot is running.**
> The bot re-reads the file on every query — no restart needed.

---

## Step 5 — Run the bot

```bash
python bot.py
```

You should see:
```
INFO │ Starting bot... Loading Excel file: orders.xlsx
INFO │ Excel file loaded OK. Bot is running.
```

Now open Telegram, find your bot, send `/start`, and test it with a tracking code!

---

## How clients use it

1. Client opens Telegram and finds the bot
2. Sends their tracking code, e.g. `TRK-10001`
3. Bot instantly replies with all shipment details

---

## Keeping it running 24/7

To keep the bot online permanently, run it on a server. Some free/cheap options:

- **Railway.app** — free tier, easy deployment
- **Render.com** — free tier for always-on services
- **A VPS** (DigitalOcean, Hetzner) — ~$5/month
- **Your own PC** — just leave it running (works fine for small volume)

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `ValueError: Please set your bot token` | Add your token (see Step 3) |
| `FileNotFoundError: orders.xlsx` | Make sure `orders.xlsx` is in the same folder as `bot.py` |
| Bot doesn't respond | Check the terminal for errors |
| Wrong data shown | Check column names in Excel match the `COL_*` variables in `bot.py` |
