# Receipt Tax Bot 🧾

A Telegram bot that reads photos of your receipts, extracts the **sales tax**,
stores every receipt in an **Excel file on your own PC**, and gives you running
totals — so you know exactly how much deductible sales tax you paid, any time
of year.

- 📸 Send a receipt photo → get store, date, subtotal, **sales tax**, total
- 📊 `/summary` → total spent + total tax, per year (your deduction number)
- 📤 `/export` → the full Excel file, any time
- ♻️ Sends the same receipt twice? Detected and rejected automatically
- 👥 Multi-user: every Telegram user gets their own private storage
- 🔒 Your data never leaves your computer (only Telegram messages do)

---

## 🚀 Setup — even if you've never coded

You need: a Windows PC and the Telegram app. About 10 minutes.

### Step 1 — Download this project

Click the green **`<> Code`** button at the top of this page → **Download ZIP**.
Right-click the downloaded file → **Extract All** → put the folder somewhere
you like (e.g. `Documents`).

### Step 2 — Create your own Telegram bot (free, 2 minutes)

1. Open Telegram and search for **@BotFather** (blue checkmark).
2. Send it: `/newbot`
3. It asks for a **name** — type anything, e.g. `My Receipt Tracker`.
4. It asks for a **username** — must end in `bot`, e.g. `janes_receipts_bot`.
5. BotFather replies with a **token** that looks like
   `1234567890:AAHfK3xyz...` — **copy it**, you'll paste it in the next step.

> Your token is the key to YOUR bot. Don't share it or post it anywhere.

### Step 3 — Run the installer

Open the extracted folder and **double-click `setup.bat`**.

It will automatically install everything the bot needs (Python, the OCR
engine, libraries), then ask you to **paste your token**, and offer to start
the bot **automatically with Windows** (say yes — recommended).

> If Windows shows "Windows protected your PC", click **More info → Run anyway**.
> If it says Python was just installed, simply double-click `setup.bat` again.

### Step 4 — Use it!

Open Telegram, search for the bot username **you** created, press **START**,
and send it a photo of a receipt. Done — it replies with the tax and saves it.

---

## 💬 Commands

| Command | What it does |
|---|---|
| send a photo 📸 | Reads & stores the receipt |
| `/summary` | Total spent + total sales tax, all time and per year |
| `/year 2026` | Totals for one year (your tax-deduction number) |
| `/export` | Sends you your Excel file |
| `/last` | Shows your last 5 receipts |
| `/add 46.77 3.56 Walmart` | Manual entry without a photo (total, tax, store) |
| `/fix tax 3.56` | Corrects the last receipt (`tax`, `total`, `subtotal`, `store`, `date`) |
| `/fix 12 store Costco` | Corrects receipt #12 |
| `/delete` | Deletes the last receipt (or `/delete 12`) |

---

## 🔌 What if my PC is off, or starts before WiFi connects?

No receipt is lost. Telegram holds your messages in a queue, and the moment
your PC starts again the bot processes them **one by one, in order**.

The bot is built to start itself reliably and stay running:

- **Auto-start with Windows** — `setup.bat` offers to launch the bot
  (minimized) every time you log in, so you rarely have to start it by hand.
- **Waits for the internet** — if the bot starts before WiFi has connected
  (common right after logging in), it doesn't crash. It waits quietly and
  begins the moment the connection is up.
- **Survives a flaky first connection** — even after WiFi connects, the very
  first request to Telegram can stall on some networks. The bot keeps retrying
  instead of giving up.
- **Auto-restarts on a crash** — if it ever stops unexpectedly, `run_bot.bat`
  restarts it automatically after 10 seconds. (It won't restart when *you*
  stop it on purpose, or when it's already running.)

One limit (Telegram's, not ours): the queue holds messages for about
**24 hours**. If your PC was off longer than that, just re-send the photos —
takes seconds, and the duplicate detector makes re-sending always safe.

---

## 📦 Where is my data?

Everything is in the `data` folder next to the bot:

- `data/users/<your-telegram-id>/receipts.xlsx` — your receipts, one row each,
  plus a **Summary** sheet with per-year totals (live Excel formulas).
- `data/users/<your-telegram-id>/photos/` — every receipt photo, kept as your
  audit record.

Each Telegram user gets their own folder — people can share one bot without
ever seeing each other's data. Close the Excel file before sending new
receipts (Excel locks files while they're open).

---

## 🛠 Troubleshooting

| Problem | Fix |
|---|---|
| `python is not recognized` | Run `setup.bat` again; if it persists, install Python from python.org and tick **"Add python.exe to PATH"** |
| Bot doesn't reply | Is the bot window running? Double-click `run_bot.bat`. Check you messaged YOUR bot's username |
| "The bot is already running" | Fine — it's already on. There's nothing to do |
| Started before WiFi connected | Nothing to do — the bot waits for the internet and starts on its own once you're online |
| Window says "Restarting in 10 seconds" | Normal self-healing after a hiccup — leave it; it comes back by itself |
| Misread a number | `/fix tax 3.56` (or `total`, `store`, `date`) |
| Can't read receipt at all | Flatten it, good light, shoot from straight above — or `/add <total> <tax> <store>` |
| Want only certain people to use your bot | Put their numeric IDs (from @userinfobot) in `ALLOWED_USER_IDS=` in `.env`, then restart the bot |

---

## ☁️ Optional: cloud hosting (advanced)

Want the bot online 24/7 even with the PC off? A `Dockerfile` is included for
Railway / Render / Fly.io / any Docker host (~$5/mo). Set `TELEGRAM_BOT_TOKEN`
and `BACKUP_KEY` as environment variables, mount a volume at `/app/data`, and
use `sync_from_cloud.ps1` + `setup_sync_task.ps1` to mirror the data back to
your PC automatically. Run the bot in only ONE place at a time.

---

## ⚖️ Tax note (US)

If you itemize deductions, you may deduct **either** state income tax **or**
sales tax — this bot gives you the actual-receipts sales tax number to compare.
Keep the photos (the bot does it for you) as your records. Not tax advice —
ask a tax professional about your situation.
