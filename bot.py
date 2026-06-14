"""Receipt Tax Bot — send a receipt photo, get the sales tax back,
everything stored in an Excel file for year-end deductions.

Multi-user: every Telegram user gets their own isolated Excel workbook and
photo folder under data/users/<id>/ — nobody can see anyone else's data.

A tiny HTTP backup endpoint (enabled when BACKUP_KEY is set) lets the
owner's PC pull a zip of all data when it comes online — used for the
cloud-to-PC sync.

Commands:
  /summary          totals (all time + per year)
  /year 2026        totals for one year
  /export           sends the Excel file
  /last             last 5 receipts
  /add <total> <tax> [store]   manual entry, no photo
  /fix tax 3.56     correct the last receipt (also: total, subtotal, store, date)
  /fix 12 tax 3.56  correct receipt #12
  /delete [id]      delete last (or given) receipt
"""

import asyncio
import hashlib
import io
import logging
import os
import socket
import threading
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from dotenv import load_dotenv
from telegram import Update
from telegram.error import NetworkError, RetryAfter
from telegram.ext import (Application, CommandHandler, ContextTypes,
                          MessageHandler, filters)

import storage
from ocr import extract_text
from receipt_parser import parse_receipt

load_dotenv()
os.makedirs(storage.DATA_DIR, exist_ok=True)
logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s",
                    level=logging.INFO,
                    handlers=[logging.StreamHandler(),
                              logging.FileHandler(os.path.join(storage.DATA_DIR, "bot.log"),
                                                  encoding="utf-8")])
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("receipt-bot")

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
# Empty ALLOWED_USER_IDS = bot is open to everyone (each user isolated).
ALLOWED = {int(x) for x in os.environ.get("ALLOWED_USER_IDS", "").replace(",", " ").split() if x.strip()}
BACKUP_KEY = os.environ.get("BACKUP_KEY", "")
PORT = int(os.environ.get("PORT", "8080"))

FIX_FIELDS = ("tax", "total", "subtotal", "store", "date")


def authorized(update: Update) -> bool:
    return not ALLOWED or (update.effective_user and update.effective_user.id in ALLOWED)


def fmt_money(v) -> str:
    return f"${v:,.2f}" if isinstance(v, (int, float)) else "—"


def receipt_card(rid: int, store, date, subtotal, tax, total) -> str:
    return (f"🧾 *Receipt #{rid}*\n"
            f"🏪 Store: {store or 'Unknown'}\n"
            f"📅 Date: {date or '—'}\n"
            f"💵 Subtotal: {fmt_money(subtotal)}\n"
            f"🏛 Sales tax: *{fmt_money(tax)}*\n"
            f"💰 Total: {fmt_money(total)}")


# ------------------------------------------------------- resilient sending

async def _retry(factory, attempts: int = 4):
    """Run an awaitable factory, retrying on network blips with backoff.

    The connection to api.telegram.org sometimes stalls mid-handshake,
    so a single failed attempt must never lose a message.
    """
    delay = 2.0
    for i in range(attempts):
        try:
            return await factory()
        except RetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
        except NetworkError as e:  # includes TimedOut
            if i == attempts - 1:
                raise
            log.warning("Telegram send failed (%s) — retrying in %.0fs",
                        type(e).__name__, delay)
            await asyncio.sleep(delay)
            delay *= 2


async def reply(msg, text: str, **kwargs):
    return await _retry(lambda: msg.reply_text(text, **kwargs))


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE):
    err = context.error
    if isinstance(err, NetworkError):
        log.warning("Network hiccup (%s): %s — the bot keeps running.",
                    type(err).__name__, err)
    else:
        log.error("Unhandled error while processing an update", exc_info=err)


# ------------------------------------------------------- backup endpoint

class _BackupHandler(BaseHTTPRequestHandler):
    """GET /            -> health check (used by the cloud host)
       GET /backup?key= -> zip of the whole data dir (owner's PC sync)"""

    def log_message(self, fmt, *args):  # route to our logger, not stderr
        log.info("backup-server: " + fmt, *args)

    def _send(self, body: bytes, ctype: str, filename: str | None = None):
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        if filename:
            self.send_header("Content-Disposition",
                             f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        url = urlparse(self.path)
        if url.path == "/":
            self._send(b"receipt-tax-bot ok", "text/plain")
            return
        if url.path == "/backup":
            key = (parse_qs(url.query).get("key") or [""])[0]
            if not BACKUP_KEY or key != BACKUP_KEY:
                self.send_error(403, "bad key")
                return
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for root, _, files in os.walk(storage.DATA_DIR):
                    for fn in files:
                        if fn == "bot.log":
                            continue
                        p = os.path.join(root, fn)
                        z.write(p, os.path.relpath(p, storage.DATA_DIR))
            self._send(buf.getvalue(), "application/zip", "receipts_backup.zip")
            return
        self.send_error(404)


def start_backup_server() -> None:
    if not BACKUP_KEY:
        log.info("BACKUP_KEY not set — backup endpoint disabled.")
        return
    try:
        srv = ThreadingHTTPServer(("0.0.0.0", PORT), _BackupHandler)
    except OSError as e:
        log.warning("Backup endpoint disabled (port %s busy: %s)", PORT, e)
        return
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    log.info("Backup endpoint listening on port %s", PORT)


# ---------------------------------------------------------------- commands

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    await reply(update.message,
        "👋 Hi! I track your receipts and sales tax for deductions.\n\n"
        "📸 Just send me a *photo of a receipt* — I'll read the sales tax, "
        "store it in your own private Excel file, and keep running totals. "
        "Your data is yours alone — every user has separate storage.\n\n"
        "*Commands*\n"
        "/summary — total spent & tax (all time + per year)\n"
        "/year 2026 — totals for one year\n"
        "/export — get your Excel file\n"
        "/last — last 5 receipts\n"
        "/add 46.77 3.56 Walmart — manual entry without a photo\n"
        "/fix tax 3.56 — correct the last receipt (tax/total/subtotal/store/date)\n"
        "/delete — delete the last receipt\n\n"
        "💡 Tips for best results: flatten the receipt, good light, "
        "shoot straight from above.",
        parse_mode="Markdown")


async def cmd_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    s = storage.get_summary(uid)
    if not s["count"]:
        await reply(update.message, "No receipts yet — send me a photo of one! 📸")
        return
    lines = ["📊 *All-time summary*",
             f"🧾 Receipts: {s['count']}",
             f"💰 Total spent: {fmt_money(s['spent'])}",
             f"🏛 Total sales tax: *{fmt_money(s['tax'])}*", ""]
    for y, d in s["per_year"].items():
        lines.append(f"*{y}* — {d['count']} receipts · spent {fmt_money(d['spent'])} "
                     f"· tax *{fmt_money(d['tax'])}*")
    lines.append("\n📤 /export for the full Excel file")
    await reply(update.message, "\n".join(lines), parse_mode="Markdown")


async def cmd_year(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    try:
        year = int(context.args[0])
    except (IndexError, ValueError):
        year = datetime.now().year
    s = storage.get_summary(uid, year)
    if not s["count"]:
        await reply(update.message, f"No receipts for {year}.")
        return
    await reply(update.message,
        f"📊 *{year}*\n🧾 Receipts: {s['count']}\n"
        f"💰 Total spent: {fmt_money(s['spent'])}\n"
        f"🏛 Total sales tax: *{fmt_money(s['tax'])}*\n\n"
        f"That's your deductible sales tax for {year}. 📤 /export for details.",
        parse_mode="Markdown")


async def cmd_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    if not storage.excel_exists(uid):
        await reply(update.message, "No receipts yet — nothing to export.")
        return

    def _send():
        f = open(storage.excel_path(uid), "rb")
        return update.message.reply_document(
            f, filename="receipts.xlsx",
            caption="📊 All your receipts. The Summary sheet has per-year totals.")

    await _retry(_send)


async def cmd_last(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    receipts = storage.get_last(uid, 5)
    if not receipts:
        await reply(update.message, "No receipts yet — send me a photo of one! 📸")
        return
    lines = ["🧾 *Recent receipts*"]
    for r in receipts:
        lines.append(f"#{r['ID']} · {r['Receipt Date'] or '—'} · {r['Store']} · "
                     f"total {fmt_money(r['Total'])} · tax *{fmt_money(r['Sales Tax'])}*")
    await reply(update.message, "\n".join(lines), parse_mode="Markdown")


async def cmd_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    args = context.args or []
    try:
        total = float(args[0].replace("$", "").replace(",", ""))
        tax = float(args[1].replace("$", "").replace(",", ""))
        store = " ".join(args[2:]) or "Manual entry"
    except (IndexError, ValueError):
        await reply(update.message,
            "Usage: /add <total> <tax> [store]\nExample: /add 46.77 3.56 Walmart")
        return
    rid = storage.add_receipt(uid,
        store=store, date=datetime.now().strftime("%m/%d/%Y"),
        subtotal=round(total - tax, 2), tax=tax, total=total,
        photo="", user=update.effective_user.first_name or str(uid))
    await reply(update.message,
        receipt_card(rid, store, datetime.now().strftime("%m/%d/%Y"),
                     round(total - tax, 2), tax, total) + "\n\n✅ Saved.",
        parse_mode="Markdown")


async def cmd_fix(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    args = list(context.args or [])
    usage = ("Usage: /fix <field> <value> — fixes the last receipt\n"
             "or /fix <id> <field> <value>\n"
             f"Fields: {', '.join(FIX_FIELDS)}\n"
             "Examples: /fix tax 3.56 · /fix 12 store Walmart")
    rid = None
    if args and args[0].isdigit() and len(args) >= 2 and args[1] in FIX_FIELDS:
        rid = int(args.pop(0))
    if len(args) < 2 or args[0] not in FIX_FIELDS:
        await reply(update.message, usage)
        return
    field, raw = args[0], " ".join(args[1:])
    if rid is None:
        rid = storage.last_receipt_id(uid)
    if rid is None:
        await reply(update.message, "No receipts to fix yet.")
        return
    if field in ("tax", "total", "subtotal"):
        try:
            value = float(raw.replace("$", "").replace(",", ""))
        except ValueError:
            await reply(update.message, f"That's not a number: {raw}")
            return
    else:
        value = raw
    try:
        ok = storage.update_receipt(uid, rid, field, value)
    except RuntimeError as e:
        await reply(update.message, f"⚠️ {e}")
        return
    if not ok:
        await reply(update.message, f"Receipt #{rid} not found.")
        return
    r = storage.get_receipt(uid, rid)
    await reply(update.message,
        receipt_card(rid, r["Store"], r["Receipt Date"], r["Subtotal"],
                     r["Sales Tax"], r["Total"]) + "\n\n✅ Updated.",
        parse_mode="Markdown")


async def cmd_delete(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    uid = update.effective_user.id
    args = context.args or []
    rid = int(args[0]) if args and args[0].isdigit() else storage.last_receipt_id(uid)
    if rid is None:
        await reply(update.message, "No receipts to delete.")
        return
    try:
        ok = storage.delete_receipt(uid, rid)
    except RuntimeError as e:
        await reply(update.message, f"⚠️ {e}")
        return
    await reply(update.message,
        f"🗑 Receipt #{rid} deleted." if ok else f"Receipt #{rid} not found.")


# ---------------------------------------------------------------- photos

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    msg = update.message
    uid = update.effective_user.id

    # Progress note is nice-to-have; never let it kill the receipt processing.
    try:
        note = await reply(msg, "🔍 Reading your receipt…")
    except NetworkError:
        note = None

    async def respond(text: str, **kwargs):
        if note is not None:
            try:
                await _retry(lambda: note.edit_text(text, **kwargs))
                return
            except NetworkError:
                pass
        await reply(msg, text, **kwargs)

    # Largest photo size, or an image sent as a file.
    if msg.photo:
        tg_file = await _retry(msg.photo[-1].get_file)
        ext = ".jpg"
    else:
        tg_file = await _retry(msg.document.get_file)
        ext = os.path.splitext(msg.document.file_name or "r.jpg")[1] or ".jpg"

    os.makedirs(storage.photos_dir(uid), exist_ok=True)
    fname = f"receipt_{datetime.now():%Y%m%d_%H%M%S}{ext}"
    path = os.path.join(storage.photos_dir(uid), fname)
    await _retry(lambda: tg_file.download_to_drive(path))

    # Duplicate check #1: exact same image file sent before.
    with open(path, "rb") as f:
        file_hash = hashlib.sha256(f.read()).hexdigest()
    dup = storage.find_duplicate(uid, file_hash=file_hash)
    if dup:
        os.remove(path)
        await respond(
            f"♻️ Duplicate! This is the exact same photo as receipt "
            f"*#{dup[0]}* — I didn't save it again.\n"
            f"Check it with /last.", parse_mode="Markdown")
        return

    try:
        text = await asyncio.to_thread(extract_text, path)
    except Exception:
        log.exception("OCR failed")
        await respond(
            "😕 I couldn't read that image. Try a sharper, well-lit photo, "
            "or enter it manually: /add <total> <tax> [store]")
        return

    parsed = parse_receipt(text)
    if parsed.total is None and parsed.tax is None:
        await respond(
            "😕 I couldn't find any amounts on that receipt.\n"
            "Try: flatten it, good light, shoot straight from above.\n"
            "Or enter it manually: /add <total> <tax> [store]")
        return

    # Duplicate check #2: same receipt re-photographed (store/date/amounts).
    dup = storage.find_duplicate(uid, store=parsed.store, date=parsed.date,
                                 tax=parsed.tax, total=parsed.total)
    if dup:
        os.remove(path)
        await respond(
            f"♻️ Duplicate! This looks like receipt *#{dup[0]}* "
            f"({dup[1]}) — I didn't save it again.\n"
            f"Check it with /last. If it really is a different purchase, "
            f"add it manually: /add {parsed.total} {parsed.tax} {parsed.store}",
            parse_mode="Markdown")
        return

    try:
        rid = storage.add_receipt(uid,
            store=parsed.store, date=parsed.date, subtotal=parsed.subtotal,
            tax=parsed.tax, total=parsed.total, photo=fname, file_hash=file_hash,
            user=update.effective_user.first_name or str(uid))
    except RuntimeError as e:
        await respond(f"⚠️ {e}")
        return

    out = receipt_card(rid, parsed.store, parsed.date, parsed.subtotal,
                       parsed.tax, parsed.total)
    if parsed.warnings:
        out += "\n\n⚠️ " + "\n⚠️ ".join(parsed.warnings)
    out += "\n\n✏️ Wrong? `/fix tax 1.23` · `/fix total 45.67` · /delete"
    await respond(out, parse_mode="Markdown")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not authorized(update):
        return
    await reply(update.message,
        "📸 Send me a photo of a receipt, or use /summary, /export, /add. "
        "/start for help.")


def acquire_single_instance_lock() -> socket.socket:
    """Only one copy of the bot may run (two copies fight over Telegram
    polling). Holding a local port is a simple cross-process lock."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", 47200))
        s.listen(1)
        return s
    except OSError:
        print("The bot is already running — no need to start it twice.")
        raise SystemExit(3)


def wait_for_network(host: str = "api.telegram.org", port: int = 443) -> None:
    """Block until Telegram is reachable.

    When the bot auto-starts at login, WiFi is often not connected yet.
    Instead of crashing, we wait quietly and start the moment the internet
    comes up. Retries forever with a gentle backoff.
    """
    import time
    waited, delay = 0, 3
    while True:
        try:
            socket.create_connection((host, port), timeout=5).close()
            if waited:
                log.info("Internet is up after %ss of waiting — starting now.", waited)
            return
        except OSError:
            if waited == 0:
                log.info("No internet yet — waiting for WiFi to connect…")
            time.sleep(delay)
            waited += delay
            delay = min(int(delay * 1.5), 30)


def main() -> None:
    if not TOKEN or "PASTE" in TOKEN:
        print("No bot token. Open the .env file and set TELEGRAM_BOT_TOKEN "
              "to the token @BotFather gave you.")
        raise SystemExit(4)
    lock = acquire_single_instance_lock()  # noqa: F841 — held for process lifetime
    wait_for_network()
    start_backup_server()
    app = (Application.builder().token(TOKEN)
           # Generous timeouts: this connection sometimes stalls on TLS setup.
           .connect_timeout(25).read_timeout(30).write_timeout(30)
           .media_write_timeout(120).pool_timeout(15)
           .get_updates_connect_timeout(25).get_updates_read_timeout(30)
           .build())
    app.add_error_handler(on_error)
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_start))
    app.add_handler(CommandHandler("summary", cmd_summary))
    app.add_handler(CommandHandler("total", cmd_summary))
    app.add_handler(CommandHandler("year", cmd_year))
    app.add_handler(CommandHandler("export", cmd_export))
    app.add_handler(CommandHandler("excel", cmd_export))
    app.add_handler(CommandHandler("last", cmd_last))
    app.add_handler(CommandHandler("add", cmd_add))
    app.add_handler(CommandHandler("fix", cmd_fix))
    app.add_handler(CommandHandler("delete", cmd_delete))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    log.info("Receipt Tax Bot running — press Ctrl+C to stop.")
    # bootstrap_retries=-1: if the FIRST connection to Telegram stalls (common
    #   on just-connected WiFi), keep retrying forever instead of crashing.
    # drop_pending_updates=False: receipts sent while the PC was off sit in
    #   Telegram's queue (~24h) and are processed one by one on startup.
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=False,
                    bootstrap_retries=-1)


if __name__ == "__main__":
    main()
