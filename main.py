import os
import re
import sqlite3
import logging
import datetime
import pandas as pd

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ======================================================
# LOGGING
# ======================================================

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# ======================================================
# CONFIG
# ======================================================

TOKEN = os.getenv("BOT_TOKEN")

DB_FILE = "transactions.db"

if not TOKEN:
    print("❌ BOT_TOKEN NOT FOUND")
    exit()

# ======================================================
# DATABASE
# ======================================================

def init_db():

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS transfers (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        transfer_date TEXT,
        sender_name TEXT,
        amount REAL,
        currency TEXT,

        UNIQUE(
            transfer_date,
            sender_name,
            amount
        )
    )
    """)

    conn.commit()
    conn.close()

# ======================================================
# SAVE TRANSACTION
# ======================================================

def save_transaction(
    transfer_date,
    sender_name,
    amount,
    currency
):

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    try:

        cursor.execute("""
        INSERT INTO transfers (

            transfer_date,
            sender_name,
            amount,
            currency

        )
        VALUES (?, ?, ?, ?)
        """, (
            transfer_date,
            sender_name,
            amount,
            currency
        ))

        conn.commit()

    except:
        pass

    conn.close()

# ======================================================
# CLEAN TEXT
# ======================================================

def clean_text(text):

    text = str(text)

    text = re.sub(r"\s+", " ", text)

    return text.strip()

# ======================================================
# EXTRACT DATE
# ======================================================

def extract_date(text):

    text = clean_text(text)

    patterns = [

        r'(\d{4}-\d{2}-\d{2})',

        r'(\d{2}/\d{2}/\d{4})',

        r'(\d{2}-\d{2}-\d{4})'

    ]

    for pattern in patterns:

        match = re.search(pattern, text)

        if match:

            raw = match.group(1)

            try:

                if "/" in raw:

                    dt = datetime.datetime.strptime(
                        raw,
                        "%d/%m/%Y"
                    )

                elif len(raw.split("-")[0]) == 4:

                    dt = datetime.datetime.strptime(
                        raw,
                        "%Y-%m-%d"
                    )

                else:

                    dt = datetime.datetime.strptime(
                        raw,
                        "%d-%m-%Y"
                    )

                return dt.strftime(
                    "%Y-%m-%d"
                )

            except:
                pass

    return datetime.datetime.now().strftime(
        "%Y-%m-%d"
    )

# ======================================================
# EXTRACT AMOUNT
# ======================================================

def extract_amount(text):

    matches = re.findall(
        r'(\d[\d,]*\.\d{2})',
        text
    )

    amounts = []

    for amt in matches:

        try:

            value = float(
                amt.replace(",", "")
            )

            if value > 0:
                amounts.append(value)

        except:
            pass

    if not amounts:
        return None

    return min(amounts)

# ======================================================
# EXTRACT NAME
# ======================================================

def extract_name(text):

    text = clean_text(text)

    # Khmer Name

    kh_match = re.search(
        r'([\u1780-\u17FF]{2,}(?:\s+[\u1780-\u17FF]{2,})+)',
        text
    )

    if kh_match:

        return kh_match.group(1)

    # English Name

    en_match = re.search(
        r'([A-Z][a-z]+\s+[A-Z][a-z]+)',
        text
    )

    if en_match:

        return en_match.group(1)

    return "Unknown"

# ======================================================
# DETECT CURRENCY
# ======================================================

def detect_currency(text):

    upper = text.upper()

    if "KHR" in upper or "៛" in text:
        return "KHR"

    return "USD"

# ======================================================
# PARSE FILE
# ======================================================

def parse_file(file_path):

    results = []

    try:

        # Excel

        if file_path.endswith((
            ".xlsx",
            ".xls"
        )):

            df = pd.read_excel(
                file_path,
                engine="openpyxl",
                header=None
            )

        # CSV

        elif file_path.endswith(".csv"):

            df = pd.read_csv(
                file_path,
                header=None
            )

        else:

            return []

        for _, row in df.iterrows():

            row_text = " ".join([
                str(x)
                for x in row.values
                if pd.notna(x)
            ])

            row_text = clean_text(
                row_text
            )

            if len(row_text) < 5:
                continue

            amount = extract_amount(
                row_text
            )

            if not amount:
                continue

            sender_name = extract_name(
                row_text
            )

            transfer_date = extract_date(
                row_text
            )

            currency = detect_currency(
                row_text
            )

            results.append({

                "date": transfer_date,

                "name": sender_name,

                "amount": amount,

                "currency": currency

            })

        return results

    except Exception as e:

        print("PARSE ERROR:", e)

        return []

# ======================================================
# START
# ======================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    text = """
🤖 ABA STATEMENT BOT

📥 Supported Files:
• .xlsx
• .xls
• .csv

📊 Commands:

/summary

or

/summary 2026-05-08
"""

    await update.message.reply_text(
        text
    )

# ======================================================
# PING
# ======================================================

async def ping(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    await update.message.reply_text(
        "✅ BOT WORKING"
    )

# ======================================================
# HANDLE FILE
# ======================================================

async def handle_file(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    try:

        document = update.message.document

        if not document:

            await update.message.reply_text(
                "❌ No file"
            )

            return

        filename = (
            document.file_name.lower()
        )

        if not filename.endswith((
            ".xlsx",
            ".xls",
            ".csv"
        )):

            await update.message.reply_text(
                "❌ Unsupported file"
            )

            return

        loading = await update.message.reply_text(
            "⏳ Reading file..."
        )

        telegram_file = await context.bot.get_file(
            document.file_id
        )

        temp_file = (
            f"temp_{document.file_id}"
        )

        await telegram_file.download_to_drive(
            temp_file
        )

        results = parse_file(
            temp_file
        )

        if not results:

            await loading.edit_text(
                "❌ No transaction found"
            )

            return

        total_usd = 0
        total_khr = 0

        count = 0

        report = "📊 REPORT\n\n"

        for item in results:

            save_transaction(

                item["date"],
                item["name"],
                item["amount"],
                item["currency"]

            )

            count += 1

            if item["currency"] == "USD":

                total_usd += item["amount"]

                report += (
                    f"👤 {item['name']}: "
                    f"{item['amount']:,.2f}$\n"
                )

            else:

                total_khr += item["amount"]

                report += (
                    f"👤 {item['name']}: "
                    f"{item['amount']:,.0f}៛\n"
                )

        report += "\n━━━━━━━━━━━━━━\n"

        report += (
            f"👥 Transfers: {count}\n"
        )

        if total_usd > 0:

            report += (
                f"💰 USD: "
                f"{total_usd:,.2f}$\n"
            )

        if total_khr > 0:

            report += (
                f"💰 KHR: "
                f"{total_khr:,.0f}៛\n"
            )

        await loading.edit_text(
            report[:4000]
        )

        if os.path.exists(temp_file):
            os.remove(temp_file)

    except Exception as e:

        await update.message.reply_text(
            f"❌ ERROR:\n{e}"
        )

# ======================================================
# SUMMARY
# ======================================================

async def summary(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if context.args:

        target_date = context.args[0]

    else:

        target_date = datetime.datetime.now().strftime(
            "%Y-%m-%d"
        )

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    cursor.execute("""
    SELECT

        sender_name,
        amount,
        currency

    FROM transfers

    WHERE transfer_date=?

    ORDER BY amount ASC
    """, (target_date,))

    rows = cursor.fetchall()

    conn.close()

    if not rows:

        await update.message.reply_text(
            f"❌ No data for {target_date}"
        )

        return

    report = (
        f"📊 DATE {target_date}\n\n"
    )

    total_usd = 0
    total_khr = 0

    for name, amount, currency in rows:

        amount = float(amount)

        if currency == "USD":

            total_usd += amount

            report += (
                f"👤 {name}: "
                f"{amount:,.2f}$\n"
            )

        else:

            total_khr += amount

            report += (
                f"👤 {name}: "
                f"{amount:,.0f}៛\n"
            )

    report += "\n━━━━━━━━━━━━━━\n"

    report += (
        f"👥 Transfers: {len(rows)}\n"
    )

    if total_usd > 0:

        report += (
            f"💰 USD: "
            f"{total_usd:,.2f}$\n"
        )

    if total_khr > 0:

        report += (
            f"💰 KHR: "
            f"{total_khr:,.0f}៛\n"
        )

    await update.message.reply_text(
        report[:4000]
    )

# ======================================================
# MAIN
# ======================================================

def main():

    init_db()

    app = (
        Application.builder()
        .token(TOKEN)
        .build()
    )

    app.add_handler(
        CommandHandler(
            "start",
            start
        )
    )

    app.add_handler(
        CommandHandler(
            "ping",
            ping
        )
    )

    app.add_handler(
        CommandHandler(
            "summary",
            summary
        )
    )

    app.add_handler(
        MessageHandler(
            filters.Document.ALL,
            handle_file
        )
    )

    print("🚀 BOT RUNNING...")

    app.run_polling()

# ======================================================

if __name__ == "__main__":

    main()
