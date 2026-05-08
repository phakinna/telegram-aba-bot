import os
import re
import sqlite3
import datetime
import pandas as pd

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# =========================================================
# CONFIG
# =========================================================

TOKEN = os.getenv("BOT_TOKEN")  # ដាក់ Token ក្នុង Environment Variable
DB_FILE = "aba_database.db"

# =========================================================
# DATABASE
# =========================================================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transfers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            sender_name TEXT,
            amount REAL,
            currency TEXT
        )
    """)

    conn.commit()
    conn.close()

# =========================================================
# CLEAN TEXT
# =========================================================

def clean_text(text):
    text = str(text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# =========================================================
# EXTRACT NAME
# =========================================================

def extract_name(text):

    text = clean_text(text)

    patterns = [
        r'(?:FROM|BY|TRFR|PAYMENT FROM|TRANSFER FROM|ពី)\s+([A-Za-z\s]{3,})',
        r'([\u1780-\u17FF\s]{3,})'
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.I)

        if match:
            name = clean_text(match.group(1))

            # កាត់ពាក្យដែលមិនចាំបាច់
            blacklist = [
                "USD",
                "KHR",
                "ABA",
                "TRANSFER",
                "PAYMENT"
            ]

            valid = True

            for bad in blacklist:
                if bad.lower() in name.lower():
                    valid = False
                    break

            if valid and len(name) >= 3:
                return name

    return "Unknown Sender"

# =========================================================
# EXTRACT DATE
# =========================================================

def extract_date(text):

    patterns = [
        r'(\d{4}-\d{2}-\d{2})',
        r'(\d{2}/\d{2}/\d{4})',
        r'(\d{2}-\d{2}-\d{4})'
    ]

    for pattern in patterns:
        match = re.search(pattern, text)

        if match:
            return match.group(1).replace("/", "-")

    return datetime.datetime.now().strftime("%Y-%m-%d")

# =========================================================
# EXTRACT AMOUNT
# =========================================================

def extract_amount(row):

    numbers = []

    for val in row.values:

        if pd.isna(val):
            continue

        text = str(val).replace(",", "").strip()

        try:
            num = float(text)

            if num > 0:
                numbers.append(num)

        except:
            pass

    if not numbers:
        return None

    # យកលេខធំបំផុតក្នុង row
    return max(numbers)

# =========================================================
# DETECT CURRENCY
# =========================================================

def detect_currency(text):

    upper = text.upper()

    if "KHR" in upper or "៛" in text or "រៀល" in text:
        return "KHR"

    return "USD"

# =========================================================
# PARSE ABA STATEMENT
# =========================================================

def parse_aba_statement(file_path):

    results = []

    try:

        df = pd.read_excel(
            file_path,
            engine="openpyxl",
            header=None
        )

        for _, row in df.iterrows():

            row_text = " ".join(
                [str(x) for x in row.values if pd.notna(x)]
            )

            row_text = clean_text(row_text)

            # Amount
            amount = extract_amount(row)

            if not amount:
                continue

            # Name
            name = extract_name(row_text)

            # Date
            date = extract_date(row_text)

            # Currency
            currency = detect_currency(row_text)

            results.append((
                date,
                name,
                amount,
                currency
            ))

        return results

    except Exception as e:
        print("PARSE ERROR:", e)
        return []

# =========================================================
# SAVE TO DATABASE
# =========================================================

def save_to_db(data):

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    new_count = 0

    for dt, name, amount, currency in data:

        cursor.execute("""
            SELECT 1
            FROM transfers
            WHERE created_at=?
            AND sender_name=?
            AND amount=?
        """, (dt, name, amount))

        exists = cursor.fetchone()

        if not exists:

            cursor.execute("""
                INSERT INTO transfers (
                    created_at,
                    sender_name,
                    amount,
                    currency
                )
                VALUES (?, ?, ?, ?)
            """, (dt, name, amount, currency))

            new_count += 1

    conn.commit()
    conn.close()

    return new_count

# =========================================================
# SUMMARY COMMAND
# =========================================================

async def show_summary(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    target_date = (
        context.args[0]
        if context.args
        else datetime.datetime.now().strftime("%Y-%m-%d")
    )

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT sender_name, amount, currency
        FROM transfers
        WHERE created_at LIKE ?
    """, (f"%{target_date}%",))

    rows = cursor.fetchall()

    conn.close()

    if not rows:

        await update.message.reply_text(
            f"📊 ថ្ងៃ {target_date} មិនមានទិន្នន័យទេ។"
        )

        return

    report = f"📊 <b>ថ្ងៃ {target_date}</b>\n"
    report += "━━━━━━━━━━━━━━━━━━\n\n"

    total_usd = 0
    total_khr = 0

    for name, amount, currency in rows:

        amount = float(amount)

        if currency == "USD":

            total_usd += amount

            report += (
                f"👤 {name}: "
                f"<code>{amount:,.2f}$</code>\n"
            )

        else:

            total_khr += amount

            report += (
                f"👤 {name}: "
                f"<code>{amount:,.0f}៛</code>\n"
            )

    report += "\n━━━━━━━━━━━━━━━━━━\n"

    report += "💰 <b>សរុប:</b> "

    if total_usd > 0:
        report += f"<code>{total_usd:,.2f}$</code> "

    if total_khr > 0:
        report += f"+ <code>{total_khr:,.0f}៛</code>"

    await update.message.reply_text(
        report,
        parse_mode="HTML"
    )

# =========================================================
# RECEIVE FILE
# =========================================================

async def on_receive_file(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    doc = update.message.document

    if not doc.file_name.lower().endswith((".xlsx", ".xls")):

        await update.message.reply_text(
            "❌ សូមផ្ញើ Excel File (.xlsx/.xls)"
        )

        return

    loading = await update.message.reply_text(
        "⏳ កំពុងវិភាគ ABA Statement..."
    )

    temp_file = f"temp_{doc.file_id}.xlsx"

    try:

        file = await context.bot.get_file(doc.file_id)

        await file.download_to_drive(temp_file)

        results = parse_aba_statement(temp_file)

        if not results:

            await loading.edit_text(
                "❌ មិនអាចអានទិន្នន័យបានទេ។"
            )

            return

        new_count = save_to_db(results)

        await loading.edit_text(
            f"✅ បានរកឃើញ {new_count} ប្រតិបត្តិការថ្មី"
        )

    except Exception as e:

        await loading.edit_text(
            f"❌ Error: {e}"
        )

    finally:

        if os.path.exists(temp_file):
            os.remove(temp_file)

# =========================================================
# START COMMAND
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    text = """
🤖 ABA Statement Bot

📥 ផ្ញើ Excel ABA Statement
📊 ប្រើ /summary ដើម្បីមើលរបាយការណ៍

ឧទាហរណ៍:
    /summary
    /summary 2026-05-08
"""

    await update.message.reply_text(text)

# =========================================================
# MAIN
# =========================================================

def main():

    init_db()

    if not TOKEN:
        print("❌ BOT_TOKEN not found")
        return

    app = (
        ApplicationBuilder()
        .token(TOKEN)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("summary", show_summary))

    app.add_handler(
        MessageHandler(
            filters.Document.ALL,
            on_receive_file
        )
    )

    print("🚀 ABA BOT RUNNING...")

    app.run_polling()

# =========================================================

if __name__ == "__main__":
    main()
