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

TOKEN = "YOUR_BOT_TOKEN"
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
        currency TEXT,

        UNIQUE(created_at, sender_name, amount)

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
# EXTRACT DATE
# =========================================================

def extract_date(value):

    try:

        # Excel datetime
        if isinstance(value, (datetime.datetime, pd.Timestamp)):

            return value.strftime("%Y-%m-%d")

        text = str(value).strip()

        # YYYY-MM-DD
        match = re.search(r'(\d{4}-\d{2}-\d{2})', text)

        if match:
            return match.group(1)

        # DD/MM/YYYY
        match = re.search(r'(\d{2}/\d{2}/\d{4})', text)

        if match:

            dt = datetime.datetime.strptime(
                match.group(1),
                "%d/%m/%Y"
            )

            return dt.strftime("%Y-%m-%d")

    except:
        pass

    return datetime.datetime.now().strftime("%Y-%m-%d")

# =========================================================
# DETECT CURRENCY
# =========================================================

def detect_currency(text):

    upper = text.upper()

    if "KHR" in upper or "៛" in text:
        return "KHR"

    return "USD"

# =========================================================
# EXTRACT NAME
# =========================================================

def extract_name(text):

    text = clean_text(text)

    # Khmer Name
    kh_match = re.search(
        r'([\u1780-\u17FF]{2,}(?:\s+[\u1780-\u17FF]{2,})+)',
        text
    )

    if kh_match:

        return kh_match.group(1).strip()

    # English Name
    en_match = re.search(
        r'([A-Z][A-Z\s]{3,})',
        text
    )

    if en_match:

        return clean_text(en_match.group(1))

    return "Unknown Sender"

# =========================================================
# EXTRACT AMOUNT
# =========================================================

def extract_amount(text):

    matches = re.findall(
        r'(\d[\d,]*\.\d{2})',
        text
    )

    if not matches:
        return None

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

    # យក Amount តូចជាងគេ
    # ព្រោះ ABA ជាញឹកញាប់ Balance ធំជាង

    return min(amounts)

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

            row_text = " ".join([
                str(x)
                for x in row.values
                if pd.notna(x)
            ])

            row_text = clean_text(row_text)

            # Skip Empty
            if len(row_text) < 5:
                continue

            # Extract Amount
            amount = extract_amount(row_text)

            if not amount:
                continue

            # Extract Name
            name = extract_name(row_text)

            # Extract Date
            date = None

            for val in row.values:

                date = extract_date(val)

                if date:
                    break

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
# SAVE DATABASE
# =========================================================

def save_to_db(data):

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    count = 0

    for dt, name, amount, currency in data:

        try:

            cursor.execute("""
            INSERT INTO transfers (

                created_at,
                sender_name,
                amount,
                currency

            )
            VALUES (?, ?, ?, ?)
            """, (
                dt,
                name,
                amount,
                currency
            ))

            count += 1

        except sqlite3.IntegrityError:

            pass

    conn.commit()
    conn.close()

    return count

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
    SELECT

        sender_name,
        amount,
        currency

    FROM transfers

    WHERE DATE(created_at)=DATE(?)

    ORDER BY amount ASC
    """, (target_date,))

    rows = cursor.fetchall()

    conn.close()

    if not rows:

        await update.message.reply_text(
            f"📊 ថ្ងៃ {target_date} មិនមានទិន្នន័យទេ។"
        )

        return

    report = (
        f"📊 <b>ថ្ងៃ {target_date}</b>\n"
    )

    report += (
        "━━━━━━━━━━━━━━━━━━\n\n"
    )

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

    report += (
        "\n━━━━━━━━━━━━━━━━━━\n"
    )

    report += "💰 <b>សរុប:</b> "

    if total_usd > 0:

        report += (
            f"<code>{total_usd:,.2f}$</code> "
        )

    if total_khr > 0:

        report += (
            f"+ <code>{total_khr:,.0f}៛</code>"
        )

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

    if not doc.file_name.lower().endswith((
        ".xlsx",
        ".xls"
    )):

        await update.message.reply_text(
            "❌ សូមផ្ញើ Excel File"
        )

        return

    loading = await update.message.reply_text(
        "⏳ កំពុងវិភាគ ABA Statement..."
    )

    temp_file = f"temp_{doc.file_id}.xlsx"

    try:

        file = await context.bot.get_file(
            doc.file_id
        )

        await file.download_to_drive(
            temp_file
        )

        results = parse_aba_statement(
            temp_file
        )

        if not results:

            await loading.edit_text(
                "❌ មិនអាចអានទិន្នន័យបានទេ"
            )

            return

        count = save_to_db(results)

        await loading.edit_text(
            f"✅ បានស្រង់ {count} ប្រតិបត្តិការ!"
        )

    except Exception as e:

        await loading.edit_text(
            f"❌ Error: {e}"
        )

    finally:

        if os.path.exists(temp_file):
            os.remove(temp_file)

# =========================================================
# START
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    text = """
🤖 ABA STATEMENT BOT

📥 ផ្ញើ ABA Excel Statement

📊 Commands:

/summary

ឬ

/summary 2026-05-08
"""

    await update.message.reply_text(text)

# =========================================================
# MAIN
# =========================================================

def main():

    init_db()

    app = (
        ApplicationBuilder()
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
            "summary",
            show_summary
        )
    )

    app.add_handler(
        MessageHandler(
            filters.Document.ALL,
            on_receive_file
        )
    )

    print("🚀 BOT RUNNING...")

    app.run_polling()

# =========================================================

if __name__ == "__main__":

    main()
