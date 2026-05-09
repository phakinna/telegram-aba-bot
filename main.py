import os
import re
import sqlite3
import datetime
import pandas as pd

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes
)

# =========================================================
# CONFIG
# =========================================================

TOKEN = "PUT_YOUR_BOT_TOKEN"

DB_FILE = "transactions.db"

# =========================================================
# DATABASE
# =========================================================

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

# =========================================================
# SAVE TO DATABASE
# =========================================================

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

        saved = True

    except:

        saved = False

    conn.close()

    return saved

# =========================================================
# CLEAN TEXT
# =========================================================

def clean_text(text):

    text = str(text)

    text = re.sub(r"\s+", " ", text)

    return text.strip()

# =========================================================
# EXTRACT DATE
# =========================================================

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

                elif raw.count("-") == 2:

                    if len(raw.split("-")[0]) == 4:

                        dt = datetime.datetime.strptime(
                            raw,
                            "%Y-%m-%d"
                        )

                    else:

                        dt = datetime.datetime.strptime(
                            raw,
                            "%d-%m-%Y"
                        )

                return dt.strftime("%Y-%m-%d")

            except:
                pass

    return datetime.datetime.now().strftime(
        "%Y-%m-%d"
    )

# =========================================================
# EXTRACT AMOUNT
# =========================================================

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

    # យកលេខតូចជាងគេ
    # ព្រោះ Balance ជាញឹកញាប់ធំជាង Amount

    return min(amounts)

# =========================================================
# EXTRACT NAME
# =========================================================

def extract_name(text):

    text = clean_text(text)

    # Khmer name
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

# =========================================================
# DETECT CURRENCY
# =========================================================

def detect_currency(text):

    text = text.upper()

    if "KHR" in text or "៛" in text:
        return "KHR"

    return "USD"

# =========================================================
# PARSE FILE
# =========================================================

def parse_file(file_path):

    results = []

    try:

        # Excel
        if file_path.endswith((".xlsx", ".xls")):

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

            row_text = clean_text(row_text)

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

# =========================================================
# RECEIVE FILE
# =========================================================

async def receive_file(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    doc = update.message.document

    filename = doc.file_name.lower()

    allowed = (
        ".xlsx",
        ".xls",
        ".csv"
    )

    if not filename.endswith(allowed):

        await update.message.reply_text(
            "❌ សូមផ្ញើ .xlsx .xls ឬ .csv"
        )

        return

    loading = await update.message.reply_text(
        "⏳ កំពុងវិភាគ File..."
    )

    temp_file = f"temp_{doc.file_id}"

    file = await context.bot.get_file(
        doc.file_id
    )

    await file.download_to_drive(
        temp_file
    )

    data = parse_file(temp_file)

    if not data:

        await loading.edit_text(
            "❌ មិនអាចស្រង់ទិន្នន័យបាន"
        )

        return

    report = "📊 របាយការណ៍\n\n"

    total = 0
    count = 0

    for item in data:

        saved = save_transaction(

            item["date"],
            item["name"],
            item["amount"],
            item["currency"]

        )

        if saved:

            count += 1

            total += item["amount"]

            if item["currency"] == "USD":

                report += (
                    f"👤 {item['name']}: "
                    f"{item['amount']:,.2f}$\n"
                )

            else:

                report += (
                    f"👤 {item['name']}: "
                    f"{item['amount']:,.0f}៛\n"
                )

    report += (
        "\n━━━━━━━━━━━━━━\n"
    )

    report += (
        f"👥 ចំនួនអ្នកផ្ទេរ: {count}\n"
    )

    report += (
        f"💰 សរុប: {total:,.2f}\n"
    )

    await loading.edit_text(report)

    os.remove(temp_file)

# =========================================================
# SUMMARY
# =========================================================

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
            f"❌ ថ្ងៃ {target_date} មិនមានទិន្នន័យ"
        )

        return

    report = (
        f"📊 ថ្ងៃ {target_date}\n"
    )

    report += (
        "━━━━━━━━━━━━━━\n\n"
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

    report += (
        "\n━━━━━━━━━━━━━━\n"
    )

    report += (
        f"👥 អ្នកផ្ទេរ: {len(rows)}\n"
    )

    if total_usd > 0:

        report += (
            f"💰 USD: {total_usd:,.2f}$\n"
        )

    if total_khr > 0:

        report += (
            f"💰 KHR: {total_khr:,.0f}៛\n"
        )

    await update.message.reply_text(
        report
    )

# =========================================================
# START
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    text = """
🤖 ABA STATEMENT BOT

📥 Support:
• .xlsx
• .xls
• .csv

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

    app = Application.builder().token(
        TOKEN
    ).build()

    app.add_handler(
        CommandHandler(
            "start",
            start
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
            receive_file
        )
    )

    print("🚀 BOT RUNNING...")

    app.run_polling()

# =========================================================

if __name__ == "__main__":

    main()
