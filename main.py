import os
import re
import sqlite3
import pandas as pd
from decimal import Decimal
from datetime import datetime

from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler

# =========================================================
# CONFIG
# =========================================================
TOKEN = os.getenv("BOT_TOKEN")

if not TOKEN:
    raise ValueError("❌ BOT_TOKEN not found!")

# =========================================================
# DATABASE
# =========================================================
DB_FILE = "payway_bot.db"
conn = sqlite3.connect(DB_FILE, check_same_thread=False)
cursor = conn.cursor()
cursor.execute("""
CREATE TABLE IF NOT EXISTS transfers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    sender_name TEXT,
    amount REAL,
    currency TEXT
)
""")
conn.commit()

# =========================================================
# PARSER (សម្រាប់សារ + XLSX)
# =========================================================
def parse_text(text: str):
    transactions = []
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # ស្រង់ទឹកប្រាក់
        amount_match = re.search(r'ORIGINAL AMOUNT\s+([\d,]+\.?\d*)\s*USD', line, re.IGNORECASE)
        if not amount_match:
            amount_match = re.search(r'([\$៛])\s?([\d,]+(?:\.\d+)?)', line)

        # ស្រង់ឈ្មោះ
        name_match = re.search(r'FROM\s+(.+?)\s*\(', line, re.IGNORECASE)
        if not name_match:
            name_match = re.search(r'PAYMENT FROM\s+(.+?)\s', line, re.IGNORECASE)
        if not name_match:
            name_match = re.search(r'paid by\s+(.+?)\s', line, re.IGNORECASE)

        if not amount_match or not name_match:
            continue

        amount_str = amount_match.group(1).replace(",", "") if 'ORIGINAL AMOUNT' in line else amount_match.group(2).replace(",", "")
        name = name_match.group(1).strip()

        try:
            amount = Decimal(amount_str)
            transactions.append({"name": name, "amount": amount, "currency": "USD"})
        except:
            continue
    return transactions

# =========================================================
# HANDLE FILE XLSX
# =========================================================
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document.file_name.lower().endswith('.xlsx'):
        await update.message.reply_text("❌ គាំទ្រតែ File .xlsx បច្ចុប្បន្ន។")
        return

    await update.message.reply_text("📥 កំពុងវិភាគ Statement...")

    file = await context.bot.get_file(document.file_id)
    temp_path = "temp.xlsx"
    await file.download_to_drive(temp_path)

    try:
        df = pd.read_excel(temp_path)
        count = 0

        for _, row in df.iterrows():
            name = str(row.get('sender_name', row.get('Name', row.get(1, 'Unknown'))))
            amount_str = str(row.get('amount', row.get('Amount', row.get(2, '0'))))

            amount_clean = re.sub(r'[^0-9.]', '', amount_str)
            try:
                amount = Decimal(amount_clean)
                if amount > 0:
                    cursor.execute("""
                    INSERT INTO transfers (date, sender_name, amount, currency)
                    VALUES (?, ?, ?, ?)
                    """, (datetime.now().strftime("%Y-%m-%d"), name, float(amount), "USD"))
                    count += 1
            except:
                continue

        conn.commit()
        await update.message.reply_text(f"✅ **បានរក្សាទុក {count} ប្រតិបត្តិការ** ពី Statement XLSX រួចរាល់។")

    except Exception as e:
        await update.message.reply_text(f"❌ មានបញ្ហា៖ {str(e)}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

# =========================================================
# SUMMARY
# =========================================================
async def summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.now().strftime("%Y-%m-%d")
    cursor.execute("SELECT sender_name, amount FROM transfers WHERE date = ? ORDER BY id", (today,))
    rows = cursor.fetchall()

    if not rows:
        await update.message.reply_text("📊 ថ្ងៃនេះមិនទាន់មានការផ្ទេរ។")
        return

    total = 0
    msg = f"📊 **សរុបថ្ងៃនេះ ({today})**\n\n"
    for i, (name, amt) in enumerate(rows, 1):
        amount = Decimal(amt)
        msg += f"{i}. {name} — {amount:,.2f} USD\n"
        total += amount

    msg += f"\n━━━━━━━━━━━━━━\n"
    msg += f"👥 សរុប {len(rows)} នាក់\n"
    msg += f"💰 **សរុបទាំងអស់:** {total:,.2f} USD"

    await update.message.reply_text(msg, parse_mode='Markdown')

# =========================================================
# MAIN
# =========================================================
def main():
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler(["start", "summary", "today"], summary))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, lambda u, c: u.message.reply_text("ផ្ញើ Statement XLSX ឬសារធនាគារ មកខ្ញុំ")))

    print("🚀 Payway ABA Bot ដំណើរការរួចរាល់!")
    app.run_polling()

if __name__ == "__main__":
    main()
