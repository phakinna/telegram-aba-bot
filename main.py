import os
import sqlite3
import pandas as pd
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes
)

# =========================
# TOKEN
# =========================

TOKEN = "ដាក់_BOT_TOKEN_នៅទីនេះ"

DB_FILE = "aba.db"

# =========================
# DATABASE
# =========================

def init_db():

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS transfers (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        sender TEXT,
        amount REAL
    )
    """)

    conn.commit()
    conn.close()

# =========================
# SAVE DATA
# =========================

def save_data(name, amount):

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    cursor.execute("""
    INSERT INTO transfers (
        sender,
        amount
    )
    VALUES (?, ?)
    """, (name, amount))

    conn.commit()
    conn.close()

# =========================
# READ EXCEL
# =========================

def parse_excel(file_path):

    data = []

    df = pd.read_excel(
        file_path,
        engine="openpyxl"
    )

    for _, row in df.iterrows():

        row_text = " ".join([
            str(x)
            for x in row.values
        ])

        # រក Amount
        amount = None

        for val in row.values:

            try:

                num = float(
                    str(val).replace(",", "")
                )

                if num > 0:

                    amount = num
                    break

            except:
                pass

        if not amount:
            continue

        # ឈ្មោះ
        words = row_text.split()

        name = "Unknown"

        if len(words) >= 2:
            name = words[0] + " " + words[1]

        data.append((
            name,
            amount
        ))

    return data

# =========================
# RECEIVE FILE
# =========================

async def receive_file(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    doc = update.message.document

    if not doc.file_name.endswith(".xlsx"):

        await update.message.reply_text(
            "សូមផ្ញើ Excel File"
        )

        return

    msg = await update.message.reply_text(
        "កំពុងវិភាគ..."
    )

    file = await context.bot.get_file(
        doc.file_id
    )

    temp = "temp.xlsx"

    await file.download_to_drive(temp)

    rows = parse_excel(temp)

    total = 0

    report = "📊 របាយការណ៍\n\n"

    for name, amount in rows:

        save_data(name, amount)

        total += amount

        report += (
            f"👤 {name}: "
            f"{amount:.2f}$\n"
        )

    report += (
        f"\n💰 សរុប: {total:.2f}$"
    )

    await msg.edit_text(report)

    os.remove(temp)

# =========================
# SUMMARY
# =========================

async def summary(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    conn = sqlite3.connect(DB_FILE)

    cursor = conn.cursor()

    cursor.execute("""
    SELECT sender, amount
    FROM transfers
    """)

    rows = cursor.fetchall()

    conn.close()

    if not rows:

        await update.message.reply_text(
            "មិនមានទិន្នន័យ"
        )

        return

    total = 0

    text = "📊 Summary\n\n"

    for name, amount in rows:

        total += amount

        text += (
            f"👤 {name}: "
            f"{amount:.2f}$\n"
        )

    text += (
        f"\n💰 សរុប: {total:.2f}$"
    )

    await update.message.reply_text(text)

# =========================
# START
# =========================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    await update.message.reply_text(
        "ផ្ញើ ABA Excel Statement"
    )

# =========================
# MAIN
# =========================

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

    print("BOT RUNNING...")

    app.run_polling()

if __name__ == "__main__":
    main()
