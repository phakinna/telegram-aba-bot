import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes
)

# ===============================
# BOT TOKEN
# ===============================

TOKEN = "YOUR_REAL_BOT_TOKEN"

# ===============================
# START
# ===============================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    await update.message.reply_text(
        "🤖 Bot Working!\n\nSend Excel File (.xlsx)"
    )

# ===============================
# HANDLE FILE
# ===============================

async def handle_file(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    try:

        document = update.message.document

        if not document.file_name.endswith(".xlsx"):

            await update.message.reply_text(
                "❌ Please send .xlsx file"
            )

            return

        await update.message.reply_text(
            "⏳ Reading file..."
        )

        telegram_file = await context.bot.get_file(
            document.file_id
        )

        await telegram_file.download_to_drive(
            "temp.xlsx"
        )

        df = pd.read_excel(
            "temp.xlsx"
        )

        total = 0

        text = "📊 REPORT\n\n"

        for _, row in df.iterrows():

            row_text = " ".join(
                [str(x) for x in row.values]
            )

            text += f"{row_text}\n"

            # Find amount
            for val in row.values:

                try:

                    amount = float(
                        str(val).replace(",", "")
                    )

                    if amount > 0:

                        total += amount
                        break

                except:
                    pass

        text += (
            f"\n💰 TOTAL: {total:,.2f}"
        )

        await update.message.reply_text(text[:4000])

    except Exception as e:

        await update.message.reply_text(
            f"❌ ERROR:\n{e}"
        )

# ===============================
# MAIN
# ===============================

def main():

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
        MessageHandler(
            filters.Document.ALL,
            handle_file
        )
    )

    print("BOT RUNNING...")

    app.run_polling()

# ===============================

if __name__ == "__main__":

    main()
