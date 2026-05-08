import os
import re
import datetime
import sqlite3
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# ផ្លូវ Database សម្រាប់ Railway
DB_FILE = "aba_database.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS transfers 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, 
                       sender_name TEXT, amount TEXT, currency TEXT)''')
    conn.commit()
    conn.close()

def parse_aba_statement(file_path):
    try:
        # អាន Excel ដោយមិនយក Header ដើម្បីស្កេនគ្រប់ជួរ
        df = pd.read_excel(file_path, engine='openpyxl', header=None)
        data_list = []
        
        for idx, row in df.iterrows():
            line_text = " ".join([str(val) for val in row.values if pd.notna(val)])
            
            # ស្វែងរកទឹកប្រាក់ (Amount)
            amt_match = re.search(r'(\d{1,3}(?:,\d{3})*(?:\.\d{2}))', line_text)
            if not amt_match: continue
            
            amount = amt_match.group(1).replace(",", "")
            if float(amount) <= 0: continue
            
            # ស្វែងរកឈ្មោះ (Sender Name) - Fix សម្រាប់ការស្រង់ឈ្មោះម្នាក់ៗ
            name = "Unknown"
            # ស្កេនរកអក្សរធំដែលនៅជាប់ពាក្យគន្លឹះ ឬអក្សរខ្មែរ
            name_match = re.search(r'(?:FROM|BY|TRANSFER|PAYMENT)\s+([A-Z\s]{3,})|([\u1780-\u17FF\s]{3,})', line_text, re.I)
            if name_match:
                name = (name_match.group(1) or name_match.group(2)).strip()

            # ស្វែងរកថ្ងៃខែ
            dt = datetime.datetime.now().strftime("%Y-%m-%d")
            date_match = re.search(r'(\d{2}[-/]\d{2}[-/]\d{4})|(\d{4}-\d{2}-\d{2})', line_text)
            if date_match: dt = date_match.group(0).replace("/", "-")

            curr = "KHR" if "KHR" in line_text or "៛" in line_text else "USD"
            data_list.append((dt, name, amount, curr))
        return data_list
    except: return []

async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    # យកថ្ងៃនេះជាគោល
    target_date = context.args[0] if context.args else datetime.datetime.now().strftime("%Y-%m-%d")

    cursor.execute("SELECT sender_name, amount, currency FROM transfers WHERE created_at LIKE ?", (f"%{target_date}%",))
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        await update.message.reply_text(f"📊 ថ្ងៃ `{target_date}` មិនទាន់មានទិន្នន័យទេ។")
        return

    # រៀបចំទម្រង់សារឱ្យដូចរូបភាពគំរូ
    report = f"📊 **ថ្ងៃ {target_date}**\n"
    report += "━━━━━━━━━━━━━━━━━━━━\n\n"
    
    total_val = 0.0
    for name, amt, curr in rows:
        val = float(amt)
        total_val += val
        symbol = "$" if curr == "USD" else "៛"
        report += f"👤 {name}:  `{val:,.2f}{symbol}`\n"

    report += "\n" + "━" * 15 + "\n"
    report += f"💰 **សរុប:  {total_val:,.2f}$**" # ឧទាហរណ៍ជា USD

    await update.message.reply_text(report, parse_mode="Markdown")

async def on_receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    temp_name = f"temp_{doc.file_id}.xlsx"
    file = await context.bot.get_file(doc.file_id)
    await file.download_to_drive(temp_name)

    results = parse_aba_statement(temp_name)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    for dt, nm, am, cr in results:
        cursor.execute("INSERT INTO transfers (created_at, sender_name, amount, currency) VALUES (?,?,?,?)", (dt, nm, am, cr))
    conn.commit()
    conn.close()
    
    await update.message.reply_text(f"✅ បានរក្សាទុក {len(results)} ប្រតិបត្តិការ រួចរាល់!")
    if os.path.exists(temp_name): os.remove(temp_name)

if __name__ == '__main__':
    init_db()
    TOKEN = os.getenv("BOT_TOKEN") # ទាញយកពី Railway Environment Variable
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("summary", show_summary))
    app.add_handler(MessageHandler(filters.Document.ALL, on_receive_file))
    app.run_polling()
