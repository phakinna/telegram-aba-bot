import os
import re
import datetime
import sqlite3
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# --- ផ្លូវ Database (សម្រាប់ Railway) ---
DB_FILE = "aba_database.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS transfers 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                       created_at TEXT, 
                       sender_name TEXT, 
                       amount TEXT, 
                       currency TEXT)''')
    conn.commit()
    conn.close()

def parse_aba_statement(file_path):
    try:
        # អាន Excel ដោយមិនយក Header ដើម្បីកុំឱ្យខុសជួរ
        df = pd.read_excel(file_path, engine='openpyxl', header=None)
        data_list = []
        
        for idx, row in df.iterrows():
            # បំប្លែងជួរដេកទាំងមូលទៅជាអត្ថបទ (Text) ដើម្បីស្កេន
            line_text = " ".join([str(val) for val in row.values if pd.notna(val)])
            
            # ១. ស្វែងរកទឹកប្រាក់ (Amount) - ABA ច្រើនតែនៅ Column "C" ឬ "D" ក្នុង Excel
            # យើងស្វែងរកលេខដែលមានចុច .00
            amt_match = re.search(r'(\d{1,3}(?:,\d{3})*(?:\.\d{2}))', line_text)
            if not amt_match:
                continue
            
            amount = amt_match.group(1).replace(",", "")
            if float(amount) <= 0:
                continue
            
            # ២. ស្វែងរកឈ្មោះ (Sender Name)
            # យើងស្កេនរកអក្សរធំ (English Name) ឬអក្សរខ្មែរដែលនៅជិតពាក្យគន្លឹះ
            name = "Unknown Sender"
            # ស្វែងរកឈ្មោះបន្ទាប់ពីពាក្យ From, TRFR, ឬ ពី
            name_match = re.search(r'(?:FROM|BY|TRFR|PAYMENT)\s+([A-Z\s]{3,})|([\u1780-\u17FF\s]{3,})', line_text, re.I)
            if name_match:
                found_name = (name_match.group(1) or name_match.group(2)).strip()
                if len(found_name) > 3:
                    name = re.sub(r'\s+', ' ', found_name)

            # ៣. ស្វែងរកកាលបរិច្ឆេទ (Date)
            dt = datetime.datetime.now().strftime("%Y-%m-%d")
            date_match = re.search(r'(\d{2}[-/]\d{2}[-/]\d{4})|(\d{4}-\d{2}-\d{2})|(\d{2}[-/][A-Za-z]{3}[-/]\d{4})', line_text)
            if date_match:
                dt_raw = date_match.group(0).replace("/", "-")
                dt = dt_raw # រក្សាសភាពដើមដើម្បីបង្ហាញក្នុង Summary

            # ៤. រកប្រភេទលុយ (Currency)
            curr = "KHR" if any(x in line_text.upper() for x in ["KHR", "៛", "រៀល"]) else "USD"
            
            data_list.append((dt, name, amount, curr))
        
        return data_list
    except Exception as e:
        print(f"Error parsing: {e}")
        return []

async def on_receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')): return
    
    msg = await update.message.reply_text("⏳ កំពុងស្រង់ទិន្នន័យម្នាក់ៗពី Statement...")
    temp_name = f"temp_{doc.file_id}.xlsx"
    file = await context.bot.get_file(doc.file_id)
    await file.download_to_drive(temp_name)

    results = parse_aba_statement(temp_name)
    if not results:
        await msg.edit_text("❌ មិនអាចអានឈ្មោះ ឬទឹកប្រាក់បានទេ។ សូមប្រាកដថា File មិនមាន Password។")
        return

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    count = 0
    for dt, nm, am, cr in results:
        # បង្ការ Duplicate
        cursor.execute("SELECT 1 FROM transfers WHERE sender_name=? AND amount=? AND created_at=?", (nm, am, dt))
        if not cursor.fetchone():
            cursor.execute("INSERT INTO transfers (created_at, sender_name, amount, currency) VALUES (?,?,?,?)", (dt, nm, am, cr))
            count += 1
    conn.commit()
    conn.close()
    
    await msg.edit_text(f"✅ រកឃើញ និងរក្សាទុកបាន `{count}` ប្រតិបត្តិការ!")
    if os.path.exists(temp_name): os.remove(temp_name)

async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    target_date = context.args[0] if context.args else datetime.datetime.now().strftime("%Y-%m-%d")

    cursor.execute("SELECT sender_name, amount, currency FROM transfers WHERE created_at LIKE ?", (f"%{target_date}%",))
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        await update.message.reply_text(f"📊 ថ្ងៃ `{target_date}` មិនមានទិន្នន័យទេ។")
        return

    report = f"📊 **ថ្ងៃ {target_date}**\n"
    report += "━━━━━━━━━━━━━━━━━━━━\n\n"
    
    total_usd = 0.0
    total_khr = 0.0
    
    for name, amt, curr in rows:
        val = float(amt)
        if curr == "USD":
            total_usd += val
            report += f"👤 {name}: `{val:,.2f}$`\n"
        else:
            total_khr += val
            report += f"👤 {name}: `{val:,.0f}៛`\n"

    report += "\n━━━━━━━━━━━━━━━━━━━━\n"
    report += f"💰 **សរុប:** "
    if total_usd > 0: report += f"`{total_usd:,.2f}$` "
    if total_khr > 0: report += f"+ `{total_khr:,.0f}៛`"

    await update.message.reply_text(report, parse_mode="Markdown")

if __name__ == '__main__':
    init_db()
    # ប្រើ Token ដែលអ្នកបានឱ្យមក
    TOKEN = "8663484036:AAEZmsFkVkZdxXNHy4G1Vzj66NScGnIYNpE"
    app = ApplicationBuilder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("summary", show_summary))
    app.add_handler(MessageHandler(filters.Document.ALL, on_receive_file))
    app.run_polling()
