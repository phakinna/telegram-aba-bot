import os
import re
import sqlite3
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

DB_FILE = "/app/data/aba_database.db"

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
        df = pd.read_excel(file_path, engine='openpyxl', header=None)
        data_list = []
        
        for _, row in df.iterrows():
            line = " ".join([str(val) for val in row.values if pd.notna(val)])
            
            # រកទឹកប្រាក់
            amt_match = re.search(r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', line)
            if not amt_match:
                continue
            
            amount = amt_match.group(1).replace(",", "")
            if float(amount) <= 0:
                continue
            
            # រកប្រភេទលុយ
            curr = "KHR" if any(x in line.upper() for x in ["KHR", "៛", "រៀល"]) else "USD"
            
            # រកឈ្មោះអ្នកផ្ទេរ
            name = "Unknown Sender"
            name_match = re.search(r'(?:FROM|BY|TRANSFER|PAYMENT)\s+([A-Z\s]{3,30})|([\u1780-\u17FF\s]{3,30})', line, re.I)
            if name_match:
                name = (name_match.group(1) or name_match.group(2)).strip()
                name = re.sub(r'\s+', ' ', name)
            
            # រកថ្ងៃខែ
            date_match = re.search(r'(\d{2}[-/]\w{3}[-/]\d{4})|(\d{4}-\d{2}-\d{2})|(\d{2}/\d{2}/\d{4})', line)
            dt = date_match.group(0) if date_match else "2026-05-08"
            
            data_list.append((dt, name, amount, curr))
        
        return data_list
    except Exception as e:
        print(f"Error parsing file: {e}")
        return []

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 **ABA Smart Bot ដំណើរការហើយ!**\n\n"
        "📤 ផ្ញើឯកសារ ABA Statement (.xlsx) ដើម្បីរក្សាទុក\n"
        "/summary - មើលសរុបប្រចាំថ្ងៃ\n"
        "/summary 2026-05-08 - មើលលម្អិតថ្ងៃជាក់លាក់"
    )

async def on_receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ សូមផ្ញើតែឯកសារ Excel (.xlsx) ប៉ុណ្ណោះ។")
        return
    
    msg = await update.message.reply_text("⏳ កំពុងវិភាគ...")
    
    temp_name = f"temp_{doc.file_id}.xlsx"
    file = await context.bot.get_file(doc.file_id)
    await file.download_to_drive(temp_name)
    
    try:
        results = parse_aba_statement(temp_name)
        if not results:
            await msg.edit_text("❌ មិនអាចអានទិន្នន័យបានទេ។")
            return
        
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        count = 0
        
        for dt, nm, am, cr in results:
            cursor.execute("SELECT 1 FROM transfers WHERE sender_name=? AND amount=? AND created_at=?", (nm, am, dt))
            if not cursor.fetchone():
                cursor.execute("INSERT INTO transfers (created_at, sender_name, amount, currency) VALUES (?,?,?,?)", 
                              (dt, nm, am, cr))
                count += 1
        
        conn.commit()
        conn.close()
        
        await msg.edit_text(f"✅ រក្សាទុក `{count}` ប្រតិបត្តិការ")
    
    except Exception as e:
        await msg.edit_text(f"❌ Error: {str(e)}")
    
    finally:
        if os.path.exists(temp_name):
            os.remove(temp_name)

async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # ប្រសិនបើមានថ្ងៃក្នុង argument
    date_filter = None
    if context.args:
        date_filter = context.args[0]
    
    if date_filter:
        # លម្អិតថ្ងៃជាក់លាក់
        cursor.execute("""
        SELECT sender_name, amount, currency
        FROM transfers 
        WHERE created_at = ?
        ORDER BY sender_name
        """, (date_filter,))
        
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            await update.message.reply_text(f"📭 ថ្ងៃ {date_filter} មិនទាន់មានទិន្នន័យ។")
            return
        
        report = f"📊 **ថ្ងៃ {date_filter}៖**\n\n"
        total_amount = 0
        for sender, amount, curr in rows:
            symbol = "$" if curr == "USD" else "៛"
            report += f"👤 {sender}: `{amount}{symbol}`\n"
            total_amount += float(amount)
        
        report += f"\n💰 **សរុប:** `{total_amount:,.2f}`"
    else:
        # សរុបប្រចាំថ្ងៃ (7 ថ្ងៃចុងក្រោយ)
        cursor.execute("""
        SELECT created_at, currency, SUM(CAST(amount AS DECIMAL))
        FROM transfers 
        GROUP BY created_at, currency 
        ORDER BY created_at DESC LIMIT 7
        """)
        
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            await update.message.reply_text("📭 មិនទាន់មានទិន្នន័យ។")
            return
        
        report = "📊 **សរុបប្រចាំថ្ងៃ៖**\n\n"
        for dt, curr, total in rows:
            symbol = "$" if curr == "USD" else "៛"
            report += f"📅 {dt}: `{total:,.2f}{symbol}`\n"
    
    await update.message.reply_text(report, parse_mode="Markdown")

if __name__ == '__main__':
    init_db()
    
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise ValueError("❌ BOT_TOKEN environment variable not set")
    
    app = ApplicationBuilder().token(token).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("summary", show_summary))
    app.add_handler(MessageHandler(filters.Document.ALL, on_receive_file))
    
    print("🚀 Bot is starting...")
    app.run_polling()
