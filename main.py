import os
import re
import datetime
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
        
        # Skip header row (row 0), start from row 1
        for idx, row in df.iterrows():
            if idx == 0:  # Skip header
                continue
            
            # Column A (0) = Date, B (1) = Transaction Details, C (2) = Money in, D (3) = Ccy
            if len(row) < 3:
                continue
            
            # --- Date (column A, index 0) ---
            raw_date = row.iloc[0]
            if pd.isna(raw_date):
                continue
            
            dt = "2026-05-08"  # fallback
            if hasattr(raw_date, 'strftime'):
                dt = raw_date.strftime("%Y-%m-%d")
            else:
                date_str = str(raw_date).strip()
                # Try YYYY-MM-DD
                m = re.match(r'^(\d{4}-\d{2}-\d{2})', date_str)
                if m:
                    dt = m.group(1)
                else:
                    # Try DD/MM/YYYY
                    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', date_str)
                    if m:
                        dt = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
                    else:
                        # Try DD-Mon-YYYY
                        m = re.match(r'^(\d{2})-([A-Za-z]{3})-(\d{4})$', date_str)
                        if m:
                            try:
                                dt = datetime.datetime.strptime(date_str, "%d-%b-%Y").strftime("%Y-%m-%d")
                            except ValueError:
                                dt = date_str
            
            # --- Sender Name (column B, index 1) ---
            raw_name = row.iloc[1]
            if pd.isna(raw_name) or str(raw_name).strip() == "":
                continue
            name = str(raw_name).strip()
            name = re.sub(r'\s+', ' ', name)
            
            # --- Amount (column C, index 2) ---
            raw_amount = row.iloc[2]
            if pd.isna(raw_amount):
                continue
            
            amount_str = str(raw_amount).replace(",", "").strip()
            try:
                amount_val = float(amount_str)
            except ValueError:
                continue
            
            if amount_val <= 0:
                continue
            
            amount = str(amount_val)
            
            # --- Currency (column D, index 3) ---
            curr = "USD"  # default
            if len(row) > 3 and pd.notna(row.iloc[3]):
                raw_curr = str(row.iloc[3]).strip().upper()
                if raw_curr in ("KHR", "USD"):
                    curr = raw_curr
                elif any(x in raw_curr for x in ["KHR", "៛", "រៀល"]):
                    curr = "KHR"
            
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
    
    date_filter = None
    if context.args:
        date_filter = context.args[0]
    
    if date_filter:
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
