import os
import re
import sqlite3
import pandas as pd
import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# --- កំណត់ផ្លូវ Database (សម្រាប់ Railway គួរប្រើ /app/data/ បើមាន Volume) ---
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

# --- មុខងារអាន ABA Excel ---
def parse_aba_statement(file_path):
    try:
        # អាន Excel ដោយប្រើ engine openpyxl
        df = pd.read_excel(file_path, engine='openpyxl', header=None)
        data_list = []
        
        for idx, row in df.iterrows():
            # បំប្លែងជួរនីមួយៗជា Text ដើម្បីស្កេនរកទិន្នន័យ
            line = " ".join([str(val) for val in row.values if pd.notna(val)])
            
            # ១. រកទឹកប្រាក់ (Amount)
            amt_match = re.search(r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', line)
            if not amt_match: continue
            
            amount = amt_match.group(1).replace(",", "")
            if float(amount) <= 0: continue
            
            # ២. រកប្រភេទលុយ (Currency)
            curr = "KHR" if any(x in line.upper() for x in ["KHR", "៛", "រៀល"]) else "USD"
            
            # ៣. រកឈ្មោះអ្នកផ្ញើ (English & Khmer)
            name = "Unknown Sender"
            name_pat = re.search(r'(?:FROM|BY|TRANSFER|PAYMENT)\s+([A-Z\s]{3,30})|([\u1780-\u17FF\s]{3,30})', line, re.I)
            if name_pat:
                name = (name_pat.group(1) or name_pat.group(2)).strip()
                name = re.sub(r'\s+', ' ', name)

            # ៤. រកថ្ងៃខែ
            date_match = re.search(r'(\d{2}[-/]\w{3}[-/]\d{4})|(\d{4}-\d{2}-\d{2})|(\d{2}/\d{2}/\d{4})', line)
            dt = date_match.group(0) if date_match else datetime.datetime.now().strftime("%Y-%m-%d")

            data_list.append((dt, name, amount, curr))
        
        return data_list
    except Exception as e:
        print(f"Error: {e}")
        return []

# --- បញ្ជា /start ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 **ABA Smart Bot ដំណើរការហើយ!**\n\n"
        "📤 សូមផ្ញើឯកសារ ABA Excel (.xlsx) មកកាន់ខ្ញុំ\n"
        "📊 ប្រើ /summary ដើម្បីមើលរបាយការណ៍"
    )

# --- ទទួល File Excel ---
async def on_receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        return
    
    msg = await update.message.reply_text("⏳ កំពុងវិភាគ Statement...")
    temp_name = f"temp_{doc.file_id}.xlsx"
    file = await context.bot.get_file(doc.file_id)
    await file.download_to_drive(temp_name)

    try:
        results = parse_aba_statement(temp_name)
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        count = 0
        for dt, nm, am, cr in results:
            cursor.execute("SELECT 1 FROM transfers WHERE sender_name=? AND amount=? AND created_at=?", (nm, am, dt))
            if not cursor.fetchone():
                cursor.execute("INSERT INTO transfers (created_at, sender_name, amount, currency) VALUES (?,?,?,?)", (dt, nm, am, cr))
                count += 1
        conn.commit()
        conn.close()
        await msg.edit_text(f"✅ រក្សាទុកបាន `{count}` ប្រតិបត្តិការថ្មី!")
    except Exception as e:
        await msg.edit_text(f"❌ កំហុស៖ {str(e)}")
    finally:
        if os.path.exists(temp_name): os.remove(temp_name)

# --- បង្ហាញរបាយការណ៍ស្អាត (Summary) ---
async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # យកថ្ងៃនេះជាគោល បើអ្នកប្រើមិនបានវាយបញ្ជាក់ថ្ងៃ
    target_date = context.args[0] if context.args else datetime.datetime.now().strftime("%Y-%m-%d")

    cursor.execute("SELECT sender_name, amount, currency FROM transfers WHERE created_at = ?", (target_date,))
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        await update.message.reply_text(f"📭 ថ្ងៃ `{target_date}` មិនទាន់មានទិន្នន័យទេ។")
        return

    # បែងចែក USD និង KHR
    usd_list = [(n, float(a)) for n, a, c in rows if c == "USD"]
    khr_list = [(n, float(a)) for n, a, c in rows if c == "KHR"]

    # បង្កើតសាររបាយការណ៍
    report = f"📊 **ថ្ងៃ {target_date}**\n"
    report += "━━━━━━━━━━━━━━━━━━━━\n\n"

    total_usd = 0
    for name, amt in usd_list:
        total_usd += amt
        report += f"👤 {name}:  `{amt:,.2f}$`\n"

    if khr_list:
        if usd_list: report += "\n" + "━" * 15 + "\n"
        total_khr = 0
        for name, amt in khr_list:
            total_khr += amt
            report += f"👤 {name}:  `{amt:,.0f}៛`\n"

    report += "\n━━━━━━━━━━━━━━━━━━━━\n"
    
    # បូកសរុបចុងក្រោយ
    if usd_list and not khr_list:
        report += f"💰 **សរុប:  {total_usd:,.2f}$**"
    elif khr_list and not usd_list:
        report += f"💰 **សរុប:  {total_khr:,.0f}៛**"
    else:
        report += f"💰 **សរុប:** `{total_usd:,.2f}$` + `{total_khr:,.0f}៛`"

    await update.message.reply_text(report, parse_mode="Markdown")

if __name__ == '__main__':
    init_db()
    # ទាញយក Token ពី Environment Variable ក្នុង Railway (ឈ្មោះ BOT_TOKEN)
    TOKEN = os.getenv("BOT_TOKEN") 
    app = ApplicationBuilder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("summary", show_summary))
    app.add_handler(MessageHandler(filters.Document.ALL, on_receive_file))
    
    print("🚀 Bot is running...")
    app.run_polling()
