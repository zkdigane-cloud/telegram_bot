import os
import time
from datetime import datetime, timedelta
import openpyxl

from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

# 🌐 KEEP ALIVE (24hr)
from flask import Flask
from threading import Thread

app_web = Flask('')

@app_web.route('/')
def home():
    return "Bot is running ✅"

def run_web():
    port = int(os.environ.get("PORT", 8080))
    app_web.run(host='0.0.0.0', port=port)

Thread(target=run_web).start()

# 🔐 ENV
TOKEN = os.getenv("TOKEN")
GROUP_ID = int(os.getenv("GROUP_ID"))

# 📦 Data
user_breaks = {}
user_attendance = {}
FILE_NAME = "attendance.xlsx"

# 🇲🇲 Time
def get_myanmar_time():
    return datetime.utcnow() + timedelta(hours=6, minutes=30)

# ⏰ Late
def get_late_minutes():
    now = get_myanmar_time()
    current = now.hour * 60 + now.minute
    start = 21 * 60 + 30

    if now.hour < 12:
        current += 24 * 60

    return max(0, current - start)

# ⏰ Early leave
def early_off():
    now = get_myanmar_time()
    current = now.hour * 60 + now.minute
    off_time = 9 * 60 + 30

    if now.hour >= 21:
        return False

    return current < off_time

# 💰 Fine rule
def calculate_fine(minutes):
    if minutes == 0:
        return 0, "On Time"
    elif 1 <= minutes <= 5:
        return minutes * 100, f"{minutes} min late"
    elif 6 <= minutes <= 14:
        return 1000, "6-14 min late"
    elif minutes == 15:
        return 3000, "15 min late"
    elif minutes >= 30:
        return 0, "❌ 3 days salary cut"
    else:
        return 0, f"{minutes} min late"

# 📊 Excel
def save_full(user_id, check_in, check_out, late_min, fine, work_time, early_leave, note):
    try:
        wb = openpyxl.load_workbook(FILE_NAME)
        ws = wb.active
    except:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["User ID","Check In","Check Out","Late(min)","Fine","Note","Work(min)","Early"])

    ws.append([user_id, check_in, check_out, late_min, fine, note, work_time, early_leave])
    wb.save(FILE_NAME)

# 🔘 Keyboard
keyboard = [
    ["✅ Check In", "🛑 Off Work"],
    ["🚻 Toilet", "🚬 Smoke", "🍱 Eat"],
    ["🔙 Back to Seat"]
]
reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

# ▶️ Start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(update.effective_chat.id, "Welcome 👋", reply_markup=reply_markup)

# ⏰ Auto report
async def shift_and_report(context: ContextTypes.DEFAULT_TYPE):
    now = get_myanmar_time().strftime("%H:%M")

    if now == "21:30":
        await context.bot.send_message(GROUP_ID, "⏰ Check In Time!")

    if now == "09:30":
        await context.bot.send_message(GROUP_ID, "⏰ Off Work Time!")

        try:
            with open(FILE_NAME, "rb") as f:
                await context.bot.send_document(GROUP_ID, f, filename="attendance.xlsx")
        except:
            await context.bot.send_message(GROUP_ID, "No report")

# 💬 Handle
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    text = update.message.text

    # ✅ Check In
    if text == "✅ Check In":
        if user_id in user_attendance:
            await context.bot.send_message(update.effective_chat.id, "Already checked in")
            return

        now = get_myanmar_time()
        late = get_late_minutes()
        fine, note = calculate_fine(late)

        user_attendance[user_id] = {
            "time": time.time(),
            "str": now.strftime("%H:%M:%S"),
            "late": late,
            "fine": fine,
            "note": note
        }

        msg = f"Check In: {now.strftime('%H:%M:%S')}\nLate: {late} min\nFine: {fine}"
        await context.bot.send_message(update.effective_chat.id, msg)

    # 🛑 Off Work
    elif text == "🛑 Off Work":
        if user_id not in user_attendance:
            await context.bot.send_message(update.effective_chat.id, "No check in")
            return

        data = user_attendance[user_id]
        work_min = int((time.time() - data["time"]) / 60)
        now = get_myanmar_time().strftime("%H:%M:%S")

        early = "YES" if early_off() else "NO"

        save_full(user_id, data["str"], now, data["late"], data["fine"], work_min, early, data["note"])

        del user_attendance[user_id]

        await context.bot.send_message(update.effective_chat.id, f"Off Work\nWork: {work_min} min")

    # 🚻 Break
    elif text in ["🚻 Toilet","🚬 Smoke","🍱 Eat"]:
        user_breaks[user_id] = time.time()
        await context.bot.send_message(update.effective_chat.id, "Break started")

    # 🔙 Back
    elif text == "🔙 Back to Seat":
        if user_id not in user_breaks:
            await context.bot.send_message(update.effective_chat.id, "No break")
            return

        total = int((time.time() - user_breaks[user_id]) / 60)
        late = max(0, total - 10)

        del user_breaks[user_id]

        await context.bot.send_message(update.effective_chat.id, f"Back\nLate: {late} min")

# 📊 Manual report
async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        with open(FILE_NAME, "rb") as f:
            await context.bot.send_document(update.effective_chat.id, f)
    except:
        await context.bot.send_message(update.effective_chat.id, "No file")

# 🚀 Run
app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CommandHandler("report", report))
app.add_handler(MessageHandler(filters.TEXT, handle))

app.job_queue.run_repeating(shift_and_report, interval=60, first=10)

print("Bot running...")
app.run_polling()