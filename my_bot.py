import logging
import os
import re
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, CallbackQueryHandler, ConversationHandler, filters
from docx import Document
import openpyxl

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

user_last_files = {}
user_meetings = {}        

# States សម្រាប់ប្រព័ន្ធ Step-by-Step Meeting
GET_TOPIC, GET_DATETIME, GET_REMIND = range(3)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROCESSED_FOLDER = os.path.join(BASE_DIR, "processed_files")

if not os.path.exists(PROCESSED_FOLDER):
    os.makedirs(PROCESSED_FOLDER)

KHMER_DICTIONARY_CORRECTIONS = {
    "សំលាប់": "សម្លាប់",
    "សំលៀកបំពាក់": "សម្លៀកបំពាក់",
    "សំលេង": "សម្លេង",
    "សំបុត្រ": "សំបុត្រ",
    "ត្រួតពិនត្យ": "ត្រួតពិនិត្យ",
    "អនុញ្ញាត្តិ": "អនុញ្ញាត",
    "កំរិត": "កម្រិត",
    "កំលាំង": "កម្លាំង",
    "ទំនេរ": "ទំនេរ",
    "ដំនើរការ": "ដំណើរការ",
    "ចំនុច": "ចំណុច",
    "ចំនាយ": "ចំណាយ",
    "ចំនូល": "ចំណូល",
    "រៀបចំរៀង": "រៀបរៀង",
    "ពត៌មាន": "ព័ត៌មាន",
    "ឧបករណ៏": "ឧបករណ៍"
}

def clean_khmer_text(text):
    if not text:
        return text
    for wrong, correct in KHMER_DICTIONARY_CORRECTIONS.items():
        text = text.replace(wrong, correct)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\s+([,.\?!៖៕៕])', r'\1', text)
    return text.strip()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_name = update.effective_user.first_name
    welcome_message = (
        f"សួស្តី {user_name}! 🙏\n"
        "ខ្ញុំជា Bot ជំនួយការការងារ (My Work Assistant)។\n\n"
        "📁 **Attach ហ្វាល Word** - សម្រាប់គ្រប់គ្រងហ្វាលតាមប៊ូតុងអន្តរកម្ម។\n"
        "📅 `/meeting` - កត់ត្រាកាលវិភាគប្រជុំ និងตั้งเวลาแจ้งเตือนដោយស្វ័យប្រវត្តិ។"
    )
    await update.message.reply_text(welcome_message)

# --- មុខងារ Step-by-Step Meeting ---
async def meeting_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📌 សូម **វាយបញ្ចូលប្រធានបទប្រជុំ** របស់អ្នក៖")
    return GET_TOPIC

async def meeting_receive_topic(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['topic'] = update.message.text.strip()
    await update.message.reply_text(
        "✅ **បានកត់ត្រាប្រធានបទប្រជុំ**\n\n"
        "📅 បន្ត **សូមបញ្ចូល ថ្ងៃខែឆ្នាំ និងម៉ោង** តាមទម្រង់នេះ៖\n"
        "`DD-MM-YYYY HH:MM` (ឧ. `26-08-2026 14:30`)"
    )
    return GET_DATETIME

async def meeting_receive_datetime(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        target_dt = datetime.strptime(text, "%d-%m-%Y %H:%M")
        now = datetime.now()
        if target_dt <= now:
            await update.message.reply_text("❌ ថ្ងៃ និងម៉ោងដែលបានកំណត់បានន្លងផុតទៅហើយ។ សូមវាយបញ្ចូលថ្មីតាមទម្រង់៖ `DD-MM-YYYY HH:MM`")
            return GET_DATETIME

        context.user_data['datetime_str'] = text
        context.user_data['target_dt'] = target_dt
        await update.message.reply_text(
            "✅ **បានកត់ត្រាថ្ងៃខែឆ្នាំ ម៉ោង**\n\n"
            "⏰ បន្ត **សូមបញ្ចូលការរម្លឹកមុន** (ជាតួលេខនាទី ឧ. `10`, `15`, `20`, `30`...)"
        )
        return GET_REMIND
    except ValueError:
        await update.message.reply_text("❌ ទម្រង់ថ្ងៃខែខុស! (ត្រូវមានចន្លោះ Space រវាងថ្ងៃ និងម៉ោង) ឧ. `26-08-2026 14:30`")
        return GET_DATETIME

async def meeting_receive_remind(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id
    remind_text = update.message.text.strip()
    
    try:
        remind_mins = int(remind_text)
    except ValueError:
        await update.message.reply_text("⚠️ សូមបញ្ចូលជាតួលេខនាទីត្រឹមត្រូវ (ឧ. `10`)។ សូមវាយបញ្ចូលម្តងទៀត៖")
        return GET_REMIND

    topic = context.user_data.get('topic')
    datetime_str = context.user_data.get('datetime_str')
    target_dt = context.user_data.get('target_dt')

    # คำนวณเวลาแจ้งเตือนล่วงหน้า
    reminder_dt = target_dt - timedelta(minutes=remind_mins)
    now = datetime.now()
    due_seconds = (reminder_dt - now).total_seconds()
    if due_seconds < 0:
        due_seconds = 0

    async def alarm_callback(ctx: ContextTypes.DEFAULT_TYPE):
        await ctx.bot.send_message(
            chat_id=chat_id,
            text=f"⏰ **ការរំលឹកការប្រជុំ (Meeting Reminder)!**\n\n📌 **ប្រធានបទ៖** {topic}\n📆 **ថ្ងៃខែឆ្នាំ ម៉ោង៖** {datetime_str}",
            parse_mode="Markdown"
        )

    context.job_queue.run_once(alarm_callback, due_seconds, chat_id=chat_id)

    # រក្សាទុកក្នុងបញ្ជីប្រជុំរបស់ User
    if user_id not in user_meetings:
        user_meetings[user_id] = []

    new_meeting = {
        "topic": topic,
        "date_time": datetime_str,
        "remind": f"{remind_mins} នាទីមុន"
    }
    user_meetings[user_id].append(new_meeting)

    # បង្ហាញជា Table List នៅចុងបញ្ចប់តាមតម្រូវការ
    table_text = (
        f"✅ **បានកត់ត្រារម្លឹកមុន {remind_mins}នាទី និងរក្សាទុកជោគជ័យ!**\n\n"
        "📊 **[ តារាងបញ្ជីការប្រជុំ (Meeting Table List) ]**\n"
        "| ល.រ | ប្រធានបទ | ថ្ងៃខែឆ្នាំ ម៉ោង | រំលឹកមុន |\n"
        "|:---:|:---|:---:|:---:|\n"
    )
    for idx, m in enumerate(user_meetings[user_id], 1):
        table_text += f"| {idx} | {m['topic']} | {m['date_time']} | {m['remind']} |\n"

    await update.message.reply_text(table_text, parse_mode="Markdown")
    return ConversationHandler.END

async def cancel_meeting(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ បានបោះបង់ការបង្កើតប្រជុំ។")
    return ConversationHandler.END

# --- មុខងារគ្រប់គ្រងប៊ូតុងហ្វាល Word ---
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if data.startswith("btn_"):
        if user_id not in user_last_files or not os.path.exists(user_last_files[user_id]):
            await query.message.reply_text("⚠️ សូម Attach ហ្វាល Word មកជាមុនសិន។")
            return
        target_file = user_last_files[user_id]
        
        try:
            if data == "btn_summary":
                doc = Document(target_file)
                paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                summary_items = paragraphs[:4]
                summary_text = "\n".join([f"• {item}" for item in summary_items])
                if not summary_text:
                    summary_text = "(ហ្វាលនេះទទេ ឬគ្មានអត្ថបទច្បាស់លាស់)"
                file_name = os.path.basename(target_file)
                await query.message.reply_text(
                    f"📋 **សេចក្តីសង្ខេបរបាយការណ៍ (`{file_name}`)៖**\n"
                    f"-----------------------------------\n"
                    f"{summary_text}\n"
                    f"-----------------------------------",
                    parse_mode="Markdown"
                )
            elif data in ["btn_preview1", "btn_preview2"]:
                page_num = 1 if data == "btn_preview1" else 2
                doc = Document(target_file)
                paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
                page_size = 12
                total_pages = (len(paragraphs) + page_size - 1) // page_size
                if page_num > total_pages:
                    await query.message.reply_text(f"⚠️ ឯកសារនេះមានត្រឹមតែ **ទំព័រទី {total_pages}** ប៉ុណ្ណោះ។")
                    return
                start_idx = (page_num - 1) * page_size
                end_idx = start_idx + page_size
                page_content = "\n\n".join(paragraphs[start_idx:end_idx])
                await query.message.reply_text(
                    f"📄 **[ទំព័រទី {page_num} នៃ {total_pages}]**\n"
                    f"-----------------------------------\n"
                    f"{page_content}\n"
                    f"-----------------------------------",
                    parse_mode="Markdown"
                )
            elif data == "btn_spellcheck":
                doc = Document(target_file)
                fixed_count = 0
                for p in doc.paragraphs:
                    for run in p.runs:
                        orig = run.text
                        cleaned = clean_khmer_text(orig)
                        if orig != cleaned:
                            run.text = cleaned
                            fixed_count += 1
                for t in doc.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                for run in p.runs:
                                    orig = run.text
                                    cleaned = clean_khmer_text(orig)
                                    if orig != cleaned:
                                        run.text = cleaned
                                        fixed_count += 1
                original_base_name = os.path.basename(target_file)
                fixed_file_name = os.path.join(PROCESSED_FOLDER, original_base_name)
                doc.save(fixed_file_name)
                await query.message.reply_text(f"✨ បានកែអក្ខរាវិរុទ្ធ និងដកឃ្លាជោគជ័យចំនួន `{fixed_count}` កន្លែង។", parse_mode="Markdown")
                await query.message.reply_document(document=open(fixed_file_name, "rb"), caption="✅ ហ្វាល Word ស្អាត រក្សា Font ដើម!")
            elif data == "btn_editword_info":
                await query.message.reply_text("✏️ សូមប្រើពាក្យបញ្ជា៖ `/editword ពាក្យចាស់ | ពាក្យថ្មី`")
        except Exception as e:
            await query.message.reply_text(f"❌ មានបញ្ហា៖ {str(e)}")

# --- មុខងារຮັບហ្វាល Word ---
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        user_id = update.effective_user.id
        document = update.message.document
        file_name = document.file_name
        
        saved_path = os.path.join(BASE_DIR, file_name)
        file = await context.bot.get_file(document.file_id)
        await file.download_to_drive(saved_path)
        user_last_files[user_id] = saved_path
        
        if file_name.endswith('.docx'):
            doc = Document(saved_path)
            total_paragraphs = len([p.text for p in doc.paragraphs if p.text.strip()])
            
            keyboard = [
                [InlineKeyboardButton("📋 សង្ខេប (Summary)", callback_data="btn_summary")],
                [InlineKeyboardButton("📄 អានទំព័រទី១ (Preview 1)", callback_data="btn_preview1"),
                 InlineKeyboardButton("📄 អានទំព័រទី២ (Preview 2)", callback_data="btn_preview2")],
                [InlineKeyboardButton("✍️ កែអក្ខរាវិរុទ្ធ (Spellcheck)", callback_data="btn_spellcheck")],
                [InlineKeyboardButton("✏️ កែពាក្យ (Editword)", callback_data="btn_editword_info")]
            ]
            reply_markup = InlineKeyboardMarkup(keyboard)

            await update.message.reply_text(
                f"📥 **ទទួលបាន និងចងចាំឯកសារ Word (`{file_name}`) រួចរាល់!**\n"
                f"📊 ចំនួនកថាខណ្ឌសរុប៖ {total_paragraphs} ឃ្លា\n\n"
                f"👇 **សូមចុចជ្រើសរើសមុខងារខាងក្រោមនេះ៖**",
                reply_markup=reply_markup,
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"📥 ទទួលបានឯកសារ `{file_name}` រួចរាល់!")
    except Exception as e:
        await update.message.reply_text(f"❌ មានបញ្ហា៖ {str(e)}")

def main():
    app = ApplicationBuilder().token("8900404018:AAEKN28HJjDuZf0bOvKwEz754Zqs8kPVaKk").build()

    meeting_conv = ConversationHandler(
        entry_points=[CommandHandler('meeting', meeting_start)],
        states={
            GET_TOPIC: [MessageHandler(filters.TEXT & ~filters.COMMAND, meeting_receive_topic)],
            GET_DATETIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, meeting_receive_datetime)],
            GET_REMIND: [MessageHandler(filters.TEXT & ~filters.COMMAND, meeting_receive_remind)],
        },
        fallbacks=[CommandHandler('cancel', cancel_meeting)]
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(meeting_conv)
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("Bot @Borann_bot កំពុងដំណើរការជាមួយ Table List Output...")
    app.run_polling()

if __name__ == '__main__':
    main()
