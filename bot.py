import logging
import os
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
from pptx import Presentation
from datetime import datetime

TOKEN = os.environ.get("BOT_TOKEN", "8990357714:AAGgbLbTHKTi2rQ5FkcKTTJ8pO1d41rb4Kw")
TEMPLATE_PATH = "template.pptx"

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

TITLE, SUBTITLE, POINTS, CONFIRM = range(4)


def replace_text_in_shape(shape, replacements):
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            for placeholder, value in replacements.items():
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)


def generate_pptx(data):
    prs = Presentation(TEMPLATE_PATH)
    title_text    = data.get("title", "عنوان التقرير")
    subtitle_text = data.get("subtitle", "")
    points        = data.get("points", [])
    date_text     = datetime.now().strftime("%Y-%m-%d")

    replacements = {
        "{{TITLE}}":    title_text,
        "{{SUBTITLE}}": subtitle_text,
        "{{DATE}}":     date_text,
        "{{POINTS}}":   "\n".join(f"- {p}" for p in points),
    }

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

    output_path = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    prs.save(output_path)
    return output_path


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "مرحباً! أرسل /report لإنشاء تقرير جديد."
    )


async def report_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("ما هو عنوان التقرير؟")
    return TITLE


async def get_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["title"] = update.message.text.strip()
    await update.message.reply_text("ما هو العنوان الفرعي؟ (أو أرسل - لتخطيه)")
    return SUBTITLE


async def get_subtitle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    context.user_data["subtitle"] = "" if text == "-" else text
    context.user_data["points"] = []
    await update.message.reply_text("أرسل نقاط المحتوى واحدة تلو الأخرى. عند الانتهاء أرسل /done")
    return POINTS


async def get_points(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["points"].append(update.message.text.strip())
    await update.message.reply_text(f"تمت إضافة النقطة. أرسل نقطة أخرى أو /done")
    return POINTS


async def points_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get("points"):
        await update.message.reply_text("أضف نقطة واحدة على الأقل.")
        return POINTS
    keyboard = ReplyKeyboardMarkup([["نعم", "لا"]], resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text(
        f"العنوان: {context.user_data['title']}\nالنقاط: {len(context.user_data['points'])}\n\nهل تريد إنشاء الملف؟",
        reply_markup=keyboard,
    )
    return CONFIRM


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "نعم" in update.message.text:
        await update.message.reply_text("جارٍ إنشاء الملف...", reply_markup=ReplyKeyboardRemove())
        try:
            path = generate_pptx(context.user_data)
            with open(path, "rb") as f:
                await update.message.reply_document(document=f, filename=os.path.basename(path), caption="تم إنشاء التقرير!")
            os.remove(path)
        except Exception as e:
            logger.error(e)
            await update.message.reply_text(f"حدث خطأ: {e}")
    else:
        await update.message.reply_text("تم الإلغاء.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("تم الإلغاء.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


if __name__ == "__main__":
    app = ApplicationBuilder().token(TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("report", report_start)],
        states={
            TITLE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, get_title)],
            SUBTITLE:[MessageHandler(filters.TEXT & ~filters.COMMAND, get_subtitle)],
            POINTS:  [MessageHandler(filters.TEXT & ~filters.COMMAND, get_points), CommandHandler("done", points_done)],
            CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    logger.info("البوت يعمل...")
    app.run_polling()
