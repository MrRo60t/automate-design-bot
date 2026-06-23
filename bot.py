import os
import logging
import io
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
from ai_parser import parse_document
from pptx_generator import generate_pptx

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

WAITING_FOR_TEXT = 1

WELCOME = (
    "👋 Привет! Я автоматически создаю презентации по вашему шаблону.\n\n"
    "Просто отправьте мне текст аудита или стратегии — и я верну готовый .pptx файл.\n\n"
    "📝 *Поддерживаемые типы:*\n"
    "• Google Ads аудит\n"
    "• Стратегия\n\n"
    "Текст может быть на *любом языке*."
)

WAITING_MSG = "⏳ Генерирую презентацию, это займёт 15–30 секунд..."

ERROR_MSG = (
    "❌ Что-то пошло не так. Попробуйте ещё раз или проверьте, "
    "что текст содержит достаточно информации для аудита/стратегии."
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(WELCOME, parse_mode="Markdown")
    return WAITING_FOR_TEXT


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user_text = update.message.text.strip()

    if len(user_text) < 100:
        await update.message.reply_text(
            "⚠️ Текст слишком короткий. Пожалуйста, вставьте полный текст аудита или стратегии."
        )
        return WAITING_FOR_TEXT

    status_msg = await update.message.reply_text(WAITING_MSG)

    try:
        parsed = parse_document(user_text)
        pptx_bytes = generate_pptx(parsed)

        client_name = parsed.get("cover", {}).get("client_name", "presentation")
        doc_type = parsed.get("cover", {}).get("document_type", "document")
        filename = f"{client_name}_{doc_type}.pptx".replace(" ", "_")

        await context.bot.delete_message(
            chat_id=update.effective_chat.id,
            message_id=status_msg.message_id,
        )

        await update.message.reply_document(
            document=io.BytesIO(pptx_bytes),
            filename=filename,
            caption=f"✅ *{client_name}* — {doc_type} готов!",
            parse_mode="Markdown",
        )

    except Exception as e:
        logger.error(f"Error generating presentation: {e}", exc_info=True)
        await context.bot.edit_message_text(
            chat_id=update.effective_chat.id,
            message_id=status_msg.message_id,
            text=ERROR_MSG,
        )

    return WAITING_FOR_TEXT


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Отменено. Отправьте текст когда будете готовы.")
    return WAITING_FOR_TEXT


def main():
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    app = Application.builder().token(token).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text),
        ],
        states={
            WAITING_FOR_TEXT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
