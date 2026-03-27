#!/usr/bin/env python3
"""
Telegram Invoice Bot — полная версия
Обрабатывает PDF счета и обновляет Excel файлы объектов
"""

import os
import json
import logging
import shutil
from pathlib import Path
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, filters, ContextTypes
)

from database import Database
from invoice_processor import InvoiceProcessor
from excel_exporter import ExcelExporter
from excel_updater import ExcelUpdater

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[logging.FileHandler('bot.log'), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

WAITING_OBJECT = 1
WAITING_SECTION = 2
WAITING_XLSX_CONFIRM = 3

OBJECTS_DIR = Path("objects")
OBJECTS_DIR.mkdir(exist_ok=True)

db = Database()
processor = InvoiceProcessor()
exporter = ExcelExporter(db)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "👋 *Привет! Я бот для учёта счетов.*\n\n"
        "📄 *Что умею:*\n"
        "• Читать PDF счета и извлекать данные\n"
        "• Добавлять строки в ваш Excel (Kokku)\n"
        "• Обновлять остатки в спецификациях\n"
        "• Предупреждать о перерасходе и росте цен\n\n"
        "🗂 *Как начать:*\n"
        "1. Загрузите Excel файл объекта — /setxlsx\n"
        "2. Отправьте PDF счёт — бот всё обновит сам\n\n"
        "📊 *Команды:*\n"
        "/objects — список объектов\n"
        "/report — отчёт по объекту\n"
        "/export — выгрузить аналитику в Excel\n"
        "/setxlsx — привязать Excel к объекту\n"
        "/help — справка"
    )
    await update.message.reply_text(text, parse_mode='Markdown')


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📖 *Справка:*\n\n"
        "*Загрузка счёта:*\n"
        "Отправьте PDF — бот извлечёт данные,\n"
        "уточнит объект и раздел если нужно,\n"
        "обновит ваш Excel и пришлёт его обратно.\n\n"
        "*Привязка Excel:*\n"
        "/setxlsx → отправьте xlsx → выберите объект\n\n"
        "*Предупреждения:*\n"
        "🔴 Перерасход по позиции спецификации\n"
        "🟠 Остаток < 10% от нормы\n"
        "📈 Цена выросла vs прошлой закупки\n"
        "⚠️ Использовано > 80% бюджета\n"
        "🚨 Бюджет превышен"
    )
    await update.message.reply_text(text, parse_mode='Markdown')


async def list_objects(update: Update, context: ContextTypes.DEFAULT_TYPE):
    objects = db.get_all_objects()
    if not objects:
        await update.message.reply_text(
            "📭 Объектов пока нет.\nЗагрузите Excel через /setxlsx"
        )
        return
    text = "🏗 *Объекты в базе:*\n\n"
    for obj in objects:
        xlsx_ok = "📊 Excel ✅" if obj.get('xlsx_path') else "📊 Excel ❌"
        text += (f"• *{obj['name']}*\n"
                 f"  {xlsx_ok} | Счетов: {obj['invoice_count']} | {obj['total']:,.2f} €\n\n")
    await update.message.reply_text(text, parse_mode='Markdown')


async def set_xlsx_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📎 Отправьте .xlsx файл объекта:",
        parse_mode='Markdown'
    )
    context.user_data['waiting_xlsx'] = True
    return WAITING_XLSX_CONFIRM


async def handle_xlsx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    filename = update.message.document.file_name
    tmp_path = f"/tmp/{filename}"
    await file.download_to_drive(tmp_path)
    context.user_data['pending_xlsx'] = tmp_path
    context.user_data['pending_xlsx_name'] = filename

    objects = db.get_all_objects()
    keyboard = [[InlineKeyboardButton(o['name'], callback_data=f"xlsx_{o['name']}")]
                for o in objects]
    keyboard.append([InlineKeyboardButton("➕ Новый объект", callback_data="xlsx_NEW")])

    await update.message.reply_text(
        f"📁 *{filename}* получен.\nК какому объекту привязать?",
        parse_mode='Markdown',
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return WAITING_XLSX_CONFIRM


async def xlsx_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    obj = query.data[5:]  # убрать "xlsx_"

    if obj == "NEW":
        await query.edit_message_text("Введите название нового объекта:")
        context.user_data['new_obj_for_xlsx'] = True
        return WAITING_XLSX_CONFIRM

    await _bind_xlsx(query, context, obj)
    return ConversationHandler.END


async def xlsx_new_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get('new_obj_for_xlsx'):
        return ConversationHandler.END
    context.user_data['new_obj_for_xlsx'] = False
    await _bind_xlsx(update, context, update.message.text.strip())
    return ConversationHandler.END


async def _bind_xlsx(upd, context, obj_name):
    tmp = context.user_data.get('pending_xlsx')
    fname = context.user_data.get('pending_xlsx_name', 'object.xlsx')
    if tmp and Path(tmp).exists():
        dest = OBJECTS_DIR / f"{obj_name}_{fname}"
        shutil.copy2(tmp, dest)
        Path(tmp).unlink(missing_ok=True)
        db.get_or_create_object(obj_name, str(dest))
        db.set_object_xlsx(obj_name, str(dest))
        txt = f"✅ Excel привязан к *{obj_name}*"
    else:
        txt = "❌ Файл не найден, попробуйте снова."
    fn = upd.edit_message_text if hasattr(upd, 'edit_message_text') else upd.message.reply_text
    await fn(txt, parse_mode='Markdown')
    context.user_data.clear()


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("⏳ Читаю счёт...")
    try:
        file = await update.message.document.get_file()
        pdf_path = f"/tmp/inv_{update.message.message_id}.pdf"
        await file.download_to_drive(pdf_path)

        await msg.edit_text("🤖 Анализирую через AI...")
        invoice_data = await processor.extract(pdf_path)

        if not invoice_data:
            await msg.edit_text("❌ Не удалось извлечь данные из файла.")
            return ConversationHandler.END

        context.user_data.update({'invoice': invoice_data, 'pdf_path': pdf_path})

        if not invoice_data.get('object'):
            objects = db.get_all_objects()
            keyboard = [[InlineKeyboardButton(o['name'], callback_data=f"iobj_{o['name']}")]
                        for o in objects]
            keyboard.append([InlineKeyboardButton("➕ Новый объект", callback_data="iobj_NEW")])
            await msg.edit_text(
                f"📄 *№{invoice_data.get('number','?')}* от {invoice_data.get('date','?')}\n"
                f"🏢 {invoice_data.get('supplier','?')}\n\n❓ *К какому объекту?*",
                parse_mode='Markdown', reply_markup=InlineKeyboardMarkup(keyboard)
            )
            return WAITING_OBJECT

        return await _ask_section_or_save(update, context, msg)

    except Exception as e:
        logger.error(f"PDF error: {e}", exc_info=True)
        await msg.edit_text(f"❌ Ошибка: {e}")
        return ConversationHandler.END


async def inv_object_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    obj = query.data[5:]
    if obj == "NEW":
        await query.edit_message_text("Введите название объекта:")
        context.user_data['new_inv_obj'] = True
        return WAITING_OBJECT
    context.user_data['invoice']['object'] = obj
    return await _ask_section_or_save(query, context, query)


async def inv_obj_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get('new_inv_obj'):
        return ConversationHandler.END
    context.user_data['new_inv_obj'] = False
    context.user_data['invoice']['object'] = update.message.text.strip()
    msg = await update.message.reply_text("✅")
    return await _ask_section_or_save(update, context, msg)


async def _ask_section_or_save(upd, context, msg):
    inv = context.user_data['invoice']
    if not inv.get('section'):
        sections = ['Kanalisatsioon', 'Vesi', 'MÄRG TORU', 'Sadevee', 'Прочее']
        kb = [[InlineKeyboardButton(s, callback_data=f"isec_{s}")] for s in sections]
        txt = (f"📄 *№{inv.get('number','?')}* | {inv.get('object')}\n"
               f"💰 {sum(i.get('amount',0) for i in inv.get('items',[])): ,.2f} €\n\n"
               f"❓ *Раздел:*")
        fn = (msg.edit_message_text if hasattr(msg, 'edit_message_text')
              else msg.edit_text if hasattr(msg, 'edit_text')
              else upd.message.reply_text)
        await fn(txt, parse_mode='Markdown', reply_markup=InlineKeyboardMarkup(kb))
        return WAITING_SECTION
    else:
        await _finalize(upd, context, msg)
        return ConversationHandler.END


async def inv_section_cb(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data['invoice']['section'] = query.data[5:]
    await _finalize(query, context, query)
    return ConversationHandler.END


async def _finalize(upd, context, msg):
    inv      = context.user_data['invoice']
    obj_name = inv.get('object', '')
    total    = sum(i.get('amount', 0) for i in inv.get('items', []))

    db.save_invoice(inv)

    xlsx_path    = db.get_object_xlsx(obj_name)
    excel_result = {}
    if xlsx_path and Path(xlsx_path).exists():
        try:
            excel_result = ExcelUpdater(xlsx_path).update(inv)
        except Exception as e:
            excel_result = {'error': str(e)}

    # ── Основная информация ──────────────────────────────────────────────────
    text = (
        f"✅ *Счёт обработан*\n\n"
        f"📄 №*{inv.get('number','?')}* от {inv.get('date','?')}\n"
        f"🏗 Объект: *{obj_name}*\n"
        f"📂 Раздел: {inv.get('section','?')}\n"
        f"🏢 Поставщик: {inv.get('supplier','?')}\n"
        f"📦 Позиций: {len(inv.get('items',[]))}  |  💰 Итого: *{total:,.2f} €*\n"
    )

    # ── Статус Excel ─────────────────────────────────────────────────────────
    if not xlsx_path:
        text += "\n💡 _Excel не привязан — используйте /setxlsx_\n"
    elif excel_result.get('error'):
        text += f"\n⚠️ _Ошибка обновления Excel: {excel_result['error']}_\n"
    else:
        rows = len(excel_result.get('rows_added', []))
        text += f"\n📊 *Excel обновлён:* Kokku ✅  |  Строк добавлено: {rows}\n"

    # ── Перерасход по количеству ─────────────────────────────────────────────
    overruns = excel_result.get('overrun_items', [])
    if overruns:
        text += "\n🔴 *ПЕРЕРАСХОД ПО МАТЕРИАЛАМ:*\n"
        for w in overruns:
            text += (f"  ❗ *{w['item']}*"
                     f"{' Ø'+w['diameter'] if w['diameter'] else ''}\n"
                     f"     Норма: {w['norm']:.1f} {w['unit']}  |  "
                     f"Закуплено: {w['used']:.1f}  |  "
                     f"Перерасход: *+{w['overrun']:.1f} {w['unit']}*\n")

    # ── Бюджет ───────────────────────────────────────────────────────────────
    budget_warns = excel_result.get('budget_warnings', [])
    if budget_warns:
        text += "\n"
        for w in budget_warns:
            if w['level'] == 'КРИТИЧНО':
                text += (f"🚨 *БЮДЖЕТ ПРЕВЫШЕН — {w['section']}*\n"
                         f"   Бюджет: {w['budget']:,.2f} €  |  "
                         f"Потрачено: {w['spent']:,.2f} € (*{w['pct']:.1f}%*)\n"
                         f"   Перерасход: *{abs(w['remaining']):,.2f} €*\n")
            else:
                text += (f"⚠️ *Бюджет {w['pct']:.0f}% — {w['section']}*\n"
                         f"   Потрачено: {w['spent']:,.2f} €  |  "
                         f"Остаток: *{w['remaining']:,.2f} €*\n")

    # ── Изменение цен ────────────────────────────────────────────────────────
    price_changes = excel_result.get('price_changes', [])
    if price_changes:
        text += "\n📈 *Изменение цен:*\n"
        for ch in price_changes:
            arrow = "📈" if ch['pct'] > 0 else "📉"
            text += (f"  {arrow} *{ch['item']}*"
                     f"{' Ø'+ch['diameter'] if ch['diameter'] else ''}: "
                     f"{ch['old_price']:,.2f} → {ch['new_price']:,.2f} € "
                     f"(*{ch['pct']:+.1f}%*)\n")

    # ── Отправить текст ──────────────────────────────────────────────────────
    fn = (msg.edit_message_text if hasattr(msg, 'edit_message_text')
          else msg.edit_text     if hasattr(msg, 'edit_text')
          else None)
    try:
        if fn:
            await fn(text, parse_mode='Markdown')
        else:
            await upd.message.reply_text(text, parse_mode='Markdown')
    except Exception:
        if hasattr(upd, 'message'):
            await upd.message.reply_text(text, parse_mode='Markdown')

    # ── Отправить обновлённый Excel ──────────────────────────────────────────
    if xlsx_path and Path(xlsx_path).exists() and not excel_result.get('error'):
        chat_id = (upd.message.chat_id if hasattr(upd, 'message')
                   else upd.effective_chat.id)
        with open(xlsx_path, 'rb') as f:
            await context.bot.send_document(
                chat_id, f,
                caption=f"📊 {obj_name} — обновлён",
                filename=Path(xlsx_path).name
            )

    pdf = context.user_data.get('pdf_path')
    context.user_data.clear()
    if pdf:
        Path(pdf).unlink(missing_ok=True)


async def report_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    objects = db.get_all_objects()
    if not objects:
        await update.message.reply_text("📭 Нет объектов.")
        return
    kb = [[InlineKeyboardButton(o['name'], callback_data=f"rep_{o['name']}")] for o in objects]
    await update.message.reply_text("Выберите объект:", reply_markup=InlineKeyboardMarkup(kb))


async def export_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    objects = db.get_all_objects()
    if not objects:
        await update.message.reply_text("📭 Нет объектов.")
        return
    kb = [[InlineKeyboardButton(o['name'], callback_data=f"exp_{o['name']}")] for o in objects]
    kb.append([InlineKeyboardButton("📊 Все объекты", callback_data="exp_ALL")])
    await update.message.reply_text("Выберите объект:", reply_markup=InlineKeyboardMarkup(kb))


async def general_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data.startswith("rep_"):
        obj_name = data[4:]
        d = db.get_object_report(obj_name)
        if not d:
            await query.edit_message_text("❌ Не найден.")
            return
        text = f"📊 *{obj_name}*\n\nСчетов: {d['invoice_count']} | Итого: {d['total']:,.2f} €\n"
        if d['categories']:
            text += "\n📂 *Разделы:*\n"
            for cat, amt in d['categories'].items():
                text += f"  • {cat}: {amt:,.2f} €\n"
        if d['recent_invoices']:
            text += "\n🗓 *Последние:*\n"
            for inv in d['recent_invoices'][:5]:
                text += f"  • №{inv['number']} {inv['date']} — {inv['total']:,.2f} €\n"
        await query.edit_message_text(text, parse_mode='Markdown')

    elif data.startswith("exp_"):
        obj_name = data[4:]
        await query.edit_message_text("⏳ Создаю аналитику...")
        fp = exporter.export(None if obj_name == "ALL" else obj_name)
        if not fp:
            await query.edit_message_text("❌ Нет данных.")
            return
        with open(fp, 'rb') as f:
            await query.message.reply_document(f, caption=f"📊 {obj_name}",
                                                filename=Path(fp).name)
        Path(fp).unlink(missing_ok=True)
        await query.delete_message()


def main():
    token = os.environ.get('TELEGRAM_BOT_TOKEN')
    if not token:
        raise ValueError("Не задан TELEGRAM_BOT_TOKEN в .env!")

    app = Application.builder().token(token).build()

    pdf_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Document.PDF, handle_pdf)],
        states={
            WAITING_OBJECT: [
                CallbackQueryHandler(inv_object_cb, pattern="^iobj_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, inv_obj_text)
            ],
            WAITING_SECTION: [
                CallbackQueryHandler(inv_section_cb, pattern="^isec_")
            ],
        },
        fallbacks=[CommandHandler('start', start)],
        allow_reentry=True
    )

    xlsx_conv = ConversationHandler(
        entry_points=[CommandHandler('setxlsx', set_xlsx_start)],
        states={
            WAITING_XLSX_CONFIRM: [
                MessageHandler(
                    filters.Document.MimeType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ), handle_xlsx),
                CallbackQueryHandler(xlsx_callback, pattern="^xlsx_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, xlsx_new_name)
            ],
        },
        fallbacks=[CommandHandler('start', start)],
        allow_reentry=True
    )

    app.add_handler(pdf_conv)
    app.add_handler(xlsx_conv)
    app.add_handler(CommandHandler('start', start))
    app.add_handler(CommandHandler('help', help_cmd))
    app.add_handler(CommandHandler('objects', list_objects))
    app.add_handler(CommandHandler('report', report_cmd))
    app.add_handler(CommandHandler('export', export_cmd))
    app.add_handler(CallbackQueryHandler(general_callback))

    logger.info("🤖 Бот запущен!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == '__main__':
    main()
