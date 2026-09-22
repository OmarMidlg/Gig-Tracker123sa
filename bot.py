#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
بوت تيليجرام لتسجيل أيام الشغل (Gig Tracker)
Made by OmarMidlg

يسجّل كل مستخدم بياناته الخاصة (رقم الحدث، التاريخ، اليوم، مكان الحدث،
نوع الشغل، مين اشتغل معه، والمبلغ)، ويحسب ملخص الحساب تلقائياً.
"""

import asyncio
import logging
import os
import sqlite3
from datetime import datetime
from contextlib import closing

from telegram import ReplyKeyboardMarkup, KeyboardButton, Update
from telegram.ext import (
    Application,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ----------------------------------------------------------------------
# إعدادات عامة
# ----------------------------------------------------------------------

BOT_TOKEN = os.environ.get("BOT_TOKEN", "8789370643:AAFlsbbu3GCllvundSrCPqBpGe5o08VD4RY")
DB_PATH = os.environ.get("DB_PATH", "gig_tracker.db")
DEFAULT_FEE = 120.00  # الأجرة الافتراضية لكل حدث بالشيكل (NIS)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

ARABIC_DAYS = {
    6: "الأحد",     # Monday=0 ... Sunday=6 في بايثون
    0: "الإثنين",
    1: "الثلاثاء",
    2: "الأربعاء",
    3: "الخميس",
    4: "الجمعة",
    5: "السبت",
}

# حالات المحادثة - إضافة يومية جديدة
ADD_DATE, ADD_LOCATION, ADD_JOB, ADD_PARTNER, ADD_MONEY = range(5)
# حالة تعديل قيمة إعداد (رصيد سابق / خصم / مبلغ مستلم / أجرة افتراضية)
SETTING_VALUE = 100
# حالة تأكيد حذف يومية
DELETE_CONFIRM = 200

MAIN_MENU = ReplyKeyboardMarkup(
    [
        [KeyboardButton("➕ إضافة يومية")],
        [KeyboardButton("📋 آخر الأيام"), KeyboardButton("📊 ملخص الحساب")],
        [KeyboardButton("🗑️ حذف يومية"), KeyboardButton("⚙️ الإعدادات")],
    ],
    resize_keyboard=True,
)

SETTINGS_MENU = ReplyKeyboardMarkup(
    [
        [KeyboardButton("💰 تعديل الرصيد من الحساب السابق")],
        [KeyboardButton("➖ تعديل مبلغ الخصم")],
        [KeyboardButton("✅ تعديل المبلغ المستلم")],
        [KeyboardButton("🏷️ تعديل الأجرة الافتراضية")],
        [KeyboardButton("⬅️ رجوع للقائمة الرئيسية")],
    ],
    resize_keyboard=True,
)

CANCEL_HINT = "\n\nأرسل /cancel في أي وقت للإلغاء."


# ----------------------------------------------------------------------
# قاعدة البيانات
# ----------------------------------------------------------------------

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                full_name TEXT,
                previous_balance REAL NOT NULL DEFAULT 0,
                discount REAL NOT NULL DEFAULT 0,
                amount_received REAL NOT NULL DEFAULT 0,
                default_fee REAL NOT NULL DEFAULT %f
            )
            """
            % DEFAULT_FEE
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                event_number INTEGER NOT NULL,
                event_date TEXT NOT NULL,
                day_name TEXT NOT NULL,
                location TEXT,
                job TEXT,
                partner TEXT,
                money REAL NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(user_id)
            )
            """
        )


def ensure_user(user_id: int, full_name: str = ""):
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """
            INSERT INTO users (user_id, full_name, previous_balance, discount,
                                amount_received, default_fee)
            VALUES (?, ?, 0, 0, 0, ?)
            ON CONFLICT(user_id) DO NOTHING
            """,
            (user_id, full_name, DEFAULT_FEE),
        )


def get_user_settings(user_id: int):
    with closing(get_conn()) as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE user_id = ?", (user_id,)
        ).fetchone()
        return row


def get_next_event_number(user_id: int) -> int:
    with closing(get_conn()) as conn:
        row = conn.execute(
            "SELECT COALESCE(MAX(event_number), 0) + 1 AS n FROM events WHERE user_id = ?",
            (user_id,),
        ).fetchone()
        return row["n"]


def add_event(user_id, event_number, event_date, day_name, location, job, partner, money):
    with closing(get_conn()) as conn, conn:
        conn.execute(
            """
            INSERT INTO events (user_id, event_number, event_date, day_name,
                                 location, job, partner, money, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_id, event_number, event_date, day_name,
                location, job, partner, money,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )


def get_events(user_id, limit=None):
    with closing(get_conn()) as conn:
        q = "SELECT * FROM events WHERE user_id = ? ORDER BY event_number DESC"
        if limit:
            q += f" LIMIT {int(limit)}"
        return conn.execute(q, (user_id,)).fetchall()


def get_event_stats(user_id):
    with closing(get_conn()) as conn:
        row = conn.execute(
            "SELECT COUNT(*) AS n, COALESCE(SUM(money), 0) AS total FROM events WHERE user_id = ?",
            (user_id,),
        ).fetchone()
        return row["n"], row["total"]


def delete_event(user_id, event_number) -> bool:
    with closing(get_conn()) as conn, conn:
        cur = conn.execute(
            "DELETE FROM events WHERE user_id = ? AND event_number = ?",
            (user_id, event_number),
        )
        return cur.rowcount > 0


def update_user_field(user_id, field, value):
    assert field in {"previous_balance", "discount", "amount_received", "default_fee"}
    with closing(get_conn()) as conn, conn:
        conn.execute(f"UPDATE users SET {field} = ? WHERE user_id = ?", (value, user_id))


# ----------------------------------------------------------------------
# أدوات مساعدة
# ----------------------------------------------------------------------

def fmt_money(x: float) -> str:
    return f"{x:,.2f} ₪"


def parse_date(text: str):
    """يقبل التاريخ بصيغة يوم/شهر/سنة أو يوم-شهر-سنة، أو كلمة 'اليوم'."""
    text = text.strip()
    if text in ("اليوم", "اليوم.", "today"):
        return datetime.now()
    text = text.replace(".", "/").replace("-", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_number(text: str):
    text = text.strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def build_summary_text(user_id: int) -> str:
    settings = get_user_settings(user_id)
    n_events, total_money = get_event_stats(user_id)

    previous_balance = settings["previous_balance"]
    discount = settings["discount"]
    amount_received = settings["amount_received"]

    net_amount = total_money
    remaining_balance = previous_balance + net_amount - discount - amount_received

    text = (
        "📊 *ملخص الحساب*\n"
        "━━━━━━━━━━━━━━\n"
        f"الرصيد من الحساب السابق: {fmt_money(previous_balance)}\n"
        f"عدد الأيام (الأحداث): {n_events}\n"
        f"الصافي من الأيام المسجّلة: {fmt_money(net_amount)}\n"
        f"مبلغ الخصم: {fmt_money(discount)}\n"
        f"المبلغ المستلم: {fmt_money(amount_received)}\n"
        "━━━━━━━━━━━━━━\n"
        f"💰 *الرصيد المتبقي: {fmt_money(remaining_balance)}*\n"
    )
    return text


# ----------------------------------------------------------------------
# أوامر عامة
# ----------------------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    ensure_user(user.id, user.full_name or "")
    await update.message.reply_text(
        "أهلاً بك 👋\n"
        "هذا بوت لتسجيل أيام الشغل مع محمد زيدان (وأي حد بتشتغل معه).\n\n"
        "اختر من القائمة تحت:",
        reply_markup=MAIN_MENU,
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "الأوامر المتاحة:\n"
        "➕ إضافة يومية - لتسجيل حدث/يوم شغل جديد\n"
        "📋 آخر الأيام - عرض آخر الأيام المسجّلة\n"
        "📊 ملخص الحساب - عرض ملخص الحساب الكامل\n"
        "🗑️ حذف يومية - حذف يومية برقمها\n"
        "⚙️ الإعدادات - تعديل الرصيد السابق / الخصم / المبلغ المستلم / الأجرة الافتراضية\n\n"
        "/cancel - لإلغاء أي عملية جارية\n\n"
        "🛠️ Made by OmarMidlg",
        reply_markup=MAIN_MENU,
    )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("تم الإلغاء ❌", reply_markup=MAIN_MENU)
    return ConversationHandler.END


async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ensure_user(update.effective_user.id, update.effective_user.full_name or "")
    text = build_summary_text(update.effective_user.id)
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=MAIN_MENU)


async def show_events(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ensure_user(user_id, update.effective_user.full_name or "")
    events = get_events(user_id, limit=10)
    if not events:
        await update.message.reply_text("لا يوجد أي أيام مسجّلة بعد.", reply_markup=MAIN_MENU)
        return

    lines = ["📋 *آخر الأيام المسجّلة (حتى 10):*\n"]
    for e in events:
        lines.append(
            f"#{e['event_number']} | {e['event_date']} ({e['day_name']})\n"
            f"  📍 {e['location'] or '-'} | 🛠️ {e['job'] or '-'}\n"
            f"  👤 معي: {e['partner'] or '-'} | 💵 {fmt_money(e['money'])}\n"
        )
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown", reply_markup=MAIN_MENU)


async def open_settings_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⚙️ الإعدادات - اختر ما تريد تعديله:", reply_markup=SETTINGS_MENU)


async def back_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("القائمة الرئيسية:", reply_markup=MAIN_MENU)


# ----------------------------------------------------------------------
# محادثة: إضافة يومية جديدة
# ----------------------------------------------------------------------

async def add_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ensure_user(update.effective_user.id, update.effective_user.full_name or "")
    context.user_data["new_event"] = {}
    await update.message.reply_text(
        "📅 أرسل تاريخ اليومية بصيغة يوم/شهر/سنة (مثال: 21/09/2026)\n"
        "أو اكتب 'اليوم' لاستخدام تاريخ اليوم." + CANCEL_HINT,
        reply_markup=ReplyKeyboardMarkup([["اليوم"]], resize_keyboard=True),
    )
    return ADD_DATE


async def add_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    dt = parse_date(update.message.text)
    if not dt:
        await update.message.reply_text(
            "صيغة التاريخ غير صحيحة، حاول مثلاً: 21/09/2026 أو اكتب 'اليوم'."
        )
        return ADD_DATE

    day_name = ARABIC_DAYS[dt.weekday()]
    context.user_data["new_event"]["event_date"] = dt.strftime("%d/%m/%Y")
    context.user_data["new_event"]["day_name"] = day_name

    await update.message.reply_text(
        f"اليوم: {day_name} ✅\n\n📍 وين كان مكان الحدث؟" + CANCEL_HINT,
        reply_markup=ReplyKeyboardMarkup([[]], resize_keyboard=True),
    )
    return ADD_LOCATION


async def add_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_event"]["location"] = update.message.text.strip()
    await update.message.reply_text("🛠️ شو كان نوع الشغل؟" + CANCEL_HINT)
    return ADD_JOB


async def add_job(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_event"]["job"] = update.message.text.strip()
    await update.message.reply_text(
        "👤 مين اشتغل معك بهاليوم؟ (اكتب '-' إذا اشتغلت لحالك)" + CANCEL_HINT
    )
    return ADD_PARTNER


async def add_partner(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_event"]["partner"] = update.message.text.strip()
    settings = get_user_settings(update.effective_user.id)
    default_fee = settings["default_fee"]
    await update.message.reply_text(
        f"💵 كم المبلغ لهاليومية؟\n"
        f"الأجرة الافتراضية هي {fmt_money(default_fee)} - اضغط الزر لاستخدامها أو اكتب مبلغ آخر."
        + CANCEL_HINT,
        reply_markup=ReplyKeyboardMarkup(
            [[f"{default_fee:.2f}"]], resize_keyboard=True
        ),
    )
    return ADD_MONEY


async def add_money(update: Update, context: ContextTypes.DEFAULT_TYPE):
    money = parse_number(update.message.text)
    if money is None or money < 0:
        await update.message.reply_text("الرجاء إرسال رقم صحيح للمبلغ.")
        return ADD_MONEY

    user_id = update.effective_user.id
    ev = context.user_data["new_event"]
    event_number = get_next_event_number(user_id)

    add_event(
        user_id=user_id,
        event_number=event_number,
        event_date=ev["event_date"],
        day_name=ev["day_name"],
        location=ev["location"],
        job=ev["job"],
        partner=ev["partner"],
        money=money,
    )

    context.user_data.pop("new_event", None)

    await update.message.reply_text(
        "✅ تم تسجيل اليومية بنجاح!\n\n"
        f"رقم الحدث: #{event_number}\n"
        f"التاريخ: {ev['event_date']} ({ev['day_name']})\n"
        f"المكان: {ev['location']}\n"
        f"الشغل: {ev['job']}\n"
        f"اشتغل معي: {ev['partner']}\n"
        f"المبلغ: {fmt_money(money)}",
        reply_markup=MAIN_MENU,
    )
    return ConversationHandler.END


# ----------------------------------------------------------------------
# محادثة: حذف يومية
# ----------------------------------------------------------------------

async def delete_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🗑️ أرسل رقم الحدث (event number) الذي تريد حذفه." + CANCEL_HINT
    )
    return DELETE_CONFIRM


async def delete_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().lstrip("#")
    if not text.isdigit():
        await update.message.reply_text("الرجاء إرسال رقم حدث صحيح.")
        return DELETE_CONFIRM

    event_number = int(text)
    ok = delete_event(update.effective_user.id, event_number)
    if ok:
        await update.message.reply_text(f"✅ تم حذف اليومية رقم #{event_number}.", reply_markup=MAIN_MENU)
    else:
        await update.message.reply_text(f"❌ لم يتم العثور على يومية برقم #{event_number}.", reply_markup=MAIN_MENU)
    return ConversationHandler.END


# ----------------------------------------------------------------------
# محادثة: تعديل الإعدادات (رصيد سابق / خصم / مستلم / أجرة افتراضية)
# ----------------------------------------------------------------------

SETTINGS_FIELD_MAP = {
    "💰 تعديل الرصيد من الحساب السابق": ("previous_balance", "الرصيد من الحساب السابق"),
    "➖ تعديل مبلغ الخصم": ("discount", "مبلغ الخصم"),
    "✅ تعديل المبلغ المستلم": ("amount_received", "المبلغ المستلم"),
    "🏷️ تعديل الأجرة الافتراضية": ("default_fee", "الأجرة الافتراضية لكل حدث"),
}


async def settings_field_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    field, label = SETTINGS_FIELD_MAP[update.message.text]
    context.user_data["setting_field"] = field
    context.user_data["setting_label"] = label

    settings = get_user_settings(update.effective_user.id)
    current_value = settings[field]

    await update.message.reply_text(
        f"القيمة الحالية لـ «{label}»: {fmt_money(current_value)}\n\n"
        f"أرسل القيمة الجديدة (رقم فقط):" + CANCEL_HINT,
        reply_markup=ReplyKeyboardMarkup([[]], resize_keyboard=True),
    )
    return SETTING_VALUE


async def settings_field_save(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = parse_number(update.message.text)
    if value is None or value < 0:
        await update.message.reply_text("الرجاء إرسال رقم صحيح.")
        return SETTING_VALUE

    field = context.user_data.pop("setting_field")
    label = context.user_data.pop("setting_label")
    update_user_field(update.effective_user.id, field, value)

    await update.message.reply_text(
        f"✅ تم تحديث «{label}» إلى {fmt_money(value)}.", reply_markup=MAIN_MENU
    )
    return ConversationHandler.END


# ----------------------------------------------------------------------
# تشغيل البوت
# ----------------------------------------------------------------------

def main():
    if not BOT_TOKEN or BOT_TOKEN == "PUT-YOUR-TELEGRAM-BOT-TOKEN-HERE":
        raise SystemExit(
            "الرجاء ضبط متغير البيئة BOT_TOKEN بتوكن البوت الخاص بك قبل التشغيل.\n"
            "مثال: export BOT_TOKEN=123456:ABC-yourtoken"
        )

    init_db()

    app = Application.builder().token(BOT_TOKEN).build()

    add_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^➕ إضافة يومية$"), add_start)],
        states={
            ADD_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_date)],
            ADD_LOCATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_location)],
            ADD_JOB: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_job)],
            ADD_PARTNER: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_partner)],
            ADD_MONEY: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_money)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    delete_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^🗑️ حذف يومية$"), delete_start)],
        states={
            DELETE_CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, delete_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    settings_conv = ConversationHandler(
        entry_points=[
            MessageHandler(
                filters.Regex(
                    "^(💰 تعديل الرصيد من الحساب السابق|➖ تعديل مبلغ الخصم|"
                    "✅ تعديل المبلغ المستلم|🏷️ تعديل الأجرة الافتراضية)$"
                ),
                settings_field_start,
            )
        ],
        states={
            SETTING_VALUE: [MessageHandler(filters.TEXT & ~filters.COMMAND, settings_field_save)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("cancel", cancel))

    app.add_handler(add_conv)
    app.add_handler(delete_conv)
    app.add_handler(settings_conv)

    app.add_handler(MessageHandler(filters.Regex("^📊 ملخص الحساب$"), show_summary))
    app.add_handler(MessageHandler(filters.Regex("^📋 آخر الأيام$"), show_events))
    app.add_handler(MessageHandler(filters.Regex("^⚙️ الإعدادات$"), open_settings_menu))
    app.add_handler(MessageHandler(filters.Regex("^⬅️ رجوع للقائمة الرئيسية$"), back_to_main))

    # Python 3.14 removed implicit event-loop creation in the main thread,
    # which python-telegram-bot's run_polling() still relies on internally.
    # Creating and setting the loop explicitly here keeps this working on
    # both Python 3.14+ and older versions.
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        try:
            asyncio.get_event_loop()
        except RuntimeError:
            asyncio.set_event_loop(asyncio.new_event_loop())

    logger.info("Bot starting... Made by OmarMidlg")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
