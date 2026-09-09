import io
import logging
import os
import re
from datetime import date, datetime, timedelta, time
from typing import Optional

import openpyxl
import requests

from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse

from telegram import (
    Update,
    ReplyKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# ============================================================
# НАСТРОЙКИ
# ============================================================

# Для компьютера можно вставить токен прямо сюда.
# Для Render токен лучше хранить в Environment Variables.
TOKEN = os.getenv("BOT_TOKEN")
AUTO_CHAT_ID = os.getenv("AUTO_CHAT_ID")

# ID твоей Google-таблицы
SPREADSHEET_ID = "1UmJKEHIy7vYc3matpjZ9ZC7NqiVY3QHETa4f4DotAlY"

# Наша группа
GROUP = "451"

# В таблице:
# K = предмет / преподаватель
# L = кабинет
SUBJECT_COL = 11
ROOM_COL = 12

# В таблице:
# A = дата
# C = номер пары
# D = время
DATE_COL = 1
PAIR_COL = 3
TIME_COL = 4

# Часовой пояс школы.Q
# Барнаул / Алтайский край = UTC+7.
TIMEZONE = "Asia/Barnaul"

# Во сколько автоматически присылать расписание.
AUTO_SEND_TIME = time(hour=7, minute=0)


# ============================================================
# ЛОГИ
# ============================================================

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)

logger = logging.getLogger(__name__)


# ============================================================
# КНОПКИ
# ============================================================

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["📅 Сегодня", "➡️ Завтра"],
        ["📆 Эта неделя", "🔄 Обновить"],
        ["⚙️ Авто-расписание"],
    ],
    resize_keyboard=True,
)


# ============================================================
# GOOGLE SHEETS
# ============================================================

def download_google_sheet():
    """
    Скачивает актуальную версию Google Sheets в формате XLSX.
    """

    url = (
        f"https://docs.google.com/spreadsheets/d/"
        f"{SPREADSHEET_ID}/export?format=xlsx"
    )

    response = requests.get(
        url,
        timeout=30,
    )

    response.raise_for_status()

    return openpyxl.load_workbook(
        io.BytesIO(response.content),
        data_only=True,
    )


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def clean(value) -> Optional[str]:
    """
    Приводит значение ячейки к аккуратному тексту.
    """

    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    # Убираем лишние переводы строк
    text = re.sub(r"\s+", " ", text)

    return text


def normalize_date(value):
    """
    Преобразует содержимое ячейки в date.
    """

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    return None


def day_name(d: date):
    names = [
        "Понедельник",
        "Вторник",
        "Среда",
        "Четверг",
        "Пятница",
        "Суббота",
        "Воскресенье",
    ]

    return names[d.weekday()]


# ============================================================
# ПОЛУЧЕНИЕ РАСПИСАНИЯ
# ============================================================

def get_schedule(target_date: date):
    """
    Возвращает расписание группы 451 на указанную дату.
    """

    workbook = download_google_sheet()

    result = []

    for sheet_name in workbook.sheetnames:

        ws = workbook[sheet_name]

        current_date = None

        # Ищем даты и расписание внутри листа
        for row in range(1, ws.max_row + 1):

            # ------------------------------------------------
            # Определяем текущую дату
            # ------------------------------------------------

            value = ws.cell(
                row=row,
                column=DATE_COL,
            ).value

            parsed_date = normalize_date(value)

            if parsed_date:
                current_date = parsed_date

            if current_date != target_date:
                continue

            # ------------------------------------------------
            # Читаем основные значения
            # ------------------------------------------------

            pair_value = ws.cell(
                row=row,
                column=PAIR_COL,
            ).value

            lesson_time = clean(
                ws.cell(
                    row=row,
                    column=TIME_COL,
                ).value
            )

            subject = clean(
                ws.cell(
                    row=row,
                    column=SUBJECT_COL,
                ).value
            )

            room = clean(
                ws.cell(
                    row=row,
                    column=ROOM_COL,
                ).value
            )

            # ------------------------------------------------
            # Обычная пара
            # ------------------------------------------------

            if pair_value is not None and subject:

                try:
                    pair = int(float(pair_value)) # type: ignore
                except (ValueError, TypeError):
                    pair = None

                teacher = None

                # В твоей таблице преподаватель находится
                # на следующей строке в колонке K.
                next_row = row + 1

                if next_row <= ws.max_row:

                    possible_teacher = clean(
                        ws.cell(
                            row=next_row,
                            column=SUBJECT_COL,
                        ).value
                    )

                    if possible_teacher:

                        # Если это не такой же текст,
                        # считаем это преподавателем.
                        if possible_teacher != subject:
                            teacher = possible_teacher

                result.append(
                    {
                        "pair": pair,
                        "time": lesson_time,
                        "subject": subject,
                        "teacher": teacher,
                        "room": room,
                    }
                )

            # ------------------------------------------------
            # Особые строки без номера пары
            #
            # Например:
            # "Классный час"
            # ------------------------------------------------

            elif lesson_time and subject:

                # Проверяем, не является ли значение
                # преподавателем.
                lower_subject = subject.lower()

                teacher_words = [
                    "преподаватель:",
                ]

                if not any(
                    word in lower_subject
                    for word in teacher_words
                ):

                    result.append(
                        {
                            "pair": None,
                            "time": lesson_time,
                            "subject": subject,
                            "teacher": None,
                            "room": room,
                        }
                    )

    # Убираем дубли
    unique = []

    seen = set()

    for item in result:

        key = (
            item["pair"],
            item["time"],
            item["subject"],
            item["room"],
        )

        if key not in seen:

            seen.add(key)
            unique.append(item)

    # Сортируем по номеру пары
    unique.sort(
        key=lambda x: (
            x["pair"] is None,
            x["pair"] if x["pair"] is not None else 999,
        )
    )

    return unique


# ============================================================
# ФОРМАТИРОВАНИЕ
# ============================================================

def format_day(target_date: date, lessons):
    """
    Красиво формирует сообщение Telegram.
    """

    text = (
        f"📅 <b>{day_name(target_date)}</b>\n"
        f"📆 {target_date.strftime('%d.%m.%Y')}\n"
        f"🎓 <b>Группа 451</b>\n"
        f"\n"
    )

    if not lessons:

        text += "😎 <b>Пар нет.</b> Можно отдыхать!"

        return text

    for lesson in lessons:

        if lesson["pair"] is not None:

            text += (
                f"🔔 <b>{lesson['pair']} пара</b>"
            )

        else:

            text += "🔔 <b>Дополнительно</b>"

        if lesson["time"]:
            text += f"  ⏰ {lesson['time']}"

        text += "\n"

        text += (
            f"📚 <b>{lesson['subject']}</b>\n"
        )

        if lesson["teacher"]:

            text += (
                f"👨‍🏫 {lesson['teacher']}\n"
            )

        if lesson["room"]:

            text += (
                f"🚪 Кабинет: {lesson['room']}\n"
            )

        text += "\n"

    return text


# ============================================================
# ПОЛУЧЕНИЕ СЕГОДНЯ
# ============================================================

# ============================================================
# СЕГОДНЯ
# ============================================================

async def send_today(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    try:
        target_date = date.today()
        lessons = get_schedule(target_date)
        text = format_day(target_date, lessons)

        assert update.message is not None
        await update.message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=MAIN_KEYBOARD,
        )

    except Exception:
        logger.exception("Ошибка получения расписания на сегодня")

        assert update.message is not None
        await update.message.reply_text(
            "❌ Не удалось получить расписание.",
            reply_markup=MAIN_KEYBOARD,
        )


async def send_tomorrow(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):

    try:

        target_date = (
            date.today()
            + timedelta(days=1)
        )

        lessons = get_schedule(target_date)

        text = format_day(
            target_date,
            lessons,
        )

        assert update.message is not None
        await update.message.reply_text(
            text,
            parse_mode="HTML",
            reply_markup=MAIN_KEYBOARD,
        )

    except Exception:

        logger.exception(
            "Ошибка получения расписания"
        )

        assert update.message is not None
        await update.message.reply_text(
            "❌ Не удалось получить расписание.",
            reply_markup=MAIN_KEYBOARD,
        )


# ============================================================
# НЕДЕЛЯ
# ============================================================

async def send_week(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):

    try:

        today_date = date.today()

        monday = (
            today_date
            - timedelta(
                days=today_date.weekday()
            )
        )

        full_text = (
            "🎓 <b>Расписание группы 451</b>\n"
            "📆 <b>Текущая неделя</b>\n\n"
        )

        for i in range(6):

            target_date = (
                monday
                + timedelta(days=i)
            )

            lessons = get_schedule(
                target_date
            )

            full_text += format_day(
                target_date,
                lessons,
            )

            full_text += (
                "\n"
                "━━━━━━━━━━━━━━\n\n"
            )

        # Telegram имеет ограничение на размер сообщения.
        # Отправляем большими кусками.
        for start in range(
            0,
            len(full_text),
            3900,
        ):

            assert update.message is not None
            await update.message.reply_text(
                full_text[
                    start:start + 3900
                ],
                parse_mode="HTML",
                reply_markup=MAIN_KEYBOARD,
            )

    except Exception:

        logger.exception(
            "Ошибка получения недели"
        )

        assert update.message is not None
        await update.message.reply_text(
            "❌ Не удалось получить расписание.",
            reply_markup=MAIN_KEYBOARD,
        )


# ============================================================
# АВТОМАТИЧЕСКАЯ ОТПРАВКА
# ============================================================

async def automatic_schedule(
    context: ContextTypes.DEFAULT_TYPE,
):

    job = context.job

    assert job is not None
    chat_id = job.chat_id

    try:

        target_date = date.today()

        lessons = get_schedule(
            target_date
        )

        text = format_day(
            target_date,
            lessons,
        )

        await context.bot.send_message(
            chat_id=chat_id, # type: ignore
            text=(
                "🌅 <b>Доброе утро!</b>\n\n"
                + text
            ),
            parse_mode="HTML",
            reply_markup=MAIN_KEYBOARD,
        )

    except Exception:

        logger.exception(
            "Ошибка автоматической отправки"
        )


async def enable_auto(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    assert update.message is not None
    await update.message.reply_text(
        "✅ <b>Авто-расписание включено.</b>\n\n"
        "Каждый день в <b>07:00 по времени Барнаула</b> расписание будет отправляться автоматически.\n\n"
        "На сервере это выполняется отдельным бесплатным планировщиком.",
        parse_mode="HTML",
        reply_markup=MAIN_KEYBOARD,
    )

async def disable_auto(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    assert update.message is not None
    await update.message.reply_text(
        "🛑 Чтобы отключить ежедневную отправку, убери AUTO_CHAT_ID из настроек сервера/планировщика.",
        reply_markup=MAIN_KEYBOARD,
    )

# ============================================================
# START
# ============================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):

    assert update.message is not None
    await update.message.reply_text(
        "👋 <b>Привет!</b>\n\n"
        "Я бот расписания группы <b>451</b>.\n\n"
        "Выбери нужный пункт ниже 👇",
        parse_mode="HTML",
        reply_markup=MAIN_KEYBOARD,
    )


# ============================================================
# ОБНОВИТЬ
# ============================================================

async def refresh(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):

    assert update.message is not None
    await update.message.reply_text(
        "🔄 Загружаю самое свежее расписание..."
    )

    await send_today(
        update,
        context,
    )


# ============================================================
# ОБРАБОТКА КНОПОК
# ============================================================

async def button_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):

    assert update.message is not None
    text = update.message.text

    if text == "📅 Сегодня":

        await send_today(
            update,
            context,
        )

    elif text == "➡️ Завтра":

        await send_tomorrow(
            update,
            context,
        )

    elif text == "📆 Эта неделя":

        await send_week(
            update,
            context,
        )

    elif text == "🔄 Обновить":

        await refresh(
            update,
            context,
        )

    elif text == "⚙️ Авто-расписание":

        await enable_auto(
            update,
            context,
        )


# ============================================================
# MAIN
# ============================================================

application = Application.builder().token(TOKEN or "").build()

application.add_handler(CommandHandler("start", start))
application.add_handler(CommandHandler("myid", myid))
application.add_handler(CommandHandler("today", send_today))
application.add_handler(CommandHandler("tomorrow", send_tomorrow))
application.add_handler(CommandHandler("week", send_week))
application.add_handler(CommandHandler("auto", enable_auto))
application.add_handler(CommandHandler("stop", disable_auto))
application.add_handler(
    MessageHandler(filters.TEXT & ~filters.COMMAND, button_handler)
)

app = FastAPI()


@app.get("/", response_class=PlainTextResponse)
async def health():
    return "Schedule bot is running"


@app.post("/telegram/webhook")
async def telegram_webhook(request: Request):
    data = await request.json()
    update = Update.de_json(data, application.bot)
    await application.process_update(update)
    return {"ok": True}


@app.on_event("startup")
async def startup():
    if not TOKEN:
        raise RuntimeError("BOT_TOKEN не задан в Render Environment Variables")
    if not os.getenv("RENDER_EXTERNAL_URL"):
        raise RuntimeError("RENDER_EXTERNAL_URL не найден")

    await application.initialize()
    await application.start()

    webhook_url = os.getenv("RENDER_EXTERNAL_URL").rstrip("/") + "/telegram/webhook"
    await application.bot.set_webhook(url=webhook_url)
    logger.info("Webhook установлен: %s", webhook_url)


@app.on_event("shutdown")
async def shutdown():
    await application.stop()
    await application.shutdown()


# ============================================================
