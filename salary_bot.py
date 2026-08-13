import asyncio
import os
import re
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

DATA_DIR = Path(__file__).resolve().parent / "data"
DATA_DIR.mkdir(exist_ok=True)

MONTH_PATTERN = re.compile(r"^\d{4}-\d{2}$")
DATE_PATTERN = re.compile(r"^(\d{2}|\d{1,2})[-.,](\d{2}|\d{1,2})[-.,](\d{4})$")
WAIT_MONTH, WAIT_BIG, WAIT_SMALL, WAIT_CONFIRM, WAIT_NEXT_ACTION, WAIT_MONTH_SUM = range(6)
BERLIN_TZ = ZoneInfo("Europe/Berlin")
WORKPLACE_1 = "place_1"
WORKPLACE_2 = "place_2"


def money(value) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def format_money(value) -> str:
    return f"{money(value):.2f}"


def parse_date_input(raw_date: str):
    normalized = raw_date.strip().replace('.', '-').replace(',', '-')
    try:
        return datetime.strptime(normalized, "%d-%m-%Y")
    except ValueError:
        return None


def get_today_in_berlin():
    return datetime.now(BERLIN_TZ).date()


def build_main_menu(month: str | None = None, location: str | None = None) -> InlineKeyboardMarkup:
    if month is None:
        month = datetime.now(BERLIN_TZ).strftime("%Y-%m")
    if location is None:
        location = WORKPLACE_1

    has_data = get_month_file(month, location).exists()

    buttons = [[
        InlineKeyboardButton("➕ Додати день", callback_data="add_more"),
        InlineKeyboardButton("📊 Сума за місяць", callback_data="show_month_total" if has_data else "show_month_total_disabled"),
    ], [
        InlineKeyboardButton("📝 Редагувати день", callback_data="edit_day_menu"),
        InlineKeyboardButton("🏠 Меню", callback_data="close_entry"),
    ]]

    return InlineKeyboardMarkup(buttons)


def get_month_file(month: str, location: str | None = None) -> Path:
    file_name = f"{month}.xlsx"
    if location:
        file_name = f"{location}_{file_name}"
    return DATA_DIR / file_name


def get_available_months(location: str | None = None) -> list[str]:
    months = set()
    for file_path in sorted(DATA_DIR.glob("*.xlsx")):
        file_name = file_path.name
        if location:
            prefix = f"{location}_"
            if not file_name.startswith(prefix):
                continue
            month = file_name[len(prefix):-5]
        else:
            month = file_name[:-5]
        if MONTH_PATTERN.match(month):
            months.add(month)
    return sorted(months)


def build_month_selection_keyboard(location: str | None = None) -> InlineKeyboardMarkup:
    months = get_available_months(location)
    if not months:
        return InlineKeyboardMarkup([[InlineKeyboardButton("✍️ Ввести вручну", callback_data="manual_month_total")]])

    rows = []
    for month in months:
        rows.append([InlineKeyboardButton(month, callback_data=f"month_total_{month}")])
    rows.append([InlineKeyboardButton("✍️ Ввести вручну", callback_data="manual_month_total")])
    return InlineKeyboardMarkup(rows)


def build_month_result_keyboard(month: str, location: str | None = None) -> InlineKeyboardMarkup:
    buttons = [[
        InlineKeyboardButton("📥 Завантажити файл", callback_data=f"download_{month}"),
        InlineKeyboardButton("🏠 Меню", callback_data="close_entry"),
    ]]
    return InlineKeyboardMarkup(buttons)


def ensure_month_file(month: str, location: str | None = None) -> Path:
    file_path = get_month_file(month, location)
    if file_path.exists():
        return file_path

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Salary"
    sheet.append(["Date", "Big boxes", "Big pay", "Small boxes", "Small pay", "Day total"])
    workbook.save(file_path)
    return file_path


def save_day_entry(month: str, day: int, big_count: int, small_count: int, location: str | None = None) -> Decimal:
    file_path = ensure_month_file(month, location)

    workbook = load_workbook(file_path)
    sheet = workbook["Salary"]
    date_value = f"{month}-{day:02d}"

    big_total = money(Decimal(big_count) * Decimal("0.9") / Decimal("2"))
    small_total = money(Decimal(small_count) * Decimal("0.7") / Decimal("2"))
    day_total = money(big_total + small_total)

    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value == date_value:
            sheet.cell(row=row, column=2, value=int(big_count))
            sheet.cell(row=row, column=3, value=float(big_total))
            sheet.cell(row=row, column=4, value=int(small_count))
            sheet.cell(row=row, column=5, value=float(small_total))
            sheet.cell(row=row, column=6, value=float(day_total))
            workbook.save(file_path)
            return day_total

    sheet.append([
        date_value,
        int(big_count),
        float(big_total),
        int(small_count),
        float(small_total),
        float(day_total),
    ])
    workbook.save(file_path)
    return day_total


def read_month_total(month: str, location: str | None = None) -> Decimal:
    file_path = get_month_file(month, location)
    if not file_path.exists():
        return Decimal("0.00")

    workbook = load_workbook(file_path)
    sheet = workbook["Salary"]

    total = Decimal("0.00")
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        if row[5] is not None:
            total += Decimal(str(row[5]))
    return money(total)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    if not context.user_data.get("location"):
        context.user_data["location"] = WORKPLACE_1
    await update.message.reply_text(
        "Введи дату: 15-08-2026\nМожна: 15.08.2026 або 15,08,2026",
        reply_markup=build_main_menu(location=context.user_data["location"]),
    )
    return WAIT_MONTH


async def process_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw_date = update.message.text.strip()

    if context.user_data.get("need_month_input"):
        month = raw_date
        if not MONTH_PATTERN.match(month):
            await update.message.reply_text("Невірний формат місяця. Введи YYYY-MM.")
            return WAIT_MONTH_SUM

        location = context.user_data.get("location", WORKPLACE_1)
        total = read_month_total(month, location=location)
        file_path = get_month_file(month, location=location)
        context.user_data.pop("need_month_input", None)

        if not file_path.exists():
            await update.message.reply_text(f"За місяць {month} ще немає записів.")
            return WAIT_MONTH

        await update.message.reply_text(
            f"{location}: {format_money(total)} €\nФайл: {file_path.name}",
            reply_markup=build_main_menu(month=month, location=location),
        )
        return WAIT_NEXT_ACTION

    if DATE_PATTERN.match(raw_date):
        parsed_date = parse_date_input(raw_date)
        if parsed_date is None:
            await update.message.reply_text("Невірна дата. Введи: 15-08-2026")
            return WAIT_MONTH

        selected_date = parsed_date.date()
        today = get_today_in_berlin()
        if selected_date > today:
            await update.message.reply_text(
                f"Не можна вносити дані за майбутні дати. Сьогодні: {today.strftime('%d-%m-%Y')}."
            )
            return WAIT_MONTH

        month = parsed_date.strftime("%Y-%m")
        day = parsed_date.day
        context.user_data["month"] = month
        context.user_data["day"] = day
        context.user_data["date_text"] = selected_date.strftime("%d-%m-%Y")

        if context.user_data.get("edit_mode") == "edit":
            context.user_data.pop("edit_mode", None)
            edit_date = selected_date.strftime("%Y-%m-%d")
            file_path = get_month_file(month, context.user_data.get("location", WORKPLACE_1))
            if not file_path.exists():
                await update.message.reply_text(f"Для {month} ще немає записів.")
                return WAIT_MONTH
            workbook = load_workbook(file_path)
            sheet = workbook["Salary"]
            for row in range(2, sheet.max_row + 1):
                if str(sheet.cell(row=row, column=1).value) == edit_date:
                    big_count = int(sheet.cell(row=row, column=2).value)
                    small_count = int(sheet.cell(row=row, column=4).value)
                    await update.message.reply_text(
                        f"Запис: {edit_date}\nВеликі: {big_count}\nМалі: {small_count}\nНові значення: big small",
                        reply_markup=build_main_menu(location=context.user_data.get("location", WORKPLACE_1)),
                    )
                    context.user_data["edit_date"] = edit_date
                    context.user_data["edit_month"] = month
                    return WAIT_MONTH
            await update.message.reply_text(f"Для {edit_date} записів не знайдено.")
            return WAIT_MONTH

        await update.message.reply_text("Великі коробки?")
        return WAIT_BIG

    await update.message.reply_text("Невірний формат. Введи: 15-08-2026")
    return WAIT_MONTH


async def process_month_total_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    month = update.message.text.strip()
    if not MONTH_PATTERN.match(month):
        await update.message.reply_text("Невірний формат місяця. Введи YYYY-MM.")
        return WAIT_MONTH_SUM

    location = context.user_data.get("location", WORKPLACE_1)
    total = read_month_total(month, location=location)
    file_path = get_month_file(month, location=location)
    if not file_path.exists():
        await update.message.reply_text(f"За місяць {month} ще немає записів.")
        context.user_data.pop("need_month_input", None)
        return WAIT_MONTH

    context.user_data.pop("need_month_input", None)
    await update.message.reply_text(
        f"{location}: {format_money(total)} €\nФайл: {file_path.name}",
        reply_markup=build_month_result_keyboard(month, location=location),
    )
    return WAIT_NEXT_ACTION


async def process_big_boxes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = int(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Введи корректне число для великих коробок.")
        return WAIT_BIG

    if value < 0:
        await update.message.reply_text("Кількість не може бути від'ємною.")
        return WAIT_BIG

    context.user_data["big_count"] = value
    await update.message.reply_text("Малі коробки?")
    return WAIT_SMALL


async def process_small_boxes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        value = int(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Введи корректне число для малих коробок.")
        return WAIT_SMALL

    if value < 0:
        await update.message.reply_text("Кількість не може бути від'ємною.")
        return WAIT_SMALL

    month = context.user_data["month"]
    day = context.user_data["day"]
    big_count = context.user_data["big_count"]
    context.user_data["small_count"] = value

    big_total = money(Decimal(big_count) * Decimal("0.9") / Decimal("2"))
    small_total = money(Decimal(value) * Decimal("0.7") / Decimal("2"))
    day_total = money(big_total + small_total)

    keyboard = [[
        InlineKeyboardButton("💾 Зберегти", callback_data="save_entry"),
        InlineKeyboardButton("✏️ Відредагувати", callback_data="edit_entry"),
    ]]

    await update.message.reply_text(
        f"Перевірка:\n"
        f"Дата: {context.user_data['date_text']}\n"
        f"Великі коробки: {big_count}\n"
        f"Малі коробки: {value}\n"
        f"Розрахунок:\n"
        f"Великі: {big_count} × 0.90 / 2 = {format_money(big_total)} €\n"
        f"Малі: {value} × 0.70 / 2 = {format_money(small_total)} €\n"
        f"Сума за день: {format_money(day_total)} €\n\n"
        f"Все вірно?",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return WAIT_CONFIRM


async def confirm_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "save_entry":
        month = context.user_data["month"]
        day = context.user_data["day"]
        big_count = context.user_data["big_count"]
        small_count = context.user_data["small_count"]
        location = context.user_data.get("location", WORKPLACE_1)

        big_total = money(Decimal(big_count) * Decimal("0.9") / Decimal("2"))
        small_total = money(Decimal(small_count) * Decimal("0.7") / Decimal("2"))
        day_total = money(big_total + small_total)
        save_day_entry(month, day, big_count, small_count, location=location)
        monthly_total = read_month_total(month, location=location)

        keyboard = [[
            InlineKeyboardButton("➕ Додати ще один день", callback_data="add_more"),
            InlineKeyboardButton("📊 Показати місячну суму", callback_data="show_month_total"),
        ], [
            InlineKeyboardButton("📥 Завантажити файл", callback_data=f"download_{month}"),
            InlineKeyboardButton("❌ Закрити", callback_data="close_entry"),
        ]]
        context.user_data["last_month"] = month
        context.user_data["last_day"] = day
        context.user_data["location"] = location
        await query.edit_message_text(
            f"✅ Збережено\n"
            f"{month}-{day:02d}\n"
            f"Великі: {big_count} = {format_money(big_total)} €\n"
            f"Малі: {small_count} = {format_money(small_total)} €\n"
            f"Сума: {format_money(day_total)} €\n"
            f"Місяць: {format_money(monthly_total)} €",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        context.user_data.clear()
        context.user_data["last_month"] = month
        context.user_data["last_day"] = day
        context.user_data["location"] = location
        return WAIT_NEXT_ACTION

    if query.data == "edit_entry":
        context.user_data.pop("big_count", None)
        context.user_data.pop("small_count", None)
        await query.edit_message_text(
            "Ок. Введи дату ще раз у форматі DD-MM-YYYY, 15.08.2026 або 15,08,2026."
        )
        return WAIT_MONTH

    return WAIT_CONFIRM


async def next_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "add_more":
        location = context.user_data.get("location", WORKPLACE_1)
        context.user_data.clear()
        context.user_data["location"] = location
        await query.edit_message_text(
            "Ок. Введи дату ще раз у форматі DD-MM-YYYY, наприклад 15-08-2026.\n"
            "Можна також вводити 15.08.2026 або 15,08,2026."
        )
        return WAIT_MONTH

    if query.data == "show_month_total":
        location = context.user_data.get("location", WORKPLACE_1)
        months = get_available_months(location)
        if not months:
            await query.answer("Поки немає збережених місяців", show_alert=True)
            return WAIT_MONTH
        await query.edit_message_text(
            "Вибери місяць:",
            reply_markup=build_month_selection_keyboard(location=location),
        )
        return WAIT_MONTH_SUM

    if query.data == "manual_month_total":
        context.user_data["need_month_input"] = True
        await query.edit_message_text(
            "Який місяць показати? Введи YYYY-MM",
            reply_markup=build_main_menu(location=context.user_data.get("location", WORKPLACE_1)),
        )
        return WAIT_MONTH_SUM

    if query.data == "show_month_total_disabled":
        await query.answer("За цей місяць ще немає записів", show_alert=True)
        return WAIT_MONTH

    if query.data.startswith("month_total_"):
        month = query.data.replace("month_total_", "")
        if not MONTH_PATTERN.match(month):
            await query.answer("Невірний місяць", show_alert=True)
            return WAIT_MONTH

        location = context.user_data.get("location", WORKPLACE_1)
        total = read_month_total(month, location=location)
        file_path = get_month_file(month, location=location)
        if not file_path.exists():
            await query.answer("Файл для цього місяця не знайдено", show_alert=True)
            return WAIT_MONTH

        await query.edit_message_text(
            f"{location}: {format_money(total)} €\nФайл: {file_path.name}",
            reply_markup=build_month_result_keyboard(month, location=location),
        )
        return WAIT_NEXT_ACTION

    if query.data.startswith("download_"):
        month = query.data.replace("download_", "")
        file_path = get_month_file(month, context.user_data.get("location", WORKPLACE_1))
        if file_path.exists():
            await query.answer("Файл надсилається...")
            with open(file_path, "rb") as file:
                await query.message.reply_document(document=file, filename=file_path.name)
            await query.message.reply_text(
                f"Файл {file_path.name} надіслано в чат.",
                reply_markup=build_main_menu(location=context.user_data.get("location", WORKPLACE_1)),
            )
            return WAIT_NEXT_ACTION
        await query.answer("Файл не знайдено", show_alert=True)
        return WAIT_NEXT_ACTION

    if query.data == "edit_day_menu":
        context.user_data["edit_mode"] = "edit"
        await query.edit_message_text(
            "Введи дату для редагування у форматі DD-MM-YYYY, наприклад 12-08-2026.",
            reply_markup=build_main_menu(),
        )
        return WAIT_MONTH

    if query.data == "close_entry":
        location = context.user_data.get("location", WORKPLACE_1)
        context.user_data.clear()
        context.user_data["location"] = location
        await query.edit_message_text(
            "Готово. Повертаюсь в головне меню.",
            reply_markup=build_main_menu(location=context.user_data["location"]),
        )
        return WAIT_MONTH

    return WAIT_NEXT_ACTION


async def month_total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args:
        await update.message.reply_text("Який місяць показати? Введи YYYY-MM")
        return

    month = args[0].strip()
    if not MONTH_PATTERN.match(month):
        await update.message.reply_text("Невірний формат місяця. Введи YYYY-MM.")
        return

    location = context.user_data.get("location", WORKPLACE_1)
    total = read_month_total(month, location=location)
    file_path = get_month_file(month, location=location)
    if not file_path.exists():
        await update.message.reply_text(f"За місяць {month} ще немає записів.")
        return

    await update.message.reply_text(
        f"{location}: {format_money(total)} €\n"
        f"Файл: {file_path.name}"
    )


async def edit_day_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if not args:
        await update.message.reply_text("Введи дату у форматі: /edit 12-08-2026")
        return

    raw_date = args[0].strip()
    parsed_date = parse_date_input(raw_date)
    if parsed_date is None:
        await update.message.reply_text("Невірна дата. Введи: 12-08-2026")
        return

    month = parsed_date.strftime("%Y-%m")
    location = context.user_data.get("location", WORKPLACE_1)
    file_path = get_month_file(month, location)
    if not file_path.exists():
        await update.message.reply_text(f"Для {month} ще немає записів.")
        return

    workbook = load_workbook(file_path)
    sheet = workbook["Salary"]
    date_value = parsed_date.strftime("%Y-%m-%d")

    for row in range(2, sheet.max_row + 1):
        if str(sheet.cell(row=row, column=1).value) == date_value:
            big_count = int(sheet.cell(row=row, column=2).value)
            small_count = int(sheet.cell(row=row, column=4).value)
            await update.message.reply_text(
                f"Запис: {date_value}\nВеликі: {big_count}\nМалі: {small_count}\nНові значення: big small",
            )
            context.user_data["edit_date"] = date_value
            context.user_data["edit_month"] = month
            context.user_data["edit_location"] = location
            return

    await update.message.reply_text(f"Для {date_value} записів не знайдено.")


async def process_edit_values(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "edit_date" not in context.user_data:
        return

    text = update.message.text.strip()
    try:
        big_count, small_count = map(int, text.split())
    except ValueError:
        await update.message.reply_text("Введи: big small, наприклад 200 150")
        return

    date_value = context.user_data["edit_date"]
    month = context.user_data["edit_month"]
    location = context.user_data.get("edit_location", context.user_data.get("location", WORKPLACE_1))
    day = int(date_value.split('-')[2])

    save_day_entry(month, day, big_count, small_count, location=location)
    monthly_total = read_month_total(month, location=location)
    await update.message.reply_text(
        f"Оновлено {date_value}\nВеликі: {big_count}\nМалі: {small_count}\nСума: {format_money(monthly_total)} €"
    )
    context.user_data.pop("edit_date", None)
    context.user_data.pop("edit_month", None)
    context.user_data.pop("edit_location", None)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Скасовано. Для нового запису натисни /start.")
    return ConversationHandler.END


def main() -> None:
    try:
        from config import BOT_TOKEN as token
    except ImportError:
        token = os.getenv("BOT_TOKEN")

    if not token or token == "YOUR_BOT_TOKEN_HERE":
        raise RuntimeError(
            "BOT token is not set. Put your token in config.py or set BOT_TOKEN environment variable."
        )

    asyncio.set_event_loop(asyncio.new_event_loop())
    application = Application.builder().token(token).build()

    conversation = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAIT_MONTH: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_month),
                CallbackQueryHandler(next_action, pattern=r"^(add_more|show_month_total|show_month_total_disabled|edit_day_menu|close_entry|manual_month_total)$"),
            ],
            WAIT_BIG: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_big_boxes),
            ],
            WAIT_SMALL: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_small_boxes),
            ],
            WAIT_CONFIRM: [
                CallbackQueryHandler(confirm_entry, pattern="^(save_entry|edit_entry)$"),
            ],
            WAIT_NEXT_ACTION: [
                CallbackQueryHandler(next_action, pattern=r"^(add_more|show_month_total|show_month_total_disabled|download_\d{4}-\d{2}|edit_day_menu|close_entry|manual_month_total|month_total_\d{4}-\d{2})$"),
            ],
            WAIT_MONTH_SUM: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_month_total_input),
                CallbackQueryHandler(next_action, pattern=r"^(manual_month_total|month_total_\d{4}-\d{2})$"),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    application.add_handler(conversation)
    application.add_handler(CommandHandler("total", month_total))
    application.add_handler(CommandHandler("edit", edit_day_data))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, process_edit_values))
    application.add_handler(CommandHandler("menu", lambda update, context: start(update, context)))
    application.add_handler(CommandHandler("cancel", cancel))

    print("Bot started...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
