import asyncio
import os
import re
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from zoneinfo import ZoneInfo
import json
import logging
from telegram.error import BadRequest
from telegram import CallbackQuery

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

LOG_FILE = DATA_DIR / "salary_bot.log"
logger = logging.getLogger("salary_bot")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh.setFormatter(formatter)
    logger.addHandler(fh)

# Patch CallbackQuery.edit_message_text to ignore 'Message is not modified' BadRequest
_original_cq_edit = CallbackQuery.edit_message_text
async def _safe_cq_edit(self, *args, **kwargs):
    try:
        return await _original_cq_edit(self, *args, **kwargs)
    except BadRequest as e:
        if "Message is not modified" in str(e):
            try:
                await self.answer()
            except Exception:
                pass
            logger.info("Ignored Message is not modified when editing callback message")
            return None
        # otherwise re-raise (but log)
        logger.exception("BadRequest in edit_message_text: %s", e)
        raise
CallbackQuery.edit_message_text = _safe_cq_edit

MONTH_PATTERN = re.compile(r"^\d{4}-\d{2}$")
DATE_PATTERN = re.compile(r"^(\d{2}|\d{1,2})[-.,](\d{2}|\d{1,2})[-.,](\d{4})$")
WAIT_MONTH, WAIT_BIG, WAIT_SMALL, WAIT_CONFIRM, WAIT_NEXT_ACTION, WAIT_MONTH_SUM = range(6)
BERLIN_TZ = ZoneInfo("Europe/Berlin")
WORKPLACE_1 = "place_1"
WORKPLACE_2 = "place_2"
# Admin user IDs who can see all users' data
ADMIN_IDS = {442336138}


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


def build_main_menu(month: str | None = None, location: str | None = None, caller_id: int | None = None) -> InlineKeyboardMarkup:
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

    # If caller is admin, add admin menu row
    if caller_id is not None and caller_id in ADMIN_IDS:
        buttons.append([
            InlineKeyboardButton("👥 Користувачі", callback_data="admin_users"),
            InlineKeyboardButton("🗒️ Логи", callback_data="admin_logs"),
        ])

    return InlineKeyboardMarkup(buttons)


def get_month_file(month: str, location: str | None = None, user_id: int | None = None) -> Path:
    """File naming:
    - user-specific: {user_id}_{month}.xlsx
    - location-specific legacy: {location}_{month}.xlsx
    - generic: {month}.xlsx
    """
    file_name = f"{month}.xlsx"
    if user_id is not None:
        file_name = f"{user_id}_{file_name}"
    elif location:
        file_name = f"{location}_{file_name}"
    return DATA_DIR / file_name


def get_available_months(location: str | None = None, user_id: int | None = None) -> list[str]:
    months = set()
    for file_path in sorted(DATA_DIR.glob("*.xlsx")):
        file_name = file_path.name
        if user_id is not None:
            prefix = f"{user_id}_"
            if not file_name.startswith(prefix):
                continue
            month = file_name[len(prefix):-5]
        elif location:
            prefix = f"{location}_"
            if not file_name.startswith(prefix):
                continue
            month = file_name[len(prefix):-5]
        else:
            month = file_name[:-5]
        if MONTH_PATTERN.match(month):
            months.add(month)
    return sorted(months)


def build_month_selection_keyboard(location: str | None = None, user_id: int | None = None) -> InlineKeyboardMarkup:
    months = get_available_months(location=location, user_id=user_id)
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


def get_user_ids() -> list[int]:
    ids = set()
    for file_path in sorted(DATA_DIR.glob("*_????-??.xlsx")):
        parts = file_path.name.split("_")
        if len(parts) >= 2 and parts[0].isdigit():
            ids.add(int(parts[0]))
    return sorted(ids)


def users_map_file() -> Path:
    return DATA_DIR / "users.json"


def load_user_map() -> dict:
    path = users_map_file()
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_user_map(m: dict):
    path = users_map_file()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=2)


def register_user(user) -> None:
    if user is None:
        return
    uid = getattr(user, "id", None)
    if uid is None:
        return
    uname = getattr(user, "username", None) or getattr(user, "full_name", None) or ""
    m = load_user_map()
    # store string keys for JSON stability
    m[str(uid)] = uname
    save_user_map(m)


def ensure_month_file(month: str, location: str | None = None) -> Path:
    file_path = get_month_file(month, location)
    if file_path.exists():
        return file_path

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Salary"
    sheet.append(["Дата", "Великі коробки", "Малі коробки", "Загальна сума за день"])
    workbook.save(file_path)
    logger.info(f"Created new month file: {file_path}")
    return file_path


def save_day_entry(month: str, day: int, big_count: int, small_count: int, location: str | None = None, user_id: int | None = None) -> Decimal:
    file_path = ensure_month_file(month, location)
    # If user_id provided, prefer user-specific file
    if user_id is not None:
        file_path = ensure_month_file(month, location=None)
        # ensure_month_file doesn't know about user_id, so build path and create if missing
        user_file = get_month_file(month, location=None, user_id=user_id)
        if not user_file.exists():
            # create workbook for user
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Salary"
            sheet.append(["Дата", "Великі коробки", "Малі коробки", "Загальна сума за день"])
            workbook.save(user_file)
            logger.info(f"[USER:{user_id}] Created new user month file: {user_file}")
        file_path = user_file

    workbook = load_workbook(file_path)
    sheet = workbook["Salary"]
    date_value = f"{month}-{day:02d}"

    big_total = money(Decimal(big_count) * Decimal("0.9") / Decimal("2"))
    small_total = money(Decimal(small_count) * Decimal("0.7") / Decimal("2"))
    day_total = money(big_total + small_total)

    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value == date_value:
            sheet.cell(row=row, column=2, value=int(big_count))
            sheet.cell(row=row, column=3, value=int(small_count))
            sheet.cell(row=row, column=4, value=float(day_total))
            workbook.save(file_path)
            logger.info(f"[USER:{user_id}] Updated file {file_path.name} for date {date_value} big={big_count} small={small_count} total={day_total}")
            return day_total

    sheet.append([
        date_value,
        int(big_count),
        int(small_count),
        float(day_total),
    ])
    workbook.save(file_path)
    logger.info(f"[USER:{user_id}] Appended date {date_value} to {file_path.name} big={big_count} small={small_count} total={day_total}")
    return day_total


def read_month_total(month: str, location: str | None = None, user_id: int | None = None) -> Decimal:
    # If user_id provided, read user-specific file
    if user_id is not None:
        file_path = get_month_file(month, location=None, user_id=user_id)
        if not file_path.exists():
            return Decimal("0.00")
        workbook = load_workbook(file_path)
        sheet = workbook["Salary"]
        total = Decimal("0.00")
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            # Support legacy format (6 cols) where day total was column 6 (index 5)
            if len(row) >= 6 and row[5] is not None:
                total += Decimal(str(row[5]))
            # New format: day total is at column 4 (index 3)
            elif len(row) >= 4 and row[3] is not None:
                total += Decimal(str(row[3]))
        return money(total)

    # Otherwise, aggregate across matching files (location-specific or generic)
    total = Decimal("0.00")
    pattern = f"*_{month}.xlsx"
    for file_path in sorted(DATA_DIR.glob(pattern)):
        workbook = load_workbook(file_path)
        sheet = workbook["Salary"]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            # Support legacy format (6 cols) where day total was column 6 (index 5)
            if len(row) >= 6 and row[5] is not None:
                total += Decimal(str(row[5]))
            # New format: day total is at column 4 (index 3)
            elif len(row) >= 4 and row[3] is not None:
                total += Decimal(str(row[3]))
    return money(total)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    if not context.user_data.get("location"):
        context.user_data["location"] = WORKPLACE_1
    # register user for admin listing
    try:
        register_user(update.message.from_user)
    except Exception:
        pass
    await update.message.reply_text(
        "Введи дату: 15-08-2026\nМожна: 15.08.2026 або 15,08,2026",
        reply_markup=build_main_menu(location=context.user_data["location"], caller_id=update.message.from_user.id),
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
            reply_markup=build_main_menu(month=month, location=location, caller_id=update.message.from_user.id),
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
                    big_count = int(sheet.cell(row=row, column=2).value) if sheet.cell(row=row, column=2).value is not None else 0
                    # detect old format (6 columns) where small_count was column 4
                    small_col = 4 if sheet.max_column >= 6 else 3
                    small_count = int(sheet.cell(row=row, column=small_col).value) if sheet.cell(row=row, column=small_col).value is not None else 0
                    await update.message.reply_text(
                        f"Запис: {edit_date}\nВеликі: {big_count}\nМалі: {small_count}\nНові значення: big small",
                        reply_markup=build_main_menu(location=context.user_data.get("location", WORKPLACE_1), caller_id=update.message.from_user.id),
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

    caller_id = update.message.from_user.id if update.message.from_user else None
    location = context.user_data.get("location", WORKPLACE_1)

    if caller_id in ADMIN_IDS:
        total = read_month_total(month)
        files = list(DATA_DIR.glob(f"*_{month}.xlsx"))
        if not files:
            await update.message.reply_text(f"За місяць {month} ще немає записів.")
            context.user_data.pop("need_month_input", None)
            return WAIT_MONTH
        file_list = ", ".join([f.name for f in files])
        context.user_data.pop("need_month_input", None)
        await update.message.reply_text(
            f"All users: {format_money(total)} €\nФайли: {file_list}",
            reply_markup=build_main_menu(month=month, location=location),
        )
        return WAIT_NEXT_ACTION

    total = read_month_total(month, user_id=caller_id)
    file_path = get_month_file(month, user_id=caller_id)
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
        month = context.user_data.get("month")
        day = context.user_data.get("day")
        big_count = context.user_data.get("big_count", 0)
        small_count = context.user_data.get("small_count", 0)
        location = context.user_data.get("location", WORKPLACE_1)

        # If both counts are zero, do not write to Excel
        if int(big_count) == 0 and int(small_count) == 0:
            # clear temporary data and return to main menu
            context.user_data.clear()
            await query.edit_message_text(
                "Нічого не збережено: обидві кількості дорівнюють 0.",
                reply_markup=build_main_menu(location=location, caller_id=query.from_user.id if query.from_user else None),
            )
            return WAIT_NEXT_ACTION

        big_total = money(Decimal(big_count) * Decimal("0.9") / Decimal("2"))
        small_total = money(Decimal(small_count) * Decimal("0.7") / Decimal("2"))
        day_total = money(big_total + small_total)
        caller_id = query.from_user.id if query.from_user else None

        # register user (store username for admin view)
        try:
            register_user(query.from_user)
            uid_reg = getattr(query.from_user, 'id', None)
            logger.info(f"[USER:{uid_reg}] Registered user: id={uid_reg} username={getattr(query.from_user, 'username', None)}")
        except Exception as e:
                uid_reg = getattr(query.from_user, 'id', None)
                logger.exception(f"[USER:{uid_reg}] Error registering user: {e}")
        # save to user-specific file
        try:
            save_day_entry(month, day, big_count, small_count, location=location, user_id=caller_id)
        except Exception as e:
                logger.exception(f"[USER:{caller_id}] Failed to save entry for {month}-{day:02d}: {e}")

        # admin sees aggregated month, others see their own month        if caller_id in ADMIN_IDS:            monthly_total = read_month_total(month)
        else:
            monthly_total = read_month_total(month, user_id=caller_id)

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
        # keep last_* in user_data but clear other temporary fields
        temp_last_month = context.user_data.pop("last_month", month)
        temp_last_day = context.user_data.pop("last_day", day)
        temp_location = context.user_data.pop("location", location)
        context.user_data.clear()
        context.user_data["last_month"] = temp_last_month
        context.user_data["last_day"] = temp_last_day
        context.user_data["location"] = temp_location
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
    caller_id = query.from_user.id if query.from_user else None

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
        caller_id = query.from_user.id if query.from_user else None
        # admin sees aggregated months across users, others see only their months
        months = get_available_months(location=location, user_id=(None if caller_id in ADMIN_IDS else caller_id))
        if not months:
            await query.answer("Поки немає збережених місяців", show_alert=True)
            return WAIT_MONTH
        await query.edit_message_text(
            "Вибери місяць:",
            reply_markup=build_month_selection_keyboard(location=location, user_id=(None if caller_id in ADMIN_IDS else caller_id)),
        )
        return WAIT_MONTH_SUM

    if query.data == "admin_users":
        # list distinct user ids from files
        user_ids = get_user_ids()
        if not user_ids:
            await query.answer("Нема користувачів", show_alert=True)
            return WAIT_MONTH
        user_map = load_user_map()
        rows = []
        for uid in user_ids:
            uname = user_map.get(str(uid))
            if uname:
                label = f"{uid} — @{uname}"
            else:
                label = str(uid)
            rows.append([InlineKeyboardButton(label, callback_data=f"admin_user_{uid}")])
        rows.append([InlineKeyboardButton("🏠 Меню", callback_data="close_entry")])
        await query.edit_message_text("Оберіть користувача:", reply_markup=InlineKeyboardMarkup(rows))
        return WAIT_MONTH

    if query.data == "admin_logs":
        caller_id = query.from_user.id if query.from_user else None
        if caller_id not in ADMIN_IDS:
            await query.answer("Доступ заборонено", show_alert=True)
            return WAIT_MONTH
        kb = [
            [InlineKeyboardButton("Останні 100 рядків", callback_data="admin_log_tail_100"), InlineKeyboardButton("Останні 500 рядків", callback_data="admin_log_tail_500")],
            [InlineKeyboardButton("Повний файл", callback_data="admin_log_file"), InlineKeyboardButton("По користувачу", callback_data="admin_log_by_user")],
            [InlineKeyboardButton("🏠 Меню", callback_data="close_entry")],
        ]
        await query.edit_message_text("Журнал (логи): оберіть опцію:", reply_markup=InlineKeyboardMarkup(kb))
        return WAIT_MONTH

    if query.data == "admin_log_by_user":
        caller_id = query.from_user.id if query.from_user else None
        if caller_id not in ADMIN_IDS:
            await query.answer("Доступ заборонено", show_alert=True)
            return WAIT_MONTH
        # list users
        user_ids = get_user_ids()
        if not user_ids:
            await query.answer("Нема користувачів", show_alert=True)
            return WAIT_MONTH
        rows = []
        user_map = load_user_map()
        for uid in user_ids:
            uname = user_map.get(str(uid))
            label = f"{uid} — @{uname}" if uname else str(uid)
            rows.append([InlineKeyboardButton(label, callback_data=f"admin_log_user_{uid}")])
        rows.append([InlineKeyboardButton("🏠 Меню", callback_data="close_entry")])
        await query.edit_message_text("Оберіть користувача для логів:", reply_markup=InlineKeyboardMarkup(rows))
        return WAIT_MONTH

    if query.data.startswith("admin_log_user_"):
        uid_str = query.data.replace("admin_log_user_", "")
        try:
            sel_uid = int(uid_str)
        except ValueError:
            await query.answer("Невірний користувач", show_alert=True)
            return WAIT_MONTH
        # show options for this user
        kb = [
            [InlineKeyboardButton("Останні 100 рядків", callback_data=f"admin_log_user_tail_{sel_uid}_100"), InlineKeyboardButton("Останні 500 рядків", callback_data=f"admin_log_user_tail_{sel_uid}_500")],
            [InlineKeyboardButton("Повний файл", callback_data=f"admin_log_user_file_{sel_uid}"), InlineKeyboardButton("🏠 Меню", callback_data="close_entry")],
        ]
        await query.edit_message_text(f"Логи для користувача {sel_uid}:", reply_markup=InlineKeyboardMarkup(kb))
        return WAIT_MONTH

    if query.data.startswith("admin_log_user_tail_"):
        # format: admin_log_user_tail_<uid>_<lines>
        parts = query.data.split("_")
        try:
            sel_uid = int(parts[4])
            lines = int(parts[5])
        except Exception:
            await query.answer("Невірний запит", show_alert=True)
            return WAIT_NEXT_ACTION
        if not LOG_FILE.exists():
            await query.answer("Файл лога не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        import re as _re
        pattern = _re.compile(rf"\b{sel_uid}\b")
        with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
            content_lines = f.read().splitlines()
        filtered = [l for l in content_lines if pattern.search(l) or f"{sel_uid}_" in l]
        tail_lines = filtered[-lines:]
        tail_text = "\n".join(tail_lines)
        if not tail_text:
            await query.answer("Записів для цього користувача не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        if len(tail_text) > 3500 or len(tail_lines) > 200:
            tmp = DATA_DIR / f"salary_bot_log_user_{sel_uid}_{caller_id}.txt"
            with open(tmp, "w", encoding="utf-8") as tf:
                tf.write(tail_text)
            with open(tmp, "rb") as tf:
                await query.message.reply_document(document=tf, filename=tmp.name)
            try:
                tmp.unlink()
            except Exception:
                pass
            logger.info(f"[ADMIN:{caller_id}] Sent filtered log tail ({lines}) for [USER:{sel_uid}] as document")
            return WAIT_NEXT_ACTION
        await query.message.reply_text(f"Останні {len(tail_lines)} рядків лога для {sel_uid}:\n\n{tail_text}")
        logger.info(f"[ADMIN:{caller_id}] Sent filtered log tail ({lines}) for [USER:{sel_uid}] as message")
        return WAIT_NEXT_ACTION

    if query.data.startswith("admin_log_user_file_"):
        # format: admin_log_user_file_<uid>
        parts = query.data.split("_")
        try:
            sel_uid = int(parts[4])
        except Exception:
            await query.answer("Невірний запит", show_alert=True)
            return WAIT_NEXT_ACTION
        if not LOG_FILE.exists():
            await query.answer("Файл лога не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        import re as _re
        pattern = _re.compile(rf"\b{sel_uid}\b")
        with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
            content_lines = f.read().splitlines()
        filtered = [l for l in content_lines if pattern.search(l) or f"{sel_uid}_" in l]
        if not filtered:
            await query.answer("Записів для цього користувача не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        tmp = DATA_DIR / f"salary_bot_log_user_{sel_uid}_{caller_id}.txt"
        with open(tmp, "w", encoding="utf-8") as tf:
            tf.write("\n".join(filtered))
        with open(tmp, "rb") as tf:
            await query.message.reply_document(document=tf, filename=tmp.name)
        try:
            tmp.unlink()
        except Exception:
            pass
        logger.info(f"[ADMIN:{caller_id}] Sent filtered full log file for [USER:{sel_uid}]")
        return WAIT_NEXT_ACTION

    if query.data.startswith("admin_user_"):
        uid_str = query.data.replace("admin_user_", "")
        try:
            sel_uid = int(uid_str)
        except ValueError:
            await query.answer("Невірний користувач", show_alert=True)
            return WAIT_MONTH
        # store selection for admin view
        context.user_data["admin_view_user"] = sel_uid
        await query.edit_message_text(
            f"Оберіть місяць для користувача {sel_uid}:",
            reply_markup=build_month_selection_keyboard(user_id=sel_uid),
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

        caller_id = query.from_user.id if query.from_user else None
        location = context.user_data.get("location", WORKPLACE_1)
        if caller_id in ADMIN_IDS:
            # If admin selected a specific user earlier, show that user's file
            admin_sel = context.user_data.get("admin_view_user")
            if admin_sel:
                total = read_month_total(month, user_id=admin_sel)
                file_path = get_month_file(month, user_id=admin_sel)
                if not file_path.exists():
                    await query.answer("Файл для цього місяця не знайдено", show_alert=True)
                    return WAIT_MONTH
                await query.edit_message_text(
                    f"User {admin_sel}: {format_money(total)} €\nФайл: {file_path.name}",
                    reply_markup=build_month_result_keyboard(month, location=location),
                )
                context.user_data.pop("admin_view_user", None)
                return WAIT_NEXT_ACTION

            # otherwise aggregate across all users
            total = read_month_total(month)
            files = list(DATA_DIR.glob(f"*_{month}.xlsx"))
            file_list = ", ".join([f.name for f in files]) if files else "(файлів немає)"
            await query.edit_message_text(
                f"All users: {format_money(total)} €\nФайли: {file_list}",
                reply_markup=build_main_menu(month=month, location=location, caller_id=query.from_user.id),
            )
            return WAIT_NEXT_ACTION

        total = read_month_total(month, user_id=caller_id)
        file_path = get_month_file(month, user_id=caller_id)
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
        caller_id = query.from_user.id if query.from_user else None
        if caller_id in ADMIN_IDS:
            files = list(DATA_DIR.glob(f"*_{month}.xlsx"))
            logger.info(f"[ADMIN:{caller_id}] Requested download for month {month}. Found files: {[f.name for f in files]}")
            if not files:
                await query.answer("Файлів не знайдено", show_alert=True)
                return WAIT_NEXT_ACTION
            await query.answer("Файли надсилаються...")
            for file_path in files:
                try:
                    with open(file_path, "rb") as file:
                        await query.message.reply_document(document=file, filename=file_path.name)
                        logger.info(f"[ADMIN:{caller_id}] Sent file {file_path.name}")
                except Exception as e:
                    logger.exception(f"[ADMIN:{caller_id}] Failed to send file {file_path} to admin: {e}")
                    # notify admin about the failure for this file
                    await query.message.reply_text(f"Не вдалося надіслати файл {file_path.name}: {e}")
            await query.message.reply_text(
                f"Всі файли за {month} надіслані.",
                reply_markup=build_main_menu(location=context.user_data.get("location", WORKPLACE_1), caller_id=query.from_user.id),
            )
            return WAIT_NEXT_ACTION

        file_path = get_month_file(month, user_id=caller_id)
        logger.info(f"[USER:{caller_id}] Requested download for month {month}. File expected: {file_path.name}")
        if file_path.exists():
            try:
                await query.answer("Файл надсилається...")
                with open(file_path, "rb") as file:
                    await query.message.reply_document(document=file, filename=file_path.name)
                await query.message.reply_text(
                    f"Файл {file_path.name} надіслано в чат.",
                    reply_markup=build_main_menu(location=context.user_data.get("location", WORKPLACE_1), caller_id=query.from_user.id),
                )
                logger.info(f"[USER:{caller_id}] Sent file {file_path.name}")
                return WAIT_NEXT_ACTION
            except Exception as e:
                logger.exception(f"[USER:{caller_id}] Failed to send file {file_path} to user: {e}")
                await query.answer("Помилка при відправці файлу", show_alert=True)
                return WAIT_NEXT_ACTION
        await query.answer("Файл не знайдено", show_alert=True)
        return WAIT_NEXT_ACTION

    if query.data.startswith("admin_log_tail_"):
        # send tail of log
        try:
            lines = int(query.data.replace("admin_log_tail_", ""))
        except Exception:
            lines = 100
        if not LOG_FILE.exists():
            await query.answer("Файл лога не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
            content_lines = f.read().splitlines()
        tail_lines = content_lines[-lines:]
        tail_text = "\n".join(tail_lines)
        if len(tail_text) > 3500 or len(tail_lines) > 200:
            tmp = DATA_DIR / f"salary_bot_log_tail_{caller_id}.txt"
            with open(tmp, "w", encoding="utf-8") as tf:
                tf.write(tail_text)
            with open(tmp, "rb") as tf:
                await query.message.reply_document(document=tf, filename=tmp.name)
            try:
                tmp.unlink()
            except Exception:
                pass
            logger.info(f"[ADMIN:{caller_id}] Sent log tail ({lines}) as document")
            return WAIT_NEXT_ACTION
        await query.message.reply_text(f"Останні {len(tail_lines)} рядків лога:\n\n{tail_text}")
        logger.info(f"[ADMIN:{caller_id}] Sent log tail ({lines}) as message")
        return WAIT_NEXT_ACTION

    if query.data == "admin_log_file":
        if not LOG_FILE.exists():
            await query.answer("Файл лога не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        with open(LOG_FILE, "rb") as f:
            await query.message.reply_document(document=f, filename=LOG_FILE.name)
        logger.info(f"Sent full log file to admin {caller_id}")
        return WAIT_NEXT_ACTION

    if query.data == "edit_day_menu":
        context.user_data["edit_mode"] = "edit"
        await query.edit_message_text(
            "Введи дату для редагування у форматі DD-MM-YYYY, наприклад 12-08-2026.",
            reply_markup=build_main_menu(caller_id=query.from_user.id),
        )
        return WAIT_MONTH

    if query.data == "close_entry":
        location = context.user_data.get("location", WORKPLACE_1)
        context.user_data.clear()
        context.user_data["location"] = location
        await query.edit_message_text(
            "Готово. Повертаюсь в головне меню.",
            reply_markup=build_main_menu(location=context.user_data["location"], caller_id=query.from_user.id),
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


async def logtail(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin command: /logtail [N|file]
    - /logtail 100  -> send last 100 lines
    - /logtail file -> send full log file as document
    If not admin, respond unauthorized.
    """
    user = update.message.from_user if update.message else None
    uid = getattr(user, "id", None)
    if uid not in ADMIN_IDS:
        await update.message.reply_text("Доступ заборонено.")
        return

    args = context.args
    mode = args[0].strip().lower() if args else "100"
    try:
        if mode == "file":
            # send full log file
            if not LOG_FILE.exists():
                await update.message.reply_text("Файл лога не знайдено.")
                return
            with open(LOG_FILE, "rb") as f:
                await update.message.reply_document(document=f, filename=LOG_FILE.name)
            logger.info(f"Admin {uid} downloaded full log file")
            return

        # else treat as number of lines
        lines = int(mode)
    except Exception:
        lines = 100

    if not LOG_FILE.exists():
        await update.message.reply_text("Файл лога не знайдено.")
        return

    try:
        with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
            content_lines = f.read().splitlines()
    except Exception as e:
        logger.exception(f"Failed to read log file: {e}")
        await update.message.reply_text(f"Не вдалося прочитати лог: {e}")
        return

    tail_lines = content_lines[-lines:]
    tail_text = "\n".join(tail_lines)

    # If text too long, send as document
    if len(tail_text) > 3500 or len(tail_lines) > 200:
        tmp = DATA_DIR / f"salary_bot_log_tail_{uid}.txt"
        try:
            with open(tmp, "w", encoding="utf-8") as tf:
                tf.write(tail_text)
            with open(tmp, "rb") as tf:
                await update.message.reply_document(document=tf, filename=tmp.name)
            logger.info(f"Admin {uid} requested log tail {lines} lines; sent as document {tmp.name}")
        except Exception as e:
            logger.exception(f"Failed to send log tail to admin {uid}: {e}")
            await update.message.reply_text(f"Не вдалося надіслати лог: {e}")
        finally:
            try:
                tmp.unlink()
            except Exception:
                pass
        return

    await update.message.reply_text(f"Останні {len(tail_lines)} рядків лога:\n\n{tail_text}")
    logger.info(f"Admin {uid} requested log tail {lines} lines sent as message")


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
                CallbackQueryHandler(next_action, pattern=r"^(add_more|show_month_total|show_month_total_disabled|edit_day_menu|close_entry|manual_month_total|admin_users|admin_user_\d+|admin_logs|admin_log_by_user|admin_log_user_\d+|admin_log_user_tail_\d+_\d+|admin_log_user_file_\d+)$"),
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
                CallbackQueryHandler(next_action, pattern=r"^(add_more|show_month_total|show_month_total_disabled|download_\d{4}-\d{2}|edit_day_menu|close_entry|manual_month_total|month_total_\d{4}-\d{2}|admin_users|admin_user_\d+|admin_logs|admin_log_tail_\d+|admin_log_file|admin_log_by_user|admin_log_user_\d+|admin_log_user_tail_\d+_\d+|admin_log_user_file_\d+)$"),
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
    # admin log command
    application.add_handler(CommandHandler("logtail", logtail))

    print("Bot started...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
