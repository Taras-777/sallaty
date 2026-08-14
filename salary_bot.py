"""Telegram-бот обліку заробітку за коробки.

Схема зберігання: один Excel-файл на користувача на місяць — data/{user_id}_{YYYY-MM}.xlsx
Колонки: Дата | Великі коробки | Малі коробки | Загальна сума за день

Запуск:
    python3 salary_bot.py
Токен береться з config.py (BOT_TOKEN) або зі змінної оточення BOT_TOKEN.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import re
from collections import deque
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, Optional
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from telegram import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.error import BadRequest, Conflict, NetworkError
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

# --------------------------------------------------------------------------- #
# Конфігурація
# --------------------------------------------------------------------------- #

try:
    import config as _config
except ImportError:
    _config = None


def _setting(name: str, default):
    """config.py -> змінна оточення -> значення за замовчуванням."""
    if _config is not None and hasattr(_config, name):
        return getattr(_config, name)
    env_value = os.getenv(name)
    return env_value if env_value is not None else default


BOT_TOKEN: Optional[str] = _setting("BOT_TOKEN", None)

# Ставки. Виносяться в config.py, якщо треба змінити без правки коду.
RATE_BIG = Decimal(str(_setting("RATE_BIG", "0.90")))
RATE_SMALL = Decimal(str(_setting("RATE_SMALL", "0.70")))
SPLIT = Decimal(str(_setting("SPLIT", "2")))  # ділимо заробіток навпіл

_raw_admins = _setting("ADMIN_IDS", "442336138")
if isinstance(_raw_admins, (set, list, tuple)):
    ADMIN_IDS = {int(x) for x in _raw_admins}
else:
    ADMIN_IDS = {int(x) for x in re.findall(r"\d+", str(_raw_admins))}

BERLIN_TZ = ZoneInfo(str(_setting("TIMEZONE", "Europe/Berlin")))

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

LOG_FILE = DATA_DIR / "salary_bot.log"
USERS_FILE = DATA_DIR / "users.json"

SHEET_NAME = "Salary"
COLUMNS = ["Дата", "Великі коробки", "Малі коробки", "Загальна сума за день"]

TELEGRAM_TEXT_LIMIT = 3500  # запас до ліміту 4096

# --------------------------------------------------------------------------- #
# Логування
# --------------------------------------------------------------------------- #

logger = logging.getLogger("salary_bot")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    _handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    _handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(_handler)
    _console = logging.StreamHandler()
    _console.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(_console)

# --------------------------------------------------------------------------- #
# Стани діалогу та шаблони
# --------------------------------------------------------------------------- #

(
    WAIT_DATE,
    WAIT_BIG,
    WAIT_SMALL,
    WAIT_CONFIRM,
    WAIT_NEXT_ACTION,
    WAIT_MONTH_INPUT,
    WAIT_EDIT_VALUES,
) = range(7)

MONTH_PATTERN = re.compile(r"^\d{4}-\d{2}$")
DATE_PATTERN = re.compile(r"^(\d{1,2})[-.,](\d{1,2})[-.,](\d{4})$")
USER_FILE_PATTERN = re.compile(r"^(\d+)_(\d{4}-\d{2})\.xlsx$")
DATE_CELL_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")

RE_MONTH_TOTAL = re.compile(r"^month_total_(\d{4}-\d{2})$")
RE_DOWNLOAD = re.compile(r"^download_(\d{4}-\d{2})$")
RE_ADMIN_USER = re.compile(r"^admin_user_(\d+)$")
RE_ADMIN_LOG_TAIL = re.compile(r"^admin_log_tail_(\d+)$")
RE_ADMIN_LOG_USER = re.compile(r"^admin_log_user_(\d+)$")
RE_ADMIN_LOG_USER_TAIL = re.compile(r"^admin_log_user_tail_(\d+)_(\d+)$")
RE_ADMIN_LOG_USER_FILE = re.compile(r"^admin_log_user_file_(\d+)$")

CALLBACK_PATTERN = re.compile(
    r"^(add_more|show_month_total|edit_day_menu|close_entry|manual_month_total"
    r"|month_total_\d{4}-\d{2}|download_\d{4}-\d{2}"
    r"|admin_users|admin_user_\d+|admin_logs|admin_log_by_user|admin_log_file"
    r"|admin_log_tail_\d+|admin_log_user_\d+|admin_log_user_tail_\d+_\d+"
    r"|admin_log_user_file_\d+)$"
)

# --------------------------------------------------------------------------- #
# Гроші та дати
# --------------------------------------------------------------------------- #


def money(value) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def format_money(value) -> str:
    return f"{money(value):.2f}"


def calc_day_total(big_count: int, small_count: int) -> tuple[Decimal, Decimal, Decimal]:
    """Повертає (сума за великі, сума за малі, разом)."""
    big_total = money(Decimal(int(big_count)) * RATE_BIG / SPLIT)
    small_total = money(Decimal(int(small_count)) * RATE_SMALL / SPLIT)
    return big_total, small_total, money(big_total + small_total)


def parse_date_input(raw_date: str) -> Optional[datetime]:
    normalized = raw_date.strip().replace(".", "-").replace(",", "-")
    try:
        return datetime.strptime(normalized, "%d-%m-%Y")
    except ValueError:
        return None


def today_in_berlin():
    return datetime.now(BERLIN_TZ).date()


def current_month() -> str:
    return datetime.now(BERLIN_TZ).strftime("%Y-%m")


# --------------------------------------------------------------------------- #
# Шар зберігання (синхронні функції — викликаються через asyncio.to_thread)
# --------------------------------------------------------------------------- #

_locks: dict[str, asyncio.Lock] = {}


def _lock_for(key: str) -> asyncio.Lock:
    """Один Lock на файл — щоб паралельні записи не губили дані."""
    lock = _locks.get(key)
    if lock is None:
        lock = asyncio.Lock()
        _locks[key] = lock
    return lock


def user_file(user_id: int, month: str) -> Path:
    return DATA_DIR / f"{user_id}_{month}.xlsx"


def _ensure_workbook(path: Path) -> None:
    if path.exists():
        return
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    sheet.append(COLUMNS)
    workbook.save(path)
    logger.info("Created new month file: %s", path.name)


def _as_int(value) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def _open_sheet(path: Path):
    """Відкриває аркуш і за потреби нормалізує його до 4 колонок.

    Обробляє два випадки:
      * старий формат на 6 колонок (малі коробки в кол. 4, сума в кол. 6);
      * пошкоджені файли (порожні рядки, дубльований заголовок посеред даних).
    Аркуш перебудовується з нуля — правити наявний через delete_cols не можна,
    бо openpyxl не скидає внутрішній лічильник рядків і дані з'їжджають униз.
    """
    workbook = load_workbook(path)
    sheet = workbook[SHEET_NAME] if SHEET_NAME in workbook.sheetnames else workbook.active
    sheet.title = SHEET_NAME

    legacy = sheet.max_column > len(COLUMNS)
    records: dict[str, list] = {}

    for row in sheet.iter_rows(min_row=1, values_only=True):
        if not row or row[0] is None:
            continue
        date_value = str(row[0]).strip()
        if not DATE_CELL_PATTERN.match(date_value):
            continue  # заголовок або сміття

        big = _as_int(row[1]) if len(row) > 1 else 0
        small_index, total_index = (3, 5) if legacy else (2, 3)
        small = _as_int(row[small_index]) if len(row) > small_index else 0

        total = row[total_index] if len(row) > total_index else None
        try:
            total = float(money(total)) if total is not None else None
        except (InvalidOperation, TypeError, ValueError):
            total = None
        if total is None:
            total = float(calc_day_total(big, small)[2])

        records[date_value] = [date_value, big, small, total]  # дублікати: лишається останній

    rows = list(records.values())
    is_clean = (
        not legacy
        and str(sheet.cell(row=1, column=1).value) == COLUMNS[0]
        and sheet.max_row == len(rows) + 1
    )
    if is_clean:
        return workbook, sheet, False

    workbook.remove(sheet)
    new_sheet = workbook.create_sheet(SHEET_NAME, 0)
    new_sheet.append(COLUMNS)
    for row in rows:
        new_sheet.append(row)
    logger.info(
        "Normalized sheet %s (%s, %d rows)",
        path.name, "legacy 6-column" if legacy else "damaged layout", len(rows),
    )
    return workbook, new_sheet, True


def _row_total(row: Iterable) -> Decimal:
    values = list(row)
    if len(values) < 4 or values[0] is None or values[3] is None:
        return Decimal("0.00")
    if not DATE_CELL_PATTERN.match(str(values[0]).strip()):
        return Decimal("0.00")
    try:
        return Decimal(str(values[3]))
    except InvalidOperation:
        logger.warning("Skipped unparsable total in row %s", values[0])
        return Decimal("0.00")


def _save_day_sync(path: Path, date_value: str, big_count: int, small_count: int) -> Decimal:
    _ensure_workbook(path)
    workbook, sheet, _ = _open_sheet(path)
    day_total = calc_day_total(big_count, small_count)[2]

    for row_idx in range(2, sheet.max_row + 1):
        if str(sheet.cell(row=row_idx, column=1).value) == date_value:
            sheet.cell(row=row_idx, column=2, value=int(big_count))
            sheet.cell(row=row_idx, column=3, value=int(small_count))
            sheet.cell(row=row_idx, column=4, value=float(day_total))
            workbook.save(path)
            return day_total

    sheet.append([date_value, int(big_count), int(small_count), float(day_total)])
    workbook.save(path)
    return day_total


def _delete_day_sync(path: Path, date_value: str) -> tuple[bool, int]:
    """Видаляє запис за дату. Повертає (чи видалено, скільки днів лишилось).

    Якщо після видалення в місяці не лишається жодного дня — файл видаляється,
    щоб порожній місяць не з'являвся в меню та в списках.
    """
    if not path.exists():
        return False, 0

    workbook, sheet, changed = _open_sheet(path)
    rows = [
        list(row)[:len(COLUMNS)]
        for row in sheet.iter_rows(min_row=2, values_only=True)
        if row and row[0] is not None and DATE_CELL_PATTERN.match(str(row[0]).strip())
    ]
    kept = [row for row in rows if str(row[0]).strip() != date_value]

    if len(kept) == len(rows):
        if changed:
            workbook.save(path)
        return False, len(kept)

    if not kept:
        workbook.close()
        path.unlink()
        logger.info("Removed empty month file: %s", path.name)
        return True, 0

    workbook.remove(sheet)
    new_sheet = workbook.create_sheet(SHEET_NAME, 0)
    new_sheet.append(COLUMNS)
    for row in kept:
        new_sheet.append(row)
    workbook.save(path)
    return True, len(kept)


def _read_day_sync(path: Path, date_value: str) -> Optional[tuple[int, int]]:
    if not path.exists():
        return None
    workbook, sheet, changed = _open_sheet(path)
    if changed:
        workbook.save(path)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        if str(row[0]) == date_value:
            return int(row[1] or 0), int(row[2] or 0)
    return None


def _read_total_sync(path: Path) -> Decimal:
    if not path.exists():
        return Decimal("0.00")
    workbook, sheet, changed = _open_sheet(path)
    if changed:
        workbook.save(path)
    total = Decimal("0.00")
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        total += _row_total(row)
    return money(total)


# Кеш "чи є в файлі дані": ключ — ім'я файлу, значення — (mtime, size, рядків, є заробіток).
# Перевірка робиться по stat(), тож відкриття книги відбувається лише після зміни файлу.
_stats_cache: dict[str, tuple[float, int, int, bool]] = {}


def _month_stats_sync(path: Path) -> tuple[int, bool]:
    """Повертає (кількість рядків з датами, чи є хоч один ненульовий день)."""
    try:
        stat = path.stat()
    except OSError:
        _stats_cache.pop(path.name, None)
        return 0, False

    cached = _stats_cache.get(path.name)
    if cached is not None and cached[0] == stat.st_mtime and cached[1] == stat.st_size:
        return cached[2], cached[3]

    rows, has_earnings = 0, False
    try:
        workbook = load_workbook(path, read_only=True)
        sheet = workbook[SHEET_NAME] if SHEET_NAME in workbook.sheetnames else workbook.active
        for row in sheet.iter_rows(min_row=1, values_only=True):
            if not row or row[0] is None:
                continue
            if not DATE_CELL_PATTERN.match(str(row[0]).strip()):
                continue
            rows += 1
            if _as_int(row[1] if len(row) > 1 else 0) or _as_int(row[2] if len(row) > 2 else 0):
                has_earnings = True
        workbook.close()
    except Exception as error:  # noqa: BLE001 — пошкоджений файл краще показати, ніж сховати
        logger.warning("Cannot inspect %s: %s", path.name, error)
        return 1, True

    _stats_cache[path.name] = (stat.st_mtime, stat.st_size, rows, has_earnings)
    return rows, has_earnings


def _month_has_data_sync(path: Path) -> bool:
    return _month_stats_sync(path)[1]


def cleanup_empty_files() -> int:
    """Прибирає файли без жодного рядка даних (лише заголовок). Викликається при старті."""
    removed = 0
    for path in sorted(DATA_DIR.glob("*.xlsx")):
        if not USER_FILE_PATTERN.match(path.name):
            continue
        if _month_stats_sync(path)[0] > 0:
            continue
        try:
            path.unlink()
            _stats_cache.pop(path.name, None)
            removed += 1
            logger.info("Removed empty month file at startup: %s", path.name)
        except OSError as error:
            logger.warning("Cannot remove %s: %s", path.name, error)
    return removed


def _list_user_files_sync(
    month: Optional[str] = None,
    user_id: Optional[int] = None,
    require_data: bool = True,
) -> list[Path]:
    result = []
    for path in sorted(DATA_DIR.glob("*.xlsx")):
        match = USER_FILE_PATTERN.match(path.name)
        if not match:
            continue
        if user_id is not None and int(match.group(1)) != user_id:
            continue
        if month is not None and match.group(2) != month:
            continue
        if require_data and not _month_has_data_sync(path):
            continue
        result.append(path)
    return result


def _list_months_sync(user_id: Optional[int] = None) -> list[str]:
    months = set()
    for path in _list_user_files_sync(user_id=user_id):
        match = USER_FILE_PATTERN.match(path.name)
        if match:
            months.add(match.group(2))
    return sorted(months)


def _list_user_ids_sync() -> list[int]:
    ids = set()
    for path in _list_user_files_sync():
        match = USER_FILE_PATTERN.match(path.name)
        if match:
            ids.add(int(match.group(1)))
    return sorted(ids)


def _read_all_totals_sync(month: str) -> tuple[Decimal, list[str]]:
    total = Decimal("0.00")
    names = []
    for path in _list_user_files_sync(month=month):
        total += _read_total_sync(path)
        names.append(path.name)
    return money(total), names


def _load_user_map_sync() -> dict:
    if not USERS_FILE.exists():
        return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        logger.warning("users.json is unreadable, treating as empty")
        return {}


def _register_user_sync(user_id: int, display_name: str) -> None:
    mapping = _load_user_map_sync()
    if mapping.get(str(user_id)) == display_name:
        return
    mapping[str(user_id)] = display_name
    with open(USERS_FILE, "w", encoding="utf-8") as fh:
        json.dump(mapping, fh, ensure_ascii=False, indent=2)


def _tail_log_sync(lines: int, marker: Optional[str] = None) -> list[str]:
    if not LOG_FILE.exists():
        return []
    with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as fh:
        if marker is None:
            return [line.rstrip("\n") for line in deque(fh, maxlen=lines)]
        matched = (line.rstrip("\n") for line in fh if marker in line)
        return list(deque(matched, maxlen=lines))


def _filter_log_sync(marker: str) -> list[str]:
    if not LOG_FILE.exists():
        return []
    with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as fh:
        return [line.rstrip("\n") for line in fh if marker in line]


# --------------------------------------------------------------------------- #
# Асинхронні обгортки (блокуючий I/O — в окремому потоці, під локом)
# --------------------------------------------------------------------------- #


async def save_day(user_id: int, month: str, day: int, big: int, small: int) -> Decimal:
    path = user_file(user_id, month)
    date_value = f"{month}-{day:02d}"
    async with _lock_for(path.name):
        total = await asyncio.to_thread(_save_day_sync, path, date_value, big, small)
    logger.info(
        "[USER:%s] Saved %s big=%s small=%s total=%s -> %s",
        user_id, date_value, big, small, total, path.name,
    )
    return total


async def delete_day(user_id: int, month: str, day: int) -> tuple[bool, int]:
    path = user_file(user_id, month)
    date_value = f"{month}-{day:02d}"
    async with _lock_for(path.name):
        removed, remaining = await asyncio.to_thread(_delete_day_sync, path, date_value)
    logger.info(
        "[USER:%s] Deleted %s (removed=%s, days left=%s)", user_id, date_value, removed, remaining
    )
    return removed, remaining


async def read_day(user_id: int, month: str, day: int) -> Optional[tuple[int, int]]:
    path = user_file(user_id, month)
    async with _lock_for(path.name):
        return await asyncio.to_thread(_read_day_sync, path, f"{month}-{day:02d}")


async def read_month_total(user_id: int, month: str) -> Decimal:
    path = user_file(user_id, month)
    async with _lock_for(path.name):
        return await asyncio.to_thread(_read_total_sync, path)


async def read_all_totals(month: str) -> tuple[Decimal, list[str]]:
    return await asyncio.to_thread(_read_all_totals_sync, month)


async def list_months(user_id: Optional[int]) -> list[str]:
    return await asyncio.to_thread(_list_months_sync, user_id)


async def list_user_ids() -> list[int]:
    return await asyncio.to_thread(_list_user_ids_sync)


async def list_user_files(month: str) -> list[Path]:
    return await asyncio.to_thread(_list_user_files_sync, month, None)


async def load_user_map() -> dict:
    return await asyncio.to_thread(_load_user_map_sync)


_registered_users: dict[int, str] = {}


async def register_user(user) -> None:
    """Записує користувача в users.json. Повторні виклики з тим самим іменем безкоштовні."""
    if user is None or getattr(user, "id", None) is None:
        return
    display_name = getattr(user, "username", None) or getattr(user, "full_name", "") or ""
    if _registered_users.get(user.id) == display_name:
        return
    async with _lock_for(USERS_FILE.name):
        await asyncio.to_thread(_register_user_sync, user.id, display_name)
    _registered_users[user.id] = display_name


async def known_users() -> dict[int, tuple[str, bool]]:
    """Усі відомі користувачі: {id: (ім'я, чи є дані)}.

    Джерела об'єднуються — users.json (усі, хто хоч раз запускав бота)
    і файли даних (на випадок, якщо users.json загубився).
    """
    with_data = set(await list_user_ids())
    user_map = await load_user_map()

    result: dict[int, tuple[str, bool]] = {}
    for key, name in user_map.items():
        if str(key).isdigit():
            uid = int(key)
            result[uid] = (name or "", uid in with_data)
    for uid in with_data:
        result.setdefault(uid, ("", True))
    return dict(sorted(result.items()))


async def tail_log(lines: int, marker: Optional[str] = None) -> list[str]:
    return await asyncio.to_thread(_tail_log_sync, lines, marker)


async def filter_log(marker: str) -> list[str]:
    return await asyncio.to_thread(_filter_log_sync, marker)


# --------------------------------------------------------------------------- #
# Клавіатури та безпечне редагування повідомлень
# --------------------------------------------------------------------------- #


def is_admin(user_id: Optional[int]) -> bool:
    return user_id is not None and user_id in ADMIN_IDS


def build_main_menu(user_id: Optional[int] = None, month: Optional[str] = None) -> InlineKeyboardMarkup:
    month = month or current_month()
    has_data = user_id is not None and _month_has_data_sync(user_file(user_id, month))

    buttons = [
        [
            InlineKeyboardButton("➕ Додати день", callback_data="add_more"),
            InlineKeyboardButton("📊 Сума за місяць", callback_data="show_month_total"),
        ],
        [
            InlineKeyboardButton("📝 Редагувати день", callback_data="edit_day_menu"),
            InlineKeyboardButton("🏠 Меню", callback_data="close_entry"),
        ],
    ]
    if has_data:
        buttons.insert(1, [InlineKeyboardButton("📥 Завантажити файл", callback_data=f"download_{month}")])
    if is_admin(user_id):
        buttons.append(
            [
                InlineKeyboardButton("👥 Користувачі", callback_data="admin_users"),
                InlineKeyboardButton("🗒️ Логи", callback_data="admin_logs"),
            ]
        )
    return InlineKeyboardMarkup(buttons)


def build_month_selection_keyboard(months: list[str]) -> InlineKeyboardMarkup:
    rows = [[InlineKeyboardButton(month, callback_data=f"month_total_{month}")] for month in months]
    rows.append([InlineKeyboardButton("✍️ Ввести вручну", callback_data="manual_month_total")])
    rows.append([InlineKeyboardButton("🏠 Меню", callback_data="close_entry")])
    return InlineKeyboardMarkup(rows)


def build_month_result_keyboard(month: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("📥 Завантажити файл", callback_data=f"download_{month}"),
                InlineKeyboardButton("🏠 Меню", callback_data="close_entry"),
            ]
        ]
    )


async def safe_edit(query: CallbackQuery, text: str, reply_markup=None) -> None:
    """edit_message_text, який не падає на 'Message is not modified'."""
    try:
        await query.edit_message_text(text, reply_markup=reply_markup)
    except BadRequest as error:
        if "Message is not modified" in str(error):
            logger.debug("Ignored 'Message is not modified'")
            return
        raise


async def send_lines(query: CallbackQuery, lines: list[str], header: str, filename: str) -> None:
    """Довгий текст надсилаємо файлом, короткий — повідомленням."""
    text = "\n".join(lines)
    if len(text) > TELEGRAM_TEXT_LIMIT or len(lines) > 200:
        tmp = DATA_DIR / filename
        try:
            await asyncio.to_thread(tmp.write_text, text, encoding="utf-8")
            with open(tmp, "rb") as fh:
                await query.message.reply_document(document=fh, filename=tmp.name)
        finally:
            try:
                tmp.unlink()
            except OSError:
                pass
        return
    await query.message.reply_text(f"{header}\n\n{text}")


# --------------------------------------------------------------------------- #
# Хендлери: введення дня
# --------------------------------------------------------------------------- #

DATE_HINT = "Введи дату: 15-08-2026\nМожна також: 15.08.2026 або 15,08,2026"


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    user = update.message.from_user
    await register_user(user)
    logger.info("[USER:%s] /start (@%s)", user.id, getattr(user, "username", None))
    await update.message.reply_text(DATE_HINT, reply_markup=build_main_menu(user_id=user.id))
    return WAIT_DATE


async def process_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw_date = update.message.text.strip()
    if not DATE_PATTERN.match(raw_date):
        await update.message.reply_text("Невірний формат. Введи: 15-08-2026")
        return WAIT_DATE

    parsed = parse_date_input(raw_date)
    if parsed is None:
        await update.message.reply_text("Такої дати не існує. Введи: 15-08-2026")
        return WAIT_DATE

    selected = parsed.date()
    today = today_in_berlin()
    if selected > today:
        await update.message.reply_text(
            f"Не можна вносити дані за майбутні дати. Сьогодні: {today.strftime('%d-%m-%Y')}."
        )
        return WAIT_DATE

    user_id = update.message.from_user.id
    await register_user(update.message.from_user)
    month = parsed.strftime("%Y-%m")
    context.user_data["month"] = month
    context.user_data["day"] = selected.day
    context.user_data["date_text"] = selected.strftime("%d-%m-%Y")

    if context.user_data.pop("edit_mode", False):
        record = await read_day(user_id, month, selected.day)
        if record is None:
            await update.message.reply_text(
                f"Для {selected.strftime('%d-%m-%Y')} записів не знайдено.",
                reply_markup=build_main_menu(user_id=user_id, month=month),
            )
            return WAIT_DATE
        big_count, small_count = record
        await update.message.reply_text(
            f"Запис {selected.strftime('%d-%m-%Y')}:\n"
            f"Великі: {big_count}\n"
            f"Малі: {small_count}\n\n"
            f"Введи нові значення через пробіл: великі малі\n"
            f"Наприклад: 200 150\n"
            f"Щоб видалити запис за цей день — введи: 0 0"
        )
        return WAIT_EDIT_VALUES

    await update.message.reply_text("Великі коробки?")
    return WAIT_BIG


async def _read_count(update: Update) -> Optional[int]:
    try:
        value = int(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Введи коректне число.")
        return None
    if value < 0:
        await update.message.reply_text("Кількість не може бути від'ємною.")
        return None
    return value


async def process_big_boxes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = await _read_count(update)
    if value is None:
        return WAIT_BIG
    context.user_data["big_count"] = value
    await update.message.reply_text("Малі коробки?")
    return WAIT_SMALL


async def process_small_boxes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = await _read_count(update)
    if value is None:
        return WAIT_SMALL

    context.user_data["small_count"] = value
    big_count = context.user_data.get("big_count", 0)
    big_total, small_total, day_total = calc_day_total(big_count, value)

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("💾 Зберегти", callback_data="save_entry"),
                InlineKeyboardButton("✏️ Відредагувати", callback_data="edit_entry"),
            ]
        ]
    )
    await update.message.reply_text(
        "Перевірка:\n"
        f"Дата: {context.user_data['date_text']}\n"
        f"Великі коробки: {big_count}\n"
        f"Малі коробки: {value}\n\n"
        "Розрахунок:\n"
        f"Великі: {big_count} × {RATE_BIG} / {SPLIT} = {format_money(big_total)} €\n"
        f"Малі: {value} × {RATE_SMALL} / {SPLIT} = {format_money(small_total)} €\n"
        f"Сума за день: {format_money(day_total)} €\n\n"
        "Все вірно?",
        reply_markup=keyboard,
    )
    return WAIT_CONFIRM


async def confirm_entry(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "edit_entry":
        context.user_data.pop("big_count", None)
        context.user_data.pop("small_count", None)
        await safe_edit(query, f"Ок. {DATE_HINT}")
        return WAIT_DATE

    month = context.user_data.get("month")
    day = context.user_data.get("day")
    big_count = int(context.user_data.get("big_count", 0))
    small_count = int(context.user_data.get("small_count", 0))

    if month is None or day is None:
        await safe_edit(query, f"Дані втрачено. {DATE_HINT}", reply_markup=build_main_menu(user_id=user_id))
        return WAIT_DATE

    if big_count == 0 and small_count == 0:
        context.user_data.clear()
        await safe_edit(
            query,
            "Нічого не збережено: обидві кількості дорівнюють 0.",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    await register_user(query.from_user)
    big_total, small_total, day_total = calc_day_total(big_count, small_count)
    await save_day(user_id, month, day, big_count, small_count)
    monthly_total = await read_month_total(user_id, month)

    context.user_data.clear()
    await safe_edit(
        query,
        "✅ Збережено\n"
        f"{month}-{day:02d}\n"
        f"Великі: {big_count} = {format_money(big_total)} €\n"
        f"Малі: {small_count} = {format_money(small_total)} €\n"
        f"Сума за день: {format_money(day_total)} €\n"
        f"Разом за місяць: {format_money(monthly_total)} €",
        reply_markup=build_main_menu(user_id=user_id, month=month),
    )
    return WAIT_NEXT_ACTION


async def process_edit_values(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    parts = update.message.text.split()
    if len(parts) != 2:
        await update.message.reply_text("Введи два числа через пробіл: великі малі, наприклад 200 150")
        return WAIT_EDIT_VALUES
    try:
        big_count, small_count = int(parts[0]), int(parts[1])
    except ValueError:
        await update.message.reply_text("Обидва значення мають бути числами. Наприклад: 200 150")
        return WAIT_EDIT_VALUES
    if big_count < 0 or small_count < 0:
        await update.message.reply_text("Кількість не може бути від'ємною.")
        return WAIT_EDIT_VALUES

    user_id = update.message.from_user.id
    month = context.user_data.get("month")
    day = context.user_data.get("day")
    if month is None or day is None:
        await update.message.reply_text(DATE_HINT, reply_markup=build_main_menu(user_id=user_id))
        return WAIT_DATE

    # 0 0 означає "видалити запис за цей день"
    if big_count == 0 and small_count == 0:
        removed, remaining = await delete_day(user_id, month, day)
        context.user_data.clear()
        if not removed:
            text = f"Запис за {month}-{day:02d} не знайдено — нічого не змінено."
        elif remaining == 0:
            text = (
                f"🗑 Запис {month}-{day:02d} видалено.\n"
                f"За {month} більше немає записів — місяць прибрано з меню."
            )
        else:
            monthly_total = await read_month_total(user_id, month)
            text = (
                f"🗑 Запис {month}-{day:02d} видалено.\n"
                f"Днів у місяці лишилось: {remaining}\n"
                f"Разом за місяць: {format_money(monthly_total)} €"
            )
        await update.message.reply_text(
            text, reply_markup=build_main_menu(user_id=user_id, month=month)
        )
        return WAIT_NEXT_ACTION

    day_total = await save_day(user_id, month, day, big_count, small_count)
    monthly_total = await read_month_total(user_id, month)
    context.user_data.clear()

    await update.message.reply_text(
        f"✏️ Оновлено {month}-{day:02d}\n"
        f"Великі: {big_count}\n"
        f"Малі: {small_count}\n"
        f"Сума за день: {format_money(day_total)} €\n"
        f"Разом за місяць: {format_money(monthly_total)} €",
        reply_markup=build_main_menu(user_id=user_id, month=month),
    )
    return WAIT_NEXT_ACTION


# --------------------------------------------------------------------------- #
# Хендлери: кнопки
# --------------------------------------------------------------------------- #


async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    admin = is_admin(user_id)
    await register_user(query.from_user)

    # --- звичайні дії ---------------------------------------------------- #
    if data == "add_more":
        context.user_data.clear()
        await safe_edit(query, f"Ок. {DATE_HINT}")
        return WAIT_DATE

    if data == "edit_day_menu":
        context.user_data.clear()
        context.user_data["edit_mode"] = True
        await safe_edit(query, "Введи дату для редагування, наприклад 12-08-2026.")
        return WAIT_DATE

    if data == "close_entry":
        context.user_data.clear()
        await safe_edit(query, DATE_HINT, reply_markup=build_main_menu(user_id=user_id))
        return WAIT_DATE

    if data == "show_month_total":
        target = context.user_data.get("admin_view_user") if admin else user_id
        months = await list_months(None if (admin and target is None) else target)
        if not months:
            await query.answer("Поки немає збережених місяців", show_alert=True)
            return WAIT_NEXT_ACTION
        await safe_edit(query, "Вибери місяць:", reply_markup=build_month_selection_keyboard(months))
        return WAIT_MONTH_INPUT

    if data == "manual_month_total":
        await safe_edit(query, "Який місяць показати? Введи YYYY-MM")
        return WAIT_MONTH_INPUT

    match = RE_MONTH_TOTAL.match(data)
    if match:
        return await show_month_total(query, context, match.group(1))

    match = RE_DOWNLOAD.match(data)
    if match:
        return await send_month_files(query, context, match.group(1))

    # --- адмінські дії ---------------------------------------------------- #
    if data.startswith("admin_") and not admin:
        await query.answer("Доступ заборонено", show_alert=True)
        logger.warning("[USER:%s] Denied admin action: %s", user_id, data)
        return WAIT_NEXT_ACTION

    if data in ("admin_users", "admin_log_by_user"):
        users = await known_users()
        if not users:
            await query.answer("Нема користувачів", show_alert=True)
            return WAIT_NEXT_ACTION
        prefix = "admin_user_" if data == "admin_users" else "admin_log_user_"
        rows = []
        for uid, (name, has_data) in users.items():
            label = f"{uid} — @{name}" if name else str(uid)
            if not has_data:
                label += " · без даних"
            rows.append([InlineKeyboardButton(label, callback_data=f"{prefix}{uid}")])
        rows.append([InlineKeyboardButton("🏠 Меню", callback_data="close_entry")])
        title = "Оберіть користувача:" if data == "admin_users" else "Логи якого користувача показати?"
        await safe_edit(query, title, reply_markup=InlineKeyboardMarkup(rows))
        return WAIT_NEXT_ACTION

    match = RE_ADMIN_USER.match(data)
    if match:
        selected = int(match.group(1))
        context.user_data["admin_view_user"] = selected
        months = await list_months(selected)
        if not months:
            await query.answer("У цього користувача немає даних", show_alert=True)
            return WAIT_NEXT_ACTION
        await safe_edit(
            query,
            f"Оберіть місяць для користувача {selected}:",
            reply_markup=build_month_selection_keyboard(months),
        )
        return WAIT_MONTH_INPUT

    if data == "admin_logs":
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton("Останні 100", callback_data="admin_log_tail_100"),
                    InlineKeyboardButton("Останні 500", callback_data="admin_log_tail_500"),
                ],
                [
                    InlineKeyboardButton("Повний файл", callback_data="admin_log_file"),
                    InlineKeyboardButton("По користувачу", callback_data="admin_log_by_user"),
                ],
                [InlineKeyboardButton("🏠 Меню", callback_data="close_entry")],
            ]
        )
        await safe_edit(query, "Журнал (логи): оберіть опцію:", reply_markup=keyboard)
        return WAIT_NEXT_ACTION

    match = RE_ADMIN_LOG_TAIL.match(data)
    if match:
        lines = await tail_log(min(int(match.group(1)), 2000))
        if not lines:
            await query.answer("Лог порожній", show_alert=True)
            return WAIT_NEXT_ACTION
        await send_lines(query, lines, f"Останні {len(lines)} рядків лога:", f"log_tail_{user_id}.txt")
        logger.info("[ADMIN:%s] Sent log tail (%d lines)", user_id, len(lines))
        return WAIT_NEXT_ACTION

    if data == "admin_log_file":
        if not LOG_FILE.exists():
            await query.answer("Файл лога не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        with open(LOG_FILE, "rb") as fh:
            await query.message.reply_document(document=fh, filename=LOG_FILE.name)
        logger.info("[ADMIN:%s] Downloaded full log file", user_id)
        return WAIT_NEXT_ACTION

    match = RE_ADMIN_LOG_USER.match(data)
    if match:
        selected = int(match.group(1))
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton("Останні 100", callback_data=f"admin_log_user_tail_{selected}_100"),
                    InlineKeyboardButton("Останні 500", callback_data=f"admin_log_user_tail_{selected}_500"),
                ],
                [
                    InlineKeyboardButton("Повний файл", callback_data=f"admin_log_user_file_{selected}"),
                    InlineKeyboardButton("🏠 Меню", callback_data="close_entry"),
                ],
            ]
        )
        await safe_edit(query, f"Логи користувача {selected}:", reply_markup=keyboard)
        return WAIT_NEXT_ACTION

    match = RE_ADMIN_LOG_USER_TAIL.match(data)
    if match:
        selected, lines = int(match.group(1)), min(int(match.group(2)), 2000)
        result = await tail_log(lines, marker=f"[USER:{selected}]")
        if not result:
            await query.answer("Записів для цього користувача не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        await send_lines(
            query, result, f"Останні {len(result)} рядків для {selected}:", f"log_user_{selected}.txt"
        )
        logger.info("[ADMIN:%s] Sent log tail for [USER:%s]", user_id, selected)
        return WAIT_NEXT_ACTION

    match = RE_ADMIN_LOG_USER_FILE.match(data)
    if match:
        selected = int(match.group(1))
        result = await filter_log(f"[USER:{selected}]")
        if not result:
            await query.answer("Записів для цього користувача не знайдено", show_alert=True)
            return WAIT_NEXT_ACTION
        await send_lines(query, result, f"Лог користувача {selected}:", f"log_user_{selected}_full.txt")
        logger.info("[ADMIN:%s] Sent full filtered log for [USER:%s]", user_id, selected)
        return WAIT_NEXT_ACTION

    return WAIT_NEXT_ACTION


async def show_month_total(query: CallbackQuery, context: ContextTypes.DEFAULT_TYPE, month: str) -> int:
    user_id = query.from_user.id
    admin_target = context.user_data.pop("admin_view_user", None) if is_admin(user_id) else None
    target_id = admin_target if admin_target is not None else user_id

    path = user_file(target_id, month)
    if not _month_has_data_sync(path):
        await query.answer("За цей місяць немає записів", show_alert=True)
        return WAIT_NEXT_ACTION
    total = await read_month_total(target_id, month)

    label = f"Користувач {target_id}" if admin_target is not None else "Разом"
    context.user_data["download_user"] = target_id
    await safe_edit(
        query,
        f"{label} за {month}: {format_money(total)} €\nФайл: {path.name}",
        reply_markup=build_month_result_keyboard(month),
    )
    return WAIT_NEXT_ACTION


async def send_month_files(query: CallbackQuery, context: ContextTypes.DEFAULT_TYPE, month: str) -> int:
    user_id = query.from_user.id
    target_id = context.user_data.pop("download_user", user_id)
    if target_id != user_id and not is_admin(user_id):
        target_id = user_id

    path = user_file(target_id, month)
    if not path.exists():
        await query.answer("Файл не знайдено", show_alert=True)
        return WAIT_NEXT_ACTION

    try:
        with open(path, "rb") as fh:
            await query.message.reply_document(document=fh, filename=path.name)
        logger.info("[USER:%s] Sent file %s", user_id, path.name)
    except (OSError, BadRequest) as error:
        logger.exception("[USER:%s] Failed to send %s: %s", user_id, path.name, error)
        await query.message.reply_text("Не вдалося надіслати файл. Спробуй ще раз.")
    return WAIT_NEXT_ACTION


async def process_month_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    month = update.message.text.strip()
    if not MONTH_PATTERN.match(month):
        await update.message.reply_text("Невірний формат місяця. Введи YYYY-MM, наприклад 2026-08.")
        return WAIT_MONTH_INPUT

    user_id = update.message.from_user.id
    admin_target = context.user_data.pop("admin_view_user", None) if is_admin(user_id) else None

    if is_admin(user_id) and admin_target is None:
        total, names = await read_all_totals(month)
        if not names:
            await update.message.reply_text(
                f"За місяць {month} ще немає записів.",
                reply_markup=build_main_menu(user_id=user_id, month=month),
            )
            return WAIT_NEXT_ACTION
        await update.message.reply_text(
            f"Усі користувачі за {month}: {format_money(total)} €\nФайли: {', '.join(names)}",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    target_id = admin_target if admin_target is not None else user_id
    path = user_file(target_id, month)
    if not _month_has_data_sync(path):
        await update.message.reply_text(
            f"За місяць {month} ще немає записів.",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    total = await read_month_total(target_id, month)
    context.user_data["download_user"] = target_id
    await update.message.reply_text(
        f"{month}: {format_money(total)} €\nФайл: {path.name}",
        reply_markup=build_month_result_keyboard(month),
    )
    return WAIT_NEXT_ACTION


# --------------------------------------------------------------------------- #
# Команди
# --------------------------------------------------------------------------- #


async def cmd_total(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    args = context.args
    month = args[0].strip() if args else current_month()
    if not MONTH_PATTERN.match(month):
        await update.message.reply_text("Формат: /total 2026-08")
        return

    user_id = update.message.from_user.id
    if is_admin(user_id):
        total, names = await read_all_totals(month)
        if not names:
            await update.message.reply_text(f"За місяць {month} ще немає записів.")
            return
        await update.message.reply_text(
            f"Усі користувачі за {month}: {format_money(total)} €\nФайли: {', '.join(names)}"
        )
        return

    path = user_file(user_id, month)
    if not _month_has_data_sync(path):
        await update.message.reply_text(f"За місяць {month} ще немає записів.")
        return
    total = await read_month_total(user_id, month)
    await update.message.reply_text(f"{month}: {format_money(total)} €\nФайл: {path.name}")


async def cmd_logtail(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/logtail [N|file] — тільки для адмінів."""
    user_id = update.message.from_user.id
    if not is_admin(user_id):
        await update.message.reply_text("Доступ заборонено.")
        logger.warning("[USER:%s] Denied /logtail", user_id)
        return

    mode = context.args[0].strip().lower() if context.args else "100"

    if mode == "file":
        if not LOG_FILE.exists():
            await update.message.reply_text("Файл лога не знайдено.")
            return
        with open(LOG_FILE, "rb") as fh:
            await update.message.reply_document(document=fh, filename=LOG_FILE.name)
        logger.info("[ADMIN:%s] Downloaded full log file", user_id)
        return

    try:
        lines = min(int(mode), 2000)
    except ValueError:
        lines = 100

    result = await tail_log(lines)
    if not result:
        await update.message.reply_text("Лог порожній.")
        return

    text = "\n".join(result)
    if len(text) > TELEGRAM_TEXT_LIMIT or len(result) > 200:
        tmp = DATA_DIR / f"log_tail_{user_id}.txt"
        try:
            await asyncio.to_thread(tmp.write_text, text, encoding="utf-8")
            with open(tmp, "rb") as fh:
                await update.message.reply_document(document=fh, filename=tmp.name)
        finally:
            try:
                tmp.unlink()
            except OSError:
                pass
    else:
        await update.message.reply_text(f"Останні {len(result)} рядків лога:\n\n{text}")
    logger.info("[ADMIN:%s] Requested log tail (%d lines)", user_id, len(result))


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Скасовано. Для нового запису натисни /start.")
    return ConversationHandler.END


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    error = context.error

    if isinstance(error, Conflict):
        logger.error(
            "Conflict: з цим токеном уже працює інший екземпляр бота. "
            "Зупини другий процес (systemctl stop salarybot) або візьми окремий тестовий токен."
        )
        return

    if isinstance(error, NetworkError) and not isinstance(error, BadRequest):
        logger.warning("Мережева помилка: %s", error)
        return

    logger.error("Unhandled exception", exc_info=error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "Сталася помилка. Спробуй ще раз або натисни /start."
            )
        except Exception:  # noqa: BLE001 — не даємо помилці в обробнику помилок впасти далі
            logger.exception("Failed to notify user about error")


# --------------------------------------------------------------------------- #
# Точка входу
# --------------------------------------------------------------------------- #


def main() -> None:
    token = BOT_TOKEN or os.getenv("BOT_TOKEN")
    if not token or token == "YOUR_BOT_TOKEN_HERE":
        raise RuntimeError(
            "BOT_TOKEN не задано. Вкажи його в config.py або в змінній оточення BOT_TOKEN."
        )

    # Починаючи з Python 3.12 asyncio.get_event_loop() не створює цикл сам,
    # а в 3.14 кидає RuntimeError. PTB викликає його всередині run_polling(),
    # тому створюємо цикл явно.
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())

    application = Application.builder().token(token).build()

    callback_handler = CallbackQueryHandler(on_callback, pattern=CALLBACK_PATTERN)

    conversation = ConversationHandler(
        entry_points=[CommandHandler("start", start), CommandHandler("menu", start)],
        states={
            WAIT_DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_date),
                callback_handler,
            ],
            WAIT_BIG: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_big_boxes),
                callback_handler,
            ],
            WAIT_SMALL: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_small_boxes),
                callback_handler,
            ],
            WAIT_CONFIRM: [
                CallbackQueryHandler(confirm_entry, pattern=r"^(save_entry|edit_entry)$"),
                callback_handler,
            ],
            WAIT_EDIT_VALUES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_edit_values),
                callback_handler,
            ],
            WAIT_MONTH_INPUT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_month_input),
                callback_handler,
            ],
            WAIT_NEXT_ACTION: [callback_handler],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("start", start),
            CommandHandler("menu", start),
        ],
        allow_reentry=True,
    )

    application.add_handler(conversation)
    application.add_handler(CommandHandler("total", cmd_total))
    application.add_handler(CommandHandler("logtail", cmd_logtail))
    application.add_error_handler(on_error)

    removed = cleanup_empty_files()
    if removed:
        logger.info("Startup cleanup: removed %d empty month file(s)", removed)

    logger.info("Bot started (admins: %s)", sorted(ADMIN_IDS))
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()