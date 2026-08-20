"""Telegram-бот обліку заробітку за коробки.

Схема зберігання: один Excel-файл на користувача на місяць — data/{user_id}_{YYYY-MM}.xlsx
Колонки: Дата | Коробки 177 | Коробки 161 | Загальна сума за день
(назви типів коробок задаються константами BIG_LABEL/SMALL_LABEL, можна
перевизначити в config.py, якщо номери коробок знову зміняться)

Запуск:
    python3 salary_bot.py
Токен береться з config.py (BOT_TOKEN) або зі змінної оточення BOT_TOKEN.

Реєстрація:
    При першому /start (якщо ім'я ще не збережено) бота просить людину
    написати ім'я та прізвище — це ім'я потім показується адміну в
    запитах на підтвердження і в списку користувачів. Зберігається в
    data/users.json разом із Telegram-юзернеймом.

Підтвердження адміном:
    Коли звичайний користувач (не з ADMIN_IDS) додає новий день або
    редагує/видаляє вже внесений запис, дані НЕ зберігаються одразу —
    запит стає в чергу. Адміну летить лише легке сповіщення "🔔 Новий
    запит на підтвердження" без деталей і кнопок — щойно один такий
    прийшов, наступні (від того самого чи інших користувачів) вже
    мовчазні, поки адмін не розгребе чергу. Деталі та дії — тільки через
    кнопку меню «⏳ Очікують підтвердження (N)»: список користувачів із
    відкритими запитами → підтвердити/відхилити/відредагувати по черзі.
    Хто з адмінів відповість першим — того рішення і застосовується.
    Користувачу приходить окреме повідомлення з результатом. Записи, які
    вносить сам адмін, зберігаються одразу, без підтвердження. Черга
    запитів зберігається в пам'яті процесу і не переживає перезапуск бота
    (systemctl restart) — це відоме обмеження поточної реалізації.
"""

from __future__ import annotations

import asyncio
import json
import logging
from logging.handlers import RotatingFileHandler
import os
import re
import uuid
from collections import deque
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, Optional
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from telegram import (
    BotCommand,
    BotCommandScopeChat,
    BotCommandScopeDefault,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Update,
)
from telegram.error import BadRequest, Conflict, NetworkError
from telegram.ext import (
    Application,
    BaseUpdateProcessor,
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

# Назви типів коробок, як їх бачить користувач (можна перевизначити в config.py).
BIG_LABEL = str(_setting("BIG_LABEL", "Коробки 177"))
SMALL_LABEL = str(_setting("SMALL_LABEL", "Коробки 161"))

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
COLUMNS = ["Дата", BIG_LABEL, SMALL_LABEL, "Загальна сума за день"]

TELEGRAM_TEXT_LIMIT = 3500  # запас до ліміту 4096

# --------------------------------------------------------------------------- #
# Логування
# --------------------------------------------------------------------------- #

logger = logging.getLogger("salary_bot")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    # Ротація: до 5 МБ на файл, 5 старих копій (salary_bot.log.1 … .5). Для
    # команди 20-30 людей цього з запасом вистачає на місяці історії, а без
    # ротації лог ріс би необмежено роками роботи сервісу.
    _handler = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
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
    WAIT_EDIT_CONFIRM,
    WAIT_REGISTER_NAME,
    WAIT_ADMIN_EDIT_BIG,
    WAIT_ADMIN_EDIT_SMALL,
    WAIT_ADMIN_EDIT_CONFIRM,
) = range(12)

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
RE_ADMIN_SUMMARY = re.compile(r"^admin_summary_(\d{4}-\d{2})$")

# Запити на підтвердження запису адміном (review_<дія>_<id>, id — 8 hex символів).
RE_REVIEW_APPROVE = re.compile(r"^review_approve_([0-9a-f]{8})$")
RE_REVIEW_REJECT = re.compile(r"^review_reject_([0-9a-f]{8})$")
RE_REVIEW_EDIT = re.compile(r"^review_edit_([0-9a-f]{8})$")
RE_REVIEW_USER = re.compile(r"^review_user_(\d+)$")

CALLBACK_PATTERN = re.compile(
    r"^(add_more|show_month_total|edit_day_menu|close_entry"
    r"|month_total_\d{4}-\d{2}|download_\d{4}-\d{2}"
    r"|admin_users|admin_user_\d+|admin_logs|admin_log_by_user|admin_log_file"
    r"|admin_log_tail_\d+|admin_log_user_\d+|admin_log_user_tail_\d+_\d+"
    r"|admin_log_user_file_\d+"
    r"|admin_monthly_summary|admin_summary_\d{4}-\d{2}"
    r"|review_approve_[0-9a-f]{8}|review_reject_[0-9a-f]{8}|review_edit_[0-9a-f]{8}"
    r"|review_user_\d+"
    r"|pending_reviews|my_pending_reviews)$"
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


def _monthly_summary_rows_sync(month: str) -> list[dict]:
    """Підсумок за місяць по кожному користувачу: скільки коробок кожного
    типу зробив і на яку суму — для зведеної Excel-таблиці для адміна."""
    user_map = _load_user_map_sync()
    rows = []
    for path in _list_user_files_sync(month=month):
        match = USER_FILE_PATTERN.match(path.name)
        if not match:
            continue
        uid = int(match.group(1))

        workbook, sheet, changed = _open_sheet(path)
        if changed:
            workbook.save(path)

        big_sum = 0
        small_sum = 0
        money_sum = Decimal("0.00")
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            if not DATE_CELL_PATTERN.match(str(row[0]).strip()):
                continue
            big_sum += _as_int(row[1] if len(row) > 1 else 0)
            small_sum += _as_int(row[2] if len(row) > 2 else 0)
            money_sum += _row_total(row)
        workbook.close()

        entry = user_map.get(str(uid), {})
        name = (entry.get("registered_name") or entry.get("telegram_label") or "").strip() or str(uid)
        rows.append(
            {
                "user_id": uid,
                "name": name,
                "big": big_sum,
                "small": small_sum,
                "total_boxes": big_sum + small_sum,
                "money": money(money_sum),
            }
        )

    # У "Зведенні" — хто заробив більше, той згори (за рівних сум — за алфавітом).
    rows.sort(key=lambda r: (-r["money"], r["name"].lower()))
    return rows


def _daily_matrix_sync(month: str) -> tuple[list[str], list[dict], dict[tuple[str, int], tuple[int, int, Decimal]]]:
    """Дані для щоденної зведеної таблиці: список дат за місяць (де хоч у когось
    є запис), список користувачів (за алфавітом) і мапа (дата, user_id) ->
    (великі, малі, сума за день)."""
    user_map = _load_user_map_sync()
    users: list[dict] = []
    cells: dict[tuple[str, int], tuple[int, int, Decimal]] = {}
    dates: set[str] = set()

    for path in _list_user_files_sync(month=month):
        match = USER_FILE_PATTERN.match(path.name)
        if not match:
            continue
        uid = int(match.group(1))

        workbook, sheet, changed = _open_sheet(path)
        if changed:
            workbook.save(path)

        entry = user_map.get(str(uid), {})
        name = (entry.get("registered_name") or entry.get("telegram_label") or "").strip() or str(uid)
        users.append({"user_id": uid, "name": name})

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            date_value = str(row[0]).strip()
            if not DATE_CELL_PATTERN.match(date_value):
                continue
            big = _as_int(row[1] if len(row) > 1 else 0)
            small = _as_int(row[2] if len(row) > 2 else 0)
            day_money = _row_total(row)
            dates.add(date_value)
            cells[(date_value, uid)] = (big, small, day_money)
        workbook.close()

    users.sort(key=lambda u: u["name"].lower())
    return sorted(dates), users, cells


_GRANDTOTAL_FILL = PatternFill(start_color="FFD6E4F5", end_color="FFD6E4F5", fill_type="solid")


def _write_daily_detail_sheet(
    workbook: Workbook,
    dates: list[str],
    users: list[dict],
    cells: dict[tuple[str, int], tuple[int, int, Decimal]],
) -> None:
    """Деталізація за днями: для кожної дати — рядок на кожного користувача, який
    того дня щось вносив (великі/малі шт., сума за великі/малі окремо, разом за
    день). Комірка «Дата» об'єднана по всьому блоку дня. Внизу — один підсумковий
    рядок «Разом за місяць» по всіх людях і днях разом."""
    sheet = workbook.create_sheet("Деталізація за днями")
    headers = [
        "Дата",
        "Користувач",
        f"{BIG_LABEL}, шт",
        f"{SMALL_LABEL}, шт",
        f"{BIG_LABEL} — сума, €",
        f"{SMALL_LABEL} — сума, €",
        "Разом за день, €",
    ]
    sheet.append(headers)
    sheet.row_dimensions[1].height = 30
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    month_big = month_small = 0
    month_big_money = month_small_money = Decimal("0.00")
    row_idx = 2

    for date_value in dates:
        date_label = datetime.strptime(date_value, "%Y-%m-%d").strftime("%d-%m-%Y")
        day_users = [u for u in users if (date_value, u["user_id"]) in cells]
        if not day_users:
            continue

        block_start = row_idx
        for user in day_users:
            big, small, _stored_total = cells[(date_value, user["user_id"])]
            big_money, small_money, _ = calc_day_total(big, small)
            sheet.append(
                [date_label, user["name"], big, small, float(big_money), float(small_money), float(big_money + small_money)]
            )
            for col in (5, 6, 7):
                sheet.cell(row=row_idx, column=col).number_format = '0.00" €"'
            month_big += big
            month_small += small
            month_big_money += big_money
            month_small_money += small_money
            row_idx += 1

        if len(day_users) > 1:
            sheet.merge_cells(start_row=block_start, start_column=1, end_row=row_idx - 1, end_column=1)
            sheet.cell(row=block_start, column=1).alignment = Alignment(vertical="center")

    month_total = month_big_money + month_small_money
    sheet.append(
        [
            "Разом за місяць", "", month_big, month_small,
            float(month_big_money), float(month_small_money), float(month_total),
        ]
    )
    for cell in sheet[row_idx]:
        cell.font = Font(bold=True)
        cell.fill = _GRANDTOTAL_FILL
    sheet.cell(row=row_idx, column=5).number_format = '0.00" €"'
    sheet.cell(row=row_idx, column=6).number_format = '0.00" €"'
    sheet.cell(row=row_idx, column=7).number_format = '0.00" €"'

    sheet.freeze_panes = "A2"
    for col_idx, width in enumerate([13, 26, 12, 12, 20, 20, 16], start=1):
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def _build_monthly_summary_sync(month: str) -> Optional[Path]:
    """Будує зведену книгу за місяць: підсумок по кожному користувачу плюс
    деталізація по днях (хто скільки 177/161 зробив і на яку суму кожного дня,
    з підсумком за день і за місяць). Зберігає у тимчасовий xlsx у DATA_DIR.
    Повертає None, якщо за місяць немає даних."""
    rows = _monthly_summary_rows_sync(month)
    if not rows:
        return None

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Зведення"

    headers = ["№", "Користувач", "ID", BIG_LABEL, SMALL_LABEL, "Разом коробок", "Сума, €"]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for idx, row in enumerate(rows, start=1):
        sheet.append(
            [idx, row["name"], row["user_id"], row["big"], row["small"], row["total_boxes"], float(row["money"])]
        )
        sheet.cell(row=idx + 1, column=7).number_format = '0.00" €"'

    total_big = sum(row["big"] for row in rows)
    total_small = sum(row["small"] for row in rows)
    total_boxes = sum(row["total_boxes"] for row in rows)
    total_money = money(sum(row["money"] for row in rows))

    total_row_idx = len(rows) + 2
    sheet.append(["", "Разом за місяць", "", total_big, total_small, total_boxes, float(total_money)])
    for cell in sheet[total_row_idx]:
        cell.font = Font(bold=True)
    sheet.cell(row=total_row_idx, column=7).number_format = '0.00" €"'

    sheet.freeze_panes = "A2"
    for col_idx, width in enumerate([4, 26, 12, 14, 14, 15, 13], start=1):
        sheet.column_dimensions[get_column_letter(col_idx)].width = width

    dates, users, cells = _daily_matrix_sync(month)
    if dates and users:
        _write_daily_detail_sheet(workbook, dates, users, cells)

    tmp_path = DATA_DIR / f"summary_{month}.xlsx"
    workbook.save(tmp_path)
    return tmp_path


def _load_user_map_sync() -> dict:
    """Повертає {user_id_str: {"telegram_label": str, "registered_name": Optional[str]}}.

    Старий формат (значення — просто рядок з юзернеймом) переводиться в нову
    форму на льоту; на диску перепишеться при першому ж збереженні.
    """
    if not USERS_FILE.exists():
        return {}
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
    except (OSError, json.JSONDecodeError):
        logger.warning("users.json is unreadable, treating as empty")
        return {}

    mapping: dict = {}
    for key, value in raw.items():
        if isinstance(value, dict):
            mapping[key] = {
                "telegram_label": value.get("telegram_label", ""),
                "registered_name": value.get("registered_name"),
            }
        else:
            mapping[key] = {"telegram_label": str(value or ""), "registered_name": None}
    return mapping


def _save_user_map_sync(mapping: dict) -> None:
    with open(USERS_FILE, "w", encoding="utf-8") as fh:
        json.dump(mapping, fh, ensure_ascii=False, indent=2)


def _register_user_sync(user_id: int, display_name: str) -> None:
    mapping = _load_user_map_sync()
    entry = mapping.get(str(user_id), {"telegram_label": "", "registered_name": None})
    if entry.get("telegram_label") == display_name:
        return
    entry["telegram_label"] = display_name
    mapping[str(user_id)] = entry
    _save_user_map_sync(mapping)


def _save_registered_name_sync(user_id: int, full_name: str) -> None:
    mapping = _load_user_map_sync()
    entry = mapping.get(str(user_id), {"telegram_label": "", "registered_name": None})
    entry["registered_name"] = full_name
    mapping[str(user_id)] = entry
    _save_user_map_sync(mapping)


def _get_registered_name_sync(user_id: int) -> Optional[str]:
    mapping = _load_user_map_sync()
    entry = mapping.get(str(user_id))
    if isinstance(entry, dict):
        return entry.get("registered_name")
    return None


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


async def build_monthly_summary(month: str) -> Optional[Path]:
    return await asyncio.to_thread(_build_monthly_summary_sync, month)


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
    і файли даних (на випадок, якщо users.json загубився). Як ім'я
    пріоритетно береться те, яке людина вписала при реєстрації, інакше —
    її Telegram-юзернейм/ім'я, автоматично зафіксовані при /start.
    """
    with_data = set(await list_user_ids())
    user_map = await load_user_map()

    result: dict[int, tuple[str, bool]] = {}
    for key, entry in user_map.items():
        if str(key).isdigit():
            uid = int(key)
            label = (entry.get("registered_name") or entry.get("telegram_label") or "") if isinstance(entry, dict) else ""
            result[uid] = (label, uid in with_data)
    for uid in with_data:
        result.setdefault(uid, ("", True))
    return dict(sorted(result.items()))


async def save_registered_name(user_id: int, full_name: str) -> None:
    async with _lock_for(USERS_FILE.name):
        await asyncio.to_thread(_save_registered_name_sync, user_id, full_name)


async def get_registered_name(user_id: int) -> Optional[str]:
    return await asyncio.to_thread(_get_registered_name_sync, user_id)


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
        pending_count = _pending_review_count()
        buttons.append(
            [InlineKeyboardButton(f"⏳ Очікують підтвердження ({pending_count})", callback_data="pending_reviews")]
        )
        buttons.append(
            [
                InlineKeyboardButton("👥 Користувачі", callback_data="admin_users"),
                InlineKeyboardButton("🗒️ Логи", callback_data="admin_logs"),
            ]
        )
        buttons.append(
            [InlineKeyboardButton("📈 Зведення за місяць", callback_data="admin_monthly_summary")]
        )
    elif user_id is not None:
        my_pending = _user_pending_count(user_id)
        if my_pending:
            buttons.append(
                [
                    InlineKeyboardButton(
                        f"⏳ Очікують підтвердження ({my_pending})", callback_data="my_pending_reviews"
                    )
                ]
            )
    return InlineKeyboardMarkup(buttons)


def build_month_selection_keyboard(months: list[str]) -> InlineKeyboardMarkup:
    rows = [[InlineKeyboardButton(month, callback_data=f"month_total_{month}")] for month in months]
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


def build_admin_summary_month_keyboard(months: list[str]) -> InlineKeyboardMarkup:
    rows = [[InlineKeyboardButton(month, callback_data=f"admin_summary_{month}")] for month in months]
    rows.append([InlineKeyboardButton("🏠 Меню", callback_data="close_entry")])
    return InlineKeyboardMarkup(rows)


# Одне «активне» меню на чат: message_id останнього повідомлення з робочими
# кнопками. Перед тим як показати нове меню новим повідомленням (не редагуючи
# старе), кнопки з попереднього прибираються — щоб у чаті ніколи не лишалось
# декількох одночасно клікабельних меню.
_active_menu_message: dict[int, int] = {}


async def _clear_previous_menu(bot, chat_id: int) -> None:
    old_message_id = _active_menu_message.get(chat_id)
    if old_message_id is None:
        return
    try:
        await bot.edit_message_reply_markup(chat_id=chat_id, message_id=old_message_id, reply_markup=None)
    except Exception:  # noqa: BLE001 — повідомлення могли вже видалити чи відредагувати вручну
        pass


async def safe_edit(query: CallbackQuery, text: str, reply_markup=None) -> None:
    """edit_message_text, який не падає на 'Message is not modified'.

    Редагує те саме повідомлення на місці, тож воно й лишається єдиним
    активним меню в чаті — просто оновлюємо трекер на випадок наступного
    send_menu_reply (щоб було що прибирати)."""
    try:
        await query.edit_message_text(text, reply_markup=reply_markup)
    except BadRequest as error:
        if "Message is not modified" not in str(error):
            raise
        logger.debug("Ignored 'Message is not modified'")

    chat_id = query.message.chat_id
    if reply_markup is not None:
        _active_menu_message[chat_id] = query.message.message_id
    else:
        _active_menu_message.pop(chat_id, None)


async def send_menu_reply(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str, reply_markup=None):
    """Заміна update.message.reply_text для «менюшних» відповідей: перед тим як
    надіслати нове повідомлення, прибирає кнопки з попереднього активного меню
    в цьому чаті, щоб завжди залишалось рівно одне робоче меню, а не купа
    старих клікабельних вікон одне під одним."""
    chat_id = update.effective_chat.id
    await _clear_previous_menu(context.bot, chat_id)
    message = await update.message.reply_text(text, reply_markup=reply_markup)
    if reply_markup is not None:
        _active_menu_message[chat_id] = message.message_id
    else:
        _active_menu_message.pop(chat_id, None)
    return message


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
# Підтвердження адміном (review)
#
# Записи звичайних користувачів не пишуться в Excel одразу — створюється
# запис у _pending_reviews, а адміну летить лише легке сповіщення (без
# деталей і кнопок; повторно не турбуємо, поки черга не спорожніє — див.
# _notified_admins). Самі дії (підтвердити/відхилити/відредагувати) — тільки
# через кнопку меню «⏳ Очікують підтвердження». Перша відповідь виграє (під
# локом). Записи самих адмінів (is_admin) зберігаються без цього кроку —
# див. confirm_entry / process_edit_values.
# --------------------------------------------------------------------------- #

_pending_reviews: dict[str, dict] = {}


def _new_review_id() -> str:
    return uuid.uuid4().hex[:8]


def _user_label(user) -> str:
    username = getattr(user, "username", None)
    if username:
        return f"@{username}"
    full_name = getattr(user, "full_name", None)
    return full_name or str(getattr(user, "id", "?"))


async def _review_author_label(user) -> str:
    """Ім'я та прізвище, вписані при реєстрації — якщо їх ще немає, запасний варіант з Telegram."""
    registered_name = await get_registered_name(user.id)
    return registered_name or _user_label(user)


def _pending_review_count() -> int:
    return sum(1 for review in _pending_reviews.values() if review["status"] == "pending")


def _pending_reviews_sorted() -> list[dict]:
    """Усі запити зі статусом pending, у порядку надходження (найстаріші перші)."""
    reviews = [review for review in _pending_reviews.values() if review["status"] == "pending"]
    reviews.sort(key=lambda review: review.get("created_at", 0))
    return reviews


def _pending_reviews_by_user() -> dict[int, list[dict]]:
    """Групує запити по користувачу, зберігаючи хронологічний порядок усередині групи."""
    grouped: dict[int, list[dict]] = {}
    for review in _pending_reviews_sorted():
        grouped.setdefault(review["user_id"], []).append(review)
    return grouped


def _user_pending_reviews(user_id: int) -> list[dict]:
    """Власні запити цього користувача, що ще очікують підтвердження адміном."""
    return [review for review in _pending_reviews_sorted() if review["user_id"] == user_id]


def _user_pending_count(user_id: int) -> int:
    return len(_user_pending_reviews(user_id))


def _reviews_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "запит"
    if 2 <= count % 10 <= 4 and not (12 <= count % 100 <= 14):
        return "запити"
    return "запитів"


def _format_pending_users_list(grouped: dict[int, list[dict]]) -> str:
    if not grouped:
        return "⏳ Очікують підтвердження: немає запитів."
    total = sum(len(reviews) for reviews in grouped.values())
    return (
        f"⏳ Очікують підтвердження: {total} {_reviews_word(total)} "
        f"від {len(grouped)} користувача(ів).\n\nОбери, кого перевірити:"
    )


def _format_my_pending_reviews(reviews: list[dict]) -> str:
    """Список власних запитів користувача, що ще чекають на рішення адміна —
    без кнопок дій, це лише статус."""
    if not reviews:
        return "⏳ У тебе немає записів, що очікують підтвердження."
    lines = [f"⏳ Очікують підтвердження адміністратора: {len(reviews)} {_reviews_word(len(reviews))}.", ""]
    for review in reviews:
        if review["action"] == "delete":
            values = "видалити запис"
        else:
            _, _, day_total = calc_day_total(review["big"], review["small"])
            values = f"{BIG_LABEL} {review['big']}, {SMALL_LABEL} {review['small']} — {format_money(day_total)} €"
        lines.append(f"• {review['date_text']}: {values}")
    return "\n".join(lines)


def build_my_pending_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton("🏠 Меню", callback_data="close_entry")]])


def build_pending_users_keyboard(grouped: dict[int, list[dict]]) -> InlineKeyboardMarkup:
    rows = []
    for uid, reviews in grouped.items():
        label = reviews[0]["user_label"]
        rows.append(
            [InlineKeyboardButton(f"👤 {label} ({len(reviews)})", callback_data=f"review_user_{uid}")]
        )
    rows.append([InlineKeyboardButton("🏠 Меню", callback_data="close_entry")])
    return InlineKeyboardMarkup(rows)


async def _next_view_after_resolution(review: dict, admin_id: int) -> tuple[str, InlineKeyboardMarkup]:
    """Що показати замість щойно обробленого запиту: наступний запит цього ж
    користувача (щоб адмін підтверджував декілька дат поспіль, по черзі),
    інакше — згорнутий список користувачів, інакше — що все оброблено."""
    remaining_same_user = [r for r in _pending_reviews_sorted() if r["user_id"] == review["user_id"]]
    if remaining_same_user:
        next_review = remaining_same_user[0]
        text = _format_review_text(next_review)
        if len(remaining_same_user) > 1:
            text += f"\n\n(Ще {len(remaining_same_user) - 1} після цього для {next_review['user_label']})"
        return text, build_review_keyboard(next_review["id"], include_back=True)

    grouped = _pending_reviews_by_user()
    if grouped:
        return _format_pending_users_list(grouped), build_pending_users_keyboard(grouped)
    return "⏳ Усі запити оброблено.", build_main_menu(user_id=admin_id)


async def _refresh_after_action(query: CallbackQuery, admin_id: int, review: dict) -> None:
    """Після підтвердження/відхилення перемальовує екран перегляду на наступний
    запит цього ж користувача (або на список користувачів, якщо більше нема
    що обробляти)."""
    if query.message is None:
        return
    text, keyboard = await _next_view_after_resolution(review, admin_id)
    await safe_edit(query, text, reply_markup=keyboard)


def build_review_keyboard(review_id: str, include_back: bool = False) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton("✅ Підтвердити", callback_data=f"review_approve_{review_id}"),
            InlineKeyboardButton("❌ Відхилити", callback_data=f"review_reject_{review_id}"),
        ],
        [InlineKeyboardButton("✏️ Редагувати", callback_data=f"review_edit_{review_id}")],
    ]
    if include_back:
        rows.append([InlineKeyboardButton("⬅️ До списку користувачів", callback_data="pending_reviews")])
    return InlineKeyboardMarkup(rows)


def _format_review_text(review: dict, status_line: Optional[str] = None) -> str:
    action_label = {
        "add": "🆕 Новий запис",
        "edit": "✏️ Зміна запису",
        "delete": "🗑 Видалення запису",
    }[review["action"]]
    lines = [
        action_label,
        f"Користувач: {review['user_label']} (ID {review['user_id']})",
        f"Дата: {review['date_text']}",
    ]
    previous = review.get("previous")
    if previous is not None:
        lines.append(f"Було: {BIG_LABEL} {previous[0]}, {SMALL_LABEL} {previous[1]}")
    if review["action"] == "delete":
        lines.append("Дія: видалити запис за цей день")
    else:
        _, _, day_total = calc_day_total(review["big"], review["small"])
        label = "Стало" if previous is not None else "Значення"
        lines.append(f"{label}: {BIG_LABEL} {review['big']}, {SMALL_LABEL} {review['small']}")
        lines.append(f"Сума за день: {format_money(day_total)} €")
    if status_line:
        lines.append("")
        lines.append(status_line)
    return "\n".join(lines)


# Кому вже пінганули про наявність нових запитів — щоб не засипати адміна
# повідомленнями, поки він не розгребе чергу. Скидається, щойно черга
# спорожніє (див. _reset_admin_notifications_if_empty), і наступний новий
# запит знову надішле сповіщення.
_notified_admins: set[int] = set()


async def _notify_admins_new_review(context: ContextTypes.DEFAULT_TYPE, review: dict) -> None:
    """Легке сповіщення без деталей і без кнопок: «є новий запит, дивись
    вкладку». Деталі та дії — тільки через «⏳ Очікують підтвердження» в
    меню. Поки в адміна лишається бодай один необроблений запит, повторно
    не турбуємо його новими такими повідомленнями."""
    text = "🔔 Новий запит на підтвердження.\nПеревір вкладку «⏳ Очікують підтвердження» в меню."
    for admin_id in ADMIN_IDS:
        if admin_id in _notified_admins:
            continue
        try:
            await context.bot.send_message(admin_id, text)
            _notified_admins.add(admin_id)
        except Exception as error:  # noqa: BLE001 — збій сповіщення не має ламати основний потік
            logger.warning("Could not notify admin %s about review %s: %s", admin_id, review["id"], error)


def _reset_admin_notifications_if_empty() -> None:
    """Коли черга запитів спорожніла — знімаємо позначку «вже сповіщений»,
    щоб наступний новий запит знову підняв адмінів."""
    if _pending_review_count() == 0:
        _notified_admins.clear()


async def _notify_user_result(context: ContextTypes.DEFAULT_TYPE, review: dict, resolution: str) -> None:
    """resolution: 'approved' | 'rejected' | 'edited'."""
    user_id = review["user_id"]
    date_text = review["date_text"]

    if resolution == "rejected":
        text = f"❌ Ваш запис за {date_text} відхилено адміністратором. Дані не збережено."
    elif review["action"] == "delete":
        prefix = "🗑 Ваш запис" if resolution == "approved" else "✏️ Адміністратор видалив ваш запис"
        text = f"{prefix} за {date_text} видалено."
    else:
        _, _, day_total = calc_day_total(review["big"], review["small"])
        prefix = "✅ Ваш запис" if resolution == "approved" else "✏️ Адміністратор відредагував ваш запис"
        text = (
            f"{prefix} за {date_text} підтверджено:\n"
            f"{BIG_LABEL}: {review['big']}\n"
            f"{SMALL_LABEL}: {review['small']}\n"
            f"Сума за день: {format_money(day_total)} €"
        )

    try:
        await context.bot.send_message(
            user_id, text, reply_markup=build_main_menu(user_id=user_id, month=review["month"])
        )
    except Exception as error:  # noqa: BLE001 — користувач міг заблокувати бота
        logger.warning("Could not notify user %s about review %s result: %s", user_id, review["id"], error)


# --------------------------------------------------------------------------- #
# Хендлери: введення дня
# --------------------------------------------------------------------------- #

DATE_HINT = "Введи дату: 15-08-2026\nМожна також: 15.08.2026 або 15,08,2026"


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    user = update.message.from_user
    await register_user(user)
    logger.info("[USER:%s] /start (@%s)", user.id, getattr(user, "username", None))

    registered_name = await get_registered_name(user.id)
    if not registered_name:
        await send_menu_reply(update, context, 
            "Вітаю! Перш ніж почати, напиши, будь ласка, своє ім'я та прізвище "
            "(наприклад: Іван Петренко) — адміністратор бачитиме його в запитах на підтвердження."
        )
        return WAIT_REGISTER_NAME

    await send_menu_reply(update, context, DATE_HINT, reply_markup=build_main_menu(user_id=user.id))
    return WAIT_DATE


async def process_registration_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    full_name = re.sub(r"\s+", " ", update.message.text.strip())[:100]
    has_letters = re.search(r"[^\W\d_]", full_name, re.UNICODE)
    if len(full_name) < 3 or not has_letters:
        await send_menu_reply(update, context, 
            "Введи, будь ласка, ім'я та прізвище словами, наприклад: Іван Петренко"
        )
        return WAIT_REGISTER_NAME

    user = update.message.from_user
    await save_registered_name(user.id, full_name)
    logger.info("[USER:%s] Registered as %s", user.id, full_name)

    await send_menu_reply(update, context, 
        f"Дякую, {full_name}! Тепер можна вносити дані.\n\n{DATE_HINT}",
        reply_markup=build_main_menu(user_id=user.id),
    )
    return WAIT_DATE


async def process_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw_date = update.message.text.strip()
    if not DATE_PATTERN.match(raw_date):
        await send_menu_reply(update, context, "Невірний формат. Введи: 15-08-2026")
        return WAIT_DATE

    parsed = parse_date_input(raw_date)
    if parsed is None:
        await send_menu_reply(update, context, "Такої дати не існує. Введи: 15-08-2026")
        return WAIT_DATE

    selected = parsed.date()
    today = today_in_berlin()
    if selected > today:
        await send_menu_reply(update, context, 
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
            await send_menu_reply(update, context, 
                f"Для {selected.strftime('%d-%m-%Y')} записів не знайдено.",
                reply_markup=build_main_menu(user_id=user_id, month=month),
            )
            return WAIT_DATE
        big_count, small_count = record
        await send_menu_reply(update, context, 
            f"Запис {selected.strftime('%d-%m-%Y')}:\n"
            f"{BIG_LABEL}: {big_count}\n"
            f"{SMALL_LABEL}: {small_count}\n\n"
            f"{EDIT_VALUES_HINT}"
        )
        return WAIT_EDIT_VALUES

    await send_menu_reply(update, context, f"{BIG_LABEL}?")
    return WAIT_BIG


async def _read_count(update: Update, context: ContextTypes.DEFAULT_TYPE) -> Optional[int]:
    try:
        value = int(update.message.text.strip())
    except ValueError:
        await send_menu_reply(update, context, "Введи коректне число.")
        return None
    if value < 0:
        await send_menu_reply(update, context, "Кількість не може бути від'ємною.")
        return None
    return value


async def process_big_boxes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = await _read_count(update, context)
    if value is None:
        return WAIT_BIG
    context.user_data["big_count"] = value
    await send_menu_reply(update, context, f"{SMALL_LABEL}?")
    return WAIT_SMALL


async def process_small_boxes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = await _read_count(update, context)
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
    await send_menu_reply(update, context, 
        "Перевірка:\n"
        f"Дата: {context.user_data['date_text']}\n"
        f"{BIG_LABEL}: {big_count}\n"
        f"{SMALL_LABEL}: {value}\n\n"
        "Розрахунок:\n"
        f"{BIG_LABEL}: {big_count} × {RATE_BIG} / {SPLIT} = {format_money(big_total)} €\n"
        f"{SMALL_LABEL}: {value} × {RATE_SMALL} / {SPLIT} = {format_money(small_total)} €\n"
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
    date_text = context.user_data.get("date_text") or (f"{month}-{day:02d}" if month and day else "")

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

    if is_admin(user_id):
        await save_day(user_id, month, day, big_count, small_count)
        monthly_total = await read_month_total(user_id, month)
        context.user_data.clear()
        await safe_edit(
            query,
            "✅ Збережено\n"
            f"{month}-{day:02d}\n"
            f"{BIG_LABEL}: {big_count} = {format_money(big_total)} €\n"
            f"{SMALL_LABEL}: {small_count} = {format_money(small_total)} €\n"
            f"Сума за день: {format_money(day_total)} €\n"
            f"Разом за місяць: {format_money(monthly_total)} €",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    previous = await read_day(user_id, month, day)
    review = {
        "id": _new_review_id(),
        "user_id": user_id,
        "user_label": await _review_author_label(query.from_user),
        "month": month,
        "day": day,
        "date_text": date_text,
        "action": "add",
        "big": big_count,
        "small": small_count,
        "previous": previous,
        "status": "pending",
        "created_at": datetime.now(BERLIN_TZ).timestamp(),
    }
    _pending_reviews[review["id"]] = review
    context.user_data.clear()

    await safe_edit(
        query,
        "⏳ Надіслано адміністратору на підтвердження\n"
        f"{date_text}\n"
        f"{BIG_LABEL}: {big_count}\n"
        f"{SMALL_LABEL}: {small_count}\n"
        f"Сума за день: {format_money(day_total)} €",
        reply_markup=build_main_menu(user_id=user_id, month=month),
    )
    await _notify_admins_new_review(context, review)
    logger.info(
        "[USER:%s] Submitted review %s for %s (big=%s small=%s)",
        user_id, review["id"], date_text, big_count, small_count,
    )
    return WAIT_NEXT_ACTION


EDIT_VALUES_HINT = (
    f"Введи нові значення через пробіл: «{BIG_LABEL}» «{SMALL_LABEL}»\n"
    "Наприклад: 200 150\n"
    "Щоб видалити запис за цей день — введи: 0 0"
)


async def process_edit_values(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    parts = update.message.text.split()
    if len(parts) != 2:
        await send_menu_reply(update, context, 
            f"Введи два числа через пробіл: «{BIG_LABEL}» «{SMALL_LABEL}», наприклад 200 150"
        )
        return WAIT_EDIT_VALUES
    try:
        big_count, small_count = int(parts[0]), int(parts[1])
    except ValueError:
        await send_menu_reply(update, context, "Обидва значення мають бути числами. Наприклад: 200 150")
        return WAIT_EDIT_VALUES
    if big_count < 0 or small_count < 0:
        await send_menu_reply(update, context, "Кількість не може бути від'ємною.")
        return WAIT_EDIT_VALUES

    user_id = update.message.from_user.id
    month = context.user_data.get("month")
    day = context.user_data.get("day")
    if month is None or day is None:
        await send_menu_reply(update, context, DATE_HINT, reply_markup=build_main_menu(user_id=user_id))
        return WAIT_DATE

    date_text = context.user_data.get("date_text") or f"{month}-{day:02d}"
    previous = await read_day(user_id, month, day)
    context.user_data["edit_new_big"] = big_count
    context.user_data["edit_new_small"] = small_count
    context.user_data["edit_previous"] = previous

    lines = [f"Перевірка змін за {date_text}:", ""]
    if previous is not None:
        lines.append(f"Було: {BIG_LABEL} {previous[0]}, {SMALL_LABEL} {previous[1]}")
    if big_count == 0 and small_count == 0:
        lines.append("Стане: запис видалено")
    else:
        _, _, day_total = calc_day_total(big_count, small_count)
        lines.append(f"Стане: {BIG_LABEL} {big_count}, {SMALL_LABEL} {small_count}")
        lines.append(f"Сума за день: {format_money(day_total)} €")
    lines.append("")
    lines.append("Зберегти ці зміни?")

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("💾 Зберегти", callback_data="save_edit"),
                InlineKeyboardButton("✏️ Відредагувати", callback_data="redo_edit"),
            ]
        ]
    )
    await send_menu_reply(update, context, "\n".join(lines), reply_markup=keyboard)
    return WAIT_EDIT_CONFIRM


async def confirm_edit_values(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "redo_edit":
        context.user_data.pop("edit_new_big", None)
        context.user_data.pop("edit_new_small", None)
        context.user_data.pop("edit_previous", None)
        await safe_edit(query, f"Ок. {EDIT_VALUES_HINT}")
        return WAIT_EDIT_VALUES

    month = context.user_data.get("month")
    day = context.user_data.get("day")
    big_count = int(context.user_data.get("edit_new_big", 0))
    small_count = int(context.user_data.get("edit_new_small", 0))
    previous = context.user_data.get("edit_previous")
    date_text = context.user_data.get("date_text") or (f"{month}-{day:02d}" if month and day else "")

    if month is None or day is None:
        context.user_data.clear()
        await safe_edit(query, DATE_HINT, reply_markup=build_main_menu(user_id=user_id))
        return WAIT_DATE

    if is_admin(user_id):
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
            await safe_edit(query, text, reply_markup=build_main_menu(user_id=user_id, month=month))
            return WAIT_NEXT_ACTION

        day_total = await save_day(user_id, month, day, big_count, small_count)
        monthly_total = await read_month_total(user_id, month)
        context.user_data.clear()

        await safe_edit(
            query,
            f"✏️ Оновлено {month}-{day:02d}\n"
            f"{BIG_LABEL}: {big_count}\n"
            f"{SMALL_LABEL}: {small_count}\n"
            f"Сума за день: {format_money(day_total)} €\n"
            f"Разом за місяць: {format_money(monthly_total)} €",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    action = "delete" if (big_count == 0 and small_count == 0) else "edit"
    review = {
        "id": _new_review_id(),
        "user_id": user_id,
        "user_label": await _review_author_label(query.from_user),
        "month": month,
        "day": day,
        "date_text": date_text,
        "action": action,
        "big": big_count,
        "small": small_count,
        "previous": previous,
        "status": "pending",
        "created_at": datetime.now(BERLIN_TZ).timestamp(),
    }
    _pending_reviews[review["id"]] = review
    context.user_data.clear()

    if action == "delete":
        text = f"⏳ Запит на видалення запису за {date_text} надіслано адміністратору."
    else:
        _, _, day_total = calc_day_total(big_count, small_count)
        text = (
            f"⏳ Зміни за {date_text} надіслано адміністратору на підтвердження\n"
            f"{BIG_LABEL}: {big_count}\n"
            f"{SMALL_LABEL}: {small_count}\n"
            f"Сума за день: {format_money(day_total)} €"
        )
    await safe_edit(query, text, reply_markup=build_main_menu(user_id=user_id, month=month))
    await _notify_admins_new_review(context, review)
    logger.info(
        "[USER:%s] Submitted review %s for %s (%s)", user_id, review["id"], date_text, action,
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

    match = RE_MONTH_TOTAL.match(data)
    if match:
        return await show_month_total(query, context, match.group(1))

    match = RE_DOWNLOAD.match(data)
    if match:
        return await send_month_files(query, context, match.group(1))

    if data == "my_pending_reviews":
        reviews = _user_pending_reviews(user_id)
        await safe_edit(query, _format_my_pending_reviews(reviews), reply_markup=build_my_pending_keyboard())
        return WAIT_NEXT_ACTION

    # --- підтвердження записів адміном ------------------------------------ #
    if data == "pending_reviews":
        if not admin:
            await query.answer("Доступ заборонено", show_alert=True)
            return WAIT_NEXT_ACTION
        grouped = _pending_reviews_by_user()
        await safe_edit(query, _format_pending_users_list(grouped), reply_markup=build_pending_users_keyboard(grouped))
        return WAIT_NEXT_ACTION

    match = RE_REVIEW_USER.match(data)
    if match:
        return await handle_review_open_user(query, context, int(match.group(1)))

    match = RE_REVIEW_APPROVE.match(data)
    if match:
        return await handle_review_approve(query, context, match.group(1))

    match = RE_REVIEW_REJECT.match(data)
    if match:
        return await handle_review_reject(query, context, match.group(1))

    match = RE_REVIEW_EDIT.match(data)
    if match:
        return await handle_review_edit_start(query, context, match.group(1))

    # --- адмінські дії ---------------------------------------------------- #
    if data.startswith("admin_") and not admin:
        await query.answer("Доступ заборонено", show_alert=True)
        logger.warning("[USER:%s] Denied admin action: %s", user_id, data)
        return WAIT_NEXT_ACTION

    if data == "admin_monthly_summary":
        months = await list_months(None)
        if not months:
            await query.answer("Поки немає збережених місяців", show_alert=True)
            return WAIT_NEXT_ACTION
        await safe_edit(
            query,
            "За який місяць зробити зведену таблицю по всіх користувачах?",
            reply_markup=build_admin_summary_month_keyboard(months),
        )
        return WAIT_NEXT_ACTION

    match = RE_ADMIN_SUMMARY.match(data)
    if match:
        return await send_monthly_summary(query, match.group(1))

    if data in ("admin_users", "admin_log_by_user"):
        users = await known_users()
        if not users:
            await query.answer("Нема користувачів", show_alert=True)
            return WAIT_NEXT_ACTION
        prefix = "admin_user_" if data == "admin_users" else "admin_log_user_"
        rows = []
        for uid, (name, has_data) in users.items():
            label = f"{uid} — {name}" if name else str(uid)
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


async def handle_review_open_user(query: CallbackQuery, context: ContextTypes.DEFAULT_TYPE, target_user_id: int) -> int:
    """Показує перший (найстаріший) запит вибраного користувача з кнопками дій.
    Якщо в нього декілька дат — після кожної обробленої показується наступна,
    тож адмін підтверджує/відхиляє їх по черзі, не вертаючись щоразу у список."""
    if not is_admin(query.from_user.id):
        await query.answer("Доступ заборонено", show_alert=True)
        return WAIT_NEXT_ACTION

    reviews = [r for r in _pending_reviews_sorted() if r["user_id"] == target_user_id]
    if not reviews:
        await query.answer("У цього користувача більше немає запитів на підтвердження.", show_alert=True)
        grouped = _pending_reviews_by_user()
        await safe_edit(query, _format_pending_users_list(grouped), reply_markup=build_pending_users_keyboard(grouped))
        return WAIT_NEXT_ACTION

    review = reviews[0]
    text = _format_review_text(review)
    if len(reviews) > 1:
        text += f"\n\n(Ще {len(reviews) - 1} після цього для {review['user_label']})"
    await safe_edit(query, text, reply_markup=build_review_keyboard(review["id"], include_back=True))
    return WAIT_NEXT_ACTION


async def handle_review_approve(query: CallbackQuery, context: ContextTypes.DEFAULT_TYPE, review_id: str) -> int:
    admin_id = query.from_user.id
    if not is_admin(admin_id):
        await query.answer("Доступ заборонено", show_alert=True)
        return WAIT_NEXT_ACTION

    async with _lock_for(f"review:{review_id}"):
        review = _pending_reviews.get(review_id)
        if review is None or review["status"] != "pending":
            await query.answer("Цей запит вже оброблено або застарів.", show_alert=True)
            return WAIT_NEXT_ACTION
        review["status"] = "approved"
        review["resolved_by"] = admin_id
        del _pending_reviews[review_id]  # оброблений запит більше не тримаємо в пам'яті
    _locks.pop(f"review:{review_id}", None)

    if review["action"] == "delete":
        await delete_day(review["user_id"], review["month"], review["day"])
    else:
        await save_day(review["user_id"], review["month"], review["day"], review["big"], review["small"])

    _reset_admin_notifications_if_empty()
    await _refresh_after_action(query, admin_id, review)
    await _notify_user_result(context, review, "approved")
    logger.info(
        "[ADMIN:%s] Approved review %s for [USER:%s] %s (%s)",
        admin_id, review_id, review["user_id"], review["date_text"], review["action"],
    )
    return WAIT_NEXT_ACTION


async def handle_review_reject(query: CallbackQuery, context: ContextTypes.DEFAULT_TYPE, review_id: str) -> int:
    admin_id = query.from_user.id
    if not is_admin(admin_id):
        await query.answer("Доступ заборонено", show_alert=True)
        return WAIT_NEXT_ACTION

    async with _lock_for(f"review:{review_id}"):
        review = _pending_reviews.get(review_id)
        if review is None or review["status"] != "pending":
            await query.answer("Цей запит вже оброблено або застарів.", show_alert=True)
            return WAIT_NEXT_ACTION
        review["status"] = "rejected"
        review["resolved_by"] = admin_id
        del _pending_reviews[review_id]
    _locks.pop(f"review:{review_id}", None)

    _reset_admin_notifications_if_empty()
    await _refresh_after_action(query, admin_id, review)
    await _notify_user_result(context, review, "rejected")
    logger.info(
        "[ADMIN:%s] Rejected review %s for [USER:%s] %s",
        admin_id, review_id, review["user_id"], review["date_text"],
    )
    return WAIT_NEXT_ACTION


async def handle_review_edit_start(query: CallbackQuery, context: ContextTypes.DEFAULT_TYPE, review_id: str) -> int:
    admin_id = query.from_user.id
    if not is_admin(admin_id):
        await query.answer("Доступ заборонено", show_alert=True)
        return WAIT_NEXT_ACTION

    review = _pending_reviews.get(review_id)
    if review is None or review["status"] != "pending":
        await query.answer("Цей запит вже оброблено або застарів.", show_alert=True)
        return WAIT_NEXT_ACTION

    context.user_data.clear()
    context.user_data["admin_edit_review_id"] = review_id
    if query.message is not None:
        context.user_data["admin_edit_origin_chat_id"] = query.message.chat_id
        context.user_data["admin_edit_origin_message_id"] = query.message.message_id
    await safe_edit(
        query,
        f"Редагування запису {review['date_text']} для {review['user_label']}.\n\n{BIG_LABEL}?",
    )
    return WAIT_ADMIN_EDIT_BIG


async def process_admin_edit_big(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = await _read_count(update, context)
    if value is None:
        return WAIT_ADMIN_EDIT_BIG
    context.user_data["admin_edit_big"] = value
    await send_menu_reply(update, context, f"{SMALL_LABEL}?")
    return WAIT_ADMIN_EDIT_SMALL


async def process_admin_edit_small(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = await _read_count(update, context)
    if value is None:
        return WAIT_ADMIN_EDIT_SMALL

    admin_id = update.message.from_user.id
    review_id = context.user_data.get("admin_edit_review_id")
    review = _pending_reviews.get(review_id) if review_id else None
    if review is None or review["status"] != "pending":
        context.user_data.clear()
        await send_menu_reply(update, context, 
            "Цей запит вже оброблено або застарів.",
            reply_markup=build_main_menu(user_id=admin_id),
        )
        return WAIT_NEXT_ACTION

    big_count = int(context.user_data.get("admin_edit_big", 0))
    small_count = value
    context.user_data["admin_edit_small"] = small_count
    _, _, day_total = calc_day_total(big_count, small_count)

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("💾 Зберегти", callback_data="review_edit_save"),
                InlineKeyboardButton("❌ Скасувати", callback_data="review_edit_cancel"),
            ]
        ]
    )
    await send_menu_reply(update, context, 
        f"Новий варіант запису {review['date_text']} для {review['user_label']}:\n"
        f"{BIG_LABEL}: {big_count}\n"
        f"{SMALL_LABEL}: {small_count}\n"
        f"Сума за день: {format_money(day_total)} €\n\n"
        "Зберегти?",
        reply_markup=keyboard,
    )
    return WAIT_ADMIN_EDIT_CONFIRM


async def admin_edit_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    admin_id = query.from_user.id

    if query.data == "review_edit_cancel":
        context.user_data.clear()
        await safe_edit(query, "Скасовано. Запит усе ще очікує підтвердження в списку вище.")
        return WAIT_NEXT_ACTION

    review_id = context.user_data.get("admin_edit_review_id")
    big_count = int(context.user_data.get("admin_edit_big", 0))
    small_count = int(context.user_data.get("admin_edit_small", 0))
    origin_chat_id = context.user_data.get("admin_edit_origin_chat_id")
    origin_message_id = context.user_data.get("admin_edit_origin_message_id")
    context.user_data.clear()

    async with _lock_for(f"review:{review_id}"):
        review = _pending_reviews.get(review_id)
        if review is None or review["status"] != "pending":
            await safe_edit(query, "Цей запит вже оброблено іншим адміном.")
            return WAIT_NEXT_ACTION
        review["status"] = "edited"
        review["resolved_by"] = admin_id
        review["big"] = big_count
        review["small"] = small_count
        review["action"] = "delete" if (big_count == 0 and small_count == 0) else "edit"
        del _pending_reviews[review_id]
    _locks.pop(f"review:{review_id}", None)

    if review["action"] == "delete":
        await delete_day(review["user_id"], review["month"], review["day"])
    else:
        await save_day(review["user_id"], review["month"], review["day"], big_count, small_count)

    _reset_admin_notifications_if_empty()
    if origin_chat_id is not None and origin_message_id is not None:
        text, keyboard = await _next_view_after_resolution(review, admin_id)
        try:
            await context.bot.edit_message_text(
                chat_id=origin_chat_id, message_id=origin_message_id, text=text, reply_markup=keyboard
            )
        except Exception as error:  # noqa: BLE001 — це лише зручність навігації, не критично
            logger.debug("Could not advance origin review screen: %s", error)
    await _notify_user_result(context, review, "edited")
    logger.info(
        "[ADMIN:%s] Edited review %s for [USER:%s] %s -> big=%s small=%s",
        admin_id, review_id, review["user_id"], review["date_text"], big_count, small_count,
    )
    await safe_edit(query, "Збережено, користувача повідомлено.")
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


async def send_monthly_summary(query: CallbackQuery, month: str) -> int:
    """Зведена таблиця за місяць: один рядок на користувача (коробки кожного
    типу, разом і сума), відсортовано за алфавітом — для швидкого огляду адміном."""
    admin_id = query.from_user.id
    path = await build_monthly_summary(month)
    if path is None:
        await query.answer(f"За {month} ще немає жодних даних.", show_alert=True)
        return WAIT_NEXT_ACTION

    try:
        with open(path, "rb") as fh:
            await query.message.reply_document(document=fh, filename=f"Зведення_{month}.xlsx")
        logger.info("[ADMIN:%s] Sent monthly summary for %s", admin_id, month)
    except (OSError, BadRequest) as error:
        logger.exception("[ADMIN:%s] Failed to send monthly summary for %s: %s", admin_id, month, error)
        await query.message.reply_text("Не вдалося надіслати таблицю. Спробуй ще раз.")
    finally:
        try:
            path.unlink()
        except OSError:
            pass
    return WAIT_NEXT_ACTION


async def process_month_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    month = update.message.text.strip()
    if not MONTH_PATTERN.match(month):
        await send_menu_reply(update, context, "Невірний формат місяця. Введи YYYY-MM, наприклад 2026-08.")
        return WAIT_MONTH_INPUT

    user_id = update.message.from_user.id
    admin_target = context.user_data.pop("admin_view_user", None) if is_admin(user_id) else None

    if is_admin(user_id) and admin_target is None:
        total, names = await read_all_totals(month)
        if not names:
            await send_menu_reply(update, context, 
                f"За місяць {month} ще немає записів.",
                reply_markup=build_main_menu(user_id=user_id, month=month),
            )
            return WAIT_NEXT_ACTION
        await send_menu_reply(update, context, 
            f"Усі користувачі за {month}: {format_money(total)} €\nФайли: {', '.join(names)}",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    target_id = admin_target if admin_target is not None else user_id
    path = user_file(target_id, month)
    if not _month_has_data_sync(path):
        await send_menu_reply(update, context, 
            f"За місяць {month} ще немає записів.",
            reply_markup=build_main_menu(user_id=user_id, month=month),
        )
        return WAIT_NEXT_ACTION

    total = await read_month_total(target_id, month)
    context.user_data["download_user"] = target_id
    await send_menu_reply(update, context, 
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
        await send_menu_reply(update, context, "Формат: /total 2026-08")
        return

    user_id = update.message.from_user.id
    if is_admin(user_id):
        total, names = await read_all_totals(month)
        if not names:
            await send_menu_reply(update, context, f"За місяць {month} ще немає записів.")
            return
        await send_menu_reply(update, context, 
            f"Усі користувачі за {month}: {format_money(total)} €\nФайли: {', '.join(names)}"
        )
        return

    path = user_file(user_id, month)
    if not _month_has_data_sync(path):
        await send_menu_reply(update, context, f"За місяць {month} ще немає записів.")
        return
    total = await read_month_total(user_id, month)
    await send_menu_reply(update, context, f"{month}: {format_money(total)} €\nФайл: {path.name}")


async def cmd_logtail(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/logtail [N|file] — тільки для адмінів."""
    user_id = update.message.from_user.id
    if not is_admin(user_id):
        await send_menu_reply(update, context, "Доступ заборонено.")
        logger.warning("[USER:%s] Denied /logtail", user_id)
        return

    mode = context.args[0].strip().lower() if context.args else "100"

    if mode == "file":
        if not LOG_FILE.exists():
            await send_menu_reply(update, context, "Файл лога не знайдено.")
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
        await send_menu_reply(update, context, "Лог порожній.")
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
        await send_menu_reply(update, context, f"Останні {len(result)} рядків лога:\n\n{text}")
    logger.info("[ADMIN:%s] Requested log tail (%d lines)", user_id, len(result))


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await send_menu_reply(update, context, "Скасовано. Для нового запису натисни /start.")
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


async def _post_init(application: Application) -> None:
    """Налаштовує список команд, який Telegram показує в кнопці «Меню» біля
    поля вводу (те саме, що BotFather /setcommands). Адмінам показуємо ще й
    /logtail — окремою «областю видимості» лише для їхніх особистих чатів."""
    default_commands = [
        BotCommand("start", "Почати"),
        BotCommand("menu", "Головне меню"),
        BotCommand("total", "Сума за поточний місяць"),
        BotCommand("cancel", "Скасувати поточну дію"),
    ]
    try:
        await application.bot.set_my_commands(default_commands, scope=BotCommandScopeDefault())
    except Exception as error:  # noqa: BLE001 — відсутність меню команд не критична
        logger.warning("Could not set default command menu: %s", error)

    admin_commands = default_commands + [BotCommand("logtail", "Останні рядки логів (адмін)")]
    for admin_id in ADMIN_IDS:
        try:
            await application.bot.set_my_commands(admin_commands, scope=BotCommandScopeChat(chat_id=admin_id))
        except Exception as error:  # noqa: BLE001 — адмін міг ще жодного разу не писати боту
            logger.warning("Could not set admin command menu for %s: %s", admin_id, error)


class PerChatUpdateProcessor(BaseUpdateProcessor):
    """Апдейти різних чатів обробляються паралельно (до max_concurrent_updates
    одночасно) — щоб один активний користувач не змушував решту команди
    чекати в черзі. Апдейти ОДНОГО чату обробляються строго послідовно, щоб
    діалоговий стан того самого користувача не зіткнувся сам із собою —
    штатний concurrent_updates=True такого не гарантує (апдейти одного чату
    теж могли б піти паралельно)."""

    def __init__(self, max_concurrent_updates: int) -> None:
        super().__init__(max_concurrent_updates)
        self._chat_locks: dict[int, asyncio.Lock] = {}

    def _lock_for_chat(self, chat_id: int) -> asyncio.Lock:
        lock = self._chat_locks.get(chat_id)
        if lock is None:
            lock = asyncio.Lock()
            self._chat_locks[chat_id] = lock
        return lock

    async def do_process_update(self, update: object, coroutine) -> None:
        chat = getattr(update, "effective_chat", None)
        if chat is None:
            await coroutine
            return
        async with self._lock_for_chat(chat.id):
            await coroutine

    async def initialize(self) -> None:
        pass

    async def shutdown(self) -> None:
        pass


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

    application = (
        Application.builder()
        .token(token)
        .concurrent_updates(PerChatUpdateProcessor(32))
        .post_init(_post_init)
        .build()
    )

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
            WAIT_EDIT_CONFIRM: [
                CallbackQueryHandler(confirm_edit_values, pattern=r"^(save_edit|redo_edit)$"),
                callback_handler,
            ],
            WAIT_REGISTER_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_registration_name),
                callback_handler,
            ],
            WAIT_ADMIN_EDIT_BIG: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_admin_edit_big),
                callback_handler,
            ],
            WAIT_ADMIN_EDIT_SMALL: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_admin_edit_small),
                callback_handler,
            ],
            WAIT_ADMIN_EDIT_CONFIRM: [
                CallbackQueryHandler(admin_edit_confirm, pattern=r"^(review_edit_save|review_edit_cancel)$"),
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