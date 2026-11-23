import asyncio
import logging
import os
import re
from typing import Any, Dict, List

import openpyxl
from aiogram import Bot, Dispatcher, F, Router
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Путь к таблице из Части 1 (должна лежать в той же папке, что и бот)
EXCEL_PATH = os.getenv("EXCEL_PATH", "Часть_1_исследование_новостроек_готовая.xlsx")
SHEET_NAME = "Часть 1 — ЖК"

# Вставьте токен от @BotFather сюда или задайте переменную окружения BOT_TOKEN
BOT_TOKEN = os.getenv("BOT_TOKEN", "")


class SearchStates(StatesGroup):
    waiting_city = State()
    waiting_budget = State()
    waiting_area = State()


def load_complexes(filepath: str = EXCEL_PATH) -> List[Dict[str, Any]]:
    """Загружает данные о ЖК напрямую из Excel-таблицы Части 1."""
    if not os.path.exists(filepath):
        raise FileNotFoundError(
            f"Файл '{filepath}' не найден. Положите Excel-таблицу рядом с bot.py."
        )

    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[SHEET_NAME]

    complexes: List[Dict[str, Any]] = []
    # Заголовки находятся в 4-й строке, данные начинаются с 5-й строки
    for row in range(5, ws.max_row + 1):
        city = ws.cell(row, 1).value
        name = ws.cell(row, 2).value
        developer = ws.cell(row, 3).value
        district = ws.cell(row, 4).value
        seg_class = ws.cell(row, 5).value
        price_m2 = ws.cell(row, 6).value
        handover = ws.cell(row, 7).value
        link_cell = ws.cell(row, 8)
        url = link_cell.hyperlink.target if link_cell.hyperlink else link_cell.value

        if not city or not name or price_m2 is None:
            continue

        complexes.append(
            {
                "city": str(city).strip(),
                "name": str(name).strip(),
                "developer": str(developer).strip(),
                "district": str(district).strip(),
                "class": str(seg_class).strip(),
                "price_m2": float(price_m2),
                "handover": str(handover).strip(),
                "url": str(url).strip() if url else "",
            }
        )
    return complexes


def parse_budget_rub(text: str) -> float | None:
    """
    Распознаёт бюджет пользователя:
    - '25' или '25.5' -> трактуется как 25 000 000 ₽ (если число < 500)
    - '25 млн' -> 25 000 000 ₽
    - '25000000' -> 25 000 000 ₽
    """
    cleaned = text.lower().replace(" ", "").replace(",", ".")
    match = re.search(r"(\d+(?:\.\d+)?)", cleaned)
    if not match:
        return None
    val = float(match.group(1))
    if val <= 0:
        return None
    if "млн" in cleaned or val < 1000:
        return val * 1_000_000
    return val


def parse_area(text: str) -> float | None:
    """Распознаёт желаемую площадь в м²."""
    cleaned = text.lower().replace(" ", "").replace(",", ".")
    match = re.search(r"(\d+(?:\.\d+)?)", cleaned)
    if not match:
        return None
    val = float(match.group(1))
    return val if 10 <= val <= 1000 else None


def find_top_3_complexes(
    complexes: List[Dict[str, Any]],
    city: str,
    budget_rub: float,
    area_m2: float,
) -> tuple[List[Dict[str, Any]], bool]:
    """
    Подбирает 3 наиболее подходящих ЖК по городу, бюджету и площади.
    Возвращает (список_до_3_ЖК, флаг_укладываются_ли_в_бюджет).
    """
    city_items = [c for c in complexes if c["city"].lower() == city.lower()]
    enriched = []
    for item in city_items:
        est_price = item["price_m2"] * area_m2
        diff = budget_rub - est_price  # >= 0 значит укладывается в бюджет
        enriched.append({**item, "est_price": est_price, "diff": diff})

    within_budget = [x for x in enriched if x["diff"] >= 0]
    if within_budget:
        # Среди укладывающихся в бюджет выбираем 3 максимальных по классу/цене (ближе к бюджету)
        within_budget.sort(key=lambda x: x["est_price"], reverse=True)
        return within_budget[:3], True

    # Если бюджет меньше любого ЖК в городе — показываем 3 самых доступных варианта с минимальной доплатой
    enriched.sort(key=lambda x: x["est_price"])
    return enriched[:3], False


def city_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🌴 Сочи", callback_data="city:Сочи"),
                InlineKeyboardButton(text="🏛 Москва", callback_data="city:Москва"),
            ],
            [
                InlineKeyboardButton(
                    text="⚓ Санкт-Петербург", callback_data="city:Санкт-Петербург"
                )
            ],
        ]
    )


def restart_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🔄 Новый подбор", callback_data="restart")]
        ]
    )


router = Router()


@router.message(CommandStart())
@router.message(Command("search"))
async def cmd_start(message: Message, state: FSMContext) -> None:
    await state.clear()
    await state.set_state(SearchStates.waiting_city)
    await message.answer(
        "👋 <b>AI-подборщик новостроек (бизнес и премиум)</b>\n\n"
        "Я подберу <b>3 подходящих ЖК</b> из верифицированной базы ЦИАН "
        "и рассчитаю примерную стоимость квартиры.\n\n"
        "<b>Шаг 1 из 3.</b> Выберите город:",
        reply_markup=city_keyboard(),
        parse_mode=ParseMode.HTML,
    )


@router.callback_query(F.data == "restart")
async def cb_restart(callback: CallbackQuery, state: FSMContext) -> None:
    await callback.answer()
    await state.clear()
    await state.set_state(SearchStates.waiting_city)
    await callback.message.answer(
        "<b>Шаг 1 из 3.</b> Выберите город:",
        reply_markup=city_keyboard(),
        parse_mode=ParseMode.HTML,
    )


@router.callback_query(SearchStates.waiting_city, F.data.startswith("city:"))
async def process_city_cb(callback: CallbackQuery, state: FSMContext) -> None:
    city = callback.data.split(":", 1)[1]
    await state.update_data(city=city)
    await state.set_state(SearchStates.waiting_budget)
    await callback.answer()
    await callback.message.answer(
        f"Выбран город: <b>{city}</b>\n\n"
        "<b>Шаг 2 из 3.</b> Введите ваш бюджет на квартиру (например: <code>25</code> или <code>25.5 млн</code>):",
        parse_mode=ParseMode.HTML,
    )


@router.message(SearchStates.waiting_city)
async def process_city_text(message: Message, state: FSMContext) -> None:
    text = (message.text or "").strip().lower()
    mapping = {
        "сочи": "Сочи",
        "москва": "Москва",
        "санкт-петербург": "Санкт-Петербург",
        "спб": "Санкт-Петербург",
        "питер": "Санкт-Петербург",
    }
    city = mapping.get(text)
    if not city:
        await message.answer(
            "Пожалуйста, выберите город кнопкой ниже или напишите: Сочи, Москва или Санкт-Петербург.",
            reply_markup=city_keyboard(),
        )
        return
    await state.update_data(city=city)
    await state.set_state(SearchStates.waiting_budget)
    await message.answer(
        f"Выбран город: <b>{city}</b>\n\n"
        "<b>Шаг 2 из 3.</b> Введите ваш бюджет на квартиру (например: <code>25</code> или <code>25.5 млн</code>):",
        parse_mode=ParseMode.HTML,
    )


@router.message(SearchStates.waiting_budget)
async def process_budget(message: Message, state: FSMContext) -> None:
    budget_rub = parse_budget_rub(message.text or "")
    if not budget_rub:
        await message.answer(
            "Не удалось распознать бюджет. Введите число в млн ₽, например: <code>22</code> или <code>35.5</code>.",
            parse_mode=ParseMode.HTML,
        )
        return

    await state.update_data(budget_rub=budget_rub)
    await state.set_state(SearchStates.waiting_area)
    await message.answer(
        f"Бюджет: <b>{budget_rub / 1_000_000:,.2f} млн ₽</b>\n\n"
        "<b>Шаг 3 из 3.</b> Введите желаемую площадь квартиры в м² (например: <code>45</code>):",
        parse_mode=ParseMode.HTML,
    )


@router.message(SearchStates.waiting_area)
async def process_area(message: Message, state: FSMContext) -> None:
    area_m2 = parse_area(message.text or "")
    if not area_m2:
        await message.answer(
            "Введите корректную площадь от 10 до 1000 м², например: <code>52</code>.",
            parse_mode=ParseMode.HTML,
        )
        return

    user_data = await state.get_data()
    city = user_data["city"]
    budget_rub = user_data["budget_rub"]

    complexes = load_complexes()
    top_3, fits_budget = find_top_3_complexes(complexes, city, budget_rub, area_m2)

    header = (
        f"🏙 <b>Город:</b> {city}\n"
        f"💰 <b>Бюджет:</b> {budget_rub / 1_000_000:,.2f} млн ₽\n"
        f"📐 <b>Желаемая площадь:</b> {area_m2:g} м²\n\n"
    )

    if fits_budget:
        header += f"✅ <b>Найдено подходящих ЖК ({len(top_3)} из 3):</b>\n\n"
    else:
        header += (
            "⚠️ <b>В указанный бюджет лоты заданной площади не укладываются.</b>\n"
            "Вот <b>3 самых близких по стоимости</b> варианта в этом городе:\n\n"
        )

    cards = []
    for idx, item in enumerate(top_3, 1):
        est_mln = item["est_price"] / 1_000_000
        price_m2_fmt = f"{item['price_m2']:,.0f}".replace(",", " ")
        diff_mln = abs(item["diff"]) / 1_000_000
        status_line = (
            f"🟢 Запас бюджета: <b>{diff_mln:,.2f} млн ₽</b>"
            if item["diff"] >= 0
            else f"🟠 Превышение бюджета: <b>+{diff_mln:,.2f} млн ₽</b>"
        )
        cards.append(
            f"<b>{idx}. ЖК «{item['name']}»</b> ({item['class']})\n"
            f"• Застройщик: {item['developer']}\n"
            f"• Район: {item['district']}\n"
            f"• Цена за м²: <b>{price_m2_fmt} ₽/м²</b>\n"
            f"• Расчётная стоимость ({area_m2:g} м²): <b>{est_mln:,.2f} млн ₽</b>\n"
            f"• {status_line}\n"
            f"• Срок сдачи: {item['handover']}\n"
            f"• <a href=\"{item['url']}\">Карточка ЖК на ЦИАН</a>"
        )

    await state.clear()
    await message.answer(
        header + "\n\n".join(cards),
        reply_markup=restart_keyboard(),
        parse_mode=ParseMode.HTML,
        disable_web_page_preview=True,
    )


async def main() -> None:
    if BOT_TOKEN == "ВСТАВЬТЕ_ВАШ_ТОКЕН_ОТ_BOTFATHER":
        raise RuntimeError(
            "Укажите токен бота в переменной BOT_TOKEN (получить можно у @BotFather)."
        )
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
