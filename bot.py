import asyncio
import json
import logging
import os
import re
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, FSInputFile
from aiogram.utils.keyboard import ReplyKeyboardBuilder
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from google import genai
from google.genai import types

# ================== НАСТРОЙКИ ==================
# Заполни переменные окружения BOT_TOKEN и GEMINI_API_KEY.
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

# По текущей документации Google для новых проектов подходит Gemini 3.8 Flash.
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-3.8-flash").strip()

BASE_DIR = Path(__file__).resolve().parent
DOCS_DIR = BASE_DIR / "documents"
DOCS_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("officer_bot")

if not BOT_TOKEN:
    raise RuntimeError(
        "Не задан BOT_TOKEN. Установи переменную окружения BOT_TOKEN перед запуском."
    )
if not GEMINI_API_KEY:
    raise RuntimeError(
        "Не задан GEMINI_API_KEY. Установи переменную окружения GEMINI_API_KEY перед запуском."
    )

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
ai_client = genai.Client(api_key=GEMINI_API_KEY)


class LessonForm(StatesGroup):
    topic = State()
    lesson = State()
    duration = State()
    place = State()


def main_keyboard():
    kb = ReplyKeyboardBuilder()
    kb.button(text="📝 Создать конспект")
    kb.button(text="❌ Отмена")
    kb.adjust(1)
    return kb.as_markup(resize_keyboard=True)


def safe_filename(value: str) -> str:
    value = re.sub(r"[^\w\-]+", "_", value, flags=re.UNICODE).strip("_")
    return (value[:55] or "konspekt")


async def ask_gemini(prompt: str) -> dict:
    """Один запрос к ИИ; сетевой вызов выполняется асинхронно."""
    last_error = None
    for attempt in range(3):
        try:
            response = await ai_client.aio.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.4,
                    max_output_tokens=7000,
                ),
            )
            raw = (response.text or "").strip()
            if not raw:
                raise ValueError("ИИ вернул пустой ответ.")
            # На случай, если модель всё-таки обернула JSON в markdown.
            raw = re.sub(r"^\s*```(?:json)?\s*", "", raw)
            raw = re.sub(r"\s*```\s*$", "", raw)
            data = json.loads(raw)
            if not isinstance(data, dict):
                raise ValueError("ИИ вернул JSON неожиданного формата.")
            return data
        except Exception as exc:
            last_error = exc
            message = str(exc).upper()
            temporary = any(x in message for x in (
                "503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED",
                "500", "INTERNAL", "TIMEOUT", "DEADLINE",
            ))
            if not temporary or attempt == 2:
                break
            await asyncio.sleep(1.5 * (attempt + 1))
    raise RuntimeError(
        f"Не удалось получить ответ от Gemini. Проверь модель, ключ API и доступность сервиса. "
        f"Последняя ошибка: {last_error}"
    ) from last_error


def build_prompt(topic: str, lesson: str, duration: str, place: str) -> str:
    return f"""
Ты — методист, составляющий подробные учебные конспекты на русском языке.
Подготовь учебный конспект по указанным данным. Текст должен быть грамотным,
практичным, официальным и пригодным для оформления в Word. Не выдумывай номера
приказов, нормативных документов или цитаты. Если тема предполагает формулы,
приведи только уместные формулы и объясни обозначения и единицы измерения.
Не включай сведения, которые не относятся к теме.

ТЕМА: {topic}
ЗАНЯТИЕ: {lesson}
ПРОДОЛЖИТЕЛЬНОСТЬ: {duration}
МЕСТО: {place}

Верни ТОЛЬКО корректный JSON без markdown в такой структуре:
{{
  "goals": ["цель 1", "цель 2", "цель 3"],
  "materials": ["материал или наглядное пособие"],
  "sections": [
    {{
      "title": "Организационная часть",
      "minutes": "5 мин.",
      "leader": "Подробные действия и речь руководителя, что проверить и объяснить.",
      "trainees": "Действия обучаемых."
    }},
    {{
      "title": "Основная часть. Название вопроса",
      "minutes": "20 мин.",
      "leader": "Подробное содержание: определения, последовательность, объяснения, примеры; при необходимости формулы и пояснение обозначений.",
      "trainees": "Что делают, изучают, отрабатывают или отвечают обучаемые."
    }},
    {{
      "title": "Заключительная часть",
      "minutes": "5 мин.",
      "leader": "Подведение итогов, контрольные вопросы, выводы и задание на самостоятельную подготовку.",
      "trainees": "Отвечают на вопросы, уточняют непонятные моменты, записывают задание."
    }}
  ],
  "control_questions": ["вопрос 1", "вопрос 2", "вопрос 3", "вопрос 4", "вопрос 5"],
  "conclusion": "Краткий содержательный вывод по занятию."
}}

Требования:
- Сделай конспект подробным, не ограничивайся общими фразами.
- В sections должно быть 5–8 логически выстроенных учебных этапов, включая организационную часть, основную часть и заключение.
- Распредели минуты так, чтобы суммарно они соответствовали продолжительности занятия {duration}; если это невозможно из-за формата времени, используй разумное распределение.
- В каждом этапе опиши отдельно действия руководителя и обучаемых.
- goals — 3–5 конкретных учебных целей.
- control_questions — 5 вопросов по содержанию.
- Ответ только на русском языке.
"""


def set_document_font(doc: Document):
    for style_name in ("Normal", "Title", "Heading 1", "Heading 2"):
        style = doc.styles[style_name]
        style.font.name = "Times New Roman"
        style.font.size = Pt(12 if style_name == "Normal" else 14)
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)


def set_cell_text(cell, text, bold=False, size=8, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(str(text or ""))
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def create_word(data: dict, topic: str, lesson: str, duration: str, place: str) -> Path:
    doc = Document()
    set_document_font(doc)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("УТВЕРЖДАЮ\n")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    p.add_run("Командир __________________________\n")
    p.add_run("«____» ______________ 20____ г.")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ПЛАН-КОНСПЕКТ")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(16)

    for label, value in (
        ("Тема:", topic),
        ("Занятие:", lesson),
        ("Продолжительность:", duration),
        ("Место проведения:", place),
    ):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        rr = p.add_run(label + " ")
        rr.bold = True
        p.add_run(str(value))

    doc.add_heading("Учебные цели", level=1)
    for item in data.get("goals", []):
        doc.add_paragraph(str(item), style="List Bullet")

    materials = data.get("materials", [])
    if materials:
        doc.add_heading("Материальное обеспечение", level=1)
        doc.add_paragraph("; ".join(str(x) for x in materials))

    doc.add_heading("Ход занятия", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.5), Cm(1.2), Cm(11.0), Cm(3.5)]
    headers = [
        "Учебные вопросы",
        "Время",
        "Действия руководителя",
        "Действия обучаемых",
    ]
    for i, header in enumerate(headers):
        table.rows[0].cells[i].width = widths[i]
        set_cell_text(
            table.rows[0].cells[i], header, bold=True, size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )

    sections = data.get("sections", [])
    for item in sections:
        cells = table.add_row().cells
        for i, width in enumerate(widths):
            cells[i].width = width
        set_cell_text(cells[0], item.get("title", ""), size=8)
        set_cell_text(cells[1], item.get("minutes", ""), size=8,
                      align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_text(cells[2], item.get("leader", ""), size=8)
        set_cell_text(cells[3], item.get("trainees", ""), size=8)

    doc.add_heading("Контрольные вопросы", level=1)
    for i, question in enumerate(data.get("control_questions", []), 1):
        doc.add_paragraph(f"{i}. {question}")

    conclusion = data.get("conclusion", "")
    if conclusion:
        doc.add_heading("Заключение", level=1)
        doc.add_paragraph(str(conclusion))

    doc.add_paragraph()
    doc.add_paragraph("Руководитель занятия: __________________ /__________________/")

    filename = f"{datetime.now():%Y%m%d_%H%M%S}_{safe_filename(topic)}.docx"
    path = DOCS_DIR / filename
    doc.save(path)
    return path


@dp.message(CommandStart())
async def start(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "🪖 <b>ВОЕННЫЙ МЕТОДИСТ AI</b>\n\n"
        "Я подготовлю подробный план-конспект и оформлю его в Word.\n"
        "Нажми «Создать конспект», чтобы начать.",
        reply_markup=main_keyboard(),
        parse_mode="HTML",
    )


@dp.message(Command("cancel"))
@dp.message(F.text == "❌ Отмена")
async def cancel(message: Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "Создание конспекта отменено. Чтобы начать снова, нажми «Создать конспект».",
        reply_markup=main_keyboard(),
    )


@dp.message(Command("new"))
@dp.message(F.text == "📝 Создать конспект")
async def create_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(LessonForm.topic)
    await message.answer("Введи <b>тему занятия</b>:", parse_mode="HTML")


@dp.message(LessonForm.topic)
async def process_topic(message: Message, state: FSMContext):
    topic = (message.text or "").strip()
    if len(topic) < 3:
        await message.answer("Напиши тему подробнее, минимум 3 символа.")
        return
    await state.update_data(topic=topic)
    await state.set_state(LessonForm.lesson)
    await message.answer(
        "Теперь введи <b>номер и название занятия</b>.\n"
        "Например: <i>Занятие 1. Взвод в нападении</i>",
        parse_mode="HTML",
    )


@dp.message(LessonForm.lesson)
async def process_lesson(message: Message, state: FSMContext):
    lesson = (message.text or "").strip()
    if len(lesson) < 2:
        await message.answer("Укажи название занятия.")
        return
    await state.update_data(lesson=lesson)
    await state.set_state(LessonForm.duration)
    await message.answer("Укажи продолжительность занятия (например, 45 минут):")


@dp.message(LessonForm.duration)
async def process_duration(message: Message, state: FSMContext):
    duration = (message.text or "").strip()
    if len(duration) < 1:
        await message.answer("Укажи продолжительность занятия.")
        return
    await state.update_data(duration=duration)
    await state.set_state(LessonForm.place)
    await message.answer("Укажи место проведения занятия:")


@dp.message(LessonForm.place)
async def process_place(message: Message, state: FSMContext):
    place = (message.text or "").strip()
    if len(place) < 1:
        await message.answer("Укажи место проведения занятия.")
        return
    await state.update_data(place=place)
    data = await state.get_data()

    status = await message.answer(
        "⏳ Формирую подробный конспект через ИИ. Это может занять некоторое время..."
    )
    try:
        prompt = build_prompt(
            data["topic"], data["lesson"], data["duration"], data["place"]
        )
        generated = await ask_gemini(prompt)
        path = await asyncio.to_thread(
            create_word,
            generated,
            data["topic"],
            data["lesson"],
            data["duration"],
            data["place"],
        )
        await status.edit_text("✅ Конспект готов. Отправляю Word-документ...")
        await message.answer_document(
            FSInputFile(path),
            caption=f"📘 {data['topic']}\n{data['lesson']}",
        )
        await message.answer(
            "Готово! Для нового документа нажми «Создать конспект».",
            reply_markup=main_keyboard(),
        )
    except Exception as exc:
        logger.exception("Ошибка при создании конспекта")
        safe_error = str(exc)
        if len(safe_error) > 700:
            safe_error = safe_error[:700] + "…"
        await status.edit_text(
            "❌ Не удалось создать конспект.\n\n"
            f"Причина: {safe_error}\n\n"
            "Проверь API-ключ, доступность модели и подключение к интернету. "
            "Попробуй снова командой /new."
        )
        await message.answer("Можно начать заново:", reply_markup=main_keyboard())
    finally:
        await state.clear()


async def main():
    logger.info("Бот запускается. Модель Gemini: %s", MODEL_NAME)
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    finally:
        # Закрываем клиент, если версия SDK предоставляет метод close.
        try:
            ai_client.close()
        except Exception:
            pass
