"""
Разбор сырой xlsx-выгрузки платформы: "Групповой отчёт 360" (лист "Оценки" +
опционально лист "Комментарии") — в словарь {ФИО: Profile360}.

Это отдельный вход в тот же Profile360, что и tidy CSV из preprocessing.py —
используется, когда данные приходят прямо с платформы, а не в подготовленном
CSV. Дальше профиль уходит в ipr_generator.py как обычно.

ПРОВЕРЕНО на всех трёх файлах беты (без суффикса, консультант, неконсультант) —
структура и обработка ниже соответствуют реальным данным. Но три вещи нужно
подтвердить/решить, прежде чем пускать в прод — см. блок "ОТКРЫТЫЕ ВОПРОСЫ"
в конце файла. Без ответа на первый пункт (пороги зон риска) risk_zones()
будет пустым для всех — это уже проверено на всех 21 участнике беты.
"""

from __future__ import annotations

import re
from io import BytesIO

import openpyxl

from preprocessing import CommentItem, Profile360, ScoredItem, _clean_text, _parse_score

# ── Категории колонок листа "Оценки" ────────────────────────────────────────
# Колонки-лейблы ("Базовые компетенции", "Деструкторы" и т.п.) сами по себе —
# не показатель, а агрегат по блоку; данные начинаются со следующей колонки.
_CATEGORY_LABELS = {
    "базовые компетенции": "base_competency",
    "менеджерские компетенции": "managerial_competency",
    "деструкторы": "destructor",
}
_IGNORE_LABELS = {"востребованность в роли", "востребованность в экспертизе"}
_SKIP_LABELS = {"1. компетенции", "2. деструкторы"}

# Какая строка (роль оценщика) даёт "основной" балл по показателю.
# ОТКРЫТЫЙ ВОПРОС 2 — см. низ файла.
MAIN_SCORE_ROLE = "Все окружение"

# Роли, чьи комментарии НЕ идут в strength/growth (самооценка — не отзыв
# коллег; "Все окружение" — это агрегат, дублирующий текст остальных ролей).
# ОТКРЫТЫЙ ВОПРОС 3 — см. низ файла.
_EXCLUDE_ROLES_FROM_COMMENTS = {"Самооценка", "Все окружение"}

_NAME_ANYWHERE = re.compile(r"\(([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){0,2})\)")
_JUNK_VALUES = {"", "н/д", "нет", "-", "—"}


def _find_header_row(ws) -> int:
    """Строка с 'ФИО' в колонке A — позиция плавает (9 или 10) между файлами."""
    for r in range(1, 15):
        if str(ws.cell(row=r, column=1).value or "").strip() == "ФИО":
            return r
    raise ValueError("Не найдена строка заголовка (ФИО) в листе 'Оценки'")


def _map_columns(ws, header_row: int) -> dict[int, tuple[str, str]]:
    """
    Идёт по строке заголовка и определяет категорию каждой колонки с данными,
    отслеживая последний встреченный лейбл-заголовок ("текущая категория").
    """
    columns: dict[int, tuple[str, str]] = {}
    current: str | None = None
    for c in range(5, ws.max_column + 1):
        value = ws.cell(row=header_row, column=c).value
        if value is None:
            continue
        label = str(value).strip().lower()
        if label in _SKIP_LABELS:
            continue
        if label in _CATEGORY_LABELS:
            current = _CATEGORY_LABELS[label]
            continue  # сама колонка — агрегат по блоку, не отдельный показатель
        if label in _IGNORE_LABELS:
            current = "ignore"
            continue
        if current in ("base_competency", "managerial_competency", "destructor"):
            columns[c] = (current, str(value).strip())
    return columns


def parse_360_workbook(file_bytes: bytes) -> dict[str, Profile360]:
    """Разбирает лист 'Оценки' — базовые/менеджерские компетенции и деструкторы."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb["Оценки"]
    header_row = _find_header_row(ws)
    columns = _map_columns(ws, header_row)

    profiles: dict[str, Profile360] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        fio = _clean_text(ws.cell(row=r, column=1).value)
        role = _clean_text(ws.cell(row=r, column=3).value)
        if not fio or role != MAIN_SCORE_ROLE:
            continue
        profile = profiles.setdefault(fio, Profile360())
        for col_idx, (category, name) in columns.items():
            score = _parse_score(ws.cell(row=r, column=col_idx).value)
            if score is None:  # в том числе "н/д"
                continue
            item = ScoredItem(name=name, score=score)
            if category == "base_competency":
                profile.base_competencies.append(item)
            elif category == "managerial_competency":
                profile.managerial_competencies.append(item)
            elif category == "destructor":
                profile.destructors.append(item)
    return profiles


def clean_comment_blob(raw: str) -> str:
    """
    Вычищает блок комментария: убирает имена авторов в скобках (встречаются
    где угодно в тексте, не только в конце) и служебный перенос строки Excel
    (_x000D_).

    Намеренно НЕ пытается разбить блок на отдельные реплики нескольких
    респондентов — на реальных данных подписи есть не у всех: попытка
    разбивать по маркеру имени либо схлопывает несколько разных людей в
    одного автора, либо рвёт одну связную реплику на бессвязные обрывки
    (проверено на реальном тексте — оба варианта хуже, чем не дробить).
    Весь блок по роли передаётся как один текст.

    ИЗВЕСТНОЕ ОГРАНИЧЕНИЕ: ловит только "(Имя Фамилия)" в скобках. Голое имя
    без скобок в конце реплики (встречалось в реальных данных, напр. подпись
    просто "Маша") не вычищается — регуляркой это не отличить от обычного
    слова текста.
    """
    text = (raw or "").replace("_x000D_", "\n")
    text = _NAME_ANYWHERE.sub("", text)
    text = re.sub(r"\n{2,}", "\n", text).strip()
    return text


def attach_comments(file_bytes: bytes, profiles: dict[str, Profile360]) -> None:
    """
    Добавляет strength/growth-комментарии из листа 'Комментарии', если он есть
    (у файла для консультантов такого листа нет вообще — пропускаем молча).

    Лист "Комментарии" шире, чем группа участников в этом файле (там
    оказываются комментарии и по другим людям) — строки с ФИО не из текущей
    группы просто пропускаются.
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    if "Комментарии" not in wb.sheetnames:
        return
    ws = wb["Комментарии"]

    for r in range(2, ws.max_row + 1):
        fio = _clean_text(ws.cell(row=r, column=1).value)
        if fio not in profiles:
            continue
        role = _clean_text(ws.cell(row=r, column=3).value)
        if role in _EXCLUDE_ROLES_FROM_COMMENTS:
            continue
        profile = profiles[fio]

        strength = clean_comment_blob(str(ws.cell(row=r, column=4).value or ""))
        if strength.lower() not in _JUNK_VALUES:
            profile.strength_comments.append(CommentItem(text=strength, author_role=role))

        growth = clean_comment_blob(str(ws.cell(row=r, column=5).value or ""))
        if growth.lower() not in _JUNK_VALUES:
            profile.growth_comments.append(CommentItem(text=growth, author_role=role))


def parse_360_report(file_bytes: bytes) -> dict[str, Profile360]:
    """Полный разбор одного файла 'Групповой отчёт 360°': Оценки + Комментарии."""
    profiles = parse_360_workbook(file_bytes)
    attach_comments(file_bytes, profiles)
    return profiles


# ── ОТКРЫТЫЕ ВОПРОСЫ (до продакшена) ─────────────────────────────────────
#
# 1. РЕШЕНО. LOW_COMPETENCY_THRESHOLD = 2.0 в preprocessing.py (было
#    математически недостижимое <1 на шкале 1–4). HIGH_DESTRUCTOR_THRESHOLD
#    = 3.0 оставлен как есть — проверено на реальных данных, что это рабочий,
#    просто редкий порог (см. пункт 2 — почему это правильно).
#
# 2. MAIN_SCORE_ROLE = "Все окружение" — подтверждено, не просто догадка.
#    В промпте AI-саммари "выраженный деструктор" прямо определён как
#    "мешающее поведение, заметное ОКРУЖЕНИЮ" — то есть источником обязано
#    быть мнение коллег ("Все окружение"), а не самооценка. На реальных
#    данных именно самооценка иногда даёт деструктор >3 (до 4,0), а
#    "Все окружение" — почти никогда: это ожидаемо, раз источник —
#    восприятие коллег, а не то, что человек думает о себе сам.
#
# 3. Комментарии — какие роли считать "коллегами" для strength/growth.
#    Сейчас исключены "Самооценка" (это не отзыв коллеги) и "Все окружение"
#    (это агрегат, на реальных данных его текст дублирует текст остальных
#    ролей — иначе одна и та же мысль попадёт в промпт дважды). Если это
#    неверное решение — скажите, поправлю.
