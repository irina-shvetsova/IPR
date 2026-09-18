"""
Разбор индивидуального PDF-отчёта 360° (тот самый файл, который сотрудник
получает на почту и должен будет сам загрузить в генератор ИПР) — в Profile360.

Это третий вход в тот же Profile360, наравне с tidy CSV (preprocessing.py) и
групповым xlsx (xlsx_parser.py). Дальше профиль уходит в ipr_generator.py как
обычно.

ПРОВЕРЕНО на одном реальном отчёте (Ирина Швецова). Отчёт — это HeadlessChrome
print-to-PDF того же веб-инструмента, что и групповая xlsx-выгрузка, поэтому
таксономия показателей (названия компетенций и деструкторов) совпадает один в
один с xlsx — это и используется для очистки текста ниже.

ИЗВЕСТНЫЙ АРТЕФАКТ ЭТОГО PDF: часть слов в тексте отчёта (и в названиях
показателей, и в комментариях коллег) рендерится с лишними пробелами внутри
слова — «Про фес си о нал» вместо «Профессионал», «ин те ре су ешь ся» вместо
«интересуешься». Похоже на артефакт рендера (letter-spacing/justify в исходной
вёрстке), а не на баг самого текста.

Для НАЗВАНИЙ показателей это лечится надёжно: берём готовый список названий
компетенций/деструкторов (та же таксономия, что в xlsx) и матчим "без пробелов,
без регистра" — там конечный словарь, ошибиться некуда.

Для СВОБОДНОГО ТЕКСТА комментариев словаря нет, но есть pymorphy3: пробуем
склеить соседние короткие фрагменты и проверяем через морфологический анализ,
получилось ли настоящее русское слово (is_known). Первая версия на чистой длине
токенов ломала короткие настоящие слова («ты» + «интересуешься» схлопывались в
одно) — решилось требованием минимум 3 токенов в цепочке: у всех наблюдаемых
реальных обрывков (интересуешься, показываешь, высококлассный...) в цепочке
3+ фрагмента, а случайные ложные совпадения ("бы"+"к"="бык") — почти всегда
ровно 2. Остаточный риск не исключён (в тексте с 3+ короткими словами подряд,
случайно складывающимися в другое существующее слово, склейка ошибётся), но на
всех примерах из реального отчёта, включая этот edge case с "бык", работает
верно — см. _despace.
"""

from __future__ import annotations

import re
from io import BytesIO

import pdfplumber
import pymorphy3

from preprocessing import CommentItem, Profile360, ScoredItem

_morph = pymorphy3.MorphAnalyzer()
_known_cache: dict[str, bool] = {}


def _is_known_word(word: str) -> bool:
    word = word.lower()
    if word not in _known_cache:
        _known_cache[word] = _morph.parse(word)[0].is_known
    return _known_cache[word]


_PURE_WORD = re.compile(r"^[А-Яа-яЁё]+$")
_HYPHEN_TOKEN = re.compile(r"^[-–—‐‑]$")
_MAX_RUN_TOKENS = 10
_MAX_FRAGMENT_LEN = 8
_MIN_MERGE_TOKENS = 3  # см. докстринг модуля — отсекает случайные 2-токенные совпадения


def _despace(text: str) -> str:
    """
    Склеивает разорванные пробелами слова обратно, используя pymorphy3 как
    проверку "получилось ли настоящее слово". Жадно ищет САМОЕ ДЛИННОЕ
    распознанное слово, начиная с каждой позиции (иначе "ин" само по себе
    уже известное слово — архаичная частица — и склейка остановится, не
    дойдя до "интересуешься").
    """
    tokens = text.split(" ")
    out: list[str] = []
    i, n = 0, len(tokens)
    while i < n:
        tok = tokens[i]
        if _PURE_WORD.match(tok) and len(tok) <= _MAX_FRAGMENT_LEN:
            parts = [tok]
            best_end: int | None = None
            best_word: str | None = None
            j, steps = i + 1, 0
            while j < n and steps < _MAX_RUN_TOKENS:
                t = tokens[j]
                if _HYPHEN_TOKEN.match(t):  # перенос слова — символ пропускаем, цепочку не рвём
                    j += 1
                    steps += 1
                    continue
                if not (_PURE_WORD.match(t) and len(t) <= _MAX_FRAGMENT_LEN):
                    break
                parts.append(t)
                merged = "".join(parts)
                if _is_known_word(merged):
                    best_end, best_word = j, merged
                j += 1
                steps += 1
            if best_end is not None and best_end - i + 1 >= _MIN_MERGE_TOKENS:
                rebuilt = best_word[0].upper() + best_word[1:].lower() if tok[0].isupper() else best_word.lower()
                out.append(rebuilt)
                i = best_end + 1
                continue
        out.append(tok)
        i += 1
    return " ".join(out)

# ── Таксономия показателей — та же, что в xlsx_parser.py ────────────────────
_BASE_NAMES = [
    "Обучаемый", "Ладит с людьми", "Профессионал", "Умный", "Пахарь",
    "Инноватор", "Смелый", "Управляет рисками", "Инициативный",
    "Ориентирован на бизнес", "Уверенная самоподача",
]
_MANAGERIAL_NAMES = ["Визионер", "Вдохновитель", "Регулярный менеджер", "Hunter", "Тимбилдер", "Наставник"]
_DESTRUCTOR_NAMES = [
    "Против других", "Лучше других", "Избегает сложностей",
    "Сомневается в себе", "Соответствует требованиям", "Никому не должен",
]

# Ключ — название без пробелов и в нижнем регистре, чтобы матчить несмотря на
# то, где именно PDF вставил лишние пробелы внутри слова.
_CANONICAL = {re.sub(r"\s+", "", n).lower(): n for n in _BASE_NAMES + _MANAGERIAL_NAMES + _DESTRUCTOR_NAMES}

# Заголовок раздела — вторая строка страницы (первая — колонтитул "360° / ФИО").
_CATEGORY_HEADINGS = {
    "Базовые компетенции": "base_competency",
    "Менеджерские компетенции": "managerial_competency",
    "Востребованность в роли": "ignore",
    "Востребованность в экспертизе": "ignore",
    "Деструкторы": "destructor",
}

_SCORE_TOKEN = re.compile(r"^(?:\d+(?:[.,]\d+)?|НД)[!?]?$")


def _split_row(raw_row_text: str) -> tuple[str, float] | None:
    """
    Разбирает одну строку таблицы вида 'Профессионал 3.8 3 4' (или
    'Про фес си о нал 3.8 3 4' с артефактом) в (каноническое_название, балл_ВСЕ).

    Первая строка таблицы на странице часто содержит склеенный заголовок
    'ВСЕ Л КК\\n<первая строка данных>' — берём последнюю строку блока.
    Баллы берём с конца (их 1–3: ВСЕ/Л/КК), остаток — название. Если ВСЕ
    ("!" — реальность заметная средой) — это первый из отобранных баллов.
    Строки без распознанного названия (легенда "Мнения разделились" и т.п.)
    или без баллов вовсе (нет данных по этому показателю) отбрасываются.
    """
    line = raw_row_text.split("\n")[-1].strip()
    tokens = line.split()
    if not tokens:
        return None

    scores: list[str] = []
    i = len(tokens)
    while i > 0 and len(scores) < 3 and _SCORE_TOKEN.match(tokens[i - 1]):
        scores.insert(0, tokens[i - 1])
        i -= 1
    if not scores:
        return None

    name_part = "".join(tokens[:i])
    canonical = _CANONICAL.get(name_part.lower())
    if not canonical:
        return None

    all_score_raw = scores[0].rstrip("!?")  # балл "Все окружение" — первый столбец
    if all_score_raw == "НД":
        return None
    try:
        score = float(all_score_raw.replace(",", "."))
    except ValueError:
        return None
    return canonical, score


def _detect_category(page_text: str) -> str | None:
    lines = page_text.split("\n")
    if len(lines) < 2:
        return None
    return _CATEGORY_HEADINGS.get(lines[1].strip())


def _extract_fio(pdf: "pdfplumber.PDF") -> str:
    """ФИО из колонтитула '360° / <ФИО>', без хвостовой скобки с прежней фамилией."""
    for page in pdf.pages:
        first_line = (page.extract_text() or "").split("\n")[0].strip()
        m = re.match(r"^360°\s*/\s*(.+)$", first_line)
        if m:
            raw = m.group(1).strip()
            return re.sub(r"\s*\([^)]*\)\s*$", "", raw).strip()
    return ""


# ── Комментарии ──────────────────────────────────────────────────────────
_COMMENT_ROLE_SET = {
    "Самооценка", "Лидер", "Команда", "Коллеги по команде",
    "Коллеги из другой команды", "Все окружение",
}
_EXCLUDE_ROLES_FROM_COMMENTS = {"Самооценка"}  # самооценка — не отзыв коллеги
_NAME_ANYWHERE = re.compile(r"\(([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){0,2})\)")
_PAGE_HEADER = re.compile(r"^360°\s*/")
_PAGE_FOOTER = re.compile(r"^ЭКОПСИ\s+\d+\s*из\s*\d+$")
_JUNK_VALUES = {"", "нет комментариев"}


def _clean_comment_text(raw: str) -> str:
    text = _NAME_ANYWHERE.sub("", raw)
    text = re.sub(r"\s+", " ", text).strip()
    return _despace(text)


def _attach_comments(pdf: "pdfplumber.PDF", profile: Profile360) -> None:
    """
    Разбирает страницы 'Комментарии' построчно, накапливая многострочные
    реплики до следующего маркера "•" / смены роли / смены вопроса.

    Колонтитул страницы ("360° / ФИО") и подвал ("ЭКОПСИ N изN") отфильтрованы
    явно — без этого они приклеивались к последней реплике на странице.
    """
    bucket: str | None = None
    current_role: str | None = None
    pending: list[str] = []

    def flush() -> None:
        nonlocal pending
        parts, pending = pending, []
        if not parts or current_role is None or bucket is None:
            return
        text = _clean_comment_text(" ".join(parts))
        if text.lower() in _JUNK_VALUES:
            return
        if current_role in _EXCLUDE_ROLES_FROM_COMMENTS:
            return
        item = CommentItem(text=text, author_role=current_role)
        (profile.strength_comments if bucket == "strength" else profile.growth_comments).append(item)

    for page in pdf.pages:
        for raw_line in (page.extract_text() or "").split("\n"):
            line = raw_line.strip()
            if not line or _PAGE_HEADER.match(line) or _PAGE_FOOTER.match(line):
                continue
            lower = line.lower()
            if "особые таланты" in lower:
                flush(); bucket, current_role = "strength", None; continue
            if "получается плохо" in lower:
                flush(); bucket, current_role = "growth", None; continue
            if bucket is None:
                continue
            if line in _COMMENT_ROLE_SET:
                flush()
                current_role = line
                continue
            if current_role is None:
                continue
            if line.startswith("•"):
                flush()
                pending = [line[1:].strip()]
            else:
                pending.append(line)  # перенос строки внутри той же реплики
    flush()


def parse_360_pdf(file_bytes: bytes) -> tuple[str, Profile360]:
    """Разбирает индивидуальный PDF-отчёт 360°. Возвращает (ФИО, Profile360)."""
    profile = Profile360()
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        fio = _extract_fio(pdf)
        for page in pdf.pages:
            category = _detect_category(page.extract_text() or "")
            if category in (None, "ignore"):
                continue
            for table in page.extract_tables():
                for row in table:
                    result = _split_row(" ".join(c for c in row if c))
                    if result is None:
                        continue
                    name, score = result
                    item = ScoredItem(name=name, score=score)
                    if category == "base_competency":
                        profile.base_competencies.append(item)
                    elif category == "managerial_competency":
                        profile.managerial_competencies.append(item)
                    elif category == "destructor":
                        profile.destructors.append(item)
        _attach_comments(pdf, profile)
    return fio, profile
