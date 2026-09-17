"""
Модуль предобработки результатов опроса 360°.

Принимает CSV-выгрузку и собирает структурированный профиль сотрудника:
базовые и менеджерские компетенции, деструкторы, оценки ролей и качественные
комментарии. Этот профиль — единственный источник данных для генерации ИПР.

Ожидаемый формат CSV (колонки; порядок не важен, регистр заголовков игнорируется):

    тип,название,оценка,комментарий,роль_автора
    базовая компетенция,Профессионализм,9.5,,
    менеджерская компетенция,Делегирование,6.3,,
    деструктор,Бескомпромиссность,4.2,,
    роль,Наставник,8.3,,
    роль,Советник,10.0,,
    комментарий_сильная_сторона,,,«системность, погружённость в проект»,Лидер
    комментарий_зона_развития,,,«не хватает стратегического видения»,Команда

Колонка `тип` помогает отличить блоки друг от друга. Колонка `роль_автора` —
это роль оценщика (Лидер/Команда/Коллеги...), не путать со строками типа
"роль" — там название и оценка целевой/желаемой роли сотрудника.
Числа принимаются и с точкой, и с запятой.

Деструкторы не выводятся моделью отдельной категорией: высокий балл по
деструктору — такой же ограничитель, как низкий балл по компетенции, и оба
объединяются в "зоны риска" (см. Profile360.risk_zones). Пороги — ниже.

Комментарии не подставляются как самостоятельный источник выводов — только
как иллюстрация к выводу, уже сделанному по количественным оценкам (правило
закреплено в системном промпте ipr_generator.py).
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from io import StringIO


# ── Пороги зон риска ─────────────────────────────────────────────────────
# Шкала платформы — 1–4, середина 2,5 (та же, что в промпте AI-саммари).
#
# HIGH_DESTRUCTOR_THRESHOLD = 3.0 — рабочий порог, проверено на реальных
# данных: "выраженный деструктор" по определению из промпта AI-саммари —
# это балл по роли "Все окружение" (мнение коллег), а не самооценка.
# На такой шкале >3 срабатывает редко — это ожидаемо, а не баг: явный
# деструктор и должен быть редким сигналом, а не частым.
#
# LOW_COMPETENCY_THRESHOLD = 2.0 — на шкале 1–4 порог "<1" математически
# недостижим (это и обнаружилось на реальных 21 участнике беты — риск не
# находился ни разу). 2.0 — симметрично порогу деструкторов относительно
# середины 2,5 (2,5 - 0,5 = 2,0, как и 2,5 + 0,5 = 3,0). На тех же 21
# участниках это даёт 4 сработавших компетенции из 295 (у 2 из 21
# участников) — избирательный сигнал, а не срабатывание почти на всех.
LOW_COMPETENCY_THRESHOLD = 2.0
HIGH_DESTRUCTOR_THRESHOLD = 3.0


# Допустимые значения колонки «тип» и их канонизация
_TYPE_ALIASES = {
    "компетенция": "base_competency",          # обратная совместимость со старым CSV
    "компетенции": "base_competency",
    "competency": "base_competency",
    "базовая компетенция": "base_competency",
    "базовые компетенции": "base_competency",
    "менеджерская компетенция": "managerial_competency",
    "менеджерские компетенции": "managerial_competency",
    "деструктор": "destructor",
    "деструкторы": "destructor",
    "destructor": "destructor",
    "роль": "role",
    "роли": "role",
    "role": "role",
    "комментарий": "comment",                   # без деления — обратная совместимость
    "комментарии": "comment",
    "comment": "comment",
    "комментарий_сильная_сторона": "comment_strength",
    "сильная сторона": "comment_strength",
    "комментарий_зона_развития": "comment_growth",
    "зона развития": "comment_growth",
}

_COLUMN_ALIASES = {
    "тип": "type", "type": "type", "категория": "type",
    "название": "name", "name": "name", "компетенция": "name", "показатель": "name",
    "оценка": "score", "score": "score", "балл": "score", "значение": "score",
    "комментарий": "comment", "comment": "comment", "примечание": "comment",
    "роль_автора": "author_role", "роль автора": "author_role",
    "автор_роль": "author_role", "author_role": "author_role",
}


@dataclass
class ScoredItem:
    """Оценённый показатель: компетенция, деструктор или роль."""

    name: str
    score: float
    comment: str = ""


@dataclass
class CommentItem:
    """Комментарий коллеги — роль автора вместо имени (имена в датасете нежелательны)."""

    text: str
    author_role: str = ""


@dataclass
class Profile360:
    """Структурированный профиль по результатам опроса 360°."""

    base_competencies: list[ScoredItem] = field(default_factory=list)
    managerial_competencies: list[ScoredItem] = field(default_factory=list)
    destructors: list[ScoredItem] = field(default_factory=list)
    roles: list[ScoredItem] = field(default_factory=list)
    strength_comments: list[CommentItem] = field(default_factory=list)
    growth_comments: list[CommentItem] = field(default_factory=list)
    general_comments: list[CommentItem] = field(default_factory=list)  # без деления — старый формат CSV

    @property
    def is_empty(self) -> bool:
        return not (
            self.base_competencies or self.managerial_competencies
            or self.destructors or self.roles
        )

    def lowest_competencies(self, limit: int = 3) -> list[ScoredItem]:
        """Возвращает зоны роста — компетенции с самыми низкими оценками (база + менеджерские)."""
        all_competencies = self.base_competencies + self.managerial_competencies
        return sorted(all_competencies, key=lambda x: x.score)[:limit]

    def top_destructors(self, limit: int = 3) -> list[ScoredItem]:
        """Возвращает самые заметные деструкторы (с наибольшими оценками)."""
        return sorted(self.destructors, key=lambda x: x.score, reverse=True)[:limit]

    def risk_zones(self) -> list[tuple[str, float]]:
        """
        Объединённый список зон риска: компетенции ниже LOW_COMPETENCY_THRESHOLD
        и деструкторы выше HIGH_DESTRUCTOR_THRESHOLD — в равном статусе, без
        разделения на категории. Пороговая логика, а не топ-N: если у человека
        всё в норме, список будет пустым — риск не придумывается искусственно.
        """
        all_competencies = self.base_competencies + self.managerial_competencies
        low = [(i.name, i.score) for i in all_competencies if i.score < LOW_COMPETENCY_THRESHOLD]
        high_destructors = [
            (i.name, i.score) for i in self.destructors if i.score > HIGH_DESTRUCTOR_THRESHOLD
        ]
        return sorted(low, key=lambda x: x[1]) + sorted(high_destructors, key=lambda x: x[1], reverse=True)

    def to_prompt_block(self) -> str:
        """
        Готовит читаемый текстовый блок профиля для вставки в промпт.

        В анализ идут базовые и менеджерские компетенции, оценки ролей,
        объединённые зоны риска (низкие компетенции + высокие деструкторы —
        деструкторы отдельным списком не выводятся, только как часть зон
        риска) и комментарии коллег — как иллюстрация к выводу по цифрам, не
        как отдельный источник выводов.
        """
        lines: list[str] = []

        if self.base_competencies:
            lines.append("БАЗОВЫЕ КОМПЕТЕНЦИИ:")
            for item in sorted(self.base_competencies, key=lambda x: x.score, reverse=True):
                lines.append(f"- {item.name} — {_ru_number(item.score)}")

        if self.managerial_competencies:
            lines.append("\nМЕНЕДЖЕРСКИЕ КОМПЕТЕНЦИИ:")
            for item in sorted(self.managerial_competencies, key=lambda x: x.score, reverse=True):
                lines.append(f"- {item.name} — {_ru_number(item.score)}")

        if self.roles:
            lines.append("\nОЦЕНКИ РОЛЕЙ:")
            for item in self.roles:
                lines.append(f"- {item.name} — {_ru_number(item.score)}")

        risk = self.risk_zones()
        if risk:
            lines.append(
                "\nЗОНЫ РИСКА (низкие компетенции и высокие деструкторы — равнозначные ограничители):"
            )
            for name, score in risk:
                lines.append(f"- {name} — {_ru_number(score)}")

        def _format(items: list[CommentItem]) -> list[str]:
            out = []
            for c in items:
                prefix = f"[{c.author_role}] " if c.author_role else ""
                out.append(f"- {prefix}{c.text}")
            return out

        if self.strength_comments:
            lines.append(
                "\nКОММЕНТАРИИ КОЛЛЕГ — СИЛЬНЫЕ СТОРОНЫ (иллюстрация к выводу по цифрам, не цитировать дословно):"
            )
            lines.extend(_format(self.strength_comments))

        if self.growth_comments:
            lines.append(
                "\nКОММЕНТАРИИ КОЛЛЕГ — ЗОНЫ РАЗВИТИЯ (иллюстрация к выводу по цифрам, не цитировать дословно):"
            )
            lines.extend(_format(self.growth_comments))

        if self.general_comments:
            lines.append("\nКОММЕНТАРИИ КОЛЛЕГ (иллюстрация, не цитировать дословно):")
            lines.extend(_format(self.general_comments))

        return "\n".join(lines)


def _ru_number(value: float) -> str:
    """Форматирует число в русском десятичном формате с одним знаком: 9,5 / 9,0."""
    return f"{value:.1f}".replace(".", ",")


def _clean_text(raw) -> str:
    """Приводит значение к строке, считая nan/none/пусто пустой строкой."""
    if raw is None:
        return ""
    text = str(raw).strip()
    if text.lower() in ("nan", "none", "null", "<na>"):
        return ""
    return text


def _parse_score(raw) -> float | None:
    """Аккуратно приводит значение оценки к float, принимая запятую и точку."""
    if raw is None:
        return None
    text = str(raw).strip().replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _normalize_row(row: dict) -> dict:
    """Приводит ключи строки к каноническим именам type/name/score/comment/author_role."""
    normalized: dict = {}
    for col, value in row.items():
        if col is None:
            continue
        key = str(col).strip().lower().lstrip("\ufeff")
        normalized[_COLUMN_ALIASES.get(key, key)] = value
    return normalized


def parse_360_csv(file_bytes: bytes) -> Profile360:
    """
    Разбирает CSV-выгрузку 360° и возвращает структурированный профиль.

    Устойчив к кодировкам utf-8 и cp1251 и к разделителям «,» и «;».

    Args:
        file_bytes: Содержимое загруженного CSV-файла.

    Returns:
        Profile360 с разнесёнными по типам показателями.

    Raises:
        ValueError: Если файл не удалось прочитать ни одним из вариантов.
    """
    rows = _read_rows(file_bytes)
    rows = [_normalize_row(r) for r in rows]

    if not rows or not any("name" in r or "score" in r for r in rows):
        raise ValueError(
            "Не найдены колонки с названием и оценкой. "
            "Ожидаются колонки: тип, название, оценка, комментарий."
        )

    profile = Profile360()

    for row in rows:
        raw_type = _clean_text(row.get("type", "")).lower()
        item_type = _TYPE_ALIASES.get(raw_type, "")
        name = _clean_text(row.get("name", ""))
        comment = _clean_text(row.get("comment", ""))
        author_role = _clean_text(row.get("author_role", ""))
        score = _parse_score(row.get("score"))

        # Комментарий-строка: с делением на сильные стороны / зоны развития
        # или без деления (обратная совместимость со старым форматом CSV).
        if item_type in ("comment", "comment_strength", "comment_growth") or (not name and comment):
            if comment:
                comment_item = CommentItem(text=comment, author_role=author_role)
                if item_type == "comment_strength":
                    profile.strength_comments.append(comment_item)
                elif item_type == "comment_growth":
                    profile.growth_comments.append(comment_item)
                else:
                    profile.general_comments.append(comment_item)
            continue

        if score is None or not name:
            # Строка без оценки или без названия — собираем комментарий, если он есть
            if comment:
                profile.general_comments.append(CommentItem(text=comment, author_role=author_role))
            continue

        item = ScoredItem(name=name, score=score, comment=comment)

        if item_type == "managerial_competency":
            profile.managerial_competencies.append(item)
        elif item_type == "destructor":
            profile.destructors.append(item)
        elif item_type == "role":
            profile.roles.append(item)
        else:
            # По умолчанию (в т.ч. нераспознанный тип) — базовая компетенция.
            profile.base_competencies.append(item)

        if comment:
            profile.general_comments.append(CommentItem(text=comment, author_role=author_role))

    return profile


def _read_rows(file_bytes: bytes) -> list[dict]:
    """
    Читает CSV стандартной библиотекой, перебирая кодировки и разделители.

    Namedtuple/DataFrame не используются намеренно: pandas и pyarrow — тяжёлые
    нативные зависимости, а для разбора небольшой выгрузки достаточно csv.
    """
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            text = file_bytes.decode(encoding)
        except Exception as exc:  # noqa: BLE001 — пробуем следующую кодировку
            last_error = exc
            continue

        for sep in (",", ";", "\t"):
            try:
                reader = csv.DictReader(StringIO(text), delimiter=sep)
                if not reader.fieldnames or len(reader.fieldnames) < 2:
                    continue
                rows = [row for row in reader]
                if rows:
                    return rows
            except Exception as exc:  # noqa: BLE001 — перебираем варианты чтения
                last_error = exc
                continue

    raise ValueError(f"Не удалось прочитать CSV: {last_error or 'формат не распознан'}")
