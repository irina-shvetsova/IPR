"""
Модуль предобработки результатов опроса 360°.

Принимает CSV-выгрузку и собирает структурированный профиль сотрудника:
компетенции, деструкторы, оценки ролей и качественные комментарии. Этот
профиль — единственный количественный источник для генерации ИПР.

Ожидаемый формат CSV (колонки; порядок не важен, регистр заголовков игнорируется):

    тип,название,оценка,комментарий
    компетенция,Профессионализм,9.5,
    компетенция,Визионерство,6.3,
    деструктор,Бескомпромиссность,4.2,
    роль,Наставник,8.3,
    роль,Советник,10.0,
    комментарий,,,«системность, погружённость в проект»

Колонка `тип` помогает отличить компетенции от деструкторов, ролей и
свободных комментариев. Числа принимаются и с точкой, и с запятой.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO, StringIO

import pandas as pd


# Допустимые значения колонки «тип» и их канонизация
_TYPE_ALIASES = {
    "компетенция": "competency",
    "компетенции": "competency",
    "competency": "competency",
    "деструктор": "destructor",
    "деструкторы": "destructor",
    "destructor": "destructor",
    "роль": "role",
    "роли": "role",
    "role": "role",
    "комментарий": "comment",
    "комментарии": "comment",
    "comment": "comment",
}

_COLUMN_ALIASES = {
    "тип": "type", "type": "type", "категория": "type",
    "название": "name", "name": "name", "компетенция": "name", "показатель": "name",
    "оценка": "score", "score": "score", "балл": "score", "значение": "score",
    "комментарий": "comment", "comment": "comment", "примечание": "comment",
}


@dataclass
class ScoredItem:
    """Оценённый показатель: компетенция, деструктор или роль."""

    name: str
    score: float
    comment: str = ""


@dataclass
class Profile360:
    """Структурированный профиль по результатам опроса 360°."""

    competencies: list[ScoredItem] = field(default_factory=list)
    destructors: list[ScoredItem] = field(default_factory=list)
    roles: list[ScoredItem] = field(default_factory=list)
    qualitative_comments: list[str] = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        return not (self.competencies or self.destructors or self.roles)

    def lowest_competencies(self, limit: int = 3) -> list[ScoredItem]:
        """Возвращает зоны роста — компетенции с самыми низкими оценками."""
        return sorted(self.competencies, key=lambda x: x.score)[:limit]

    def top_destructors(self, limit: int = 3) -> list[ScoredItem]:
        """Возвращает самые заметные деструкторы (с наибольшими оценками)."""
        return sorted(self.destructors, key=lambda x: x.score, reverse=True)[:limit]

    def to_prompt_block(self) -> str:
        """Готовит читаемый текстовый блок профиля для вставки в промпт."""
        lines: list[str] = []

        if self.competencies:
            lines.append("КОМПЕТЕНЦИИ:")
            for item in sorted(self.competencies, key=lambda x: x.score, reverse=True):
                lines.append(f"- {item.name} — {_ru_number(item.score)}")

        if self.destructors:
            lines.append("\nДЕСТРУКТОРЫ:")
            for item in sorted(self.destructors, key=lambda x: x.score, reverse=True):
                lines.append(f"- {item.name} — {_ru_number(item.score)}")

        if self.roles:
            lines.append("\nОЦЕНКИ РОЛЕЙ:")
            for item in self.roles:
                lines.append(f"- {item.name} — {_ru_number(item.score)}")

        if self.qualitative_comments:
            lines.append("\nКАЧЕСТВЕННЫЕ КОММЕНТАРИИ КОЛЛЕГ:")
            for comment in self.qualitative_comments:
                lines.append(f"- {comment}")

        return "\n".join(lines)


def _ru_number(value: float) -> str:
    """Форматирует число в русском десятичном формате: 9,5 вместо 9.5."""
    if value == int(value):
        return str(int(value))
    return f"{value:.1f}".replace(".", ",")


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


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Приводит заголовки колонок к каноническим именам type/name/score/comment."""
    renamed = {}
    for col in df.columns:
        key = str(col).strip().lower()
        renamed[col] = _COLUMN_ALIASES.get(key, key)
    return df.rename(columns=renamed)


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
    df = _read_dataframe(file_bytes)
    df = _normalize_columns(df)

    if "name" not in df.columns and "score" not in df.columns:
        raise ValueError(
            "Не найдены колонки с названием и оценкой. "
            "Ожидаются колонки: тип, название, оценка, комментарий."
        )

    profile = Profile360()

    for _, row in df.iterrows():
        raw_type = str(row.get("type", "")).strip().lower()
        item_type = _TYPE_ALIASES.get(raw_type, "")
        name = str(row.get("name", "")).strip()
        comment = str(row.get("comment", "")).strip()
        score = _parse_score(row.get("score"))

        # Свободный комментарий без оценки
        if item_type == "comment" or (not name and comment):
            if comment:
                profile.qualitative_comments.append(comment)
            continue

        if score is None or not name:
            # Строка без оценки или без названия — собираем комментарий, если он есть
            if comment:
                profile.qualitative_comments.append(comment)
            continue

        item = ScoredItem(name=name, score=score, comment=comment)

        if item_type == "destructor":
            profile.destructors.append(item)
        elif item_type == "role":
            profile.roles.append(item)
        else:
            # По умолчанию считаем строку компетенцией
            profile.competencies.append(item)

        if comment:
            profile.qualitative_comments.append(comment)

    return profile


def _read_dataframe(file_bytes: bytes) -> pd.DataFrame:
    """Пытается прочитать CSV в разных кодировках и с разными разделителями."""
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        for sep in (",", ";", "\t"):
            try:
                text = file_bytes.decode(encoding)
                df = pd.read_csv(StringIO(text), sep=sep)
                if df.shape[1] >= 2:
                    return df
            except Exception as exc:  # noqa: BLE001 — перебираем варианты чтения
                last_error = exc
                continue
    # Последняя попытка — пусть pandas сам определит параметры
    try:
        return pd.read_csv(BytesIO(file_bytes))
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"Не удалось прочитать CSV: {last_error or exc}")
