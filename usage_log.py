"""
Логирование использования ИПР·AI в Google-таблицу.

Пишет по одной строке на каждое значимое действие: какие данные загрузили, что
ввели на экранах выбора и какой план получили. Основной приёмник — Google-таблица
(переживает перезапуски Streamlit Cloud). Если доступа к таблице нет, событие
дописывается в локальный JSONL-файл — чтобы локальный запуск не ломался.

ВАЖНО: в лог попадают персональные данные (имена, оценки 360°, тексты планов).
Это фактически хранилище ПДн — место хранения согласуйте с комплаенсом.

Подключение таблицы (в Streamlit Secrets):

    USAGE_LOG_SHEET_ID = "1AbC...id_таблицы..."

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "ipr-logger@....iam.gserviceaccount.com"
    ...

Таблицу нужно расшарить на client_email сервисного аккаунта с правом
редактирования. Подробнее — в usage_log_integration.md.
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone

LOG_PATH = os.environ.get("USAGE_LOG_PATH", "usage_log.jsonl")
ENABLED = os.environ.get("USAGE_LOG_ENABLED", "1") not in ("0", "false", "False")

# Заголовки таблицы. Порядок фиксирован — по нему собирается строка.
_HEADER = ["ts", "session", "event", "name", "tokens", "warning", "details"]

# Ленивая инициализация листа: подключаемся один раз и кэшируем.
_worksheet = None
_worksheet_ready = False

# Последняя загрузка (сессия, файл) — для отсева повторов из-за перерисовок Streamlit.
_last_upload = None


def _get_worksheet():
    """
    Возвращает объект листа Google-таблицы или None.

    Реквизиты берутся из Streamlit Secrets: сервисный аккаунт в секции
    [gcp_service_account] и идентификатор таблицы в USAGE_LOG_SHEET_ID.
    Подключение кэшируется, чтобы не авторизоваться на каждое событие.
    """
    global _worksheet, _worksheet_ready
    if _worksheet_ready:
        return _worksheet

    _worksheet_ready = True  # пытаемся один раз; при неудаче — работает файл
    try:
        import gspread
        import streamlit as st
        from google.oauth2.service_account import Credentials

        sheet_id = st.secrets.get("USAGE_LOG_SHEET_ID", "")
        creds_info = st.secrets.get("gcp_service_account", None)
        if not sheet_id or not creds_info:
            _worksheet = None
            return None

        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(dict(creds_info), scopes=scopes)
        client = gspread.authorize(creds)
        ws = client.open_by_key(sheet_id).sheet1

        # Проставляем заголовок, если лист пустой
        if not ws.get_all_values():
            ws.update("A1", [_HEADER], value_input_option="RAW")

        _worksheet = ws
        return ws
    except Exception:  # noqa: BLE001 — любая проблема с таблицей → откат на файл
        _worksheet = None
        return None


def _row_from(record: dict) -> list[str]:
    """Сплющивает событие в строку таблицы по схеме _HEADER."""
    data = record.get("data", {})
    return [
        record.get("ts", ""),
        record.get("session", ""),
        record.get("event", ""),
        str(data.get("full_name", "")),
        str(data.get("tokens", "")),
        str(data.get("warning", "")),
        json.dumps(data, ensure_ascii=False),
    ]


def _write_line(record: dict) -> None:
    """
    Пишет одно событие. Сначала пробует Google-таблицу, иначе — локальный файл.

    Это единственная точка ввода-вывода: чтобы сменить приёмник (например, на
    объектное хранилище Яндекса), достаточно переписать её тело.
    """
    ws = _get_worksheet()
    if ws is not None:
        try:
            # Пишем в следующую пустую строку начиная с колонки A.
            # append_row в gspread без явного диапазона «уползает» вправо по
            # диагонали, поэтому вычисляем строку сами и обновляем A{row}.
            next_row = len(ws.get_all_values()) + 1
            ws.update(f"A{next_row}", [_row_from(record)], value_input_option="RAW")
            return
        except Exception:  # noqa: BLE001 — при сбое таблицы дублируем в файл
            pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:  # noqa: BLE001 — лог не должен ронять приложение
        pass


def _event(session_id: str, event: str, data: dict) -> None:
    if not ENABLED:
        return
    _write_line({
        "ts": datetime.now(timezone.utc).isoformat(),
        "session": session_id,
        "event": event,
        "data": data,
    })


def log_upload(session_id: str, filename: str, profile) -> None:
    """
    Фиксирует загруженный профиль 360°.

    Streamlit перерисовывает страницу несколько раз после загрузки файла,
    поэтому одно и то же событие может прийти дважды. Дедупим по паре
    (сессия, имя файла), чтобы в логе была одна запись на загрузку.
    """
    global _last_upload
    key = (session_id, filename)
    if _last_upload == key:
        return
    _last_upload = key

    _event(session_id, "upload", {
        "filename": filename,
        "n_competencies": len(profile.competencies),
        "n_destructors": len(profile.destructors),
        "n_roles": len(profile.roles),
        "competencies": [{"name": c.name, "score": c.score} for c in profile.competencies],
        "roles": [{"name": r.name, "score": r.score} for r in profile.roles],
    })


def log_choices(session_id: str, intent_raw: dict) -> None:
    """Фиксирует ответы сотрудника с экранов выбора."""
    _event(session_id, "choices", {
        "full_name": intent_raw.get("full_name", ""),
        "role": intent_raw.get("role", ""),
        "direction": intent_raw.get("direction", ""),
        "direction_note": intent_raw.get("direction_note", ""),
        "expectations": [e for e in intent_raw.get("expectations", []) if e],
        "readiness": intent_raw.get("readiness", ""),
    })


def log_generation(session_id: str, full_name: str, plan: dict,
                   tokens: int, warning: str | None) -> None:
    """Фиксирует результат генерации: полный план, токены, предупреждения."""
    _event(session_id, "generation", {
        "full_name": full_name,
        "tokens": tokens,
        "warning": warning or "",
        "plan": plan,
    })


def read_log(limit: int | None = None) -> list[dict]:
    """
    Читает лог для выгрузки/просмотра.

    Если настроена Google-таблица — читает из неё, иначе из локального файла.
    Возвращает список событий; при недоступности источника — пустой список.
    """
    ws = _get_worksheet()
    if ws is not None:
        try:
            rows = ws.get_all_values()
            records = []
            for row in rows[1:]:  # пропускаем заголовок
                cells = (row + [""] * len(_HEADER))[:len(_HEADER)]
                ts, session, event, _name, _tokens, _warning, details = cells
                try:
                    data = json.loads(details) if details else {}
                except json.JSONDecodeError:
                    data = {}
                records.append({"ts": ts, "session": session, "event": event, "data": data})
            return records[-limit:] if limit else records
        except Exception:  # noqa: BLE001 — при сбое таблицы читаем файл
            pass

    if not os.path.exists(LOG_PATH):
        return []
    records = []
    with open(LOG_PATH, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    records.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    return records[-limit:] if limit else records
