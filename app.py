"""
ИПР·AI — MVP генерации индивидуального плана развития.

Сквозной сценарий в Streamlit, состыкованный с принятым пайплайном проекта:
    загрузка 360° (CSV) → экраны выбора → один промпт DeepSeek →
    структурированный вывод → рендер DOCX и ICS → превью и скачивание.

Визуальный стиль повторяет дельтовский референс (акцент #7c3aed, светлый фон,
шрифт Inter, мягкие тени, скругл-8px). Архитектура повторяет курсовой MVP:
боковая навигация по этапам, resolve_api_key, кэширование клиента.

Запуск:
    pip install -r requirements.txt
    streamlit run app.py

Ключ DeepSeek задаётся через Streamlit Secrets (DEEPSEEK_API_KEY) либо
вводится вручную в боковой панели.
"""

from __future__ import annotations

from datetime import datetime

import streamlit as st

from preprocessing import Profile360, parse_360_csv
from ipr_generator import EmployeeIntent, IPRGenerator
from renderers import render_docx, render_ics


# ---------------------------------------------------------------------------
# Конфигурация страницы и дельтовский стиль
# ---------------------------------------------------------------------------

ACCENT = "#7c3aed"

st.set_page_config(
    page_title="ИПР·AI — генерация плана развития",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        :root {{ color-scheme: light; }}
        html, body, [class*="css"] {{ font-family: 'Inter', -apple-system, sans-serif; }}

        /* Базовый светлый фон и тёмный текст — не зависим от темы деплоя */
        .stApp, [data-testid="stAppViewContainer"] {{ background: #f4f5f7; color: #1a1d23; }}
        [data-testid="stHeader"] {{ background: transparent; }}
        .main .block-container {{ padding-top: 2.2rem; max-width: 1060px; }}
        [data-testid="stSidebar"] {{ background: #ffffff; border-right: 1px solid #e2e5ea; }}

        /* Весь текст — тёмный, кроме кнопок и бейджей (у них свой цвет ниже) */
        .stApp, .stApp p, .stApp li, .stApp label,
        [data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] * ,
        [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] * {{ color: #1a1d23; }}
        h1, h2, h3, h4, h5, h6 {{ color: #1a1d23 !important; font-weight: 700; letter-spacing: -0.01em; }}
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * {{ color: #6b7280 !important; }}

        /* Поля ввода */
        .stTextInput input, .stTextArea textarea {{
            background: #ffffff !important; color: #1a1d23 !important;
            -webkit-text-fill-color: #1a1d23 !important;
            border: 1px solid #e2e5ea !important; border-radius: 8px !important;
        }}
        .stTextInput input::placeholder, .stTextArea textarea::placeholder {{
            color: #9aa1ab !important; -webkit-text-fill-color: #9aa1ab !important;
        }}
        /* Выпадающие списки (selectbox) */
        [data-baseweb="select"] > div {{
            background: #ffffff !important; color: #1a1d23 !important; border-color: #e2e5ea !important;
        }}
        [data-baseweb="select"] * {{ color: #1a1d23 !important; }}
        [data-baseweb="popover"], [data-baseweb="menu"], [role="option"] {{
            background: #ffffff !important; color: #1a1d23 !important;
        }}

        /* Радио-кнопки */
        [data-testid="stRadio"] label, [data-testid="stRadio"] label * {{ color: #1a1d23 !important; }}

        /* Загрузчик файлов */
        [data-testid="stFileUploaderDropzone"] {{
            background: #faf8ff !important; border: 1px dashed #ddd6fe !important; color: #1a1d23 !important;
        }}
        [data-testid="stFileUploaderDropzone"] * {{ color: #1a1d23 !important; }}

        /* Метрики */
        [data-testid="stMetricValue"] {{ color: #1a1d23 !important; }}
        [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * {{ color: #6b7280 !important; }}

        /* Бордерные контейнеры st.container(border=True), таблицы, экспандеры */
        [data-testid="stVerticalBlockBorderWrapper"] {{ background: #ffffff; border-radius: 8px; }}
        [data-testid="stTable"] table {{ background: #ffffff; color: #1a1d23; }}
        [data-testid="stTable"] th, [data-testid="stTable"] td {{ color: #1a1d23 !important; border-color: #e2e5ea !important; }}
        [data-testid="stExpander"] {{ background: #ffffff; border: 1px solid #e2e5ea; border-radius: 8px; }}
        [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * {{ color: #1a1d23 !important; }}

        /* Кнопки — акцентные, белый текст */
        .stButton > button, [data-testid="stDownloadButton"] button {{
            background: {ACCENT} !important; color: #fff !important; border: none !important;
            border-radius: 8px; padding: 0.55rem 1.4rem; font-weight: 600;
        }}
        .stButton > button *, [data-testid="stDownloadButton"] button * {{ color: #fff !important; }}
        .stButton > button:hover, [data-testid="stDownloadButton"] button:hover {{ background: #6d28d9 !important; }}

        /* Прочие классы демо */
        .step-badge {{
            display: inline-block; background: #ede9fe; color: #5b21b6 !important;
            border: 1px solid #ddd6fe; border-radius: 20px;
            font-size: 0.78rem; font-weight: 600; padding: 3px 12px; margin-bottom: 0.6rem;
        }}
        .card {{
            background: #fff; border: 1px solid #e2e5ea; border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.07); padding: 1.2rem 1.4rem; margin-bottom: 1rem;
        }}
        .card-accent {{ border-left: 4px solid {ACCENT}; }}
        .muted {{ color: #6b7280; font-size: 0.9rem; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Ключ DeepSeek и кэширование генератора
# ---------------------------------------------------------------------------


def resolve_credentials() -> tuple[str, str, str]:
    """
    Достаёт доступ к Яндекс Клауду: API-ключ, folder_id и id модели.

    Источники по приоритету: поля в боковой панели → Streamlit Secrets.
    Возвращает кортеж (api_key, folder_id, model_id).
    """
    def _secret(name: str, default: str = "") -> str:
        try:
            return st.secrets.get(name, default)
        except Exception:
            return default

    api_key = (st.session_state.get("manual_api_key", "") or _secret("YANDEX_API_KEY")).strip()
    folder_id = (st.session_state.get("manual_folder_id", "") or _secret("YANDEX_FOLDER_ID")).strip()
    model_id = (st.session_state.get("manual_model_id", "") or _secret("YANDEX_MODEL_ID") or "deepseek-v3").strip()
    return api_key, folder_id, model_id


@st.cache_resource(show_spinner=False)
def build_generator(api_key: str, folder_id: str, model_id: str) -> IPRGenerator:
    """Создаёт и кэширует генератор ИПР для заданных реквизитов Яндекс Клауда."""
    return IPRGenerator(api_key=api_key, folder_id=folder_id, model_id=model_id)


# ---------------------------------------------------------------------------
# Боковая панель
# ---------------------------------------------------------------------------


def render_sidebar() -> str:
    with st.sidebar:
        st.markdown("## 🎯 ИПР·AI")
        st.caption("Генерация индивидуального плана развития по данным 360°")

        section = st.radio(
            "Этап",
            options=[
                "1 · Загрузка 360°",
                "2 · Экраны выбора",
                "3 · Генерация ИПР",
            ],
            label_visibility="collapsed",
        )

        st.divider()
        st.markdown("**Доступ к DeepSeek (Яндекс Клауд)**")
        st.text_input(
            "API-ключ сервисного аккаунта",
            type="password", key="manual_api_key", placeholder="AQVN...",
            help="Ключ сервисного аккаунта с ролью ai.languageModels.user.",
        )
        st.text_input(
            "folder_id", key="manual_folder_id", placeholder="b1g...",
            help="Идентификатор каталога Яндекс Клауда.",
        )
        st.text_input(
            "id модели", key="manual_model_id", placeholder="deepseek-v3",
            help="Скопируй точный id модели DeepSeek из Model Gallery. По умолчанию deepseek-v3.",
        )
        api_key, folder_id, _ = resolve_credentials()
        if api_key and folder_id:
            st.success("Доступ настроен", icon="✅")
        else:
            st.info("Нужны ключ и folder_id", icon="ℹ️")

        st.divider()
        st.caption("Пайплайн: препроцессинг → промпт DeepSeek → DOCX / ICS")

    return section


# ---------------------------------------------------------------------------
# Этап 1. Загрузка 360°
# ---------------------------------------------------------------------------


def page_upload() -> None:
    st.markdown('<span class="step-badge">Этап 1</span>', unsafe_allow_html=True)
    st.title("Загрузка результатов 360°")
    st.write(
        "Загрузи CSV-выгрузку опроса. Колонки: тип, название, оценка, комментарий. "
        "Числа принимаются и с точкой, и с запятой."
    )

    uploaded = st.file_uploader("CSV с результатами 360°", type=["csv"])
    if uploaded is not None:
        try:
            profile = parse_360_csv(uploaded.getvalue())
        except ValueError as exc:
            st.error(f"Не удалось разобрать файл: {exc}")
            return

        if profile.is_empty:
            st.warning("В файле не нашлось оценок. Проверь колонки и формат.")
            return

        st.session_state["profile"] = profile
        st.success("Файл распознан. Проверь, что данные считаны верно.", icon="✅")
        _show_profile(profile)
    elif "profile" in st.session_state:
        st.info("Файл уже загружен ранее. Можно перейти к экранам выбора.")
        _show_profile(st.session_state["profile"])


def _show_profile(profile: Profile360) -> None:
    col1, col2, col3 = st.columns(3)
    col1.metric("Компетенции", len(profile.competencies))
    col2.metric("Деструкторы", len(profile.destructors))
    col3.metric("Оценки ролей", len(profile.roles))

    with st.expander("Что распознано", expanded=True):
        if profile.competencies:
            st.markdown("**Компетенции**")
            st.table(_items_to_rows(profile.competencies))
        if profile.destructors:
            st.markdown("**Деструкторы**")
            st.table(_items_to_rows(profile.destructors))
        if profile.roles:
            st.markdown("**Роли**")
            st.table(_items_to_rows(profile.roles))
        if profile.qualitative_comments:
            st.markdown("**Качественные комментарии** (в характеристику профиля не выводятся)")
            for c in profile.qualitative_comments:
                st.caption(f"— {c}")


def _items_to_rows(items) -> list[dict]:
    return [{"Показатель": it.name, "Оценка": str(it.score).replace(".", ",")} for it in items]


# ---------------------------------------------------------------------------
# Этап 2. Экраны выбора (Поток 2)
# ---------------------------------------------------------------------------


def page_choices() -> None:
    st.markdown('<span class="step-badge">Этап 2</span>', unsafe_allow_html=True)
    st.title("Экраны выбора")
    st.write(
        "Ответы задают контекст развития (раздел 2) и интенсивность плана. "
        "Это вводные вопросы, на которые опирается документ."
    )

    answers = st.session_state.get("intent_raw", {})

    dir_options = ["углубление", "горизонтально", "вертикально"]
    dir_captions = {
        "углубление": "Углубление — рост экспертизы и мастерства в текущей роли, без смены масштаба.",
        "горизонтально": "Горизонтально — расширение задач и смежные продукты без смены позиции.",
        "вертикально": "Вертикально — рост в уровне, больше ответственности и масштаба.",
    }
    dir_titles = {"углубление": "Углубление в роли", "горизонтально": "Горизонтально", "вертикально": "Вертикально"}

    with st.container(border=True):
        st.subheader("Для кого план")
        st.caption("Поля для ввода — заполни. Эти данные идут в шапку плана.")
        c1, c2 = st.columns(2)
        full_name = c1.text_input(
            "ФИО сотрудника", value=answers.get("full_name", ""),
            placeholder="Например: Иванова Анна",
        )
        role = c2.text_input(
            "Текущая роль", value=answers.get("role", ""),
            placeholder="Например: Руководитель направления",
        )
        period = st.selectbox(
            "Период плана", options=["Полугодие", "Квартал", "Год"],
            index=["Полугодие", "Квартал", "Год"].index(answers.get("period", "Полугодие"))
            if answers.get("period", "Полугодие") in ("Полугодие", "Квартал", "Год") else 0,
            help="Привязка к циклу, без жёстких дат.",
        )

    with st.container(border=True):
        st.subheader("1. Направление движения")
        st.caption("Куда движешься относительно текущей роли в ближайший год.")
        direction = st.radio(
            "Направление",
            options=dir_options,
            index=dir_options.index(answers.get("direction", "горизонтально"))
            if answers.get("direction", "горизонтально") in dir_options else 1,
            horizontal=True, format_func=lambda d: dir_titles[d],
            label_visibility="collapsed",
        )
        st.caption(dir_captions[direction])
        direction_note = st.text_input(
            "Пояснение — своими словами",
            value=answers.get("direction_note", ""),
            placeholder="Например: сильнее продукты и шире их принятие, без смены позиции",
            help="Один-два предложения: что это значит для тебя на практике.",
        )

    with st.container(border=True):
        st.subheader("2. Три ожидания от работы через год")
        st.caption("Принципиально не про карьерное качание — про содержание работы и её результат.")
        prev_exp = (answers.get("expectations", ["", "", ""]) + ["", "", ""])[:3]
        e1 = st.text_input("Ожидание 1", value=prev_exp[0],
                           placeholder="Например: продукты приносят больше и принимаются шире")
        e2 = st.text_input("Ожидание 2", value=prev_exp[1],
                           placeholder="Например: команда несёт продукт без постоянного моего участия")
        e3 = st.text_input("Ожидание 3", value=prev_exp[2],
                           placeholder="Например: клиенты верят в результат, а не только в модель")

    with st.container(border=True):
        st.subheader("3. Готовность вложиться в развитие")
        st.caption("Определяет темп: выше готовность — мягкий старт, дальше интенсивнее.")
        rd_options = ["Лёгкая", "Средняя", "Высокая"]
        rd_captions = {
            "Лёгкая": "Лёгкая — щадящий темп, небольшие шаги.",
            "Средняя": "Средняя — устойчивый темп без перегруза.",
            "Высокая": "Высокая — готов к интенсивной работе.",
        }
        readiness = st.radio(
            "Готовность", options=rd_options,
            index=rd_options.index(answers.get("readiness", "Средняя"))
            if answers.get("readiness", "Средняя") in rd_options else 1,
            horizontal=True, label_visibility="collapsed",
        )
        st.caption(rd_captions[readiness])
        hours = st.selectbox(
            "Сколько часов в неделю готов уделять",
            options=["1–2 ч", "3–5 ч", "5+ ч"],
            index=["1–2 ч", "3–5 ч", "5+ ч"].index(answers.get("hours_per_week", "3–5 ч"))
            if answers.get("hours_per_week", "3–5 ч") in ("1–2 ч", "3–5 ч", "5+ ч") else 1,
            help="Влияет на реалистичность ритма и нагрузку плана.",
        )

    if st.button("Сохранить ответы", type="primary"):
        if not full_name.strip():
            st.warning("Укажи ФИО сотрудника.")
            return
        st.session_state["intent_raw"] = {
            "full_name": full_name, "role": role, "period": period,
            "direction": direction, "direction_note": direction_note,
            "expectations": [e1, e2, e3],
            "readiness": readiness, "hours_per_week": hours,
        }
        st.success("Ответы сохранены. Можно переходить к генерации.", icon="✅")


# ---------------------------------------------------------------------------
# Этап 3. Генерация ИПР
# ---------------------------------------------------------------------------


def page_generate() -> None:
    st.markdown('<span class="step-badge">Этап 3</span>', unsafe_allow_html=True)
    st.title("Генерация ИПР")

    profile = st.session_state.get("profile")
    raw = st.session_state.get("intent_raw")

    if profile is None:
        st.warning("Сначала загрузи результаты 360° на этапе 1.")
        return
    if not raw:
        st.warning("Заполни экраны выбора на этапе 2.")
        return

    api_key, folder_id, model_id = resolve_credentials()
    if not api_key or not folder_id:
        st.warning("Для генерации нужны API-ключ и folder_id Яндекс Клауда — задай их слева.")
        return

    st.write("Готово к сборке. Профиль и ответы сотрудника подобраны.")
    cols = st.columns(3)
    cols[0].metric("Направление", raw["direction"].capitalize())
    cols[1].metric("Готовность · темп", f"{raw['readiness']} · {raw['hours_per_week']}")
    cols[2].metric("Компетенций", len(profile.competencies))

    if st.button("Сформировать ИПР", type="primary"):
        intent = EmployeeIntent(
            full_name=raw["full_name"], role=raw["role"], period=raw["period"],
            direction=raw["direction"], direction_note=raw["direction_note"],
            expectations=[e for e in raw["expectations"] if e.strip()],
            readiness=raw["readiness"], hours_per_week=raw["hours_per_week"],
        )
        with st.spinner("DeepSeek собирает план…"):
            generator = build_generator(api_key, folder_id, model_id)
            result = generator.generate(profile, intent)

        if not result.successful:
            st.error(result.error_message or "Не удалось сгенерировать план.")
            return

        st.session_state["ipr_data"] = result.data
        st.success(f"План готов. Токенов: {result.tokens_consumed}", icon="✅")

    if "ipr_data" in st.session_state:
        _show_result(st.session_state["ipr_data"], raw)


def _show_result(data: dict, raw: dict) -> None:
    st.divider()

    docx_bytes = render_docx(data)
    ics_bytes = render_ics(data, start_date=datetime.now())
    safe_name = raw["full_name"].replace(" ", "_") or "IPR"

    c1, c2 = st.columns(2)
    c1.download_button(
        "Скачать DOCX", data=docx_bytes,
        file_name=f"ИПР_{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    c2.download_button(
        "Скачать календарь (ICS)", data=ics_bytes,
        file_name=f"ИПР_{safe_name}.ics", mime="text/calendar",
    )

    st.subheader("Превью")
    if data.get("intro"):
        st.markdown(f"_{data['intro']}_")
    s1 = data.get("section1", {})
    if s1.get("purpose"):
        st.markdown("**1. Контекст и назначение**")
        st.write(s1["purpose"])
    s4 = data.get("section4", {})
    if s4.get("directions"):
        st.markdown("**4. Направления развития**")
        for d in s4["directions"]:
            with st.expander(f"{d.get('num','')} {d.get('title','')}"):
                st.write("**Цель:**", d.get("goal", ""))
                ras = d.get("risk_and_support", {})
                if ras:
                    st.write("**Риски и поддержка:**", ras.get("risk", ""))
                    if ras.get("psychological_support"):
                        st.caption(ras["psychological_support"])

    with st.expander("Полный JSON ответа"):
        st.json(data)


# ---------------------------------------------------------------------------
# Маршрутизация
# ---------------------------------------------------------------------------


def main() -> None:
    section = render_sidebar()
    if section.startswith("1"):
        page_upload()
    elif section.startswith("2"):
        page_choices()
    else:
        page_generate()


if __name__ == "__main__":
    main()
