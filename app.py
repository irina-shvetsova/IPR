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
from ipr_generator import (
    EmployeeIntent, IPRGenerator, checkpoint_spec,
    READINESS_OPTIONS, readiness_key, readiness_plan,
)
from plan_texts import PREAMBLE_TITLE, SECTIONS, build_preamble, section_title
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
        .main .block-container {{ padding-top: 2.2rem; max-width: 820px; }}
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
        [data-testid="stExpander"] details, [data-testid="stExpander"] summary {{ background: #ffffff !important; }}
        [data-testid="stExpander"] summary:hover {{ background: #f4f5f7 !important; }}
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


STEP_LABELS = ["1 · Загрузка 360°", "2 · Экраны выбора", "3 · Генерация ИПР"]


def render_sidebar() -> None:
    """Рисует навигацию. Активный этап хранится в st.session_state.step (1–3)."""
    st.session_state.setdefault("step", 1)
    with st.sidebar:
        st.markdown("## 🎯 ИПР·AI")
        st.caption("Генерация индивидуального плана развития по данным 360°")

        # Радио без key: индексом управляет session_state.step, клик — сверяем и перерисовываем
        choice = st.radio(
            "Этап", options=STEP_LABELS, index=st.session_state.step - 1,
            label_visibility="collapsed",
        )
        chosen = STEP_LABELS.index(choice) + 1
        if chosen != st.session_state.step:
            st.session_state.step = chosen
            st.rerun()

        st.divider()
        st.caption("Пайплайн: препроцессинг → промпт DeepSeek → DOCX / ICS")


def goto_step(step: int) -> None:
    """Переключает активный этап и перерисовывает страницу."""
    st.session_state.step = step
    st.rerun()


# ---------------------------------------------------------------------------
# Этап 1. Загрузка 360°
# ---------------------------------------------------------------------------


def page_upload() -> None:
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
        st.divider()
        if st.button("Далее → Экраны выбора", type="primary"):
            goto_step(2)
    elif "profile" in st.session_state:
        st.info("Файл уже загружен ранее. Можно перейти к экранам выбора.")
        _show_profile(st.session_state["profile"])
        st.divider()
        if st.button("Далее → Экраны выбора", type="primary"):
            goto_step(2)


def _show_profile(profile: Profile360) -> None:
    col1, col2 = st.columns(2)
    col1.metric("Компетенции", len(profile.competencies))
    col2.metric("Оценки ролей", len(profile.roles))

    with st.container(border=True):
        st.markdown("**Что распознано**")
        st.caption("В анализ идут только компетенции и оценки ролей.")
        if profile.competencies:
            st.markdown("Компетенции")
            st.markdown(_items_to_markdown(profile.competencies))
        if profile.roles:
            st.markdown("Роли")
            st.markdown(_items_to_markdown(profile.roles))


def _items_to_markdown(items) -> str:
    """
    Собирает markdown-таблицу показателей.

    Streamlit сериализует st.table/st.dataframe через pyarrow; на некоторых
    сборках это приводит к падению процесса, поэтому таблицу рисуем текстом.
    """
    lines = ["| Показатель | Оценка |", "| --- | ---: |"]
    for item in items:
        score = f"{item.score:.1f}".replace(".", ",")
        lines.append(f"| {item.name} | {score} |")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Этап 2. Экраны выбора (Поток 2)
# ---------------------------------------------------------------------------


def page_choices() -> None:
    st.title("Экраны выбора")
    st.write(
        "Ответы задают контекст развития (раздел 2) и интенсивность плана. "
        "Это вводные вопросы, на которые опирается документ."
    )
    st.info(
        "Отвечай реалистично и в пределах года. Цель — какой конкретный шаг сделать за "
        "этот срок. Формулируй как наблюдаемый результат в работе.",
        icon="🎯",
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
        st.caption("Поля для ввода — заполни. Эти данные идут в шапку плана. План составляется на год.")
        c1, c2 = st.columns(2)
        full_name = c1.text_input(
            "ФИО сотрудника", value=answers.get("full_name", ""),
            placeholder="Например: Иванова Анна",
        )
        role = c2.text_input(
            "Текущая роль", value=answers.get("role", ""),
            placeholder="Например: Руководитель направления",
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
            placeholder="Например: вырасти до руководителя группы из 3–4 человек",
            help="Реалистичный шаг на год, а не мечта на десять лет.",
        )

    with st.container(border=True):
        st.subheader("2. Три ожидания от работы через год")
        st.caption(
            "Не про должность, а про содержание работы и её видимый результат. "
            "Ожидания равнозначны: в плане они будут перечислены твоими словами, "
            "без домысливания мотивов."
        )
        prev_exp = (answers.get("expectations", ["", "", ""]) + ["", "", ""])[:3]
        e1 = st.text_input("Ожидание 1", value=prev_exp[0],
                           placeholder="Например: сам веду переговоры с крупными клиентами")
        e2 = st.text_input("Ожидание 2", value=prev_exp[1],
                           placeholder="Например: отвечаю за отдельный продукт, а не отдельные задачи")
        e3 = st.text_input("Ожидание 3", value=prev_exp[2],
                           placeholder="Например: команда решает типовые вопросы без меня")

    with st.container(border=True):
        st.subheader("3. Готовность вложиться в развитие")
        st.caption("Определяет объём плана: сколько направлений и действий он будет содержать.")
        prev_readiness = answers.get("readiness_option", READINESS_OPTIONS[1])
        readiness_option = st.radio(
            "Готовность", options=READINESS_OPTIONS,
            index=READINESS_OPTIONS.index(prev_readiness)
            if prev_readiness in READINESS_OPTIONS else 1,
            label_visibility="collapsed",
        )
        readiness = readiness_key(readiness_option)
        volume = readiness_plan(readiness)
        st.caption(
            f"{volume['hint'].capitalize()}. План будет содержать "
            f"{volume['directions']} направления, по {volume['actions']} действия в каждом."
        )

    st.divider()
    nav_back, nav_next = st.columns([1, 1])
    with nav_back:
        if st.button("← Назад к загрузке"):
            goto_step(1)
    with nav_next:
        if st.button("Далее → Генерация", type="primary"):
            if not full_name.strip():
                st.warning("Укажи ФИО сотрудника.")
                return
            st.session_state["intent_raw"] = {
                "full_name": full_name, "role": role, "period": "Год",
                "direction": direction, "direction_note": direction_note,
                "expectations": [e1, e2, e3],
                "readiness": readiness, "readiness_option": readiness_option,
            }
            goto_step(3)


# ---------------------------------------------------------------------------
# Этап 3. Генерация ИПР
# ---------------------------------------------------------------------------


def page_generate() -> None:
    st.title("Генерация ИПР")
    if st.button("← Назад к экранам выбора"):
        goto_step(2)

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
    cols[1].metric("Готовность", raw.get("readiness_option", raw["readiness"]))
    cols[2].metric("Компетенций", len(profile.competencies))

    if st.button("Сформировать ИПР", type="primary"):
        intent = EmployeeIntent(
            full_name=raw["full_name"], role=raw["role"], period="Год",
            direction=raw["direction"], direction_note=raw["direction_note"],
            expectations=[e for e in raw["expectations"] if e.strip()],
            readiness=raw["readiness"],
        )
        with st.spinner("DeepSeek собирает план…"):
            generator = build_generator(api_key, folder_id, model_id)
            result = generator.generate(profile, intent)

        if not result.successful:
            st.error(result.error_message or "Не удалось сгенерировать план.")
            return

        st.session_state["ipr_data"] = result.data
        st.success(f"План готов. Токенов: {result.tokens_consumed}", icon="✅")
        if result.error_message:
            st.warning(result.error_message)

    if "ipr_data" in st.session_state:
        _show_result(st.session_state["ipr_data"], raw)


def _show_result(data: dict, raw: dict) -> None:
    st.divider()

    docx_bytes = render_docx(data)
    _, _, step_days = checkpoint_spec("Год")
    ics_bytes = render_ics(data, start_date=datetime.now(), step_days=step_days)
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

    st.subheader("Превью полного плана")
    st.caption("Это тот же текст, что и в DOCX. Календарь — в файле ICS.")

    # Интерактивное оглавление: ссылки ведут на якоря разделов ниже
    st.markdown("**Содержание**")
    toc = "\n".join(
        f"{idx}. [{name}](#sec-{idx})" for idx, (_, name) in enumerate(SECTIONS, start=1)
    )
    st.markdown(toc)
    st.divider()

    st.markdown(f"### {PREAMBLE_TITLE}")
    for para in build_preamble():
        st.write(para)

    # 1. Контекст развития
    s2 = data.get("section2", {})
    if s2:
        st.header(section_title("section2"), anchor="sec-1")
        if s2.get("narrative"):
            st.write(s2["narrative"])
        if s2.get("strengths"):
            st.markdown("**Сильные стороны, на которые опирается план:**")
            for it in s2["strengths"]:
                st.markdown(f"- {it.get('name','')} — {it.get('score','')}. {it.get('note','')}")
        if s2.get("growth_zones"):
            st.markdown("**Зоны роста, на которые направлен план:**")
            for it in s2["growth_zones"]:
                st.markdown(f"- {it.get('name','')} — {it.get('score','')}. {it.get('note','')}")
        t = s2.get("table", {})
        if t:
            for label, key in [("Твои ожидания", "stated_expectations"),
                               ("Фокус плана", "focus"),
                               ("Ключевой сдвиг", "key_shift")]:
                if t.get(key):
                    st.markdown(f"- **{label}:** {t[key]}")
        if s2.get("source_note"):
            st.caption(s2["source_note"])

    # 2. Зоны роста
    s3 = data.get("section3", {})
    if s3.get("zones"):
        st.header(section_title("section3"), anchor="sec-2")
        if s3.get("intro"):
            st.write(s3["intro"])
        for i, z in enumerate(s3["zones"], 1):
            title = str(z.get("title", "")).strip()
            score = str(z.get("score", "")).strip()
            heading = title if (not score or score in title) else f"{title} — {score}"
            st.markdown(f"**2.{i}. {heading}**")
            st.write(z.get("text", ""))

    # 3. Направления развития
    s4 = data.get("section4", {})
    if s4.get("directions"):
        st.header(section_title("section4"), anchor="sec-3")
        if s4.get("intro"):
            st.write(s4["intro"])
        for d in s4["directions"]:
            st.markdown(f"#### {d.get('num','')} {d.get('title','')}")
            for label, key in [("Компетенция", "competency"), ("Основание (360°)", "basis_360"),
                               ("Опора на сильное", "strength_support"), ("Цель развития", "goal")]:
                if d.get(key):
                    st.markdown(f"**{label}:** {d[key]}")
            if d.get("workplace"):
                st.markdown("**На рабочем месте:**")
                for a in d["workplace"]:
                    st.markdown(f"- {a}")
            if d.get("projects"):
                st.markdown("**Развивающие проекты:**")
                for a in d["projects"]:
                    st.markdown(f"- {a}")
            if d.get("comfort_zone"):
                st.markdown(f"**Выход из зоны комфорта:** {d['comfort_zone']}")
            if d.get("criteria"):
                st.markdown("**Критерии достижения:**")
                for a in d["criteria"]:
                    st.markdown(f"- {a}")
            ras = d.get("risk_and_support", {})
            if ras:
                parts = [ras.get("risk", ""), ras.get("how_to_manage", ""), ras.get("psychological_support", "")]
                st.markdown("**Риски и поддержка:** " + " ".join(p for p in parts if p))

    # 4. Обучение
    s5 = data.get("section5", {})
    if s5:
        st.header(section_title("section5"), anchor="sec-4")
        if s5.get("intro"):
            st.write(s5["intro"])
        for grp in s5.get("books", []):
            if grp.get("theme"):
                st.markdown(f"**{grp['theme']}**")
            for bk in grp.get("items", []):
                st.markdown(f"- {bk.get('author','')} «{bk.get('title','')}» — {bk.get('note','')}")
        if s5.get("internal_formats"):
            st.markdown("**Форматы внутреннего обучения:**")
            for f in s5["internal_formats"]:
                st.markdown(f"- {f}")
        if s5.get("source_note"):
            st.caption(s5["source_note"])

    # 5. Точки контроля
    s7 = data.get("section7", {})
    if s7.get("questions"):
        st.header(section_title("section7") + " — вопросы для саморефлексии", anchor="sec-5")
        if s7.get("note"):
            st.caption(s7["note"])
        for q in s7["questions"]:
            st.markdown(f"- **{q.get('quarter','')}** ({q.get('intensity','')}): {q.get('question','')}")

    # 6. Согласование
    if data.get("section8"):
        st.header(section_title("section8"), anchor="sec-6")
        st.write(data["section8"])


# ---------------------------------------------------------------------------
# Маршрутизация
# ---------------------------------------------------------------------------


def main() -> None:
    render_sidebar()
    step = st.session_state.get("step", 1)
    if step == 1:
        page_upload()
    elif step == 2:
        page_choices()
    else:
        page_generate()


if __name__ == "__main__":
    main()
