import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_COLOR_INDEX
from openai import OpenAI
import io
import json
import re

# --- 1. ФУНКЦИИ ИЗВЛЕЧЕНИЯ ТЕКСТА ---

def get_contract_start_text(file):
    """Берем только начало контракта (первая страница), где указаны стороны и подписанты"""
    doc = Document(file)
    full_text = []
    # Сначала таблицы (там часто реквизиты в шапке)
    for table in doc.tables:
        for row in table.rows:
            full_text.append(" ".join(cell.text.strip() for cell in row.cells))
    # Затем параграфы первой страницы
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            if re.match(r"^2\.", txt): # Как и договаривались, не идем глубже 2-го раздела
                break
            full_text.append(txt)
    return "\n".join(full_text)[:2000]

def get_text_from_file(file):
    """Извлекает весь текст из ТЗ для генерации отчета"""
    doc = Document(file)
    content = []
    for p in doc.paragraphs:
        if p.text.strip(): content.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            content.append(" ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(content)

# --- 2. ЛОГИКА ТРАНСФОРМАЦИИ (УМНАЯ СБОРКА) ---

def smart_generate_step_strict(client, section_text, requirements_text):
    """Цепочка: Черновик -> Проверка -> Финальное исправление"""
    
    # Базовые правила (используем их во всех запросах, чтобы ИИ не забывал стиль)
    core_rules = """
    1. НУМЕРАЦИЯ: Сохраняй нумерацию пунктов (1.1, 1.2...) строго как в ТЗ.
    2. ЗАГОЛОВКИ: Пиши в НАСТОЯЩЕМ времени.
    3. ТЕКСТ: Пиши строго в ПРОШЕДШЕМ времени ('оказано', 'выполнено').
    4. ЗАПРЕТ: Удали слова 'должен', 'обязан', 'необходимо', 'будет'. Только свершившийся факт.
    5. ПОЛНОТА: Все цифры, объемы и характеристики из ТЗ должны быть перенесены в отчет.
    """

    # ШАГ 1: ГЕНЕРАЦИЯ
    res = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": f"Ты юридический редактор. Перепиши ТЗ в Отчет. Правила: {core_rules} Доп. требования: {requirements_text}"},
            {"role": "user", "content": f"ТРАНСФОРМИРУЙ В ОТЧЕТ:\n\n{section_text}"}
        ],
        temperature=0.1
    )
    draft = res.choices[0].message.content

    # ШАГ 2: САМОПРОВЕРКА
    verify_res = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "Ты контролер. Найди упущенные цифры/характеристики в Отчете, сравнив его с ТЗ."},
            {"role": "user", "content": f"ТЗ: {section_text}\n\nОТЧЕТ: {draft}\n\nВыдай ответ: 'ОШИБОК: 0' или список пропусков."}
        ],
        temperature=0
    )
    v_report = verify_res.choices[0].message.content

    # ШАГ 3: ИСПРАВЛЕНИЕ (если есть ошибки)
    if "ОШИБОК: 0" not in v_report:
        final_res = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": f"Исправь отчет, сохранив стиль: {core_rules}"},
                {"role": "user", "content": f"ТЗ: {section_text}\nОшибки: {v_report}\nИсправь этот текст: {draft}"}
            ],
            temperature=0.1
        )
        return final_res.choices[0].message.content
    
    return draft

# --- 3. ОФОРМЛЕНИЕ И СБОРКА DOCX ---

def apply_yellow_highlight(doc):
    keywords = ["Акт", "Фотоотчет", "Ведомость", "Скриншот", "Смета", "Резюме", "USB", "Флеш-накопитель", "Ссылка"]
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for word in keywords:
                if word.lower() in run.text.lower():
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def create_final_report(t_info, report_body, req_body):
    # Создаем титульник (центрирование внутри функции build_title_page)
    from main_logic import build_title_page # Предполагаем, что она в этом же файле
    doc = build_title_page(t_info)
    doc.add_page_break()

    # Основной заголовок отчета (По центру)
    p_name = str(t_info.get('project_name', '')).strip()
    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head.add_run(f"Отчет об оказании услуг по {p_name}").bold = True

    # Тело отчета (По ширине)
    for line in report_body.split('\n'):
        line = line.strip()
        if not line: continue
        para = doc.add_paragraph()
        run = para.add_run(line)
        if re.match(r"^\d+\.", line): # Жирный для пунктов
            run.bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # СТРОГО ПО ШИРИНЕ

    # Требования к документации
    if req_body:
        doc.add_page_break()
        doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
        p_req = doc.add_paragraph(req_body)
        p_req.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    apply_yellow_highlight(doc)
    return doc

# --- 4. ИНТЕРФЕЙС STREAMLIT ---

st.set_page_config(page_title="Генератор Отчетов 3.0", layout="wide")

# (Блок авторизации и сброса остается прежним)

col1, col2, col3 = st.columns(3)

with col1:
    st.header("📄 1. Реквизиты")
    # Логика извлечения JSON из начала контракта...

with col2:
    st.header("📝 2. Техзадание")
    f_tz = st.file_uploader("Загрузить ТЗ", type="docx")
    if f_tz: st.session_state.raw_tz = get_text_from_file(f_tz)

with col3:
    st.header("📋 3. Требования")
    # Извлечение требований...

# ГЛАВНАЯ КНОПКА
if st.button("🚀 ЗАПУСТИТЬ УМНУЮ СБОРКУ", use_container_width=True):
    if "raw_tz" in st.session_state and "t_info" in st.session_state:
        client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
        
        # Разбивка на пункты
        steps = [s.strip() for s in re.split(r'\n(?=\d+\.)', st.session_state.raw_tz) if s.strip()]
        
        final_text = ""
        progress_bar = st.progress(0) # Инициализируем
        
        for i, step in enumerate(steps):
            st.write(f"⌛ Обработка пункта {i+1} из {len(steps)}...")
            chunk = smart_generate_step_strict(client, step, st.session_state.get('raw_requirements', ''))
            final_text += chunk + "\n\n"
            # ОБНОВЛЯЕМ ПРОГРЕСС
            progress_bar.progress((i + 1) / len(steps))
        
        # Финальная сборка файла
        full_doc = create_final_report(st.session_state.t_info, final_text, st.session_state.get('raw_requirements', ''))
        buf = io.BytesIO()
        full_doc.save(buf)
        st.session_state.smart_file = buf.getvalue()
        st.success("✅ Сборка завершена!")

if "smart_file" in st.session_state:
    st.download_button("📥 СКАЧАТЬ ОТЧЕТ", st.session_state.smart_file, "Report_Final.docx", use_container_width=True)
