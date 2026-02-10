import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import re

# --- 1. НАСТРОЙКА СТРАНИЦЫ ---
st.set_page_config(page_title="Генератор Отчетов PRO", layout="wide")

# --- 2. ПОДКЛЮЧЕНИЕ СЕКРЕТОВ ---
try:
    gcp_info = dict(st.secrets["gcp_service_account"])
    gcp_info["private_key"] = gcp_info["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(gcp_info, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    gc = gspread.authorize(creds)
    
    DEEPSEEK_KEY = st.secrets["DEEPSEEK_API_KEY"].strip().strip('"')
    client_ai = OpenAI(api_key=DEEPSEEK_KEY, base_url="https://api.deepseek.com/v1")
    
    SHEET_ID = st.secrets["SHEET_ID"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфигурации: {e}")
    st.stop()

# --- 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def add_table_from_markdown(doc, markdown_text):
    lines = [line.strip() for line in markdown_text.split('\n') if '|' in line]
    if len(lines) < 3: return
    headers = [cell.strip() for cell in lines[0].split('|') if cell.strip()]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers): hdr_cells[i].text = h
    for line in lines[2:]:
        cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        if len(cells) >= len(headers):
            row_cells = table.add_row().cells
            for i in range(len(headers)): row_cells[i].text = cells[i]

def create_report_docx(report_content, title_data):
    doc = Document()
    # ТИТУЛЬНЫЙ ЛИСТ
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    auth_text = f"""УТВЕРЖДАЮ
{title_data.get('company', 'Исполнитель')}

________________ / {title_data.get('director', '_________')}
«___» _________ 2025 г."""
    run_auth = p_auth.add_run(auth_text)
    run_auth.font.size = Pt(11)

    for _ in range(7): doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("ИНФОРМАЦИОННЫЙ ОТЧЕТ\n").bold = True
    p_title.runs[-1].font.size = Pt(20)
    
    sub_text = f"по исполнению Государственного контракта\n№ {title_data.get('contract_no', '_________')} от {title_data.get('contract_date', '_________')}\n\n"
    run_sub = p_title.add_run(sub_text)
    run_sub.font.size = Pt(14)
    p_title.add_run(title_data.get('project_name', '')).italic = True

    for _ in range(10): doc.add_paragraph()
    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.add_run("Москва, 2025 г.")
    doc.add_page_break()

    # ТЕКСТ
    for block in report_content.split('\n\n'):
        if '|' in block and '-|-' in block:
            add_table_from_markdown(doc, block)
        else:
            p = doc.add_paragraph()
            if block.strip().startswith('#'):
                p.add_run(block.replace('#', '').strip()).bold = True
            else:
                parts = block.split('**')
                for i, part in enumerate(parts):
                    run = p.add_run(part.replace('*', ''))
                    if i % 2 != 0: run.bold = True
    return doc

# --- 4. ОСНОВНОЙ ИНТЕРФЕЙС ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD:
    st.info("Введите пароль.")
    st.stop()

sheet = gc.open_by_key(SHEET_ID).sheet1
df_etalons = pd.DataFrame(sheet.get_all_records())

st.title("⚖️ Автоматический генератор отчетов")
uploaded_file = st.file_uploader("Загрузите Контракт", type="docx")

if uploaded_file:
    contract_text = "\n".join([p.text for p in Document(uploaded_file).paragraphs])
    
    # ЭТАП 1: ИИ ВЫТАСКИВАЕТ РЕКВИЗИТЫ И ОПРЕДЕЛЯЕТ ЭТАЛОН
    if 'title_info' not in st.session_state:
        with st.spinner("Анализирую реквизиты сторон..."):
            all_types = df_etalons["Тип проекта"].tolist()
            extraction_prompt = f"""Проанализируй начало контракта:
            {contract_text[:4000]}
            Выдай ответ строго в формате JSON:
            {{
              "company": "полное название Исполнителя",
              "director": "ФИО директора в им. падеже",
              "contract_no": "номер контракта",
              "contract_date": "дата",
              "project_name": "краткое название предмета контракта",
              "type": "один тип из списка {all_types}"
            }}"""
            
            res_meta = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": extraction_prompt}],
                response_format={ 'type': 'json_object' }
            )
            import json
            st.session_state['title_info'] = json.loads(res_meta.choices[0].message.content)

    meta = st.session_state['title_info']
    st.success(f"Распознано: {meta['company']} | {meta['director']}")

    # ФОРМА
    with st.form("data_form"):
        col1, col2 = st.columns(2)
        q1 = col1.text_input("Кол-во участников", placeholder="100")
        q2 = col2.text_input("Письмо согласования", placeholder="№1 от 01.12.25")
        facts = st.text_area("Доп. детали реализации")
        submitted = st.form_submit_button("Сформировать отчет")

    if submitted:
        with st.spinner("Анализирую контракт на наличие утвержденной формы отчета..."):
            try:
                # ШАГ 1: Ищем форму отчета в тексте контракта
                search_form_prompt = f"""Внимательно изучи текст контракта:
                {contract_text[-15000:]} 
                (особое внимание приложениям в конце).
                
                Задание:
                1. Есть ли в контракте утвержденная 'Форма отчета' или 'Образец отчета'?
                2. Если есть, выпиши структуру этой формы (заголовки разделов).
                3. Если нет, напиши 'ФОРМА НЕ НАЙДЕНА'.
                Ответь кратко."""
                
                form_detection = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": search_form_prompt}]
                )
                contract_form = form_detection.choices[0].message.content

                # ШАГ 2: Формируем финальный отчет
                if "ФОРМА НЕ НАЙДЕНА" not in contract_form:
                    st.info("📎 Обнаружена утвержденная форма отчета в контракте. Использую её.")
                    current_struct = f"ИСПОЛЬЗУЙ ЭТУ ФОРМУ ИЗ КОНТРАКТА: {contract_form}"
                else:
                    # Берем структуру из вашей Google Таблицы (как раньше)
                    selected_row = df_etalons[df_etalons["Тип проекта"] == meta['type']].iloc[0]
                    current_struct = f"Используй эталонную структуру: {selected_row.get('ЭТАЛОННАЯ СТРУКТУРА', 'Стандартная')}"

                sys_msg = f"""Ты — эксперт-юрист. Твоя задача — составить отчет.
                ПРАВИЛА:
                1. СТРУКТУРА: {current_struct}.
                2. ТИТУЛЬНЫЙ ЛИСТ: Данные уже извлечены ({meta['company']}, {meta['director']}).
                3. ТЕКСТ: Преобразуй требования ТЗ из будущего времени ('должен оказать') в прошедшее ('оказано/выполнено').
                4. ТАБЛИЦЫ: Если в форме есть таблицы — заполни их данными из ТЗ."""

                user_msg = f"КОНТРАКТ: {contract_text[:8000]}\nУЧАСТНИКИ: {q1}\nПИСЬМО: {q2}\nФАКТЫ: {facts}"
                
                # Финальный вызов ИИ для текста отчета
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":sys_msg}, {"role":"user","content":user_msg}]
                )
                
                # Создание документа (функция create_report_docx остается прежней)
                final_doc = create_report_docx(res.choices[0].message.content, meta)
                
                buf = io.BytesIO()
                final_doc.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()
                st.success("Отчет сформирован по форме из контракта!")

            except Exception as e:
                st.error(f"Ошибка: {e}")
