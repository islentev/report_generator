import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json

# --- 1. НАСТРОЙКА И ПАМЯТЬ ---
st.set_page_config(page_title="Генератор Отчетов PRO", layout="wide")

if 'report_buffer' not in st.session_state:
    st.session_state['report_buffer'] = None
if 'title_info' not in st.session_state:
    st.session_state['title_info'] = None

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
        cells = [cell.strip() for line in line.split('|') if (cell := line.strip())]
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

    # ТЕКСТ ОТЧЕТА
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

# Загрузка базы эталонов
try:
    sheet = gc.open_by_key(SHEET_ID).sheet1
    df_etalons = pd.DataFrame(sheet.get_all_records())
except Exception as e:
    st.error(f"Ошибка Google Таблицы: {e}")
    st.stop()

st.title("⚖️ Юридический генератор отчетов")
uploaded_file = st.file_uploader("Загрузите файл Контракта (docx)", type="docx")

if uploaded_file:
    contract_text = "\n".join([p.text for p in Document(uploaded_file).paragraphs])
    
    # ЭТАП 1: Распознавание реквизитов
    if not st.session_state['title_info']:
        with st.spinner("Извлекаю реквизиты из контракта..."):
            all_types = df_etalons["Тип проекта"].tolist()
            extraction_prompt = f"""Анализируй контракт: {contract_text[:5000]}
            Выдай ответ строго в JSON:
            {{ "company": "Название компании", "director": "ФИО директора", "contract_no": "№", "contract_date": "дата", "project_name": "предмет", "type": "тип из списка {all_types}" }}"""
            
            res_meta = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": extraction_prompt}],
                response_format={ 'type': 'json_object' }
            )
            st.session_state['title_info'] = json.loads(res_meta.choices[0].message.content)

    meta = st.session_state['title_info']
    st.success(f"Распознано: {meta['company']} | {meta['director']}")

    with st.form("data_form"):
        col1, col2 = st.columns(2)
        q1 = col1.text_input("Кол-во участников", value="100")
        q2 = col2.text_input("Письмо согласования", placeholder="№123 от 01.12.25")
        facts = st.text_area("Доп. детали реализации (меню, даты, адреса)")
        submitted = st.form_submit_button("🔥 Сгенерировать отчет")

    if submitted:
        with st.spinner("Генерация документа..."):
            try:
                # Поиск формы в контракте
                search_prompt = f"Найди в тексте Приложение с формой или образцом отчета: {contract_text[-15000:]}. Если есть - опиши её структуру. Если нет - напиши 'НЕТ'."
                form_check = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": search_prompt}]
                )
                contract_form = form_check.choices[0].message.content

                # Определение структуры (используем "Техническое описание" вместо "Эталонная структура")
                if "НЕТ" not in contract_form.upper():
                    st.write("✅ Используется форма из приложения к контракту.")
                    struct_instr = f"Строго следуй форме из контракта: {contract_form}"
                else:
                    selected_row = df_etalons[df_etalons["Тип проекта"] == meta['type']].iloc[0]
                    # ПОПРАВКА ТУТ: берем данные из вашего столбца "Техническое описание"
                    struct_instr = f"Используй эталонную структуру из технического описания: {selected_row['Техническое описание']}"

                # Генерация текста
                sys_msg = f"Ты опытный юрист. Напиши отчет в прошедшем времени. {struct_instr}. Отрази все требования ТЗ как выполненные."
                user_msg = f"КОНТРАКТ: {contract_text[:8000]}\nУчастники: {q1}, Письмо: {q2}, Факты: {facts}"
                
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":sys_msg}, {"role":"user","content":user_msg}]
                )
                
                # Сохранение
                final_doc = create_report_docx(res.choices[0].message.content, meta)
                buf = io.BytesIO()
                final_doc.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()

            except Exception as e:
                st.error(f"Ошибка: {e}")

# --- 5. ВЫВОД КНОПКИ СКАЧИВАНИЯ (ИСПРАВЛЕННЫЙ) ---
if st.session_state.get('report_buffer') is not None:
    st.divider()
    st.subheader("📥 Результат")
    
    # ПРОВЕРКА: Берем номер контракта из meta, если он есть, иначе пишем "final"
    # Это предотвращает ошибку NameError / KeyError
    if st.session_state.get('title_info'):
        contract_suffix = st.session_state['title_info'].get('contract_no', 'final')
    else:
        contract_suffix = 'final'
        
    # Очищаем номер контракта от символов, которые запрещены в именах файлов
    clean_name = str(contract_suffix).replace("/", "_").replace("\\", "_")
    
    st.download_button(
        label="Скачать готовый Отчет (.docx)",
        data=st.session_state['report_buffer'],
        file_name=f"Report_{clean_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
