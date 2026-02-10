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

# --- 1. НАСТРОЙКА ---
st.set_page_config(page_title="Генератор Отчетов PRO", layout="wide")

if 'report_buffer' not in st.session_state: st.session_state['report_buffer'] = None
if 'title_info' not in st.session_state: st.session_state['title_info'] = None

# --- 2. ПОДКЛЮЧЕНИЕ (Удалены лишние детали для краткости) ---
try:
    gcp_info = dict(st.secrets["gcp_service_account"])
    gcp_info["private_key"] = gcp_info["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(gcp_info, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    gc = gspread.authorize(creds)
    client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"].strip().strip('"'), base_url="https://api.deepseek.com/v1")
    SHEET_ID = st.secrets["SHEET_ID"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфига: {e}"); st.stop()

# --- 3. ФУНКЦИИ ---

def create_report_docx(report_content, title_data, requirements_text):
    doc = Document()
    
    # 1. ТИТУЛЬНЫЙ ЛИСТ
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    auth_text = f"УТВЕРЖДАЮ\n{title_data.get('company', '')}\n\n________________ / {title_data.get('director', '')}\n«___» _________ 2025 г."
    p_auth.add_run(auth_text).font.size = Pt(11)

    for _ in range(7): doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("ИНФОРМАЦИОННЫЙ ОТЧЕТ\n").bold = True
    p_title.runs[-1].font.size = Pt(20)
    p_title.add_run(f"по Контракту № {title_data.get('contract_no', '')}\n").font.size = Pt(14)
    p_title.add_run(title_data.get('project_name', '')).italic = True

    doc.add_page_break()

    # 2. ОСНОВНОЙ ТЕКСТ ОТЧЕТА (по ТЗ)
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ РАБОТ', level=1)
    for block in report_content.split('\n\n'):
        p = doc.add_paragraph()
        parts = block.split('**')
        for i, part in enumerate(parts):
            run = p.add_run(part.replace('*', ''))
            if i % 2 != 0: run.bold = True
            
    doc.add_page_break()

    # 3. НОВАЯ СТРАНИЦА: ТРЕБОВАНИЯ К ДОКУМЕНТАЦИИ
    doc.add_heading('ТРЕБОВАНИЯ К ОТЧЕТНОЙ ДОКУМЕНТАЦИИ', level=1)
    p_req = doc.add_paragraph()
    p_req.add_run(requirements_text)

    return doc

# --- 4. ИНТЕРФЕЙС ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD: st.stop()

sheet = gc.open_by_key(SHEET_ID).sheet1
df_etalons = pd.DataFrame(sheet.get_all_records())

uploaded_file = st.file_uploader("Загрузите контракт", type="docx")

if uploaded_file:
    # Очистка кэша при новом файле
    if 'last_file' not in st.session_state or st.session_state.last_file != uploaded_file.name:
        st.session_state.title_info = None
        st.session_state.last_file = uploaded_file.name

    doc_obj = Document(uploaded_file)
    contract_text = "\n".join([p.text for p in doc_obj.paragraphs])
    
    if not st.session_state['title_info']:
        with st.spinner("Анализ реквизитов..."):
            extraction_prompt = f"Найди в тексте: {contract_text[:8000]}. Выдай JSON: {{'company','director','contract_no','contract_date','project_name','type'}}"
            res_meta = client_ai.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": extraction_prompt}], response_format={ 'type': 'json_object' })
            st.session_state['title_info'] = json.loads(res_meta.choices[0].message.content)

    meta = st.session_state['title_info']
    st.info(f"Контракт: {meta['contract_no']} | Исполнитель: {meta['company']}")

    with st.form("data"):
        q1 = st.text_input("Кол-во участников", value="100")
        facts = st.text_area("Доп. детали")
        if st.form_submit_button("Сгенерировать"):
            with st.spinner("Анализ ТЗ и формирование требований..."):
                # Шаг 1: Извлекаем требования к отчетности из ТЗ
                req_res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди в этом тексте раздел 'Требования к отчетности' или 'Отчетная документация'. Выпиши список всех документов, которые Исполнитель должен предоставить. ТЕКСТ: {contract_text[-20000:]}"}]
                )
                requirements_found = req_res.choices[0].message.content

                # Шаг 2: Генерируем отчет на основании ТЗ
                sys_msg = f"Ты юрист. Твоя задача: написать отчет, строго следуя Техническому заданию (ТЗ) из контракта. Опиши выполнение каждого пункта ТЗ в прошедшем времени."
                user_msg = f"КОНТРАКТ И ТЗ: {contract_text}\nДАННЫЕ: Участников {q1}, Факты: {facts}"
                
                # Ограничиваем передаваемый текст для API (берём только ТЗ, если оно большое)
                report_res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":sys_msg}, {"role":"user","content":user_msg[:15000]}]
                )
                
                doc_final = create_report_docx(report_res.choices[0].message.content, meta, requirements_found)
                buf = io.BytesIO()
                doc_final.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()

if st.session_state['report_buffer']:
    st.download_button("📥 Скачать готовый Отчет", st.session_state['report_buffer'], "Report.docx")
