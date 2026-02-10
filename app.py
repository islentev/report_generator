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
import re

# --- 1. НАСТРОЙКА ---
st.set_page_config(page_title="Юридический Генератор", layout="wide")

if 'report_buffer' not in st.session_state: st.session_state['report_buffer'] = None
if 'title_info' not in st.session_state: st.session_state['title_info'] = None

# --- 2. ПОДКЛЮЧЕНИЕ ---
try:
    client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"].strip().strip('"'), base_url="https://api.deepseek.com/v1")
    gc = gspread.authorize(Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']))
    SHEET_ID = st.secrets["SHEET_ID"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфига: {e}"); st.stop()

# --- 3. ФУНКЦИЯ СОЗДАНИЯ DOCX ---
def create_report_docx(report_content, title_data, requirements_list):
    doc = Document()
    
    # 1. ТИТУЛЬНЫЙ ЛИСТ
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_auth.add_run(f"УТВЕРЖДАЮ\n{title_data.get('company', '')}\n\n________________ / {title_data.get('director', '')}\n«___» _________ 2025 г.").font.size = Pt(11)

    for _ in range(7): doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("ИНФОРМАЦИОННЫЙ ОТЧЕТ\n").bold = True
    p_title.runs[-1].font.size = Pt(20)
    p_title.add_run(f"по Контракту № {title_data.get('contract_no', '')}\n").font.size = Pt(14)
    p_title.add_run(title_data.get('project_name', '')).italic = True

    doc.add_page_break()

    # 2. ОСНОВНОЙ ТЕКСТ ОТЧЕТА
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    for block in report_content.split('\n\n'):
        p = doc.add_paragraph()
        for part in block.split('**'):
            run = p.add_run(part.replace('*', ''))
            if part in block.split('**')[1::2]: run.bold = True
    
    # --- ПОДПИСЬ ДИРЕКТОРА СРАЗУ ПОСЛЕ ОТЧЕТА ---
    p_sign = doc.add_paragraph()
    p_sign.add_run(f"\n\nДиректор {title_data.get('company', '')}  _________________ / {title_data.get('director', '')}")

    doc.add_page_break()

    # 3. ОТДЕЛЬНАЯ СТРАНИЦА: ТРЕБОВАНИЯ К ДОКУМЕНТАЦИИ
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph("Перечень документов, обязательных к предоставлению Заказчику согласно условиям ТЗ:")
    p_req = doc.add_paragraph()
    p_req.add_run(requirements_list)

    return doc

# --- 4. ОСНОВНОЙ БЛОК ЛОГИКИ (ТЕПЕРЬ БЕЗ ЛИШНИХ ОТСТУПОВ) ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD: st.stop()

uploaded_file = st.file_uploader("Загрузите контракт (DOCX)", type="docx")

if uploaded_file:
    if 'last_file' not in st.session_state or st.session_state.last_file != uploaded_file.name:
        st.session_state.title_info = None
        st.session_state.report_buffer = None
        st.session_state.last_file = uploaded_file.name

    doc_obj = Document(uploaded_file)
    full_text = "\n".join([p.text for p in doc_obj.paragraphs])
    
    # 1. Распознавание реквизитов (первые 3к символов)
    if not st.session_state['title_info']:
        with st.spinner("Извлечение реквизитов из начала документа..."):
            res = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Найди Исполнителя, Директора, Номер и Предмет контракта в тексте: {full_text[:3000]}. Выдай JSON."}],
                response_format={ 'type': 'json_object' }
            )
            st.session_state['title_info'] = json.loads(res.choices[0].message.content)

    meta = st.session_state['title_info']
    st.info(f"Объект: {meta.get('project_name', 'Не определен')}")

    with st.form("main_form"):
        facts = st.text_area("Фактические детали выполнения (даты, количество и т.д.)")
        if st.form_submit_button("Сгенерировать отчет"):
            with st.spinner("Точечный анализ: Реквизиты (3к) + ТЗ..."):
                head_text = full_text[:3000] 
                
                # 2. Ищем ТЗ с конца документа
                text_upper = full_text.upper()
                tz_markers = ["ПРИЛОЖЕНИЕ № 1", "ТЕХНИЧЕСКОЕ ЗАДАНИЕ", "ОПИСАНИЕ ОБЪЕКТА ЗАКУПКИ"]
                tz_index = -1
                
                for marker in tz_markers:
                    found = text_upper.rfind(marker)
                    if found != -1 and found > tz_index:
                        tz_index = found
                
                clean_tz = full_text[tz_index:] if tz_index != -1 else full_text[-40000:]
    
                # 3. Написание отчета по ТЗ
                report_res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "Ты технический эксперт. Твоя задача — описать выполнение УСЛУГ из ТЗ. Пиши только про мероприятия, застройку, персонал и логистику. Галлюцинации запрещены."},
                        {"role": "user", "content": f"НАПИШИ ОТЧЕТ ПО ЭТОМУ ТЗ В ПРОШЕДШЕМ ВРЕМЕНИ: {clean_tz}. ФАКТЫ: {facts}"}
                    ]
                )

                # 4. Поиск требований к документации
                req_prompt = f"""Внимательно изучи текст ТЗ и выпиши ВСЕ документы, которые Исполнитель обязан предоставить по итогам работ (Акты, фотоотчеты, видео и т.д.). ТЕКСТ ТЗ: {clean_tz}"""
                
                req_res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": req_prompt}]
                )
                
                # 5. Создание документа
                doc_final = create_report_docx(
                    report_res.choices[0].message.content, 
                    meta, 
                    req_res.choices[0].message.content
                )
                
                buf = io.BytesIO()
                doc_final.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()

# Кнопка скачивания с номером контракта
if st.session_state['report_buffer']:
    c_no = re.sub(r'[\\/*?:"<>|]', "_", str(meta.get('contract_no', '')))
    st.download_button(f"📥 Скачать отчет № {c_no}", st.session_state['report_buffer'], f"отчет и № {c_no}.docx")
