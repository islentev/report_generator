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
st.set_page_config(page_title="Универсальный Генератор", layout="wide")

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

# --- 3. УНИВЕРСАЛЬНАЯ ФУНКЦИЯ СОЗДАНИЯ DOCX ---
def create_report_docx(report_content, title_data, requirements_list):
    doc = Document()
    
    # ТИТУЛЬНЫЙ ЛИСТ
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

    # СТРАНИЦА 2: ОТЧЕТ ПО ТЗ
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    for block in report_content.split('\n\n'):
        p = doc.add_paragraph()
        for part in block.split('**'):
            run = p.add_run(part.replace('*', ''))
            if part in block.split('**')[1::2]: run.bold = True
            
    doc.add_page_break()

    # СТРАНИЦА 3: ТРЕБОВАНИЯ К ДОКУМЕНТАЦИИ (ЧЕК-ЛИСТ)
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph("Ниже представлен перечень документов, обязательных к предоставлению Заказчику согласно условиям Контракта:")
    p_req = doc.add_paragraph()
    p_req.add_run(requirements_list)
    
    # ФИНАЛЬНАЯ ПОДПИСЬ
    p_sign = doc.add_paragraph()
    p_sign.add_run(f"\n\nДиректор {title_data.get('company', '')}  _________________ / {title_data.get('director', '')}")

    return doc

# --- 4. ОСНОВНОЙ ПРОЦЕСС ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD: st.stop()

uploaded_file = st.file_uploader("Загрузите контракт", type="docx")

if uploaded_file:
    # Очистка памяти при смене файла
    if 'last_file' not in st.session_state or st.session_state.last_file != uploaded_file.name:
        st.session_state.title_info = None
        st.session_state.last_file = uploaded_file.name

    doc_obj = Document(uploaded_file)
    full_text = "\n".join([p.text for p in doc_obj.paragraphs])
    
    # 1. Распознавание реквизитов (берем начало файла)
    if not st.session_state['title_info']:
        with st.spinner("Анализ сторон и реквизитов..."):
            res = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Найди Исполнителя, Директора, Номер и Предмет контракта в тексте: {full_text[:10000]}. Выдай JSON."}],
                response_format={ 'type': 'json_object' }
            )
            st.session_state['title_info'] = json.loads(res.choices[0].message.content)

    meta = st.session_state['title_info']
    st.info(f"Объект: {meta.get('project_name', 'Не определен')}")

    with st.form("main_form"):
        facts = st.text_area("Фактические детали выполнения (даты, количество и т.д.)")
        if st.form_submit_button("Сформировать универсальный отчет"):
            with st.spinner("Точечный анализ: Реквизиты + ТЗ..."):
            # 1. Выделяем фрагменты текста, чтобы не «кормить» ИИ лишним
            head_text = full_text[:10000]  # Первые страницы для реквизитов
            
            # Ищем начало ТЗ по ключевым словам
            tz_start_keywords = ["Приложение №", "Техническое задание", "Описание объекта закупки", "ТЕХНИЧЕСКОЕ ЗАДАНИЕ"]
            tz_index = -1
            for kw in tz_start_keywords:
                found = full_text.find(kw)
                if found != -1 and (tz_index == -1 or found < tz_index):
                    tz_index = found
            
            # Если ТЗ найдено, берем текст от его начала и до конца
            if tz_index != -1:
                clean_tz_text = full_text[tz_index:]
            else:
                clean_tz_text = full_text[-30000:] # Резерв: берем хвост, если ключевики не сработали
            
            # 2. Извлекаем требования к документам из этого же куска ТЗ
            req_prompt = f"Найди в ТЗ требования к отчетности. Выпиши только список документов. ТЕКСТ: {clean_tz_text[:15000]}"
            req_res = client_ai.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": req_prompt}])
            req_list = req_res.choices[0].message.content

            # 3. Пишем отчет СТРОГО ПО ПУНКТАМ ТЗ
            report_prompt = f"""Используя фрагмент ТЗ ниже, напиши отчет. 
            Действуй строго по пунктам ТЗ. Описывай только то, что есть в тексте ТЗ.
            
            ТРЕБОВАНИЯ:
            - Замени будущее время на прошедшее ('организовать' -> 'организовано').
            - Сохраняй все цифры, названия стран (Саудовская Аравия, Китай и т.д.) и объемы.
            - ЕСЛИ ИНФОРМАЦИИ НЕТ В ТЗ — НЕ ПРИДУМЫВАЙ ЕЁ.
            - Используй доп. факты: {facts}

            ФРАГМЕНТ ТЗ ДЛЯ РАБОТЫ:
            {clean_tz_text}"""
            
            report_res = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "system", "content": "Ты — юридический ассистент, который пишет отчеты строго по предоставленному ТЗ без галлюцинаций."},
                          {"role": "user", "content": report_prompt}]
            )
            
            # Формируем итоговый DOCX
            doc_final = create_report_docx(report_res.choices[0].message.content, meta, req_list)
            buf = io.BytesIO()
            doc_final.save(buf)
            st.session_state['report_buffer'] = buf.getvalue()

if st.session_state['report_buffer']:
    st.download_button("📥 Скачать Отчет", st.session_state['report_buffer'], "Report_Universal.docx")


