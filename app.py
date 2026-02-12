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
    
    def format_name(full_name):
        if not full_name: return ""
        parts = full_name.split()
        if len(parts) >= 3:
            return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
        return full_name

    # Подготовка данных
    contract_no = title_data.get('contract_no', '________________')
    contract_date = title_data.get('contract_date', '___')
    ikz = title_data.get('ikz', '________________')
    
    raw_name = title_data.get('project_name', '')
    project_name = raw_name[0].upper() + raw_name[1:] if raw_name else ""
    
    customer = title_data.get('customer', '')
    customer_signer = title_data.get('customer_signer', '________________')
    company = title_data.get('company', '')
    director = format_name(title_data.get('director', ''))

    # Стиль по умолчанию: Times New Roman 12
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # --- ТИТУЛЬНЫЙ ЛИСТ ---
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # ЖИРНЫМ: Заголовок и Контракт (первая строка)
    run1 = p_top.add_run("Информационно-аналитический отчет об исполнении условий\n")
    run1.bold = True
    run2 = p_top.add_run(f"Контракта № {contract_no} от «{contract_date}» 2025 г.\n")
    run2.bold = True
    p_top.add_run(f"Идентификационный код закупки: {ikz}.")

    for _ in range(3): doc.add_paragraph()

    p_tom = doc.add_paragraph()
    p_tom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tom.add_run("ТОМ I").bold = True

    # Блоки с ЖИРНЫМИ заголовками
    labels_values = [
        ("Наименование предмета КОНТРАКТА :", project_name),
        ("Заказчик:", customer),
        ("Исполнитель:", company)
    ]
    
    for label, value in labels_values:
        p_h = doc.add_paragraph()
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h.add_run(label).bold = True # ЖИРНЫМ заголовок
        
        p_v = doc.add_paragraph()
        p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(value).italic = True # КУРСИВОМ данные

    for _ in range(4): doc.add_paragraph()

    # --- ТАБЛИЦА ПОДПИСЕЙ ---
    table = doc.add_table(rows=2, cols=2)
    table.width = doc.sections[0].page_width
    
    # Заказчик (Жирный заголовок)
    p_l = table.rows[0].cells[0].paragraphs[0]
    p_l.add_run("Отчет принят Заказчиком").bold = True
    p_l.add_run(f"\n\n{customer_signer}\n\n_______________")
    
    # Исполнитель (Жирный заголовок, выравнивание ВЛЕВО)
    p_r = table.rows[0].cells[1].paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    p_r.add_run("Отчет передан Исполнителем").bold = True
    p_r.add_run(f"\n\nДиректор\n\n_______________ / {director}")
    
    # М.П.
    table.rows[1].cells[0].paragraphs[0].add_run("м.п.")
    table.rows[1].cells[1].paragraphs[0].add_run("м.п.")

    doc.add_page_break()

    # --- ТЕКСТ ОТЧЕТА (БЕЗ ПОДПИСЕЙ В КОНЦЕ) ---
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    for block in report_content.split('\n\n'):
        p = doc.add_paragraph()
        for part in block.split('**'):
            run = p.add_run(part.replace('*', ''))
            if part in block.split('**')[1::2]: run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

    doc.add_page_break()
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    # Здесь вставляем requirements_list аналогично блоку выше
    doc.add_paragraph(requirements_list)

    return doc
        
# --- 4. ОСНОВНОЙ БЛОК ЛОГИКИ ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD: st.stop()

uploaded_file = st.file_uploader("Загрузите контракт (DOCX)", type="docx")

if uploaded_file:
    # Если загружен новый файл — сбрасываем старые данные
    if 'last_file' not in st.session_state or st.session_state.last_file != uploaded_file.name:
        st.session_state.title_info = None
        st.session_state.report_buffer = None
        st.session_state.last_file = uploaded_file.name

    doc_obj = Document(uploaded_file)
    full_text = "\n".join([p.text for p in doc_obj.paragraphs])
    
    # 1. Извлечение реквизитов (строго один раз для файла)
    if not st.session_state.get('title_info'):
        with st.spinner("Анализ титульных данных и ИКЗ..."):
            res = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"""
                    Извлеки данные из начала и конца контракта для титульного листа.
                    ВАЖНО: 
                    1. project_name (Наименование предмета) пиши С БОЛЬШОЙ БУКВЫ.
                    2. Найди подписанта со стороны ЗАКАЗЧИКА (обычно в конце или начале): его должность и ФИО.
                    
                    Формат ответа — JSON с ключами:
                    - contract_no (номер контракта, например "39/25/ГК")
                    - contract_date (дата)
                    - ikz (36 цифр)
                    - project_name (Предмет контракта, С БОЛЬШОЙ БУКВЫ)
                    - customer (Полное название Заказчика)
                    - customer_signer (Должность и ФИО подписанта Заказчика, например: "Заместитель председателя Комитета Иванов И.И.")
                    - company (Название Исполнителя)
                    - director (ФИО директора Исполнителя)
                    
                    Текст для анализа: {full_text[:5000]} {full_text[-3000:]} 
                """}],
                response_format={ 'type': 'json_object' }
            )
            st.session_state['title_info'] = json.loads(res.choices[0].message.content)

    meta = st.session_state['title_info']
    st.info(f"Объект: {meta.get('project_name', 'Не определен')}")

    with st.form("main_form"):
        facts = st.text_area("Фактические детали выполнения (даты, количество и т.д.)")
        if st.form_submit_button("Сгенерировать отчет"):
            with st.spinner("Генерация отчета по пунктам ТЗ..."):
                # 1. Находим начало ТЗ
                text_upper = full_text.upper()
                tz_markers = ["ПРИЛОЖЕНИЕ № 1", "ТЕХНИЧЕСКОЕ ЗАДАНИЕ", "ОПИСАНИЕ ОБЪЕКТА ЗАКУПКИ"]
                tz_index = -1
                for marker in tz_markers:
                    found = text_upper.find(marker)
                    if found != -1:
                        tz_index = found
                        break
                
                if tz_index == -1:
                    tz_index = 0 
                
                # 2. Находим КОНЕЦ ТЗ (чтобы не захватить лишние приложения №2, №3 и т.д.)
                end_markers = ["ПРИЛОЖЕНИЕ № 2", "ПРИЛОЖЕНИЕ № 3", "РАСЧЕТ СТОИМОСТИ", "ПОДПИСИ СТОРОН"]
                tz_end_index = len(full_text)
                for marker in end_markers:
                    # Ищем маркер конца только ПОСЛЕ начала ТЗ
                    found_end = text_upper.find(marker, tz_index + 100)
                    if found_end != -1:
                        tz_end_index = found_end
                        break
                
                # --- ФОРМИРУЕМ БЛОКИ ---
                
                # Блок 1: Для титульника (1000 знаков с начала и 1000 с конца)
                context_title = full_text[:1000] + "\n[...]\n" + full_text[-1000:]
                
                # Блок 2 и 3: Чистое ТЗ (от начала ТЗ до следующего приложения или конца)
                context_tz_full = full_text[tz_index : tz_end_index]
                
                # --- ЗАПРОСЫ К ИИ ---
                
                # 1. Данные титульника
                res_title = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Извлеки JSON: contract_no (номер из ПЕРВОЙ строки), contract_date, ikz, project_name (предмет), customer, customer_signer (должность и ФИО из конца), company, director. Текст: {context_title}"}],
                    response_format={ 'type': 'json_object' }
                )
                title_info = json.loads(res_title.choices[0].message.content)
                
                # 2. Текст отчета (на базе полного ТЗ)
                res_report = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Напиши подробный отчет о выполненных работах, используя ВСЁ это ТЗ: {context_tz_full}"}]
                )
                report_text = res_report.choices[0].message.content
                
                # 3. Требования (на базе полного ТЗ + хвост контракта для поиска условий приемки)
                context_docs = context_tz_full + "\n" + full_text[-3000:]
                res_req = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Выпиши список отчетных документов (акты, фото, видео) из этого текста: {context_docs}"}]
                )
                requirements_text = res_req.choices[0].message.content
                
                # --- СОХРАНЕНИЕ ---
                # Вызываем вашу функцию создания документа (убедитесь, что она определена выше)
                doc_final = create_report_docx(report_text, title_info, requirements_text)
                
                buf = io.BytesIO()
                doc_final.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()
                
                # Кнопка скачивания
                if st.session_state.get('report_buffer'):
                    # Очищаем номер контракта от запрещенных символов для имени файла
                    raw_no = title_info.get('contract_no', 'бн')
                    c_no = re.sub(r'[\\/*?:"<>|]', "_", str(raw_no))
                    st.download_button(
                        label=f"📥 Скачать отчет № {c_no}",
                        data=st.session_state['report_buffer'],
                        file_name=f"отчет и № {c_no}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
