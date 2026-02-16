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

def format_fio_universal(raw_fio):
    if not raw_fio or len(raw_fio) < 5: return "________________"
    # Убираем возможный "мусор" (должности), который ИИ может случайно прихватить в ФИО
    clean = re.sub(r'(директор|министр|заместитель|начальник|председатель|генеральный)', '', raw_fio, flags=re.IGNORECASE).strip()
    parts = clean.split()
    if len(parts) >= 3: return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    if len(parts) == 2: return f"{parts[0]} {parts[1][0]}."
    return clean

# --- 1. НАСТРОЙКА ---
st.set_page_config(page_title="Юридический Генератор", layout="wide")

# ────────────────────────────────────────────────────────────────
# НАЧАЛО — ШАГ 1: только титульный лист (вставь вместо старого основного кода)
# ────────────────────────────────────────────────────────────────

if 'title_data' not in st.session_state:
    st.session_state.title_data = None
if 'title_buffer' not in st.session_state:
    st.session_state.title_buffer = None

# Пароль (оставляем как было)
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD:
    st.stop()

uploaded_file = st.file_uploader("Загрузите контракт (DOCX)", type=["docx"])

if uploaded_file is not None:
    # Новый файл → сбрасываем старые результаты
    current_file_name = uploaded_file.name
    if 'last_uploaded_name' not in st.session_state or st.session_state.last_uploaded_name != current_file_name:
        st.session_state.title_data = None
        st.session_state.title_buffer = None
        st.session_state.last_uploaded_name = current_file_name

    # Читаем документ один раз
    try:
        doc_obj = Document(uploaded_file)
        full_text = "\n".join([para.text for para in doc_obj.paragraphs])
    except Exception as e:
        st.error(f"Не удалось прочитать файл: {e}")
        st.stop()

    # Контекст для ИИ — начало + конец файла
    head = full_text[:1500]
    tail = full_text[-2200:]
    context = head + "\n\n[ ... середина опущена ... ]\n\n" + tail

    if st.session_state.title_data is None:
        with st.spinner("Извлекаем данные для титульного листа..."):
            try:
                response = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{
                        "role": "user",
                        "content": f"""Ты извлекаешь данные СТРОГО из текста. Ничего не придумывай.
                        
                        Верни ТОЛЬКО JSON. Никакого другого текста.
                        
                        Ключи:
                        - contract_no             → номер контракта (пример: 39/25/ГК)
                        - contract_date_raw       → дата как в тексте (пример: «20» октября 2025 г. или ___.10.2025)
                        - ikz                     → 36 цифр ИКЗ
                        - customer_org            → полное название заказчика
                        - customer_post           → должность подписанта заказчика (полностью)
                        - customer_basis          → основание полномочий (если есть, иначе пустая строка)
                        - customer_fio_raw        → ФИО заказчика как написано в тексте
                        - executor_org            → полное название исполнителя
                        - executor_post           → должность подписанта исполнителя
                        - executor_fio_raw        → ФИО исполнителя как написано в тексте
                        
                        Текст (начало + конец):
                        {context}
                        """
                    }],
                    response_format={"type": "json_object"},
                    temperature=0.1,
                    max_tokens=700
                )
                
                raw_data = json.loads(response.choices[0].message.content)
                
                # Постобработка
                td = {}
                td['contract_no'] = raw_data.get('contract_no', '________________')
                td['contract_date'] = raw_data.get('contract_date_raw', '________________')
                td['ikz'] = raw_data.get('ikz', '_______________________________')
                td['customer'] = raw_data.get('customer_org', '________________')
                td['customer_post_full'] = raw_data.get('customer_post', '').strip()
                if basis := raw_data.get('customer_basis', '').strip():
                    td['customer_post_full'] += f" {basis}"
                td['customer_fio'] = format_fio_universal(raw_data.get('customer_fio_raw', ''))
                td['executor'] = raw_data.get('executor_org', '________________')
                td['executor_post'] = (raw_data.get('executor_post', 'Директор') or 'Директор').capitalize()
                td['executor_fio'] = format_fio_universal(raw_data.get('executor_fio_raw', ''))
                
                st.session_state.title_data = td
                
            except Exception as e:
                st.error(f"Ошибка при запросе к DeepSeek: {str(e)}")
                st.stop()

    # Показываем результат
    if st.session_state.title_data:
        data = st.session_state.title_data
        
        st.subheader("Шаг 1 — Титульный лист")
        st.caption("Проверьте, правильно ли распознаны данные")
        
        cols = st.columns([3, 1])
        with cols[0]:
            with st.expander("Извлечённые данные", expanded=True):
                st.json(data)
        
        if st.button("Создать титульный лист → скачать для проверки"):
            buf = create_title_only_docx(data)
            st.session_state.title_buffer = buf.getvalue()
        
        if st.session_state.title_buffer:
            no_safe = re.sub(r'[^0-9а-яА-Яa-zA-Z\-_]', '_', data['contract_no'])
            st.download_button(
                label="📄 Скачать титульный лист (проверить)",
                data=st.session_state.title_buffer,
                file_name=f"Титульник_{no_safe}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_title"
            )

# ────────────────────────────────────────────────────────────────
# КОНЕЦ — ШАГ 1
# ────────────────────────────────────────────────────────────────

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
                # --- СНАЧАЛА ОПРЕДЕЛЯЕМ ИНДЕКСЫ И КОНТЕКСТ ---
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
                
                end_markers = ["ПРИЛОЖЕНИЕ № 2", "ПРИЛОЖЕНИЕ № 3", "РАСЧЕТ СТОИМОСТИ", "ПОДПИСИ СТОРОН"]
                tz_end_index = len(full_text)
                for marker in end_markers:
                    found_end = text_upper.find(marker, tz_index + 100)
                    if found_end != -1:
                        tz_end_index = found_end
                        break
                
                # --- ТЕПЕРЬ СОЗДАЕМ ПЕРЕМЕННЫЕ КОНТЕКСТА (Важно!) ---
                # Теперь NameError исчезнет, так как переменная context_title создана ДО запроса
                context_title = full_text[:1000] + "\n[...]\n" + full_text[-1000:]
                context_tz_full = full_text[tz_index : tz_end_index]

                # --- ТЕПЕРЬ ДЕЛАЕМ ЗАПРОСЫ К ИИ ---
                
                # 1. Данные титульника
                res_title = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"""
                        Извлеки СЫРЫЕ данные (как в тексте) в формате JSON. 
                        Не сокращай ФИО сам! Просто найди полное имя.
                        
                        Поля:
                        - contract_no: Номер из САМОЙ ПЕРВОЙ строки.
                        - contract_date: Дата.
                        - ikz: 36 цифр кода.
                        - project_name: Предмет контракта (С БОЛЬШОЙ БУКВЫ).
                        - customer: Название организации Заказчика.
                        - customer_post: Должность подписанта Заказчика.
                        - customer_fio: ФИО подписанта Заказчика.
                        - company: Название Исполнителя.
                        - executor_post: Должность руководителя Исполнителя.
                        - director: ПОЛНОЕ ФИО руководителя Исполнителя.

                        Текст: {context_title}
                    """}],
                    response_format={ 'type': 'json_object' }
                )
                
                title_info = json.loads(res_title.choices[0].message.content)
                st.session_state['title_info'] = title_info 
                
                # 2. Текст отчета
                res_report = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Напиши подробный отчет о выполненных работах, используя ВСЁ это ТЗ: {context_tz_full}"}]
                )
                report_text = res_report.choices[0].message.content
                
                # 3. Требования
                context_docs = context_tz_full + "\n" + full_text[-3000:]
                res_req = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Выпиши список отчетных документов (акты, фото, видео) из этого текста: {context_docs}"}]
                )
                requirements_text = res_req.choices[0].message.content
                
                # --- СОХРАНЕНИЕ ---
                doc_final = create_report_docx(report_text, title_info, requirements_text)
                
                buf = io.BytesIO()
                doc_final.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()
                st.session_state['current_no'] = title_info.get('contract_no', 'бн')
                
                st.success("Отчет готов!")

# --- КНОПКА СКАЧИВАНИЯ (ВНЕ ФОРМЫ - без отступов) ---
if st.session_state.get('report_buffer'):
    raw_no = st.session_state.get('current_no', 'бн')
    c_no = re.sub(r'[\\/*?:"<>|]', "_", str(raw_no))
    st.download_button(
        label=f"📥 Скачать отчет № {c_no}",
        data=st.session_state['report_buffer'],
        file_name=f"отчет и № {c_no}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )




