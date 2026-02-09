import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from openai import OpenAI
import io

# --- 1. НАСТРОЙКА СТРАНИЦЫ ---
st.set_page_config(page_title="Юридический Генератор Отчетов", layout="wide")

# --- 2. ПОДКЛЮЧЕНИЕ СЕКРЕТОВ И API ---
try:
    # Google Sheets (для базы эталонов)
    gcp_info = dict(st.secrets["gcp_service_account"])
    gcp_info["private_key"] = gcp_info["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(gcp_info, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    gc = gspread.authorize(creds)
    
    # DeepSeek
    DEEPSEEK_KEY = st.secrets["DEEPSEEK_API_KEY"].strip().strip('"')
    client_ai = OpenAI(api_key=DEEPSEEK_KEY, base_url="https://api.deepseek.com")
    
    SHEET_ID = st.secrets["SHEET_ID"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфигурации секретов: {e}")
    st.stop()

# --- 3. ФУНКЦИИ ---
def add_table_from_markdown(doc, markdown_text):
    """Превращает Markdown-таблицу от ИИ в реальную таблицу Word"""
    lines = [line.strip() for line in markdown_text.split('\n') if '|' in line]
    if len(lines) < 3: return
    headers = [cell.strip() for cell in lines[0].split('|') if cell.strip()]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
    for line in lines[2:]:
        cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        if len(cells) == len(headers):
            row_cells = table.add_row().cells
            for i, c in enumerate(cells):
                row_cells[i].text = c

# --- 4. АВТОРИЗАЦИЯ ---
user_pass = st.sidebar.text_input("Введите пароль доступа", type="password")
if user_pass != APP_PASSWORD:
    st.info("Пожалуйста, введите пароль в боковой панели.")
    st.stop()

# --- 5. ИНТЕРФЕЙС ---
st.title("🤖 Генератор профессиональных отчетов (v2.0)")

# Подгружаем базу из Google Sheets
try:
    sheet = gc.open_by_key(SHEET_ID).sheet1
    data_etalons = pd.DataFrame(sheet.get_all_records())
except:
    st.warning("Не удалось подключиться к Google Таблице. Используются настройки по умолчанию.")
    data_etalons = pd.DataFrame([{"Тип проекта": "Стандартный", "ЭТАЛОННАЯ СТРУКТУРА": "Стандартная"}])

uploaded_file = st.file_uploader("Загрузите Контракт (DOCX)", type=["docx"])

if uploaded_file:
    # ЧИТАЕМ ТЕКСТ КОНТРАКТА (Этого не хватало!)
    doc_input = Document(uploaded_file)
    contract_text = "\n".join([p.text for p in doc_input.paragraphs])
    
    st.success(f"Контракт загружен ({len(contract_text)} симв.)")

    with st.form("interview"):
        st.subheader("📝 Данные для наполнения отчета")
        col1, col2 = st.columns(2)
        with col1:
            q1 = st.text_input("Итоговое количество участников (цифрой)", placeholder="Напр: 80")
        with col2:
            q2 = st.text_input("Реквизиты письма согласования", placeholder="Напр: №123 от 01.12.25")
        
        additional_facts = st.text_area(
            "Дополнительные факты реализации (для полноты)", 
            help="Вставьте сюда даты заездов, адреса сбора, меню. ИИ распределит это по разделам.",
            placeholder="Напр: 2 группы по 40 чел. Заезды 8-9 и 10-11 декабря. Сбор в Реутове. Питание по меню..."
        )
        
        submitted = st.form_submit_button("🔥 Сформировать профессиональный отчет")

        if submitted:
            with st.spinner("Старший юрист DeepSeek анализирует контракт и факты..."):
                system_instruction = """Ты — ведущий юрист-аналитик. Твоя задача: на основе Контракта составить подробный ИНФОРМАЦИОННЫЙ ОТЧЕТ.
                ПРАВИЛА:
                1. ПРИНЦИП ЗЕРКАЛА: Опиши выполнение КАЖДОГО требования из ТЗ. Если в ТЗ указаны параметры оборудования или состав питания — перенеси их в отчет как выполненные.
                2. ТРАНСФОРМАЦИЯ: Контракт "должен" -> Отчет "Исполнителем обеспечено/выполнено".
                3. ТАБЛИЦЫ: ОБЯЗАТЕЛЬНО оформляй списки характеристик, меню или графики в виде Markdown-таблиц.
                4. СТРУКТУРА: Информационная справка -> Предмет -> Сроки -> Объем -> Содержательная часть (Питание, Транспорт и т.д.) -> Качество (ГОСТы)."""

                prompt_text = f"""
                КОНТРАКТ (ТЗ): {contract_text[:7000]} 
                ФАКТЫ ИЗ ИНТЕРВЬЮ: Участников: {q1}, Письмо: {q2}, Детали: {additional_facts}
                ЗАДАНИЕ: Напиши полный текст отчета. Для разделов с характеристиками используй таблицы."""

                try:
                    res = client_ai.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": system_instruction},
                            {"role": "user", "content": prompt_text}
                        ]
                    )
                    
                    report_content = res.choices[0].message.content
                    
                    # СОЗДАНИЕ ДОКУМЕНТА WORD
                    out_doc = Document()
                    blocks = report_content.split('\n\n')
                    
                    for block in blocks:
                        if '|' in block and '-|-' in block:
                            add_table_from_markdown(out_doc, block)
                        else:
                            if block.startswith('#'):
                                out_doc.add_heading(block.replace('#', '').strip(), level=2)
                            else:
                                out_doc.add_paragraph(block)
                    
                    # Сохранение в память
                    buffer = io.BytesIO()
                    out_doc.save(buffer)
                    st.session_state['report_buffer'] = buffer.getvalue()
                    st.success("Отчет готов!")
                except Exception as e:
                    st.error(f"Ошибка ИИ: {e}")

# КНОПКА СКАЧИВАНИЯ
if 'report_buffer' in st.session_state:
    st.download_button(
        label="📥 Скачать готовый Отчет (.docx)", 
        data=st.session_state['report_buffer'], 
        file_name="Final_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
