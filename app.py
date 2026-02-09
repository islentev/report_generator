import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from openai import OpenAI
import io

# --- 1. ЗАГРУЗКА СЕКРЕТОВ ИЗ ОБЛАКА ---
try:
   # Копируем данные из секретов
    gcp_info = dict(st.secrets["gcp_service_account"])
    
    if "private_key" in gcp_info:
        # 1. Убираем лишние кавычки, если они случайно попали внутрь строки
        raw_key = gcp_info["private_key"].strip('"').strip("'")
        
        # 2. Заменяем текстовые \n на реальные символы переноса строки
        # И убираем возможные пробелы вокруг
        gcp_info["private_key"] = raw_key.replace("\\n", "\n").strip()
    
    # Теперь замена сработает, так как gcp_info — это обычный словарь
    if "private_key" in gcp_info:
        gcp_info["private_key"] = gcp_info["private_key"].replace("\\n", "\n")
    
    creds = Credentials.from_service_account_info(
        gcp_info, 
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    gc = gspread.authorize(creds)
    
    # Ключи API и ID таблицы
    DEEPSEEK_API_KEY = st.secrets["DEEPSEEK_API_KEY"]
    SHEET_ID = st.secrets["SHEET_ID"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
    
    # Инициализация DeepSeek
    client_ai = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")
except Exception as e:
    st.error(f"Ошибка конфигурации: {e}")
    st.stop()

# --- 2. ЗАЩИТА ПАРОЛЕМ ---
st.sidebar.title("🔐 Доступ")
user_pass = st.sidebar.text_input("Введите пароль", type="password")
if user_pass != APP_PASSWORD:
    st.info("Введите пароль в боковой панели, чтобы начать работу.")
    st.stop()

# --- 3. ЛОГИКА ПРИЛОЖЕНИЯ ---
st.title("🤖 Генератор отчетов по госконтрактам")

try:
    sheet = gc.open_by_key(SHEET_ID).sheet1
    data = pd.DataFrame(sheet.get_all_records())
    st.success("База эталонов подключена!")
except Exception as e:
    st.error(f"Не удалось прочитать таблицу: {e}")
    st.stop()

uploaded_file = st.file_uploader("Загрузите Контракт (DOCX)", type=["docx"])

if uploaded_file:
    # Читаем DOCX
    doc = Document(uploaded_file)
    contract_text = "\n".join([p.text for p in doc.paragraphs])
    
    # Выбор эталона (например, первый)
    selected_etalon = data.iloc[0]
    st.info(f"Выбран эталон: {selected_etalon.get('Тип проекта')}")

    # Создаем пустое место для хранения документа
report_data = None

if uploaded_file:
    # ... ваш код чтения DOCX ...
    
    with st.form("interview"):
        st.subheader("Уточнение деталей")
        q1 = st.text_input("Фактическое число участников")
        q2 = st.text_input("Реквизиты письма согласования")
        
        # Кнопка внутри формы только ЗАПУСКАЕТ процесс
        submitted = st.form_submit_button("Сформировать отчет")
        
        if submitted:
            with st.spinner("DeepSeek пишет отчет в прошедшем времени..."):
                # ... ваш код генерации через DeepSeek ...
                
                # Собираем документ
                out_doc = Document()
                out_doc.add_heading(f"Отчет по проекту: {selected_etalon.get('Тип проекта')}", 0)
                out_doc.add_paragraph(res.choices[0].message.content)
                
                buffer = io.BytesIO()
                out_doc.save(buffer)
                # Сохраняем результат в session_state, чтобы он не пропал при перезагрузке страницы
                st.session_state['report_buffer'] = buffer.getvalue()
                st.success("Отчет успешно сформирован!")

    # КНОПКА СКАЧИВАНИЯ — ТЕПЕРЬ ВНЕ ФОРМЫ
    if 'report_buffer' in st.session_state:
        st.download_button(
            label="📥 Скачать готовый Отчет (.docx)",
            data=st.session_state['report_buffer'],
            file_name="Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )



