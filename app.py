import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from openai import OpenAI
import io

# --- 1. ЗАГРУЗКА СЕКРЕТОВ ИЗ ОБЛАКА ---
try:
    # Данные Google из раздела Secrets
    gcp_info = st.secrets["gcp_service_account"]
    # Исправляем переносы строк в ключе
    gcp_info["private_key"] = gcp_info["private_key"].replace("\\n", "\n")
    
    creds = Credentials.from_service_account_info(
        gcp_info, 
        scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    gc = gspread.authorize(creds)
    
    # Ключи API и ID таблицы из раздела Secrets
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

    with st.form("interview"):
        st.subheader("Уточнение деталей")
        q1 = st.text_input("Фактическое число участников")
        q2 = st.text_input("Реквизиты письма согласования")
        
        if st.form_submit_button("Сгенерировать отчет"):
            with st.spinner("DeepSeek пишет отчет в прошедшем времени..."):
                
                # Промпт для DeepSeek
                prompt = f"""Перепиши условия этого контракта в прошедшее время для отчета.
                Контракт: {contract_text[:3000]}
                Эталонная структура: {selected_etalon.get('ЭТАЛОННАЯ СТРУКТУРА')}
                Доп. данные: {q1}, {q2}"""
                
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": prompt}]
                )
                
                # Создаем файл
                out_doc = Document()
                out_doc.add_heading(f"Отчет по проекту: {selected_etalon.get('Тип проекта')}", 0)
                out_doc.add_paragraph(res.choices[0].message.content)
                
                buffer = io.BytesIO()
                out_doc.save(buffer)
                
                st.download_button("📥 Скачать готовый Отчет (.docx)", buffer.getvalue(), "Report.docx")
