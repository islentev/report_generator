import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io

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
    st.error(f"Ошибка конфигурации секретов: {e}")
    st.stop()

# --- 3. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def add_table_from_markdown(doc, markdown_text):
    """Превращает Markdown-таблицу в таблицу Word"""
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
        if len(cells) >= len(headers):
            row_cells = table.add_row().cells
            for i in range(len(headers)):
                row_cells[i].text = cells[i]

def create_report_docx(report_content, title_data):
    """Создает документ с титульным листом и чистым форматированием"""
    doc = Document()
    
    # ТИТУЛЬНЫЙ ЛИСТ
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Исправлено: использование тройных кавычек для предотвращения SyntaxError
    auth_text = f"""УТВЕРЖДАЮ
Директор ООО «{title_data.get('Исполнитель', 'ЭОМ')}»

________________ / {title_data.get('Директор', 'Д.В. Скиба')}
«___» _________ 2025 г."""
    
    run_auth = p_auth.add_run(auth_text)
    run_auth.font.size = Pt(11)

    for _ in range(7): doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_main = p_title.add_run("ИНФОРМАЦИОННЫЙ ОТЧЕТ\n")
    run_main.bold = True
    run_main.font.size = Pt(20)
    
    contract_info = f"по исполнению Государственного контракта\n№ {title_data.get('Номер контракта', '_________')} от {title_data.get('Дата контракта', '_________')}\n\n"
    run_sub = p_title.add_run(contract_info)
    run_sub.font.size = Pt(14)
    p_title.add_run(f"{title_data.get('Название проекта', '')}").italic = True

    for _ in range(10): doc.add_paragraph()

    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.add_run("Москва, 2025 г.")
    
    doc.add_page_break()

    # ОСНОВНОЙ ТЕКСТ (Очистка от звездочек и форматирование)
    blocks = report_content.split('\n\n')
    for block in blocks:
        if '|' in block and '-|-' in block:
            add_table_from_markdown(doc, block)
        else:
            p = doc.add_paragraph()
            if block.strip().startswith('#'):
                p.add_run(block.replace('#', '').strip()).bold = True
                continue
            
            parts = block.split('**')
            for i, part in enumerate(parts):
                run = p.add_run(part.replace('*', ''))
                if i % 2 != 0:
                    run.bold = True
    return doc

# --- 4. ОСНОВНОЙ ИНТЕРФЕЙС ---
user_pass = st.sidebar.text_input("Пароль доступа", type="password")
if user_pass != APP_PASSWORD:
    st.info("Введите пароль в боковой панели для начала работы.")
    st.stop()

# Загрузка базы эталонов
try:
    sheet = gc.open_by_key(SHEET_ID).sheet1
    df_etalons = pd.DataFrame(sheet.get_all_records())
    selected_name = st.selectbox("Выберите тип проекта (эталон)", df_etalons["Тип проекта"].tolist())
    selected_etalon = df_etalons[df_etalons["Тип проекта"] == selected_name].iloc[0]
except Exception as e:
    st.error(f"Не удалось загрузить данные из Google Таблицы: {e}")
    st.stop()

uploaded_file = st.file_uploader("Загрузите файл Контракта (DOCX)", type="docx")

if uploaded_file:
    # Чтение текста контракта
    doc_input = Document(uploaded_file)
    contract_text = "\n".join([p.text for p in doc_input.paragraphs])
    
    with st.form("data_form"):
        st.subheader("📝 Параметры реализации")
        col1, col2 = st.columns(2)
        with col1:
            q1 = st.text_input("Кол-во участников", placeholder="Напр: 80")
        with col2:
            q2 = st.text_input("Письмо согласования", placeholder="Напр: №123 от 01.12.25")
        
        facts = st.text_area("Дополнительные детали (даты, адреса, меню)", 
                             placeholder="Введите факты, которых нет в контракте...")
        
        submitted = st.form_submit_button("🔥 Сформировать отчет")
        
    if submitted:
        if not q1 or not q2:
            st.warning("Заполните обязательные поля (участники и письмо).")
        else:
            with st.spinner("Старший юрист DeepSeek готовит документ..."):
                try:
                    sys_msg = "Ты — ведущий юрист. Создай отчет, зеркально отражая ТЗ Контракта в прошедшем времени. Используй таблицы для характеристик. Не используй Markdown разметку кроме жирного текста через **."
                    user_msg = f"КОНТРАКТ: {contract_text[:7000]}\nДАННЫЕ: Участников: {q1}, Письмо: {q2}, Детали: {facts}"
                    
                    res = client_ai.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role":"system","content":sys_msg}, {"role":"user","content":user_msg}]
                    )
                    
                    title_info = {
                        "Исполнитель": str(selected_etalon.get("Исполнитель", "ООО «ЕОМ»")),
                        "Директор": str(selected_etalon.get("Директор", "Скиба Д.В.")),
                        "Номер контракта": str(selected_etalon.get("Номер", "_________")),
                        "Дата контракта": str(selected_etalon.get("Дата", "_________")),
                        "Название проекта": selected_name
                    }
                    
                    final_doc = create_report_docx(res.choices[0].message.content, title_info)
                    
                    buf = io.BytesIO()
                    final_doc.save(buf)
                    st.session_state['report_buffer'] = buf.getvalue()
                    st.success("Отчет успешно сформирован!")
                    
                except Exception as e:
                    st.error(f"Ошибка при генерации: {e}")

    # Кнопка скачивания вне формы
    if 'report_buffer' in st.session_state:
        st.download_button(
            label="📥 Скачать готовый Отчет (.docx)", 
            data=st.session_state['report_buffer'], 
            file_name="Final_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
