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
    # Добавили /v1 для стабильности
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
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        
    for line in lines[2:]:
        # Более надежное разделение ячеек
        cells = [cell.strip() for cell in line.split('|') if cell.strip() or line.split('|')[0] == '']
        # Фильтруем пустые элементы по краям
        if line.startswith('|'): cells = [cell.strip() for cell in line.split('|')][1:-1]
        
        if len(cells) >= len(headers):
            row_cells = table.add_row().cells
            for i in range(len(headers)):
                row_cells[i].text = cells[i]

def create_report_docx(report_content, title_data):
    doc = Document()
    
    # ТИТУЛЬНЫЙ ЛИСТ
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_auth = p_auth.add_run(f"УТВЕРЖДАЮ\nДиректор ООО «{title_data.get('Исполнитель', 'ЭОМ')}»\n\n________________ / {title_data.get('Директор', 'Д.В. Скиба')}\n«___» _________ 2025 г.")
    run_auth.font.size = Pt(11)

    for _ in range(7): doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_main = p_title.add_run("ИНФОРМАЦИОННЫЙ ОТЧЕТ\n")
    run_main.bold = True
    run_main.font.size = Pt(20)
    
    run_sub = p_title.add_run(f"по исполнению Государственного контракта\n№ {title_data.get('Номер контракта', '_________')} от {title_data.get('Дата контракта', '_________')}\n\n")
    run_sub.font.size = Pt(14)
    p_title.add_run(f"{title_data.get('Название проекта', '')}").italic = True

    for _ in range(10): doc.add_paragraph()

    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.add_run("Москва, 2025 г.")
    
    doc.add_page_break()

    # ОСНОВНОЙ ТЕКСТ
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
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD:
    st.info("Введите пароль для доступа к системе.")
    st.stop()

# Загрузка базы эталонов
try:
    sheet = gc.open_by_key(SHEET_ID).sheet1
    df_etalons = pd.DataFrame(sheet.get_all_records())
    selected_name = st.selectbox("Выберите тип проекта (эталон)", df_etalons["Тип проекта"].tolist())
    selected_etalon = df_etalons[df_etalons["Тип проекта"] == selected_name].iloc[0]
except Exception as e:
    st.error(f"Ошибка загрузки таблицы: {e}")
    st.stop()

uploaded_file = st.file_uploader("Загрузите файл Контракта", type="docx")

if uploaded_file:
    contract_text = "\n".join([p.text for p in Document(uploaded_file).paragraphs])
    
    with st.form("data_form"):
        col1, col2 = st.columns(2)
        with col1:
            q1 = st.text_input("Кол-во участников", placeholder="80")
        with col2:
            q2 = st.text_input("Письмо согласования", placeholder="№1 от 01.12.25")
        
        facts = st.text_area("Доп. детали (даты, меню, адреса)", placeholder="Заезды 8-11 дек, меню: каша...")
        
        submitted = st.form_submit_button("Сгенерировать")
        
    if submitted:
        with st.spinner("DeepSeek формирует юридический текст..."):
            try:
                sys_msg = "Ты — ведущий юрист. Создай отчет, зеркально отражая ТЗ Контракта в прошедшем времени. Используй таблицы для характеристик."
                user_msg = f"КОНТРАКТ: {contract_text[:7000]}\nДАННЫЕ: Участников: {q1}, Письмо: {q2}, Детали: {facts}"
                
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":sys_msg}, {"role":"user","content":user_msg}]
                )
                
                title_info = {
                    "Исполнитель": str(selected_etalon.get("Исполнитель", "ЕОМ")),
                    "Директор": str(selected_etalon.get("Директор", "Скиба Д.В.")),
                    "Номер контракта": str(selected_etalon.get("Номер", "0148200002625000032")),
                    "Дата контракта": str(selected_etalon.get("Дата", "01.12.2025")),
                    "Название проекта": selected_name
                }
                
                final_doc = create_report_docx(res.choices[0].message.content, title_info)
                
                buf = io.BytesIO()
                final_doc.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()
                st.success("Отчет готов!")
                
            except Exception as e:
                st.error(f"Ошибка ИИ: {e}")

    # Кнопка скачивания вне формы
    if 'report_buffer' in st.session_state:
        st.download_button(
            label="📥 Скачать Отчет .docx", 
            data=st.session_state['report_buffer'], 
            file_name="Report_Legal.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    # ТИТУЛЬНЫЙ ЛИСТ
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_auth = p_auth.add_run(f"УТВЕРЖДАЮ\n
