import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. ОЧИСТКА ТЕКСТА ОТ СИМВОЛОВ ---

def clean_markdown(text):
    """Удаляет символы разметки типа ** или #"""
    text = text.replace('**', '')
    text = text.replace('###', '')
    text = text.replace('##', '')
    text = text.replace('|', '')
    return text.strip()

def get_text_from_file(file):
    doc = Document(file)
    content = []
    for p in doc.paragraphs:
        if p.text.strip(): content.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            content.append(" ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(content)

# --- 2. СБОРКА ДОКУМЕНТА (РУКОПИСНЫЙ СТИЛЬ) ---

def create_final_report(title_data, report_body, req_body):
    doc = Document()
    t = title_data
    
    # Настройка стиля (Times New Roman)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # --- БЛОК 1: ТИТУЛЬНИК (БЕЗ ИЗМЕНЕНИЙ) ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Информационно-аналитический отчет об исполнении условий\n").bold = True
    p.add_run(f"Контракта № {t.get('contract_no', '___')} от «{t.get('contract_date', '___')}» 2025 г.\n").bold = True
    p.add_run(f"Идентификационный код закупки: {t.get('ikz', '___')}.")
    for _ in range(5): doc.add_paragraph()
    doc.add_paragraph("ТОМ I").alignment = WD_ALIGN_PARAGRAPH.CENTER
    for label, val in [("Наименование предмета КОНТРАКТА :", t.get('project_name')), ("Заказчик:", t.get('customer')), ("Исполнитель:", t.get('company'))]:
        p_l = doc.add_paragraph(); p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.add_run(label).bold = True
        p_v = doc.add_paragraph(); p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(str(val)).italic = True
    for _ in range(5): doc.add_paragraph()
    tab = doc.add_table(rows=1, cols=2)
    tab.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{t.get('customer_fio')}\n\n___________"
    tab.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{t.get('director')}\n\n___________"
    doc.add_page_break()

    # --- БЛОК 2: ОТЧЕТ (РУКОПИСНЫЙ ТЕКСТ) ---
    # Единый заголовок по центру
    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head.add_run("ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ").bold = True
    
    # Очищаем и вставляем основной текст
    cleaned_body = clean_markdown(report_body)
    doc.add_paragraph(cleaned_body)

    doc.add_page_break()

    # --- БЛОК 3: ТРЕБОВАНИЯ ---
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(clean_markdown(req_body))

    return doc

# --- 3. ИНТЕРФЕЙС ---

st.set_page_config(page_title="Генератор Отчетов 3.0", layout="wide")

# (Блок пароля остается прежним)
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    if st.text_input("Пароль", type="password") == st.secrets["APP_PASSWORD"]:
        st.session_state.auth = True
        st.rerun()
    st.stop()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Файл Контракта")
    file_contract = st.file_uploader("Загрузите контракт для реквизитов", type="docx")
    if file_contract and st.button("Собрать реквизиты"):
        client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
        text = get_text_from_file(file_contract)
        context = text[:3000] + "\n" + text[-3000:]
        res = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {context}"}],
            response_format={'type': 'json_object'}
        )
        st.session_state.title_info = json.loads(res.choices[0].message.content)
        st.success("Титульник зафиксирован")

with col2:
    st.subheader("2. Файл ТЗ")
    file_tz = st.file_uploader("Загрузите только файл ТЗ", type="docx")
    if file_tz and "title_info" in st.session_state:
        if st.button("Сформировать рукописный отчет"):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
            tz_text = get_text_from_file(file_tz)
            
            with st.spinner("Пишу отчет..."):
                # Промпт для "рукописного" стиля с главами
                res_body = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": """Ты профессиональный технический писатель. 
                        Сформируй отчет по следующим правилам:
                        1. Никаких таблиц.
                        2. Каждая услуга из ТЗ — это новая глава с нумерацией (1., 2. и т.д.).
                        3. ЗАГОЛОВОК ГЛАВЫ пиши в НАСТОЯЩЕМ времени жирным шрифтом.
                        4. ОПИСАНИЕ внутри главы пиши в ПРОШЕДШЕМ времени (выполнено, организовано, предоставлено).
                        5. Убирай любые символы разметки типа **, #, |. 
                        6. Текст должен быть связным, как будто написан человеком."""},
                        {"role": "user", "content": f"Сделай отчет из этого ТЗ:\n\n{tz_text}"}
                    ]
                )
                
                res_req = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Выпиши списком требования к фото и документам из этого ТЗ: {tz_text}"}]
                )
                
                final_docx = create_final_report(st.session_state.title_info, res_body.choices[0].message.content, res_req.choices[0].message.content)
                buf = io.BytesIO()
                final_docx.save(buf)
                st.session_state.ready_file = buf.getvalue()
                st.success("Отчет в новом стиле готов!")

if "ready_file" in st.session_state:
    st.divider()
    st.download_button("📥 Скачать готовый отчет", st.session_state.ready_file, "Handwritten_Report.docx")
