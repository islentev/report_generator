import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def clean_markdown(text):
    """Удаляет символы разметки типа ** или #"""
    text = text.replace('**', '')
    text = text.replace('###', '')
    text = text.replace('##', '')
    text = text.replace('|', '')
    return text.strip()

def format_fio_short(fio_str):
    """Иванов Иван Иванович -> Иванов И.И."""
    if not fio_str: return "___________"
    parts = fio_str.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio_str

def get_text_from_file(file):
    doc = Document(file)
    content = []
    for p in doc.paragraphs:
        if p.text.strip(): content.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            content.append(" ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(content)

def get_contract_start_text(file):
    doc = Document(file)
    start_text = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            start_text.append(txt)
            if re.match(r"^2\.", txt): 
                break
    return "\n".join(start_text)

# --- 2. ФУНКЦИИ СБОРКИ DOCX ---

def build_title_page_logic(doc, t):
    """Логика заполнения титульного листа (используется и отдельно, и в общем отчете)"""
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    p = doc.add_paragraph()
    p.alignment = 1 # Center
    p.add_run(f"Информационно-аналитический отчет об исполнении условий\n").bold = True
    p.add_run(f"Контракта № {t.get('contract_no', '___')} от «{t.get('contract_date', '___')}» 2025 г.\n").bold = True
    p.add_run(f"Идентификационный код закупки: {t.get('ikz', '___')}.")
    
    # Уменьшенные отступы до 3, чтобы влезло на 1 страницу
    for _ in range(3): doc.add_paragraph()
    doc.add_paragraph("ТОМ I").alignment = 1
    for _ in range(2): doc.add_paragraph()

    for label, val in [("Наименование предмета КОНТРАКТА :", t.get('project_name')), ("Заказчик:", t.get('customer')), ("Исполнитель:", t.get('company'))]:
        p_l = doc.add_paragraph(); p_l.alignment = 1
        p_l.add_run(label).bold = True
        p_v = doc.add_paragraph(); p_v.alignment = 1
        p_v.add_run(str(val)).italic = True

    for _ in range(3): doc.add_paragraph()
    tab = doc.add_table(rows=2, cols=2)
    
    cust_post = str(t.get('customer_post', 'Должность')).capitalize()
    exec_post = str(t.get('director_post', 'Должность')).capitalize()
    
    tab.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{cust_post}\n\n___________ / {format_fio_short(t.get('customer_fio'))}"
    tab.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{exec_post}\n\n___________ / {format_fio_short(t.get('director'))}"
    tab.rows[1].cells[0].text = "м.п."
    tab.rows[1].cells[1].text = "м.п."
    return doc

def build_report_body_logic(doc, report_body, req_body, t):
    """Логика заполнения тела отчета"""
    project_name = str(t.get('project_name', 'оказанию услуг')).strip()
    head = doc.add_paragraph()
    head.alignment = 1 # Center
    head.add_run(f"Отчет об оказании услуг по {project_name}").bold = True
    doc.add_paragraph()

    lines = clean_markdown(report_body).split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        para = doc.add_paragraph()
        if re.match(r"^\d+\.", line):
            para.add_run(line).bold = True
        else:
            para.add_run(line)
        para.alignment = 3 # Justify
        
    doc.add_page_break()
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(clean_markdown(req_body))
    return doc

# ФУНКЦИИ-ОБЕРТКИ ДЛЯ КНОПОК
def build_title_page(t):
    return build_title_page_logic(Document(), t)

def build_report_body(report_body, req_body, t):
    return build_report_body_logic(Document(), report_body, req_body, t)

def create_final_report(t, report_body, req_body):
    doc = Document()
    build_title_page_logic(doc, t)
    doc.add_page_break()
    build_report_body_logic(doc, report_body, req_body, t)
    return doc

# --- 3. ИНТЕРФЕЙС STREAMLIT ---

st.set_page_config(page_title="Генератор Отчетов 3.0", layout="wide")

with st.sidebar:
    st.title("Авторизация")
    if "auth" not in st.session_state: st.session_state.auth = False
    pwd = st.text_input("Введите пароль", type="password")
    if pwd == st.secrets["APP_PASSWORD"]: st.session_state.auth = True
    if not st.session_state.auth:
        st.warning("Доступ ограничен.")
        st.stop()
    st.success("Доступ разрешен")

col1, col2 = st.columns(2)

with col1:
    st.header("📄 1. Титульный лист")
    file_contract = st.file_uploader("Загрузите Контракт", type="docx", key="contract_loader")
    if file_contract:
        if st.button("Сформировать Титульный лист"):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
            context = get_contract_start_text(file_contract)
            res = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Верни JSON по первой странице: contract_no, contract_date, ikz, project_name, customer, customer_post, customer_fio, company, director_post, director. Текст: {context}"}],
                response_format={'type': 'json_object'}
            )
            st.session_state.t_info = json.loads(res.choices[0].message.content)
            doc_title = build_title_page(st.session_state.t_info)
            buf_t = io.BytesIO()
            doc_title.save(buf_t)
            st.session_state.file_title_only = buf_t.getvalue()
            st.success("Титульный лист готов!")
        if "file_title_only" in st.session_state:
            st.download_button("📥 Скачать Титульник", st.session_state.file_title_only, "Title_Page.docx")

with col2:
    st.header("📝 2. Отчет по ТЗ")
    file_tz = st.file_uploader("Загрузите Техзадание", type="docx", key="tz_loader")
    if file_tz:
        if st.button("Сформировать Рукописный отчет"):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
            raw_tz = get_text_from_file(file_tz)
            with st.spinner("ИИ анализирует ТЗ..."):
                res_body = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": "Ты техписатель. Главы (1., 2.) в настоящем времени, описание в прошедшем. Без символов **."},
                             {"role": "user", "content": f"Сделай отчет из ТЗ:\n{raw_tz}"}]
                )
                res_req = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди требования к фото и документам: {raw_tz}"}]
                )
                st.session_state.raw_report_body = res_body.choices[0].message.content
                st.session_state.raw_requirements = res_req.choices[0].message.content
                
                doc_rep = build_report_body(
                    st.session_state.raw_report_body, 
                    st.session_state.raw_requirements,
                    st.session_state.t_info if "t_info" in st.session_state else {}
                )
                buf_r = io.BytesIO()
                doc_rep.save(buf_r)
                st.session_state.file_report_only = buf_r.getvalue()
                st.success("Отчет сформирован!")
        if "file_report_only" in st.session_state:
            st.download_button("📥 Скачать Отчет (без титульника)", st.session_state.file_report_only, "Report_Only.docx")

if "file_title_only" in st.session_state and "file_report_only" in st.session_state:
    st.divider()
    if st.button("🚀 СОБРАТЬ ПОЛНЫЙ ОТЧЕТ", use_container_width=True):
        full_doc = create_final_report(
            st.session_state.t_info, 
            st.session_state.raw_report_body, 
            st.session_state.raw_requirements
        )
        final_buf = io.BytesIO()
        full_doc.save(final_buf)
        st.session_state.full_ready_file = final_buf.getvalue()
    if "full_ready_file" in st.session_state:
        st.download_button("🔥 СКАЧАТЬ ВЕСЬ ДОКУМЕНТ", st.session_state.full_ready_file, "Full_Final_Report.docx", use_container_width=True)
