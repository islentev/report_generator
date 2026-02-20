import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI
import io
import json
import re
import docx2txt

def get_text_from_file(file):
    # Извлекает абсолютно весь текст, включая тот, что в таблицах
    text = docx2txt.process(file)
    return text

def get_contract_start_text(file):
    doc = Document(file)
    full_text = []
    
    # Читаем таблицы (там часто № и ИКЗ)
    for table in doc.tables:
        for row in table.rows:
            full_text.append(" ".join(cell.text.strip() for cell in row.cells))
    
    # Читаем параграфы
    for p in doc.paragraphs:
        full_text.append(p.text.strip())
        
    # Склеиваем и берем первые 5000 символов (этого хватит до 3-5 страницы)
    context = "\n".join(full_text)
    return context[:1000]

    # 2. Затем добавляем обычные параграфы
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            # Проверка на начало 2-го раздела (чтобы не кормить ИИ лишним)
            if re.match(r"^2\.", txt): 
                break
            start_text.append(txt)
            
    return "\n".join(start_text)

# --- 1. ОЧИСТКА ТЕКСТА ОТ СИМВОЛОВ ---
def clean_markdown(text):
    """Удаляет символы разметки типа ** или #"""
    text = text.replace('**', '')
    text = text.replace('###', '')
    text = text.replace('##', '')
    text = text.replace('|', '')
    return text.strip()

def format_fio_short(fio_str):
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

# --- 2. СБОРКА ДОКУМЕНТА (РУКОПИСНЫЙ СТИЛЬ) ---

def build_title_page(t):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
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
    tab = doc.add_table(rows=2, cols=2)

    # Делаем первую букву заглавной
    cust_post = str(t.get('customer_post', 'Должность')).capitalize()
    exec_post = str(t.get('director_post', 'Должность')).capitalize()

    # Вставляем именно переменные cust_post и exec_post
    tab.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{cust_post}\n\n___________ / {format_fio_short(t.get('customer_fio'))}"
    tab.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{exec_post}\n\n___________ / {format_fio_short(t.get('director'))}"
    tab.rows[1].cells[0].text = "м.п."
    tab.rows[1].cells[1].text = "м.п."

    return doc

def create_final_report(t, report_body, req_body):
    doc = build_title_page(t) # Используем уже готовую логику титульника
    doc.add_page_break()

    # Добавляем тело отчета (копируем логику из build_report_body)
    project_name = str(t.get('project_name', 'оказанию услуг')).strip()
    head = doc.add_paragraph()
    head.alignment = 1 # По центру
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
        para.alignment = 3 # По ширине

    doc.add_page_break()
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(clean_markdown(req_body))

    return doc

def build_report_body(report_text, req_text, t):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # Динамический жирный заголовок по центру
    project_name = str(t.get('project_name', 'оказанию услуг')).strip()
    head = doc.add_paragraph()
    head.alignment = 1 # 1 — это центр
    head.add_run(f"Отчет об оказании услуг по {project_name}").bold = True
    doc.add_paragraph()

    # Разделение на жирные главы и обычный текст
    lines = clean_markdown(report_text).split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        para = doc.add_paragraph()
        if re.match(r"^\d+\.", line): # Если строка начинается с цифры и точки
            para.add_run(line).bold = True
        else:
            para.add_run(line)
        para.alignment = 3 # 3 — это по ширине (решает ошибку AttributeError)
    
    doc.add_page_break()
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(clean_markdown(req_text))
    
    return doc
    
# --- 3. ИНТЕРФЕЙС ---
st.set_page_config(page_title="Генератор Отчетов 3.0", layout="wide")

# --- ПАРОЛЬ (без изменений) ---
with st.sidebar:
    st.title("Авторизация")
    if "auth" not in st.session_state: 
        st.session_state.auth = False
    pwd = st.text_input("Введите пароль", type="password")
    if pwd == st.secrets["APP_PASSWORD"]:
        st.session_state.auth = True
    if not st.session_state.auth:
        st.warning("Доступ ограничен.")
        st.stop()
    st.success("Доступ разрешен")

# --- НОВАЯ СТРУКТУРА ИНТЕРФЕЙСА (3 КОЛОНКИ) ---
col1, col2, col3 = st.columns(3)

# 1. ТИТУЛЬНЫЙ ЛИСТ
with col1:
    st.header("📄 1. Титульный лист")
    t_tab1, t_tab2 = st.tabs(["📁 Файл", "⌨️ Текст"])
    
    t_context = ""
    with t_tab1:
        f_title = st.file_uploader("Контракт (DOCX)", type="docx", key="u_title")
        if f_title: t_context = get_contract_start_text(f_title)
    with t_tab2:
        m_title = st.text_area("Вставьте начало контракта", height=150, key="m_title")
        if m_title: t_context = m_title

    if st.button("🔍 Извлечь реквизиты", use_container_width=True):
        if t_context:
            with st.spinner("Ищем данные..."):
                client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": "Ты парсер. Верни JSON (contract_no, contract_date, ikz, project_name, customer, customer_post, customer_fio, company, director_post, director)."},
                              {"role": "user", "content": t_context}],
                    response_format={'type': 'json_object'}
                )
                st.session_state.t_info = json.loads(res.choices[0].message.content)
        else: st.error("Нет данных!")

    # --- ПРЕВЬЮ ТИТУЛЬНИКА (Редактируемое) ---
    if "t_info" in st.session_state:
        st.info("Проверьте данные:")
        ti = st.session_state.t_info
        ti['contract_no'] = st.text_input("№", ti.get('contract_no'))
        ti['ikz'] = st.text_input("ИКЗ", ti.get('ikz'))
        ti['customer_fio'] = st.text_input("ФИО Заказчика", ti.get('customer_fio'))
        # Кнопка скачивания только титульника
        doc_t = build_title_page(ti)
        buf_t = io.BytesIO(); doc_t.save(buf_t)
        st.download_button("📥 Скачать Титульник", buf_t.getvalue(), "Title.docx", use_container_width=True)

# 2. ОТЧЕТ (ОСНОВНОЙ ТЕКСТ)
with col2:
    st.header("📝 2. Отчет (ТЗ)")
    r_tab1, r_tab2 = st.tabs(["📁 Файл", "⌨️ Текст"])
    
    raw_tz_content = ""
    with r_tab1:
        f_tz = st.file_uploader("Техзадание (DOCX)", type="docx", key="u_tz")
        if f_tz: raw_tz_content = get_text_from_file(f_tz)
    with r_tab2:
        m_tz = st.text_area("Вставьте текст ТЗ", height=150, key="m_tz")
        if m_tz: raw_tz_content = m_tz

    if st.button("⚙️ Сгенерировать текст", use_container_width=True):
        if raw_tz_content:
            st.session_state.raw_tz_source = raw_tz_content # Сохраняем для 3-й колонки
            with st.spinner("Пишем черновик..."):
                # (Временно обычная генерация, во 2 шаге заменим на пошаговую)
                client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": "Ты техписатель. Пиши отчет в прошедшем времени."},
                              {"role": "user", "content": f"Текст ТЗ: {raw_tz_content}"}]
                )
                st.session_state.raw_report_body = res.choices[0].message.content
        else: st.error("Загрузите ТЗ!")

    # --- ПРЕВЬЮ ОТЧЕТА ---
    if "raw_report_body" in st.session_state:
        st.session_state.raw_report_body = st.text_area("Черновик текста:", st.session_state.raw_report_body, height=300)

# 3. ТРЕБОВАНИЯ К ДОКУМЕНТАЦИИ
with col3:
    st.header("📋 3. Требования")
    st.write("Использует ТЗ из 2-й колонки")
    
    if st.button("🔍 Выделить требования", use_container_width=True):
        if "raw_tz_source" in st.session_state:
            with st.spinner("Ищем правила оформления..."):
                client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": "Найди в ТЗ требования к фото, документам, количеству знаков и носителям (USB и т.д.)."},
                              {"role": "user", "content": st.session_state.raw_tz_source}]
                )
                st.session_state.raw_requirements = res.choices[0].message.content
        else: st.warning("Сначала загрузите ТЗ во 2-й колонке!")

    # --- ПРЕВЬЮ ТРЕБОВАНИЙ ---
    if "raw_requirements" in st.session_state:
        st.session_state.raw_requirements = st.text_area("Список требований:", st.session_state.raw_requirements, height=300)

# --- ФИНАЛЬНЫЙ БЛОК (НИЖНЯЯ ПАНЕЛЬ) ---
st.divider()
st.subheader("🏁 Финальная сборка")
f_col1, f_col2 = st.columns(2)

with f_col1:
    if st.button("🚀 СОБРАТЬ ПОЛНЫЙ ОТЧЕТ (КАК ЕСТЬ)", use_container_width=True):
        if "t_info" in st.session_state and "raw_report_body" in st.session_state:
            full_doc = create_final_report(st.session_state.t_info, st.session_state.raw_report_body, st.session_state.get('raw_requirements', ''))
            buf = io.BytesIO(); full_doc.save(buf)
            st.session_state.full_file = buf.getvalue()
            st.success("Готово!")
    
    if "full_file" in st.session_state:
        st.download_button("📥 Скачать всё одним файлом", st.session_state.full_file, "Full_Report.docx", use_container_width=True)

with f_col2:
    if st.button("🪄 ПРИМЕНИТЬ ТРЕБОВАНИЯ К ОТЧЕТУ", use_container_width=True):
        st.info("Здесь будет запускаться трансформация (Шаг 2)")
