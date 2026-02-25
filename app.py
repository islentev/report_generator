import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_COLOR_INDEX
from openai import OpenAI
import io
import json
import re

# --- 1. ФУНКЦИИ ПАРСИНГА (ТВОИ ОРИГИНАЛЬНЫЕ) ---

def get_contract_start_text(file):
    doc = Document(file)
    full_text = []
    for table in doc.tables:
        for row in table.rows:
            full_text.append(" ".join(cell.text.strip() for cell in row.cells))
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            if re.match(r"^2\.", txt): 
                break
            full_text.append(txt)
    return "\n".join(full_text)[:2000]

def get_text_from_file(file):
    doc = Document(file)
    content = []
    for p in doc.paragraphs:
        if p.text.strip(): content.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            content.append(" ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(content)

def format_fio_short(fio_str):
    if not fio_str: return "___________"
    parts = fio_str.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio_str

def clean_markdown(text):
    return text.replace('**', '').replace('###', '').replace('##', '').replace('|', '').strip()

# --- 2. УМНАЯ ГЕНЕРАЦИЯ (ЛОГИКА ВНУТРИ) ---

def smart_generate_step_strict(client, section_text, requirements_text):
    system_prompt = f"""Ты — юридический редактор. Перепиши пункты ТЗ в Отчет.
    ПРАВИЛА:
    1. НУМЕРАЦИЯ: Сохраняй (1.1, 1.2...).
    2. ВРЕМЯ: Заголовки в НАСТОЯЩЕМ, пункты в ПРОШЕДШЕМ ('Услуги оказаны').
    3. ЗАПРЕТ: Удали 'должен', 'обязан', 'будет'.
    4. ПОЛНОТА: Все цифры и характеристики из ТЗ должны быть в отчете.
    ТРЕБОВАНИЯ: {requirements_text}"""

    # Шаг 1: Черновик
    res = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "system", "content": system_prompt},
                  {"role": "user", "content": section_text}],
        temperature=0.1
    )
    draft = res.choices[0].message.content

    # Шаг 2: Проверка
    v_res = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "system", "content": "Ты контролер. Сравни Отчет и ТЗ. Найди пропуски."},
                  {"role": "user", "content": f"ТЗ: {section_text}\nОТЧЕТ: {draft}\nВыдай 'ОШИБОК: 0' или список."}],
        temperature=0
    )
    
    # Шаг 3: Исправление
    if "ОШИБОК: 0" not in v_res.choices[0].message.content:
        fix = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": system_prompt},
                      {"role": "user", "content": f"Исправь ошибки: {v_res.choices[0].message.content}\nТекст: {draft}"}],
            temperature=0.1
        )
        return fix.choices[0].message.content
    return draft

# --- 3. СБОРКА ДОКУМЕНТА (ТВОЕ ОФОРМЛЕНИЕ) ---

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
    c_post = str(t.get('customer_post', 'Должность')).capitalize()
    e_post = str(t.get('director_post', 'Должность')).capitalize()
    tab.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{c_post}\n\n___________ / {format_fio_short(t.get('customer_fio'))}"
    tab.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{e_post}\n\n___________ / {format_fio_short(t.get('director'))}"
    tab.rows[1].cells[0].text = "м.п."; tab.rows[1].cells[1].text = "м.п."
    return doc

def apply_yellow_highlight(doc):
    keywords = ["Акт", "Фотоотчет", "Ведомость", "Скриншот", "Смета", "Резюме", "USB", "Флеш-накопитель"]
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for word in keywords:
                if word.lower() in run.text.lower():
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def create_final_report(t, report_body, req_body):
    doc = build_title_page(t)
    doc.add_page_break()
    p_name = str(t.get('project_name', 'услуг')).strip()
    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head.add_run(f"Отчет об оказании услуг по {p_name}").bold = True
    for line in clean_markdown(report_body).split('\n'):
        line = line.strip()
        if not line: continue
        para = doc.add_paragraph()
        run = para.add_run(line)
        if re.match(r"^\d+\.", line): run.bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # ПО ШИРИНЕ
    if req_body:
        doc.add_page_break()
        doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
        doc.add_paragraph(clean_markdown(req_body)).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_yellow_highlight(doc)
    return doc

# --- 4. ИНТЕРФЕЙС (ВОЗВРАТ К ТВОЕЙ СТРУКТУРЕ) ---

st.set_page_config(page_title="Генератор Отчетов 3.0", layout="wide")

with st.sidebar:
    st.title("Авторизация")
    if "auth" not in st.session_state: st.session_state.auth = False
    pwd = st.text_input("Пароль", type="password")
    if pwd == st.secrets["APP_PASSWORD"]: st.session_state.auth = True
    if not st.session_state.auth: st.stop()
    if st.button("♻️ СБРОСИТЬ ВСЕ ДАННЫЕ", use_container_width=True, type="primary"):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()

col1, col2, col3 = st.columns(3)

# КОЛОНКА 1: ТИТУЛЬНИК
with col1:
    st.header("📄 1. Титульный лист")
    f_title = st.file_uploader("Контракт (DOCX)", type="docx")
    t_context_area = st.text_area("ИЛИ вставьте начало контракта сюда:", height=150)
    if st.button("🔍 Извлечь реквизиты", use_container_width=True):
        if f_title:
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
            txt = get_contract_start_text(f_title)
            res = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "system", "content": "Верни JSON реквизитов."}, {"role": "user", "content": txt}],
                response_format={'type': 'json_object'}
            )
            st.session_state.t_info = json.loads(res.choices[0].message.content)

    if "t_info" in st.session_state:
        ti = st.session_state.t_info
        ti['contract_no'] = st.text_input("№", ti.get('contract_no'))
        ti['customer_fio'] = st.text_input("ФИО Заказчика", ti.get('customer_fio'))

# КОЛОНКА 2: ОТЧЕТ
with col2:
    st.header("📝 2. Отчет (ТЗ)")
    f_tz = st.file_uploader("Техзадание (DOCX)", type="docx")
    m_tz_area = st.text_area("ИЛИ вставьте текст ТЗ сюда:", height=150)
    if f_tz: st.session_state.raw_tz_source = get_text_from_file(f_tz)
    
    if st.button("⚙️ Сгенерировать текст", use_container_width=True):
        if "raw_tz_source" in st.session_state:
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
            res = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Напиши черновик отчета по ТЗ: {st.session_state.raw_tz_source}"}]
            )
            st.session_state.raw_report_body = res.choices[0].message.content

    if "raw_report_body" in st.session_state:
        st.session_state.raw_report_body = st.text_area("Черновик:", st.session_state.raw_report_body, height=300)

# КОЛОНКА 3: ТРЕБОВАНИЯ
with col3:
    st.header("📋 3. Требования")
    if st.button("🔍 Выделить требования", use_container_width=True):
        if "raw_tz_source" in st.session_state:
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
            res = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Выпиши требования к документам из ТЗ: {st.session_state.raw_tz_source}"}]
            )
            st.session_state.raw_requirements = res.choices[0].message.content

    if "raw_requirements" in st.session_state:
        st.session_state.raw_requirements = st.text_area("Требования:", st.session_state.raw_requirements, height=300)

# НИЖНИЙ БЛОК: СБОРКА
st.divider()
f_col1, f_col2 = st.columns(2)

with f_col1:
    if st.button("🚀 СОБРАТЬ ПОЛНЫЙ ОТЧЕТ (КАК ЕСТЬ)", use_container_width=True):
        if "t_info" in st.session_state:
            doc = create_final_report(st.session_state.t_info, st.session_state.get('raw_report_body', ''), st.session_state.get('raw_requirements', ''))
            buf = io.BytesIO(); doc.save(buf)
            st.session_state.full_file = buf.getvalue()

with f_col2:
    if st.button("🚀 ЗАПУСТИТЬ ПОШАГОВУЮ СБОРКУ", use_container_width=True):
        if all(k in st.session_state for k in ["t_info", "raw_tz_source"]):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
            steps = [s.strip() for s in re.split(r'\n(?=\d+\.)', st.session_state.raw_tz_source) if s.strip()]
            final_text = ""
            pb = st.progress(0)
            for i, step in enumerate(steps):
                final_text += smart_generate_step_strict(client, step, st.session_state.get('raw_requirements', '')) + "\n\n"
                pb.progress((i + 1) / len(steps))
            doc = create_final_report(st.session_state.t_info, final_text, st.session_state.get('raw_requirements', ''))
            buf = io.BytesIO(); doc.save(buf)
            st.session_state.smart_file = buf.getvalue()

if "full_file" in st.session_state:
    st.download_button("📥 Скачать обычный", st.session_state.full_file, "Report.docx")
if "smart_file" in st.session_state:
    st.download_button("📥 СКАЧАТЬ УМНЫЙ ОТЧЕТ", st.session_state.smart_file, "Smart_Report.docx")

