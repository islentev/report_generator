import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. УТИЛИТЫ ДЛЯ ЧТЕНИЯ ---

def get_full_text_including_tables(doc):
    full_element_list = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = [p for p in doc.paragraphs if p._element == element]
            if para:
                full_element_list.append(para[0].text)
        elif element.tag.endswith('tbl'):
            table = [t for t in doc.tables if t._element == element]
            if table:
                table_text = []
                for row in table[0].rows:
                    row_data = " | ".join(cell.text.strip() for cell in row.cells)
                    table_text.append(row_data)
                full_element_list.append("\n[ТАБЛИЦА ТЗ]:\n" + "\n".join(table_text))
    return "\n".join(full_element_list)

def extract_tz_content(full_text):
    text_upper = full_text.upper()
    start_markers = ["ПРИЛОЖЕНИЕ № 1", "ТЕХНИЧЕСКОЕ ЗАДАНИЕ"]
    start_pos = -1
    for m in start_markers:
        found = text_upper.find(m)
        if found != -1:
            start_pos = found
            break
    if start_pos == -1: return full_text
    
    end_markers = ["ПРИЛОЖЕНИЕ № 2", "РАСЧЕТ СТОИМОСТИ", "ПОДПИСИ СТОРОН"]
    end_pos = len(full_text)
    for m in end_markers:
        found_end = text_upper.find(m, start_pos + 100)
        if found_end != -1:
            end_pos = found_end
            break
    return full_text[start_pos:end_pos]

def format_fio_universal(raw_fio):
    if not raw_fio or len(raw_fio) < 5: return "________________"
    clean = re.sub(r'(директор|министр|заместитель|начальник|председатель|генеральный)', '', raw_fio, flags=re.IGNORECASE).strip()
    parts = clean.split()
    if len(parts) >= 3: return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    if len(parts) == 2: return f"{parts[0]} {parts[1][0]}."
    return clean

# --- 2. КОНСТРУКЦИЯ ТИТУЛЬНИКА (ВАША) ---

def create_report_docx(report_content, title_data, requirements_list):
    doc = Document()
    contract_no = title_data.get('contract_no', '________________')
    contract_date = title_data.get('contract_date', '___')
    ikz = title_data.get('ikz', '________________')
    raw_name = title_data.get('project_name', '')
    project_name = raw_name[0].upper() + raw_name[1:] if raw_name else ""
    customer = title_data.get('customer', '')
    company = title_data.get('company', '')
    
    cust_post = str(title_data.get('customer_post', 'Заказчик')).capitalize()
    cust_fio = format_fio_universal(title_data.get('customer_fio', ''))
    exec_post = str(title_data.get('executor_post', 'Директор')).capitalize()
    exec_fio = format_fio_universal(title_data.get('director', ''))

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_top.add_run("Информационно-аналитический отчет об исполнении условий\n").bold = True
    p_top.add_run(f"Контракта № {contract_no} от «{contract_date}» 2025 г.\n").bold = True
    p_top.add_run(f"Идентификационный код закупки: {ikz}.")

    for _ in range(3): doc.add_paragraph()
    p_tom = doc.add_paragraph()
    p_tom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tom.add_run("ТОМ I").bold = True

    for label, value in [("Наименование предмета КОНТРАКТА :", project_name), ("Заказчик:", customer), ("Исполнитель:", company)]:
        p_h = doc.add_paragraph(); p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h.add_run(label).bold = True
        p_v = doc.add_paragraph(); p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(value).italic = True

    for _ in range(4): doc.add_paragraph()
    table = doc.add_table(rows=2, cols=2)
    table.width = doc.sections[0].page_width
    table.rows[0].cells[0].paragraphs[0].add_run("Отчет принят Заказчиком").bold = True
    table.rows[0].cells[0].paragraphs[0].add_run(f"\n\n{cust_post} {cust_fio}\n\n_______________")
    table.rows[0].cells[1].paragraphs[0].add_run("Отчет передан Исполнителем").bold = True
    table.rows[0].cells[1].paragraphs[0].add_run(f"\n\n{exec_post}\n\n_______________ / {exec_fio}")
    table.rows[1].cells[0].paragraphs[0].add_run("м.п.")
    table.rows[1].cells[1].paragraphs[0].add_run("м.п.")

    doc.add_page_break()
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
    doc.add_paragraph(requirements_list)
    return doc

# --- 3. ОСНОВНАЯ ЛОГИКА ---
st.set_page_config(page_title="Юридический Генератор", layout="wide")
if 'report_buffer' not in st.session_state: st.session_state['report_buffer'] = None
if 'title_info' not in st.session_state: st.session_state['title_info'] = None

try:
    client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфига: {e}"); st.stop()

user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD: st.stop()

uploaded_file = st.file_uploader("Загрузите контракт (DOCX)", type="docx")

if uploaded_file:
    doc_obj = Document(uploaded_file)
    full_text_with_tables = get_full_text_including_tables(doc_obj)
    
    if st.button("Шаг 1: Подготовить данные титульника"):
        context = full_text_with_tables[:3000] + "\n" + full_text_with_tables[-4000:]
        res = client_ai.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Извлеки реквизиты в JSON: contract_no, contract_date, ikz, project_name, customer, customer_post, customer_fio, company, executor_post, director. Текст: {context}"}],
            response_format={ 'type': 'json_object' }
        )
        st.session_state.title_info = json.loads(res.choices[0].message.content)
        st.success("Данные титульника готовы")

    if st.session_state.title_info:
        with st.form("report_form"):
            facts = st.text_area("Доп. факты (необязательно)")
            if st.form_submit_button("Сгенерировать отчет по ТЗ"):
                with st.spinner("Анализируем ТЗ..."):
                    pure_tz = extract_tz_content(full_text_with_tables)
                    
                    # ВОТ ЗДЕСЬ Я УСИЛИЛ ИНСТРУКЦИЮ, ЧТОБЫ ОН НЕ ГАЛЛЮЦИНИРОВАЛ
                    res_report = client_ai.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": """Ты — технический писатель. Твоя задача — составить отчет ПО КАЖДОМУ ПУНКТУ ТЗ.
                            ИНСТРУКЦИЯ:
                            1. Найди в тексте ТЗ таблицу или список услуг (Аренда, Застройка, Видеосопровождение и т.д.).
                            2. Для КАЖДОЙ позиции из ТЗ создай отдельный раздел в отчете.
                            3. Описывай характеристики услуг строго в ПРОШЕДШЕМ времени (напр. 'было обеспечено', 'смонтировано').
                            4. ЗАПРЕЩЕНО писать общие фразы. Если в ТЗ написано 'Стены 4 метра', в отчете должно быть 'Были возведены стены высотой 4 метра'.
                            5. Объем отчета должен соответствовать объему ТЗ."""},
                            {"role": "user", "content": f"Напиши детальный отчет, превратив пункты этой таблицы в выполненные работы:\n\n{pure_tz}\n\nФакты: {facts}"}
                        ]
                    )
                    
                    res_req = client_ai.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role": "user", "content": f"Выпиши из ТЗ список финальных документов (фотоотчет, акты, флешки): {pure_tz[-3000:]}"}]
                    )
                    
                    final_docx = create_report_docx(res_report.choices[0].message.content, st.session_state.title_info, res_req.choices[0].message.content)
                    buf = io.BytesIO()
                    final_docx.save(buf)
                    st.session_state.report_buffer = buf.getvalue()
                    st.success("Отчет готов!")

if st.session_state.report_buffer:
    st.download_button("📥 Скачать готовый отчет", st.session_state.report_buffer, "Report.docx")
