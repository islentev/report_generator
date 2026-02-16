import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. УТИЛИТЫ (ОБНУЛЕНИЕ И ЖЕСТКИЙ ПОИСК ТЗ) ---

def get_full_text_including_tables(doc):
    """Сборка текста и таблиц в единую структуру"""
    full_element_list = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = [p for p in doc.paragraphs if p._element == element]
            if para and para[0].text.strip():
                full_element_list.append(para[0].text)
        elif element.tag.endswith('tbl'):
            table = [t for t in doc.tables if t._element == element]
            if table:
                full_element_list.append("\n[ТАБЛИЦА ТЗ]")
                for row in table[0].rows:
                    row_data = " | ".join(cell.text.strip() for cell in row.cells)
                    full_element_list.append(row_data)
                full_element_list.append("[КОНЕЦ ТАБЛИЦЫ]\n")
    return "\n".join(full_element_list)

def extract_tz_content_v2(text):
    """Находит ТЗ через регулярные выражения (игнорируя ошибки написания)"""
    # Ищем Приложение 1 (любые пробелы, № или N)
    start_match = re.search(r"ПРИЛОЖЕНИЕ\s*[№N]?\s*1", text, re.IGNORECASE)
    # Ищем Приложение 2
    end_match = re.search(r"ПРИЛОЖЕНИЕ\s*[№N]?\s*2", text, re.IGNORECASE)
    
    if not start_match:
        return text[len(text)//2:] # Если не нашли, берем вторую половину документа (там обычно ТЗ)
    
    start_idx = start_match.start()
    end_idx = end_match.start() if end_match else len(text)
    
    return text[start_idx:end_idx]

def format_fio_universal(raw_fio):
    if not raw_fio or len(raw_fio) < 5: return "________________"
    clean = re.sub(r'(директор|министр|заместитель|начальник|председатель|генеральный)', '', raw_fio, flags=re.IGNORECASE).strip()
    parts = clean.split()
    if len(parts) >= 3: return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return clean

# --- 2. ГЕНЕРАЦИЯ DOCX (ВАШ ТИТУЛЬНИК) ---

def create_report_docx(report_content, title_data, req_list):
    doc = Document()
    t = title_data
    
    # Настройка шрифта
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # ТИТУЛЬНЫЙ ЛИСТ
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Информационно-аналитический отчет об исполнении условий\n").bold = True
    p.add_run(f"Контракта № {t.get('contract_no', '___')} от {t.get('contract_date', '___')}\n").bold = True
    p.add_run(f"Идентификационный код закупки: {t.get('ikz', '___')}")

    for _ in range(5): doc.add_paragraph()
    doc.add_paragraph("ТОМ I").alignment = WD_ALIGN_PARAGRAPH.CENTER

    for label, val in [("Предмет:", t.get('project_name')), ("Заказчик:", t.get('customer')), ("Исполнитель:", t.get('company'))]:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(label).bold = True
        p_v = doc.add_paragraph(); p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(str(val)).italic = True

    for _ in range(5): doc.add_paragraph()
    
    # Таблица подписей
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = f"Заказчик:\n{format_fio_universal(t.get('customer_fio'))}\n\n___________"
    table.rows[0].cells[1].text = f"Исполнитель:\n{format_fio_universal(t.get('director'))}\n\n___________"

    doc.add_page_break()
    doc.add_heading('ОТЧЕТ ПО ТЕХНИЧЕСКОМУ ЗАДАНИЮ', level=1)
    doc.add_paragraph(report_content)
    
    doc.add_page_break()
    doc.add_heading('ТРЕБОВАНИЯ К ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(req_list)
    
    return doc

# --- 3. ИНТЕРФЕЙС STREAMLIT ---

st.set_page_config(page_title="Генератор 2.0", layout="wide")

# ВВОД ПАРОЛЯ (Вернул)
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    pwd = st.text_input("Введите пароль", type="password")
    if pwd == st.secrets["APP_PASSWORD"]:
        st.session_state.authenticated = True
        st.rerun()
    st.stop()

# ИНИЦИАЛИЗАЦИЯ (ОБНУЛЕНИЕ ПРИ НОВОМ ФАЙЛЕ)
uploaded_file = st.file_uploader("Загрузите файл Контракта", type="docx")

if uploaded_file:
    # Если загружен новый файл - чистим старые данные в памяти
    if "last_file" not in st.session_state or st.session_state.last_file != uploaded_file.name:
        st.session_state.title_info = None
        st.session_state.report_done = None
        st.session_state.last_file = uploaded_file.name

    doc_obj = Document(uploaded_file)
    full_text = get_full_text_including_tables(doc_obj)
    
    client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")

    # ШАГ 1: ТИТУЛЬНИК
    if st.button("Шаг 1: Извлечь данные титульника"):
        with st.spinner("Анализирую начало и конец..."):
            # Берем только первые и последние 4к знаков, чтобы ИИ не путался в ТЗ на этом этапе
            ctx = full_text[:4000] + "\n" + full_text[-4000:]
            res = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Верни JSON с полями: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {ctx}"}],
                response_format={'type': 'json_object'}
            )
            st.session_state.title_info = json.loads(res.choices[0].message.content)
            st.success("Данные титульника сохранены")

    # ШАГ 2: ОТЧЕТ (КОПИРОВАНИЕ ТЗ)
    if st.session_state.title_info:
        if st.button("Шаг 2: Скопировать ТЗ в отчет"):
            with st.spinner("Вырезаю ТЗ и копирую..."):
                # 1. Жестко вырезаем кусок
                pure_tz = extract_tz_content_v2(full_text)
                
                # 2. ИИ получает ТОЛЬКО вырезанный кусок ТЗ
                res_copy = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "Ты — технический ассистент. Твоя задача: взять предоставленный текст ТЗ и переписать его в отчет ПОЛНОСТЬЮ. ЗАПРЕЩЕНО сокращать. Переноси все параметры и таблицы. Не используй информацию из начала контракта."},
                        {"role": "user", "content": f"СКОПИРУЙ ВЕСЬ ЭТОТ ТЕКСТ:\n\n{pure_tz}"}
                    ]
                )
                
                # 3. Документы (берем из хвоста вырезанного ТЗ)
                res_req = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Выпиши требования к фото и документам из этого текста: {pure_tz[-4000:]}"}]
                )
                
                # Собираем файл
                final_docx = create_report_docx(res_copy.choices[0].message.content, st.session_state.title_info, res_req.choices[0].message.content)
                
                buf = io.BytesIO()
                final_docx.save(buf)
                st.session_state.report_done = buf.getvalue()

    if st.session_state.get("report_done"):
        st.download_button("📥 Скачать полный отчет", st.session_state.report_done, "Final_Report.docx")
