import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. УТИЛИТЫ ЧТЕНИЯ (Улучшенный захват таблиц) ---

def get_full_text_including_tables(doc):
    """Превращает весь документ в текст, помечая таблицы для ИИ"""
    full_element_list = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = [p for p in doc.paragraphs if p._element == element]
            if para and para[0].text.strip():
                full_element_list.append(para[0].text)
        elif element.tag.endswith('tbl'):
            table = [t for t in doc.tables if t._element == element]
            if table:
                full_element_list.append("\n[НАЧАЛО ТАБЛИЦЫ ТЗ]")
                for row in table[0].rows:
                    row_data = " | ".join(cell.text.strip() for cell in row.cells)
                    full_element_list.append(row_data)
                full_element_list.append("[КОНЕЦ ТАБЛИЦЫ ТЗ]\n")
    return "\n".join(full_element_list)

def extract_tz_content(full_text):
    """Механически вырезает кусок ТЗ из общего текста"""
    text_upper = full_text.upper()
    # Ищем начало Приложения №1
    start_marker = "ПРИЛОЖЕНИЕ № 1"
    start_pos = text_upper.find(start_marker)
    
    if start_pos == -1:
        return "ОШИБКА: Не удалось найти заголовок 'Приложение № 1'. Проверьте текст контракта."

    # Ищем маркер конца (Приложение №2 или Расчет стоимости)
    end_marker = "ПРИЛОЖЕНИЕ № 2"
    end_pos = text_upper.find(end_marker, start_pos + 100)
    
    if end_pos == -1:
        end_pos = text_upper.find("РАСЧЕТ СТОИМОСТИ", start_pos + 100)
    
    if end_pos == -1:
        # Если конца не нашли, берем 30.000 знаков (чтобы влезло всё ТЗ)
        return full_text[start_pos:start_pos + 30000]
        
    return full_text[start_pos:end_pos]

def format_fio_universal(raw_fio):
    if not raw_fio or len(raw_fio) < 5: return "________________"
    clean = re.sub(r'(директор|министр|заместитель|начальник|председатель|генеральный)', '', raw_fio, flags=re.IGNORECASE).strip()
    parts = clean.split()
    if len(parts) >= 3: return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    if len(parts) == 2: return f"{parts[0]} {parts[1][0]}."
    return clean

# --- 2. КОНСТРУКЦИЯ ТИТУЛЬНИКА (СОХРАНЕНА) ---

def create_report_docx(report_content, title_data, requirements_list):
    doc = Document()
    # Данные для титульника
    t = title_data
    cust_fio = format_fio_universal(t.get('customer_fio', ''))
    exec_fio = format_fio_universal(t.get('director', ''))

    # Стили
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # Титульный лист (Ваш блок)
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_top.add_run(f"Информационно-аналитический отчет об исполнении условий\n").bold = True
    p_top.add_run(f"Контракта № {t.get('contract_no')} от «{t.get('contract_date')}» 2025 г.\n").bold = True
    p_top.add_run(f"Идентификационный код закупки: {t.get('ikz')}.")

    for _ in range(3): doc.add_paragraph()
    doc.add_paragraph("ТОМ I").alignment = WD_ALIGN_PARAGRAPH.CENTER

    for label, val in [("Предмет:", t.get('project_name')), ("Заказчик:", t.get('customer')), ("Исполнитель:", t.get('company'))]:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(label).bold = True
        p_v = doc.add_paragraph(); p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(str(val)).italic = True

    for _ in range(4): doc.add_paragraph()
    tab = doc.add_table(rows=1, cols=2)
    tab.rows[0].cells[0].text = f"ПРИНЯЛ: {cust_fio}"
    tab.rows[0].cells[1].text = f"ПЕРЕДАЛ: {exec_fio}"

    doc.add_page_break()

    # --- БЛОК ОТЧЕТА (Теперь здесь полное копирование) ---
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    doc.add_paragraph(report_content)

    doc.add_page_break()
    # --- БЛОК ТРЕБОВАНИЙ (СОХРАНЕН) ---
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(requirements_list)

    return doc

# --- 3. ЛОГИКА СТРИМЛИТ ---
st.set_page_config(page_title="Поэтапное обучение ИИ", layout="wide")

if 'title_info' not in st.session_state: st.session_state.title_info = None

client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")

file = st.file_uploader("Загрузите контракт", type="docx")

if file:
    doc_obj = Document(file)
    full_text = get_full_text_including_tables(doc_obj)

    if st.button("Шаг 1: Собрать Титульник"):
        res = client_ai.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {full_text[:4000]}"}],
            response_format={'type': 'json_object'}
        )
        st.session_state.title_info = json.loads(res.choices[0].message.content)
        st.json(st.session_state.title_info)

    if st.session_state.title_info:
        if st.button("Шаг 2: Полное копирование ТЗ в Отчет"):
            # 1. Находим ТЗ программно
            pure_tz = extract_tz_content(full_text)
            
            # 2. Заставляем ИИ перенести текст БЕЗ ИЗМЕНЕНИЙ
            res_copy = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "Ты — копировщик. Твоя задача: взять текст ТЗ и перенести его в отчет ПОЛНОСТЬЮ. Не меняй слова, не сокращай, не резюмируй. Сохрани все пункты и описания услуг. Если видишь таблицу — перенеси её данные текстом."},
                    {"role": "user", "content": f"СКОПИРУЙ ЭТОТ ТЕКСТ ПОЛНОСТЬЮ БЕЗ ИЗМЕНЕНИЙ:\n\n{pure_tz}"}
                ]
            )
            
            # 3. Находим требования (хвост документа)
            res_req = client_ai.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Выпиши требования к фото и документам из конца этого текста: {full_text[-5000:]}"}]
            )
            
            # Сборка
            final_doc = create_report_docx(res_copy.choices[0].message.content, st.session_state.title_info, res_req.choices[0].message.content)
            
            buf = io.BytesIO()
            final_doc.save(buf)
            st.download_button("Скачать результат копирования", buf.getvalue(), "Full_Copy_Report.docx")
