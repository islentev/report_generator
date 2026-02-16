import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. УНИВЕРСАЛЬНЫЙ ДВИЖОК ИЗВЛЕЧЕНИЯ ---

def get_text_from_docx(file):
    """Читает всё: и параграфы, и таблицы в порядке их следования"""
    doc = Document(file)
    full_text = []
    for element in doc.element.body:
        if element.tag.endswith('p'): # Параграф
            p = [p for p in doc.paragraphs if p._element == element]
            if p: full_text.append(p[0].text)
        elif element.tag.endswith('tbl'): # Таблица
            t = [t for t in doc.tables if t._element == element]
            if t:
                for row in t[0].rows:
                    full_text.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(full_text)

def extract_universal_tz(text):
    """
    Ищет ТЗ по ключевым словам, независимо от того, таблица это или текст.
    """
    # Список возможных заголовков начала ТЗ
    start_patterns = [
        r"ПРИЛОЖЕНИЕ\s*[№N]?\s*1", 
        r"ТЕХНИЧЕСКОЕ\s*ЗАДАНИЕ",
        r"ОПИСАНИЕ\s*ОБЪЕКТА\s*ЗАКУПКИ"
    ]
    
    start_idx = -1
    for pattern in start_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            start_idx = match.start()
            break
            
    if start_idx == -1:
        # Если маркеры не найдены, берем последние 60% документа (ТЗ всегда в конце)
        return text[int(len(text)*0.4):]
    
    # Ищем конец ТЗ (обычно это Приложение 2 или Расчет стоимости)
    end_patterns = [r"ПРИЛОЖЕНИЕ\s*[№N]?\s*2", r"РАСЧЕТ\s*СТОИМОСТИ"]
    end_idx = len(text)
    for pattern in end_patterns:
        match = re.search(pattern, text[start_idx + 100:], re.IGNORECASE)
        if match:
            end_idx = start_idx + 100 + match.start()
            break
            
    return text[start_idx:end_idx]

# --- 2. ВАШИ ФУНКЦИИ (БЕЗ ИЗМЕНЕНИЙ) ---

def format_fio(fio):
    if not fio: return "___________"
    parts = str(fio).split()
    if len(parts) >= 3: return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio

def create_report_docx(report_content, title_data, req_list):
    doc = Document()
    t = title_data
    # Титульник
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Информационно-аналитический отчет\nКонтракт № {t.get('contract_no')}").bold = True
    for _ in range(8): doc.add_paragraph()
    
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = f"Заказчик: {format_fio(t.get('customer_fio'))}"
    table.rows[0].cells[1].text = f"Исполнитель: {format_fio(t.get('director'))}"
    
    doc.add_page_break()
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    doc.add_paragraph(report_content)
    
    doc.add_page_break()
    doc.add_heading('ТРЕБОВАНИЯ К ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(req_list)
    return doc

# --- 3. ИНТЕРФЕЙС ---

st.set_page_config(page_title="Универсальный Генератор", layout="wide")

# Пароль
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    if st.text_input("Пароль", type="password") == st.secrets["APP_PASSWORD"]:
        st.session_state.auth = True
        st.rerun()
    st.stop()

up_file = st.file_uploader("Загрузить контракт", type="docx")

if up_file:
    # ОБНУЛЕНИЕ ПРИ НОВОМ ФАЙЛЕ
    if "fname" not in st.session_state or st.session_state.fname != up_file.name:
        st.session_state.fname = up_file.name
        st.session_state.t_info = None
        st.session_state.res_doc = None

    full_contract_text = get_text_from_docx(up_file)
    client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")

    if st.button("Шаг 1: Извлечь реквизиты"):
        # Даем ИИ только самое начало и самый конец (где реквизиты)
        ctx = full_contract_text[:4000] + full_contract_text[-4000:]
        res = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {ctx}"}],
            response_format={'type': 'json_object'}
        )
        st.session_state.t_info = json.loads(res.choices[0].message.content)
        st.success("Реквизиты получены")

    if st.session_state.get("t_info"):
        if st.button("Шаг 2: Сформировать полный отчет"):
            with st.spinner("Извлекаю ТЗ..."):
                # УНИВЕРСАЛЬНОЕ ИЗВЛЕЧЕНИЕ
                tz_body = extract_universal_tz(full_contract_text)
                
                # Показываем для проверки
                with st.expander("Проверка извлеченного ТЗ"):
                    st.text(tz_body[:1000] + "...")

                # Промпт на полное копирование
                res_rep = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "Ты — технический писатель. Тебе дали текст ТЕХНИЧЕСКОГО ЗАДАНИЯ. Перепиши его в отчет ПОЛНОСТЬЮ. Не сокращай пункты. Если это список — перенеси списком. Если это текст — перенеси текстом. Используй прошедшее время (оказано, выполнено)."},
                        {"role": "user", "content": f"ПЕРЕНЕСИ ВСЕ ПУНКТЫ УСЛУГ В ОТЧЕТ:\n\n{tz_body}"}
                    ]
                )
                
                # Требования (ищем в том же хвосте)
                res_req = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди требования к фотоотчету и документам в этом тексте: {tz_body[-5000:]}"}]
                )
                
                final_docx = create_report_docx(res_rep.choices[0].message.content, st.session_state.t_info, res_req.choices[0].message.content)
                b = io.BytesIO()
                final_docx.save(b)
                st.session_state.res_doc = b.getvalue()
                st.success("Отчет готов")

    if st.session_state.get("res_doc"):
        st.download_button("📥 Скачать отчет", st.session_state.res_doc, "Final_Report.docx")
