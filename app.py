import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. ФУНКЦИИ ИЗВЛЕЧЕНИЯ (ТОЛЬКО КОПИРОВАНИЕ) ---

def get_full_text(doc):
    """Собирает текст параграфов и таблиц в один поток"""
    full_text = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = [p for p in doc.paragraphs if p._element == element]
            if p: full_text.append(p[0].text)
        elif element.tag.endswith('tbl'):
            t = [t for t in doc.tables if t._element == element]
            if t:
                for row in t[0].rows:
                    full_text.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(full_text)

def find_only_tz_content(text):
    """Находит начало ТЗ и отрезает всё, что было ДО него"""
    # Ищем Приложение №1 или Техническое задание
    match = re.search(r"(ПРИЛОЖЕНИЕ\s*[№N]?\s*1|ТЕХНИЧЕСКОЕ\s*ЗАДАНИЕ)", text, re.IGNORECASE)
    if not match:
        return text # Если не нашли маркер, отдаем всё (страховка)
    
    start_pos = match.start()
    # Ищем конец ТЗ (Приложение №2)
    end_match = re.search(r"ПРИЛОЖЕНИЕ\s*[№N]?\s*2", text[start_pos:], re.IGNORECASE)
    
    if end_match:
        return text[start_pos : start_pos + end_match.start()]
    return text[start_pos:]

# --- 2. ФИКСИРОВАННЫЙ ТИТУЛЬНИК И СТРУКТУРА ---

def create_final_docx(report_body, title_info, requirements):
    doc = Document()
    t = title_info

    # --- БЛОК 1: ТИТУЛЬНЫЙ ЛИСТ (ЗАФИКСИРОВАНО) ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Информационно-аналитический отчет об исполнении условий\n")
    run.bold = True
    run2 = p.add_run(f"Контракта № {t.get('contract_no', '___')} от {t.get('contract_date', '___')}\n")
    run2.bold = True
    p.add_run(f"Идентификационный код закупки: {t.get('ikz', '___')}")

    for _ in range(5): doc.add_paragraph()
    doc.add_paragraph("ТОМ I").alignment = WD_ALIGN_PARAGRAPH.CENTER

    labels = [("Наименование предмета КОНТРАКТА :", t.get('project_name')), 
              ("Заказчик:", t.get('customer')), 
              ("Исполнитель:", t.get('company'))]
    
    for label, val in labels:
        p1 = doc.add_paragraph(); p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.add_run(label).bold = True
        p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.add_run(str(val)).italic = True

    for _ in range(5): doc.add_paragraph()
    
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{t.get('customer_fio', '___')}\n\n___________"
    table.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{t.get('director', '___')}\n\n___________"

    doc.add_page_break()

    # --- БЛОК 2: ОТЧЕТ (ЧИСТОЕ КОПИРОВАНИЕ ТЗ) ---
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    doc.add_paragraph(report_body)

    doc.add_page_break()

    # --- БЛОК 3: ТРЕБОВАНИЯ (ЗАФИКСИРОВАНО) ---
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(requirements)

    return doc

# --- 3. ИНТЕРФЕЙС ---

st.set_page_config(page_title="Генератор (Только Копирование)")

# Проверка пароля
if "pass_ok" not in st.session_state: st.session_state.pass_ok = False
if not st.session_state.pass_ok:
    if st.text_input("Пароль", type="password") == st.secrets["APP_PASSWORD"]:
        st.session_state.pass_ok = True
        st.rerun()
    st.stop()

file = st.file_uploader("Загрузите контракт", type="docx")

if file:
    # ОБНУЛЕНИЕ ПРИ НОВОМ ФАЙЛЕ
    if "current_fname" not in st.session_state or st.session_state.current_fname != file.name:
        st.session_state.clear()
        st.session_state.current_fname = file.name
        st.session_state.pass_ok = True # Чтобы не выкинуло
        st.rerun()

    doc_obj = Document(file)
    text_data = get_full_text(doc_obj)
    client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")

    if st.button("Шаг 1: Подготовить Титульник"):
        # Реквизиты берем из начала и конца
        ctx = text_data[:4000] + text_data[-4000:]
        res = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {ctx}"}],
            response_format={'type': 'json_object'}
        )
        st.session_state.title_data = json.loads(res.choices[0].message.content)
        st.success("Титульник готов")

    if st.session_state.get("title_data"):
        if st.button("Шаг 2: Создать отчет (Полное копирование ТЗ)"):
            with st.spinner("Ищу ТЗ и копирую..."):
                # Находим только мясо ТЗ
                pure_tz = find_only_tz_content(text_data)
                
                # Команда ИИ: просто перенести текст без изменений
                res_copy = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "Ты — технический копировщик. Твоя единственная задача: ПЕРЕНЕСТИ ТЕКСТ ТЗ ПОЛНОСТЬЮ. Не сокращай, не меняй время глаголов, не делай выводы. Просто выдай тот же текст, что тебе прислали."},
                        {"role": "user", "content": f"СКОПИРУЙ ЭТОТ ТЕКСТ БЕЗ ИЗМЕНЕНИЙ:\n\n{pure_tz}"}
                    ]
                )
                
                # Доп. требования (поиск фото и т.д.)
                res_req = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди требования к фотоотчету и количеству фото в этом тексте: {pure_tz[-5000:]}"}]
                )
                
                # Сборка итогового файла
                final_file = create_final_docx(
                    res_copy.choices[0].message.content, 
                    st.session_state.title_data, 
                    res_req.choices[0].message.content
                )
                
                buf = io.BytesIO()
                final_file.save(buf)
                st.session_state.final_out = buf.getvalue()
                st.success("Отчет собран")

    if st.session_state.get("final_out"):
        st.download_button("📥 Скачать готовый отчет", st.session_state.final_out, "Full_Copy_Report.docx")
