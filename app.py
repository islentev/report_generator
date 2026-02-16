import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. ФУНКЦИИ ИЗВЛЕЧЕНИЯ (БЕЗ ИЗМЕНЕНИЙ) ---

def get_text_ordered(doc):
    """Собирает текст документа, сохраняя последовательность параграфов и таблиц"""
    full_text = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = [p for p in doc.paragraphs if p._element == element]
            if p and p[0].text.strip():
                full_text.append(p[0].text)
        elif element.tag.endswith('tbl'):
            t = [t for t in doc.tables if t._element == element]
            if t:
                for row in t[0].rows:
                    row_data = " | ".join(cell.text.strip() for cell in row.cells)
                    full_text.append(row_data)
    return "\n".join(full_text)

def slice_only_tz(text):
    """Механически вырезает кусок, начиная с Приложения 1"""
    # Ищем маркер начала ТЗ
    start_match = re.search(r"ПРИЛОЖЕНИЕ\s*[№N]?\s*1", text, re.IGNORECASE)
    if not start_match:
        return "ОШИБКА: Заголовок 'Приложение № 1' не найден в документе."
    
    start_idx = start_match.start()
    
    # Ищем маркер конца (Приложение 2 или Расчет стоимости)
    end_match = re.search(r"(ПРИЛОЖЕНИЕ\s*[№N]?\s*2|РАСЧЕТ\s*СТОИМОСТИ)", text[start_idx:], re.IGNORECASE)
    
    if end_match:
        return text[start_idx : start_idx + end_match.start()]
    else:
        # Если конца нет, берем все до конца документа
        return text[start_idx:]

# --- 2. СБОРКА DOCX (ФИКСИРОВАННЫЙ ТИТУЛЬНИК) ---

def create_final_report(title_data, tz_content, req_content):
    doc = Document()
    t = title_data

    # ШРИФТ ПО УМОЛЧАНИЮ
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # --- БЛОК 1: ТИТУЛЬНИК (ЗАФИКСИРОВАНО) ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Информационно-аналитический отчет об исполнении условий\n").bold = True
    p.add_run(f"Контракта № {t.get('contract_no')} от «{t.get('contract_date')}» 2025 г.\n").bold = True
    p.add_run(f"Идентификационный код закупки: {t.get('ikz')}.")

    for _ in range(5): doc.add_paragraph()
    doc.add_paragraph("ТОМ I").alignment = WD_ALIGN_PARAGRAPH.CENTER

    for label, val in [("Наименование предмета КОНТРАКТА :", t.get('project_name')), 
                      ("Заказчик:", t.get('customer')), 
                      ("Исполнитель:", t.get('company'))]:
        p_l = doc.add_paragraph(); p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.add_run(label).bold = True
        p_v = doc.add_paragraph(); p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(str(val)).italic = True

    for _ in range(5): doc.add_paragraph()
    
    tab = doc.add_table(rows=1, cols=2)
    tab.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{t.get('customer_fio')}\n\n___________ / С.В. Куц"
    tab.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{t.get('director')}\n\n___________ / Е.В. Гринин"

    doc.add_page_break()

    # --- БЛОК 2: ОТЧЕТ (ПРОСТОЕ КОПИРОВАНИЕ) ---
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    doc.add_paragraph(tz_content)

    doc.add_page_break()

    # --- БЛОК 3: ТРЕБОВАНИЯ (К ДОКУМЕНТАЦИИ) ---
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(req_content)

    return doc

# --- 3. СТРИМЛИТ ИНТЕРФЕЙС ---

st.set_page_config(page_title="Генератор Отчетов (Поэтапный)")

# Пароль (возвращаем как было)
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    pwd = st.text_input("Введите пароль", type="password")
    if pwd == st.secrets["APP_PASSWORD"]:
        st.session_state.authenticated = True
        st.rerun()
    st.stop()

uploaded_file = st.file_uploader("Загрузите файл", type="docx")

if uploaded_file:
    # ПОЛНОЕ ОБНУЛЕНИЕ ПРИ ЗАГРУЗКЕ НОВОГО ФАЙЛА
    if "current_file" not in st.session_state or st.session_state.current_file != uploaded_file.name:
        st.session_state.clear()
        st.session_state.authenticated = True # Сохраняем вход
        st.session_state.current_file = uploaded_file.name
        st.rerun()

    doc_obj = Document(uploaded_file)
    text_data = get_text_ordered(doc_obj)
    client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")

    # ШАГ 1: ТИТУЛЬНИК
    if st.button("Шаг 1: Сформировать Титульник"):
        # Даем ИИ только начало и конец для реквизитов
        context = text_data[:4000] + text_data[-4000:]
        res = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {context}"}],
            response_format={'type': 'json_object'}
        )
        st.session_state.title_data = json.loads(res.choices[0].message.content)
        st.success("Титульник зафиксирован")

    # ШАГ 2: ОТЧЕТ
    if "title_data" in st.session_state:
        if st.button("Шаг 2: Скопировать ТЗ в отчет"):
            # 1. Программная вырезка (ИИ не увидит ничего кроме ТЗ)
            pure_tz = slice_only_tz(text_data)
            
            # 2. Передаем ИИ с запретом на изменения
            res_tz = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "Ты — копировальный аппарат. Твоя задача: взять текст и выдать его БЕЗ ИЗМЕНЕНИЙ. Не меняй время, не сокращай, не добавляй вводных слов. Просто копия текста."},
                    {"role": "user", "content": f"СКОПИРУЙ ЭТОТ ТЕКСТ ПОЛНОСТЬЮ:\n\n{pure_tz}"}
                ]
            )
            
            # 3. Доп требования (поиск фото)
            res_req = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Выпиши требования к количеству фото и документам для сдачи из этого текста: {pure_tz[-5000:]}"}]
            )
            
            # Генерация
            final_doc = create_final_report(st.session_state.title_data, res_tz.choices[0].message.content, res_req.choices[0].message.content)
            
            buf = io.BytesIO()
            final_doc.save(buf)
            st.session_state.result_file = buf.getvalue()
            st.success("Отчет готов (ТЗ скопировано полностью)")

    if "result_file" in st.session_state:
        st.download_button("📥 Скачать отчет", st.session_state.result_file, "Report_Fixed.docx")
