import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 1. ФУНКЦИИ ЧТЕНИЯ ---

def read_docx(file):
    doc = Document(file)
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

# --- 2. ФИКСИРОВАННЫЙ ТИТУЛЬНИК ---

def create_final_report(title_data, tz_processed, req_data):
    doc = Document()
    t = title_data
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # Титульный лист (СТРОГО ПО ВАШЕМУ ОБРАЗЦУ)
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
    tab.rows[0].cells[0].text = f"Отчет принят Заказчиком\n{t.get('customer_fio')}\n\n___________"
    tab.rows[0].cells[1].text = f"Отчет передан Исполнителем\n{t.get('director')}\n\n___________"

    doc.add_page_break()
    # БЛОК 2: ОТЧЕТ
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    doc.add_paragraph(tz_processed)

    doc.add_page_break()
    # БЛОК 3: ТРЕБОВАНИЯ
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(req_data)

    return doc

# --- 3. ИНТЕРФЕЙС ---

st.set_page_config(page_title="Генератор Отчетов (Двухфайловый)")

# Проверка пароля
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    if st.text_input("Пароль", type="password") == st.secrets["APP_PASSWORD"]:
        st.session_state.auth = True
        st.rerun()
    st.stop()

st.header("Шаг 1: Реквизиты из Контракта")
contract_file = st.file_uploader("Загрузите файл КОНТРАКТА (для титульника)", type="docx", key="contract")

if contract_file:
    if st.button("Извлечь данные титульника"):
        client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
        raw_text = read_docx(contract_file)
        # Ограничиваем ИИ только краями документа
        context = raw_text[:3000] + "\n" + raw_text[-3000:]
        
        res = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": f"Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_fio, company, director. Текст: {context}"}],
            response_format={'type': 'json_object'}
        )
        st.session_state.title_info = json.loads(res.choices[0].message.content)
        st.success("Данные титульника сохранены!")

st.divider()

st.header("Шаг 2: Работа с ТЗ")
tz_file = st.file_uploader("Загрузите файл ТЕХЗАДАНИЯ (для отчета)", type="docx", key="tz")

if tz_file and "title_info" in st.session_state:
    if st.button("Преобразовать ТЗ и создать отчет"):
        with st.spinner("Обработка ТЗ в прошедшее время..."):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
            tz_raw_text = read_docx(tz_file)
            
            # 1. Преобразование в прошедшее время
            res_tz = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "Ты — технический редактор. Твоя задача: взять текст ТЗ и переписать его в отчет ПОЛНОСТЬЮ. ГЛАВНОЕ: поменяй все глаголы на прошедшее время (сделано, оказано, выполнено, поставлено). Не сокращай текст, сохрани все детали и пункты."},
                    {"role": "user", "content": f"ПЕРЕПИШИ В ПРОШЕДШЕМ ВРЕМЕНИ:\n\n{tz_raw_text}"}
                ]
            )
            
            # 2. Поиск требований к документам
            res_req = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": f"Найди в этом ТЗ все требования к фотоотчетам, количеству фото и закрывающим документам. Выпиши списком: {tz_raw_text}"}]
            )
            
            # Сборка финального файла
            final_docx = create_final_report(
                st.session_state.title_info, 
                res_tz.choices[0].message.content, 
                res_req.choices[0].message.content
            )
            
            buf = io.BytesIO()
            final_docx.save(buf)
            st.session_state.final_file = buf.getvalue()
            st.success("Отчет успешно сформирован!")

if "final_file" in st.session_state:
    st.download_button("📥 Скачать готовый отчет", st.session_state.final_file, "Final_Report_Full.docx")
