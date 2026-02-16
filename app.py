import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json
import re

# --- 0. УТИЛИТЫ (Ваши функции) ---

def extract_tz_content(full_text):
    # Очищаем текст от лишних пробелов для поиска
    text_clean = " ".join(full_text.split())
    text_upper = text_clean.upper()
    
    # Ищем максимально точные маркеры начала
    start_markers = [
        "ПРИЛОЖЕНИЕ № 1 К КОНТРАКТУ", 
        "ТЕХНИЧЕСКОЕ ЗАДАНИЕ", 
        "ОПИСАНИЕ ОБЪЕКТА ЗАКУПКИ"
    ]
    
    start_index = -1
    for marker in start_markers:
        found = text_upper.find(marker)
        if found != -1:
            # Берем индекс из оригинального текста, чтобы не потерять форматирование
            # Ищем это же место в full_text
            start_index = full_text.upper().find(marker)
            break
    
    if start_index == -1:
        return "ОШИБКА: ТЗ не найдено. ИИ не видит раздел 'Приложение № 1'."

    # Ищем конец ТЗ - обычно это Приложение № 2 или Расчет стоимости
    end_markers = ["ПРИЛОЖЕНИЕ № 2", "РАСЧЕТ СТОИМОСТИ", "ПОДПИСИ СТОРОН"]
    end_index = len(full_text)
    
    for marker in end_markers:
        found_end = full_text.upper().find(marker, start_index + 100)
        if found_end != -1:
            end_index = found_end
            break
            
    return full_text[start_index:end_index]

def format_fio_universal(raw_fio):
    if not raw_fio or len(raw_fio) < 5: return "________________"
    clean = re.sub(r'(директор|министр|заместитель|начальник|председатель|генеральный)', '', raw_fio, flags=re.IGNORECASE).strip()
    parts = clean.split()
    if len(parts) >= 3: return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    if len(parts) == 2: return f"{parts[0]} {parts[1][0]}."
    return clean

# --- 1. НАСТРОЙКА ---
st.set_page_config(page_title="Юридический Генератор v3", layout="wide")

# Инициализация состояний
for key in ['title_info', 'report_text', 'req_text', 'report_buffer', 'last_file']:
    if key not in st.session_state: st.session_state[key] = None

try:
    client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"].strip().strip('"'), base_url="https://api.deepseek.com/v1")
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфига: {e}"); st.stop()

# --- 2. ФУНКЦИЯ СОЗДАНИЯ DOCX (Ваша конструкция Титульника) ---
def create_report_docx(report_content, title_data, requirements_list):
    doc = Document()
    
    # Подготовка данных
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

    # Стиль Times New Roman 12
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # --- ТИТУЛЬНЫЙ ЛИСТ (Ваша структура на 90%) ---
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p_top.add_run("Информационно-аналитический отчет об исполнении условий\n")
    run1.bold = True
    run2 = p_top.add_run(f"Контракта № {contract_no} от «{contract_date}» 2025 г.\n")
    run2.bold = True
    p_top.add_run(f"Идентификационный код закупки: {ikz}.")

    for _ in range(3): doc.add_paragraph()

    p_tom = doc.add_paragraph()
    p_tom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tom.add_run("ТОМ I").bold = True

    labels_values = [
        ("Наименование предмета КОНТРАКТА :", project_name),
        ("Заказчик:", customer),
        ("Исполнитель:", company)
    ]
    
    for label, value in labels_values:
        p_h = doc.add_paragraph()
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h.add_run(label).bold = True
        p_v = doc.add_paragraph()
        p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.add_run(value).italic = True

    for _ in range(4): doc.add_paragraph()

    # Таблица подписей
    table = doc.add_table(rows=2, cols=2)
    table.width = doc.sections[0].page_width
    
    p_l = table.rows[0].cells[0].paragraphs[0]
    p_l.add_run("Отчет принят Заказчиком").bold = True
    p_l.add_run(f"\n\n{cust_post} {cust_fio}\n\n_______________")
    
    p_r = table.rows[0].cells[1].paragraphs[0]
    p_r.add_run("Отчет передан Исполнителем").bold = True
    p_r.add_run(f"\n\n{exec_post}\n\n_______________ / {exec_fio}")
    
    table.rows[1].cells[0].paragraphs[0].add_run("м.п.")
    table.rows[1].cells[1].paragraphs[0].add_run("м.п.")

    doc.add_page_break()

    # --- ТЕКСТ ОТЧЕТА (Блок 2) ---
    doc.add_heading('ОТЧЕТ О ВЫПОЛНЕНИИ ТЕХНИЧЕСКОГО ЗАДАНИЯ', level=1)
    for block in report_content.split('\n\n'):
        p = doc.add_paragraph()
        for part in block.split('**'):
            run = p.add_run(part.replace('*', ''))
            if part in block.split('**')[1::2]: run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)

    doc.add_page_break()
    
    # --- ТРЕБОВАНИЯ (Блок 3) ---
    doc.add_heading('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ', level=1)
    doc.add_paragraph(requirements_list)

    return doc

# --- 3. ИНТЕРФЕЙС ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD: st.stop()

uploaded_file = st.file_uploader("Загрузите контракт (DOCX)", type="docx")

if uploaded_file:
    if st.session_state.last_file != uploaded_file.name:
        st.session_state.title_info = None
        st.session_state.report_text = None
        st.session_state.req_text = None
        st.session_state.report_buffer = None
        st.session_state.last_file = uploaded_file.name

    doc_obj = Document(uploaded_file)
    full_text = "\n".join([p.text for p in doc_obj.paragraphs])

    # Разделение по табам для пошаговой работы
    tab1, tab2, tab3 = st.tabs(["Шаг 1: Титульник", "Шаг 2: Отчет (ТЗ)", "Шаг 3: Требования"])

    with tab1:
        if st.button("Извлечь данные титульника"):
            with st.spinner("Анализ реквизитов..."):
                # Изолируем контекст: только начало и конец
                context = full_text[:2000] + "\n" + full_text[-3000:]
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди данные для титульного листа (номера, даты, ИКЗ, полные ФИО и должности). Верни JSON: contract_no, contract_date, ikz, project_name, customer, customer_post, customer_fio, company, executor_post, director. Текст: {context}"}],
                    response_format={ 'type': 'json_object' }
                )
                st.session_state.title_info = json.loads(res.choices[0].message.content)
        if st.session_state.title_info:
            st.json(st.session_state.title_info)

    with tab2:
        if st.button("Сгенерировать текст отчета по ТЗ"):
            with st.spinner("Вырезаю ТЗ из контракта..."):
                pure_tz = extract_tz_content(full_text)
                
                # Проверка для вас: выводим в консоль или лог, что именно мы нашли
                if len(pure_tz) < 500:
                    st.error("Извлеченный кусок текста слишком мал. Скорее всего, ТЗ не захвачено.")
                
                res_report = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": """Ты технический эксперт. 
                        ТЕБЕ ЗАПРЕЩЕНО использовать разделы 'Предмет контракта', 'Права и обязанности', 'Сроки'.
                        РАБОТАЙ ТОЛЬКО С ТАБЛИЦЕЙ ТЗ.
                        
                        ИНСТРУКЦИЯ:
                        1. Найди в присланном тексте перечень услуг (например: Аренда, Застройка, Монтаж).
                        2. Сделай каждую услугу заголовком (Настоящее время).
                        3. Под заголовком напиши детализацию из столбца 'Технические характеристики' в ПРОШЕДШЕМ времени (Например: 'Было обеспечено...', 'Произведен монтаж...').
                        4. Если в тексте нет конкретных услуг, напиши 'УСЛУГИ НЕ НАЙДЕНЫ'."""},
                        {"role": "user", "content": f"Вот текст Приложения №1. Напиши отчет строго по нему:\n\n{pure_tz}"}
                    ]
                )
                st.session_state.report_text = res_report.choices[0].message.content

    with tab3:
        if st.button("Собрать список документов"):
            with st.spinner("Поиск требований..."):
                # Контекст только для документов
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Выпиши список отчетных документов (акты, фото и т.д.) из текста:\n{full_text[-4000:]}"}]
                )
                st.session_state.req_text = res.choices[0].message.content
        if st.session_state.req_text:
            st.write(st.session_state.req_text)

    # Финальная сборка
    st.divider()
    if st.button("Собрать финальный файл DOCX"):
        if st.session_state.title_info and st.session_state.report_text and st.session_state.req_text:
            doc_final = create_report_docx(st.session_state.report_text, st.session_state.title_info, st.session_state.req_text)
            buf = io.BytesIO()
            doc_final.save(buf)
            st.session_state.report_buffer = buf.getvalue()
            st.success("Документ собран!")
        else:
            st.error("Сначала выполните все три шага!")

if st.session_state.report_buffer:
    st.download_button("📥 Скачать готовый отчет", st.session_state.report_buffer, "final_report.docx")



