import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI
import io
import json
import re

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

def get_contract_start_text(file):
    """Считывает текст только до начала 2-го раздела (Предмет контракта)"""
    doc = Document(file)
    start_text = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            start_text.append(txt)
            # Если строка начинается с "2." (например, 2. ЦЕНА КОНТРАКТА), стоп.
            if re.match(r"^2\.", txt): 
                break
    return "\n".join(start_text)

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

# --- ПАРОЛЬ В БОКОВОЙ ПАНЕЛИ ---
with st.sidebar:
    st.title("Авторизация")
    if "auth" not in st.session_state: 
        st.session_state.auth = False
    pwd = st.text_input("Введите пароль", type="password")
    if pwd == st.secrets["APP_PASSWORD"]:
        st.session_state.auth = True
    if not st.session_state.auth:
        st.warning("Доступ ограничен. Введите пароль в поле выше.")
        st.stop()
    st.success("Доступ разрешен")

col1, col2 = st.columns(2)

# СТОЛБЕЦ 1: ТИТУЛЬНЫЙ ЛИСТ
with col1:
    st.header("📄 1. Титульный лист")
    file_contract = st.file_uploader(
    "Загрузите Контракт", 
    type="docx", 
    key="contract_loader", 
    on_change=lambda: st.session_state.pop("t_info", None) # Удаляет данные старого контракта
    )

    if file_contract:
        if st.button("Сформировать Титульный лист"):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")

            # Читаем только начало контракта
            context = get_contract_start_text(file_contract)
            res = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {
                        "role": "system", "content": """Ты — строгий аналитик. Работай ТОЛЬКО с тем текстом, который передан в текущем запросе. Верни результат строго в формате JSON.
                        Запрещено использовать любые данные из предыдущих диалогов или внешних знаний. 
                        Если данные (номер, дата) отсутствуют в предоставленном тексте — пиши "НЕ УКАЗАНО". 
                        Не пытайся угадать или подставить значения по умолчанию."""
                    },
                    {
                        "role": "user", 
                        "content": f"""Извлеки данные из начала контракта и сформируй JSON объект с полями:
                        - contract_no: Номер контракта (ищи сразу слова 'КОНТРАКТ №' или 'ДОГОВОР №')
                        - contract_date: Дата контракта
                        - ikz: Идентификационный код закупки (полностью включая цифры)
                        - project_name: Предмет (название) контракта
                        - customer: Полное наименование Заказчика
                        - customer_post: Должность подписанта Заказчика (напр. Министр)
                        - customer_fio: ФИО подписанта Заказчика
                        - company: Полное наименование Исполнителя
                        - director_post: Должность подписанта Исполнителя (напр. Генеральный директор)
                        - director: ФИО подписанта Исполнителя

                        Текст для анализа:
                        {context}"""
                    }
                ],
                response_format={'type': 'json_object'}
            )

            st.session_state.t_info = json.loads(res.choices[0].message.content)

            # Собираем только титульник
            doc_title = build_title_page(st.session_state.t_info)
            buf_t = io.BytesIO()
            doc_title.save(buf_t)
            st.session_state.file_title_only = buf_t.getvalue()
            st.success("Титульный лист готов!")

        if "file_title_only" in st.session_state:
            st.download_button("📥 Скачать Титульник", st.session_state.file_title_only, "Title_Page.docx")

# СТОЛБЕЦ 2: РУКОПИСНЫЙ ОТЧЕТ

with col2:
    st.header("📝 2. Отчет по ТЗ")
    file_tz = st.file_uploader(
    "Загрузите Техзадание", 
    type="docx", 
    key="tz_loader", 
    on_change=lambda: st.session_state.pop("raw_report_body", None) # Удаляет старый текст отчета
    )

    if file_tz:
        if st.button("Сформировать Рукописный отчет"):
            client = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com/v1")
            raw_tz = get_text_from_file(file_tz)
 
            with st.spinner("ИИ анализирует ТЗ и пишет текст..."):
                # 1. Запрос к ИИ для формирования текста глав
                res_body = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": """Ты техписатель. Сформируй отчет по правилам:
                        1. Каждая глава начинается с номера и названия (например, 1. Организация мероприятия).
                        2. ЗАГОЛОВОК ГЛАВЫ пиши в НАСТОЯЩЕМ времени.
                        3. ОПИСАНИЕ внутри главы пиши в ПРОШЕДШЕМ времени (выполнено, оказано, организовано).
                        4. ВАЖНО: Не используй символы разметки (** или #). Весь текст должен быть чистым."""},
                        {"role": "user", "content": f"Сделай отчет из этого ТЗ:\n\n{raw_tz}"}
                    ]
                )

                # 2. Запрос к ИИ для поиска требований
                res_req = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди требования к фото и документам в этом ТЗ: {raw_tz}"}]
                )

                # Сохраняем результаты в сессию
                st.session_state.raw_report_body = res_body.choices[0].message.content
                st.session_state.raw_requirements = res_req.choices[0].message.content

                # 3. Сборка документа (передаем t_info для заголовка)
                # Убедись, что в build_report_body теперь заложен логика жирных глав
                doc_rep = build_report_body(
                    st.session_state.raw_report_body, 
                    st.session_state.raw_requirements,
                    st.session_state.t_info if "t_info" in st.session_state else {}
                )
               
                buf_r = io.BytesIO()
                doc_rep.save(buf_r)
                st.session_state.file_report_only = buf_r.getvalue()
                st.success("Отчет сформирован!")

        if "file_report_only" in st.session_state:
            st.download_button("📥 Скачать Отчет (без титульника)", st.session_state.file_report_only, "Report_Only.docx")

# --- КНОПКА ПОЛНОЙ СБОРКИ ---
if "file_title_only" in st.session_state and "file_report_only" in st.session_state:
    st.divider()
    st.subheader("🏁 Финальный шаг")

    if st.button("🚀 СОБРАТЬ ПОЛНЫЙ ОТЧЕТ", use_container_width=True):
        # Используем функцию из Кода 2.0 для сборки всего документа
        full_doc = create_final_report(
            st.session_state.t_info, 
            st.session_state.raw_report_body, 
            st.session_state.raw_requirements
        )

        final_buf = io.BytesIO()
        full_doc.save(final_buf)
        st.session_state.full_ready_file = final_buf.getvalue()

    if "full_ready_file" in st.session_state:
        st.download_button(

            label="🔥 СКАЧАТЬ ВЕСЬ ДОКУМЕНТ (ТИТУЛЬНИК + ОТЧЕТ)",
            data=st.session_state.full_ready_file,
            file_name="Full_Final_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )








