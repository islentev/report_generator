import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_COLOR_INDEX
from openai import OpenAI
client = OpenAI(
  base_url="https://openrouter.ai/api/v1",
  api_key=st.secrets["OPENROUTER_API_KEY"],
  default_headers={
    "HTTP-Referer": "https://report-generator.streamlit.app", # Твой адрес
    "X-Title": "Report Generator"
  }
)
GEMINI_MODEL = "anthropic/claude-3.5-sonnet"
import io
import json
import re

if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

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

def smart_generate_step_strict(section_text, requirements_text):
    system_prompt = f"""Ты - юридический редактор. Перепиши пункты ТЗ в Отчет.
    ПРАВИЛА:
    1. ВРЕМЯ: У заголовком - настоящее, у текста - СТРОГО ПРОШЕДШЕЕ ('организовано', 'оказано', 'размещено').
    2. ЗАПРЕТ: Слова 'должен', 'обязан', 'будет', 'необходимо' КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНЫ.
    3. ТОЧНОСТЬ: Перенеси ВСЕ цифры, площади, сроки и названия без изменений.
    4. ПУНКТУАЦИЯ: Соблюдай правила русского языка. Не обрывай предложения.
    ТРЕБОВАНИЯ К ДОКУМЕНТАМ: {requirements_text}"""

    full_prompt = f"{system_prompt}\n\nТРАНСФОРМИРУЙ ЭТОТ КУСОК ТЗ В ОТЧЕТ:\n{section_text}"

    # Шаг 1: Генерация
    res = client.chat.completions.create(
      model=GEMINI_MODEL,
      messages=[{"role": "user", "content": full_prompt}]
    )
    draft = res.choices[0].message.content

    # Шаг 2: ЖЕСТКИЙ КОНТРОЛЬ
    v_prompt = f"Сравни ТЗ и Отчет. Если есть ошибки, напиши их. Если всё ок, пиши 'ОШИБОК: 0'.\nТЗ: {section_text}\nОТЧЕТ: {draft}"
    v_res = client.chat.completions.create(
        model=GEMINI_MODEL,
        messages=[{"role": "user", "content": v_prompt}]
    )
    v_text = v_res.choices[0].message.content
    
    # Шаг 3: Исправление (если инспектор нашел брак)
    if "ОШИБОК: 0" not in v_text:
        fix_prompt = f"{system_prompt}\nИСПРАВЬ ОШИБКИ: {v_text}\nТЗ: {section_text}\nЧЕРНОВИК: {draft}"
        fix = client.chat.completions.create(
            model=GEMINI_MODEL,
            messages=[{"role": "user", "content": fix_prompt}]
        )
        return fix.choices[0].message.content
    return draft

# --- 3. СБОРКА ДОКУМЕНТА (ТВОЕ ОФОРМЛЕНИЕ) ---

def build_title_page(t):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # ИСПРАВЛЕНИЕ: Извлекаем переменные из словаря t, прежде чем использовать их
    contract_no = t.get('contract_no', '___')
    contract_date = t.get('contract_date', '___')
    ikz = t.get('ikz', '___')

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Информационно-аналитический отчет об исполнении условий\n").bold = True
    p.add_run(f"Контракта № {contract_no} от «{contract_date}» 2025 г.\n").bold = True
    p.add_run(f"Идентификационный код закупки: {ikz}").bold = True

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
    
    # Чтобы не было "Отчет об оказании услуг по услуг"
    p_name = t.get('project_name', '')
    if isinstance(p_name, dict): p_name = p_name.get('name', '')
    p_name = str(p_name).strip() if p_name else "услугам"

    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head.add_run(f"Отчет об оказании услуг по {p_name}").bold = True
    
    # Очищаем основной текст от дублей и пустых строк
    lines = clean_markdown(report_body).split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        para = doc.add_paragraph()
        run = para.add_run(line)
        if re.match(r"^\d+\.", line): run.bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
    if req_body:
        doc.add_page_break()
        p_req = doc.add_paragraph()
        p_req.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_req.add_run('ТРЕБОВАНИЯ К ПРЕДОСТАВЛЯЕМОЙ ДОКУМЕНТАЦИИ').bold = True
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
        # 1. Полная очистка session_state
        for key in list(st.session_state.keys()):
            if key != "reset_counter":
                del st.session_state[key]
        
        # 2. Явное обнуление переменных кэша текста (чтобы ИИ не подтянул старое)
        st.session_state.raw_tz_source = ""
        st.session_state.raw_report_body = ""
        st.session_state.raw_requirements = ""
        st.session_state.t_info = {}

        # Это принудительно очистит text_area в колонках
        st.session_state[f"t_area_{st.session_state.reset_counter}"] = ""
        st.session_state[f"tz_area_{st.session_state.reset_counter}"] = ""
        
        # 3. Смена ключей виджетов (то, что мы делали со счетчиком)
        st.session_state.reset_counter += 1
        
        # 4. Очистка кэша самого Streamlit (на всякий случай)
        st.cache_data.clear()
        
        st.rerun()
    
col1, col2, col3 = st.columns(3)

# КОЛОНКА 1: ТИТУЛЬНИК
with col1:
    st.header("📄 1. Титульный лист")
    t_tab1, t_tab2 = st.tabs(["📁 Файл", "⌨️ Текст"])
    
    t_context = ""
    with t_tab1:
        f_title = st.file_uploader("Контракт (DOCX)", type="docx", key="u_title")
        t_context = "" # Инициализируем пустой строкой
        if f_title: 
            t_context = get_contract_start_text(f_title)
    
    with t_tab2:
        # 1. Сначала определяем ключ текущего виджета
        area_key = f"t_area_{st.session_state.reset_counter}"
        
        # 2. Отрисовываем виджет
        m_title = st.text_area(
            "Вставьте начало контракта:", 
            value=st.session_state.get(area_key, ""), 
            height=150, 
            key=area_key
        )
    
    # 3. Если в поле что-то вписали, обновляем контекст
    if m_title: 
        t_context = m_title

    if st.button("🔍 Извлечь реквизиты", use_container_width=True):
        if t_context:
            with st.spinner("Ищем данные..."):
                res = client.chat.completions.create(
                    model=GEMINI_MODEL,
                    messages=[{
                        "role": "user", 
                        "content": f"""Извлеки данные СТРОГО в формате JSON с этими ключами:
                        'contract_no' (номер), 
                        'contract_date' (дата), 
                        'ikz' (ИКЗ), 
                        'project_name' (предмет контракта), 
                        'customer' (Заказчик), 
                        'customer_post' (должность заказчика), 
                        'customer_fio' (ФИО заказчика), 
                        'company' (Исполнитель), 
                        'director_post' (должность руководителя исполнителя), 
                        'director' (ФИО руководителя исполнителя).
                        Текст: {t_context}"""
                    }],
                    response_format={ "type": "json_object" }
                )
                st.session_state.t_info = json.loads(res.choices[0].message.content)
        else: st.error("Нет данных!")

    # --- ПРЕВЬЮ ТИТУЛЬНИКА (Редактируемое) ---
    if "t_info" in st.session_state:
        st.info("Проверьте данные:")
        ti = st.session_state.t_info
        ti['contract_no'] = st.text_input("№", ti.get('contract_no'))
        ti['ikz'] = st.text_input("ИКЗ", ti.get('ikz'))
        ti['customer_fio'] = st.text_input("ФИО Заказчика", ti.get('customer_fio'))
        # Кнопка скачивания только титульника
        doc_t = build_title_page(ti)
        buf_t = io.BytesIO(); doc_t.save(buf_t)
        st.download_button("📥 Скачать Титульник", buf_t.getvalue(), "Title.docx", use_container_width=True)
        
# КОЛОНКА 2: ОТЧЕТ
with col2:
    st.header("📝 2. Отчет (ТЗ)")
    tz_tab1, tz_tab2 = st.tabs(["📁 Файл", "⌨️ Текст"])
    
    with tz_tab1:
        # Добавляем ключ, чтобы файл тоже можно было сбросить
        f_tz = st.file_uploader("Техзадание (DOCX)", type="docx", key=f"u_tz_{st.session_state.reset_counter}")
    
    with tz_tab2:
        # Используем reset_counter для очистки при нажатии кнопки Сброс
        tz_area_key = f"tz_area_{st.session_state.reset_counter}"
        m_tz_area = st.text_area(
            "Текст техзадания:", 
            value=st.session_state.get(tz_area_key, ""), 
            height=150, 
            key=tz_area_key
        )
    
    if st.button("⚙️ Сгенерировать текст", use_container_width=True):
        # Берем текст из окна, если пусто - из файла
        tz_content = m_tz_area.strip() if m_tz_area.strip() else ""
        if not tz_content and f_tz:
            tz_content = get_text_from_file(f_tz)
            
        if tz_content:
            st.session_state.raw_tz_source = tz_content  # СОХРАНЯЕМ ДЛЯ ПОШАГОВОЙ СБОРКИ
            seg_res = client.chat.completions.create(
                model=GEMINI_MODEL,
                messages=[{"role": "user", "content": f"Раздели на блоки с тегом [END_BLOCK]: {tz_content}"}]
            )
            steps = [s.strip() for s in seg_res.choices[0].message.content.split('[END_BLOCK]') if s.strip()]

            instruction = """Роль и контекст:
            Ты — ассистент юриста по договорной работе. Перед тобой текст Технического Задания (ТЗ), который был написан в будущем времени как план работ. 
            Сейчас его нужно превратить в черновик отчета о выполнении. Твоя задача — действовать как автомат по замене времени и удалению лишних слов, не меняя структуру и терминологию документа. 

            Инструкция:
            1. ВРЕМЯ: СТРОГО ПРОШЕДШЕЕ (организовано, оказано, размещено).
            2. ЗАПРЕТ: Слова 'должен', 'обязан', 'будет', 'необходимо' КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНЫ.
            3. НЕПРИКОСНОВЕННОСТЬ ЗАГОЛОВКОВ: Заголовки пунктов оставляй как в ТЗ.
            4. ТОЧНОСТЬ: Сохраняй все цифры, площади и сроки. Не сокращай текст!
            
            Выведи только чистый текст отчета."""
            
            # Вызываем Gemini с твоей инструкцией и текстом ТЗ
            res = client.chat.completions.create(
                model=GEMINI_MODEL,
                messages=[{"role": "user", "content": f"{instruction}\n\n{tz_content}"}]
            )
            st.session_state.raw_report_body = res.choices[0].message.content
            
            # Сохраняем результат (у Gemini это res.text)
            st.session_state.raw_report_body = res.choices[0].message.content
            
            final_text_parts = []
            pb = st.progress(0)
            status_text = st.empty()
            
            for i, step in enumerate(steps):
                status_text.text(f"Обработка блока {i+1} из {len(steps)}...")
                part = smart_generate_step_strict(step, st.session_state.get('raw_requirements', ''))
                final_text_parts.append(part)
                pb.progress((i + 1) / len(steps))
            
            st.session_state.raw_report_body = res.choices[0].message.content
        else:
            st.warning("Данные ТЗ отсутствуют")

    if "raw_report_body" in st.session_state:
        st.session_state.raw_report_body = st.text_area("Черновик:", st.session_state.raw_report_body, height=300)

# КОЛОНКА 3: ТРЕБОВАНИЯ
with col3:
    st.header("📋 3. Требования")
    if st.button("🔍 Выделить требования", use_container_width=True):
        if "raw_tz_source" in st.session_state:
            res = client.chat.completions.create(
                model=GEMINI_MODEL,
                messages=[{"role": "user", "content": f"Выпиши требования к документам: {st.session_state.raw_tz_source}"}]
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
        if "t_info" in st.session_state and st.session_state.get('raw_tz_source'):
            
            # Разрезаем по пунктам типа 1.1., 2.1.
            steps = [s.strip() for s in re.split(r'\n(?=\d+\.\d+)', st.session_state.raw_tz_source) if s.strip()]
            
            final_text_parts = []
            pb = st.progress(0)
            for i, step in enumerate(steps):
                part = smart_generate_step_strict(step, st.session_state.get('raw_requirements', ''))
                final_text_parts.append(part)
                pb.progress((i + 1) / len(steps))
            
            # Соединяем один раз
            full_smart_text = "\n\n".join(final_text_parts)
            doc = create_final_report(st.session_state.t_info, full_smart_text, st.session_state.get('raw_requirements', ''))
            buf = io.BytesIO()
            doc.save(buf)
            st.session_state.smart_file = buf.getvalue()
            st.success("Умная сборка завершена!")
if "full_file" in st.session_state:
    st.download_button("📥 Скачать обычный", st.session_state.full_file, "Report.docx")
if "smart_file" in st.session_state:
    st.download_button("📥 СКАЧАТЬ УМНЫЙ ОТЧЕТ", st.session_state.smart_file, "Smart_Report.docx")














