import streamlit as st
import json
import re
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI

# ─── ФУНКЦИЯ ФОРМАТИРОВАНИЯ ФИО ────────────────────────────────────────
def format_fio_universal(raw_fio):
    if not raw_fio or len(raw_fio.strip()) < 3:
        return "________________"
    # Убираем возможные должности/мусор, которые ИИ мог прихватить
    clean = re.sub(r'(директор|министр|заместитель|начальник|председатель|генеральный|зам|и\.о\.|исполняющий|обязанности)',
                   '', raw_fio, flags=re.IGNORECASE).strip()
    parts = clean.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    if len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}."
    return clean or "________________"

# ─── ФУНКЦИЯ СОЗДАНИЯ ТОЛЬКО ТИТУЛЬНОГО ЛИСТА ────────────────────────────
def create_title_only_docx(data):
    doc = Document()

    # Базовый стиль
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    # Заголовок
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Информационно-аналитический отчет об исполнении условий\n")
    run.bold = True
    run.font.size = Pt(14)
    run = p.add_run(f"Контракта № {data.get('contract_no', '—')} от {data.get('contract_date', '—')}\n")
    run.bold = True
    run.font.size = Pt(14)
    p.add_run(f"Идентификационный код закупки: {data.get('ikz', '—')}.")

    # Отступы
    for _ in range(5):
        doc.add_paragraph()

    # ТОМ I
    p = doc.add_paragraph("ТОМ I")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].bold = True
    p.runs[0].font.size = Pt(14)

    for _ in range(4):
        doc.add_paragraph()

    # Таблица подписей (3 строки × 2 столбца)
    table = doc.add_table(rows=3, cols=2)
    table.autofit = True
    table.allow_autofit = True

    # Заголовки жирным
    table.cell(0, 0).text = "Отчет принят Заказчиком"
    table.cell(0, 1).text = "Отчет передан Исполнителем"
    for cell in table.rows[0].cells:
        run = cell.paragraphs[0].runs[0]
        run.bold = True

    # Должности
    table.cell(1, 0).text = data.get('customer_post_full', '_______________________________')
    table.cell(1, 1).text = data.get('executor_post', 'Директор')

    # Подписи
    p_left = table.cell(2, 0).add_paragraph()
    p_left.add_run("_______________ ").font.size = Pt(12)
    p_left.add_run(f"{data.get('customer_fio', '________________')} м.п.")

    p_right = table.cell(2, 1).add_paragraph()
    p_right.add_run("_______________ ").font.size = Pt(12)
    p_right.add_run(f"{data.get('executor_fio', '________________')} м.п.")

    # Выравнивание по центру во всех ячейках
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Сохраняем в буфер
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─── НАСТРОЙКА СТРАНИЦЫ ─────────────────────────────────────────────────
st.set_page_config(page_title="Генератор отчёта — Шаг 1", layout="wide")

# Секреты и клиент DeepSeek
try:
    client_ai = OpenAI(
        api_key=st.secrets["DEEPSEEK_API_KEY"].strip(),
        base_url="https://api.deepseek.com/v1"
    )
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка чтения секретов: {e}")
    st.stop()

# Инициализация session_state
if 'title_data' not in st.session_state:
    st.session_state.title_data = None
if 'title_buffer' not in st.session_state:
    st.session_state.title_buffer = None
if 'last_uploaded_name' not in st.session_state:
    st.session_state.last_uploaded_name = None

# ─── ИНТЕРФЕЙС ──────────────────────────────────────────────────────────
st.title("Шаг 1 — Формирование титульного листа")

user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD:
    st.info("Введите пароль для доступа")
    st.stop()

uploaded_file = st.file_uploader("Загрузите файл контракта (.docx)", type=["docx"])

if uploaded_file is not None:
    current_name = uploaded_file.name

    # Сброс при новом файле
    if st.session_state.last_uploaded_name != current_name:
        st.session_state.title_data = None
        st.session_state.title_buffer = None
        st.session_state.last_uploaded_name = current_name

    # Чтение текста документа
    try:
        doc_obj = Document(uploaded_file)
        full_text = "\n".join(para.text for para in doc_obj.paragraphs if para.text.strip())
    except Exception as e:
        st.error(f"Не удалось прочитать файл: {e}")
        st.stop()

    # Контекст для ИИ — начало + конец
    head = full_text[:1800]
    tail = full_text[-2500:]
    context = head + "\n\n……\n\n" + tail

    if st.session_state.title_data is None:
        with st.spinner("Анализ титульных данных..."):
            try:
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{
                        "role": "user",
                        "content": f"""Извлеки данные ТОЧНО из текста. Ничего не придумывай и не додумывай.

Верни ТОЛЬКО валидный JSON, без лишнего текста.

Ключи (если поле отсутствует — пустая строка или null):

{{
  "contract_no": "номер контракта",
  "contract_date_raw": "дата как написана в тексте",
  "ikz": "36-значный код закупки (только цифры)",
  "customer_org": "полное наименование заказчика",
  "customer_post": "должность подписанта заказчика (полностью)",
  "customer_basis": "основание полномочий (если указано, иначе пустая строка)",
  "customer_fio_raw": "ФИО подписанта заказчика как в тексте",
  "executor_org": "полное наименование исполнителя",
  "executor_post": "должность подписанта исполнителя",
  "executor_fio_raw": "ФИО подписанта исполнителя как в тексте"
}}

Текст для анализа (начало + конец документа):
{context}
"""
                    }],
                    response_format={"type": "json_object"},
                    temperature=0.15,
                    max_tokens=800
                )

                raw = json.loads(res.choices[0].message.content)

                td = {}
                td['contract_no']   = raw.get('contract_no',   '—')
                td['contract_date'] = raw.get('contract_date_raw', '—')
                td['ikz']           = raw.get('ikz',           '—')
                td['customer']      = raw.get('customer_org',  '—')
                td['customer_post_full'] = (raw.get('customer_post') or '').strip()
                if basis := (raw.get('customer_basis') or '').strip():
                    td['customer_post_full'] += f" {basis}"
                td['customer_fio']  = format_fio_universal(raw.get('customer_fio_raw', ''))
                td['executor']      = raw.get('executor_org',  '—')
                td['executor_post'] = (raw.get('executor_post') or 'Директор').strip().capitalize()
                td['executor_fio']  = format_fio_universal(raw.get('executor_fio_raw', ''))

                st.session_state.title_data = td

            except Exception as e:
                st.error(f"Ошибка при обращении к DeepSeek: {str(e)}")
                st.stop()

    # Показываем результат
    if st.session_state.title_data:
        data = st.session_state.title_data

        st.subheader("Извлечённые данные")
        st.json(data)

        if st.button("Сформировать и скачать титульный лист"):
            with st.spinner("Создаём документ..."):
                buf = create_title_only_docx(data)
                st.session_state.title_buffer = buf.getvalue()

        if st.session_state.title_buffer:
            safe_no = re.sub(r'[^0-9а-яА-Яa-zA-Z\-_]', '_', data['contract_no'])
            st.download_button(
                label="📄 Скачать титульный лист (проверка)",
                data=st.session_state.title_buffer,
                file_name=f"Титульный_лист_{safe_no}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.markdown("---")
st.caption("После проверки титульника скажите, что получилось — перейдём к шагу 2 (ТЗ и основной отчёт)")
