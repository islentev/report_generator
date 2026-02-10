import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import io
import json

# --- 1. НАСТРОЙКА ---
st.set_page_config(page_title="Генератор Отчетов", layout="wide")

if 'report_buffer' not in st.session_state: st.session_state['report_buffer'] = None
if 'title_info' not in st.session_state: st.session_state['title_info'] = None

# --- 2. ПОДКЛЮЧЕНИЕ ---
try:
    gcp_info = dict(st.secrets["gcp_service_account"])
    gcp_info["private_key"] = gcp_info["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(gcp_info, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    gc = gspread.authorize(creds)
    client_ai = OpenAI(api_key=st.secrets["DEEPSEEK_API_KEY"].strip().strip('"'), base_url="https://api.deepseek.com/v1")
    SHEET_ID = st.secrets["SHEET_ID"]
    APP_PASSWORD = st.secrets["APP_PASSWORD"]
except Exception as e:
    st.error(f"Ошибка конфига: {e}"); st.stop()

# --- 3. ФУНКЦИИ ---

def create_report_docx(report_content, title_data):
    doc = Document()
    
    # ТИТУЛЬНЫЙ ЛИСТ (Берем строго из title_data)
    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    auth_text = f"""УТВЕРЖДАЮ
{title_data.get('company', '________________')}

________________ / {title_data.get('director', '________________')}
«___» _________ 2025 г."""
    run_auth = p_auth.add_run(auth_text)
    run_auth.font.size = Pt(11)

    for _ in range(7): doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("ИНФОРМАЦИОННЫЙ ОТЧЕТ\n").bold = True
    p_title.runs[-1].font.size = Pt(20)
    
    sub_text = f"по исполнению Государственного контракта\n№ {title_data.get('contract_no', '_________')} от {title_data.get('contract_date', '_________')}\n\n"
    run_sub = p_title.add_run(sub_text)
    run_sub.font.size = Pt(14)
    p_title.add_run(title_data.get('project_name', '________________')).italic = True

    for _ in range(10): doc.add_paragraph()
    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.add_run("Москва, 2025 г.")
    doc.add_page_break()

    # ТЕКСТ ОТЧЕТА
    for block in report_content.split('\n\n'):
        p = doc.add_paragraph()
        parts = block.split('**')
        for i, part in enumerate(parts):
            run = p.add_run(part.replace('*', ''))
            if i % 2 != 0: run.bold = True
            
    # ПОДПИСЬ В КОНЦЕ (Та же компания и директор, что на титульнике)
    p_sign = doc.add_paragraph()
    p_sign.add_run(f"\n\nДиректор {title_data.get('company', '')}  _________________ / {title_data.get('director', '')}")
    
    return doc

# --- 4. ИНТЕРФЕЙС ---
user_pass = st.sidebar.text_input("Пароль", type="password")
if user_pass != APP_PASSWORD:
    st.info("Введите пароль."); st.stop()

sheet = gc.open_by_key(SHEET_ID).sheet1
df_etalons = pd.DataFrame(sheet.get_all_records())

st.title("⚖️ Универсальный генератор отчетов")
uploaded_file = st.file_uploader("Загрузите контракт", type="docx")

if uploaded_file:
    contract_text = "\n".join([p.text for p in Document(uploaded_file).paragraphs])
    
    # ИИ вытаскивает реквизиты (БЕЗ использования шаблонов имен)
    if not st.session_state['title_info']:
        with st.spinner("Извлечение реквизитов из текста контракта..."):
            all_types = df_etalons["Тип проекта"].tolist()
            extraction_prompt = f"""ПРОАНАЛИЗИРУЙ ТЕКСТ: {contract_text[:6000]}
            Найди сторону Исполнителя и выдай JSON строго по факту текста:
            {{
              "company": "Полное название компании Исполнителя",
              "director": "ФИО Директора Исполнителя (кто подписывает)",
              "contract_no": "Номер контракта",
              "contract_date": "Дата заключения",
              "project_name": "Предмет контракта (кратко)",
              "type": "Выбери один тип из списка {all_types}"
            }}
            ВАЖНО: Если данных нет в тексте, пиши 'Не указано'. Не выдумывай имена."""
            
            res_meta = client_ai.chat.completions.create(
                model="deepseek-chat", messages=[{"role": "user", "content": extraction_prompt}],
                response_format={ 'type': 'json_object' }
            )
            st.session_state['title_info'] = json.loads(res_meta.choices[0].message.content)

    meta = st.session_state['title_info']
    st.success(f"Работаем от лица: **{meta['company']}** | Директор: **{meta['director']}**")

    with st.form("data"):
        q1 = st.text_input("Кол-во участников", value="100")
        q2 = st.text_input("Письмо согласования")
        facts = st.text_area("Доп. детали")
        if st.form_submit_button("Сгенерировать"):
            with st.spinner("Генерация..."):
                # Ищем форму в контракте
                search_res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": f"Найди образец отчета в тексте: {contract_text[-12000:]}. Если есть - опиши структуру. Если нет - напиши НЕТ."}]
                )
                form_instr = search_res.choices[0].message.content
                
                # Собираем отчет
                sys_msg = f"Ты юрист компании {meta['company']}. Напиши отчет от имени директора {meta['director']}. Структура: {form_instr if 'НЕТ' not in form_instr else 'Стандартная'}"
                res = client_ai.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":sys_msg}, {"role":"user","content":f"Текст контракта: {contract_text[:8000]}\nУчастники: {q1}\nПисьмо: {q2}\nДетали: {facts}"}]
                )
                
                doc_final = create_report_docx(res.choices[0].message.content, meta)
                buf = io.BytesIO()
                doc_final.save(buf)
                st.session_state['report_buffer'] = buf.getvalue()

if st.session_state['report_buffer']:
    st.download_button("📥 Скачать готовый Отчет", st.session_state['report_buffer'], "Report.docx")
