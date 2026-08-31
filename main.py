import imaplib
import os
import smtplib
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

# ================= 1. НАЛАШТУВАННЯ СТОРІНКИ =================
st.set_page_config(page_title="UF Mail Automation", page_icon="📧", layout="wide")

st.title("📧 Автоматизація розсилки листів")
st.caption("Інструмент для персоналізованої відправки листів з можливістю (опціонально) додавати відповідні PDF-файли.")

# ================= 2. БІЧНА ПАНЕЛЬ: SMTP/IMAP ТА АВТОРИЗАЦІЯ =================
default_email = st.secrets.get("EMAIL_SENDER", "mail@ufincubator.com")
default_password = st.secrets.get("EMAIL_PASSWORD", "")

with st.sidebar:
    st.header("⚙️ Поштовий сервер")
    mail_server = st.text_input("Поштовий Сервер (SMTP/IMAP)", value="mail.adm.tools")
    smtp_port = st.number_input("SMTP Порт (відправка)", value=465, step=1)
    imap_port = st.number_input("IMAP Порт (збереження)", value=993, step=1)
    sender_email = st.text_input("Ваш Email", value=default_email)
    sender_password = st.text_input("Пароль від пошти", value=default_password, type="password")
    
    st.markdown("---")
    st.header("🛡️ Антиспам та безпека")
    delay_seconds = st.slider("Затримка між листами (сек)", min_value=1, max_value=30, value=5)
    test_mode = st.checkbox("🧪 Тестовий режим (надсилати все на мою пошту)", value=True)
    if test_mode:
        st.info("У тестовому режимі лист надсилатиметься на вашу адресу для перевірки теми та вкладення.")

# ================= 3. ФУНКЦІЯ ЗБЕРЕЖЕННЯ В "НАДІСЛАНІ" (IMAP) =================
def save_to_sent_folder(imap_host, imap_port, user, password, raw_msg_bytes):
    try:
        imap = imaplib.IMAP4_SSL(imap_host, imap_port)
        imap.login(user, password)

        status, folder_list = imap.list()
        target_folder = None
        
        for folder_entry in folder_list:
            decoded = folder_entry.decode('utf-8', errors='ignore')
            for possible_name in ['Sent', 'INBOX.Sent', 'Sent Messages', 'Надіслані', 'Отправленные']:
                if f'"{possible_name}"' in decoded or f' {possible_name}' in decoded:
                    target_folder = possible_name
                    break
            if target_folder:
                break
        
        if not target_folder:
            target_folder = "Sent"

        imap.append(target_folder, '\\Seen', imaplib.Time2Internaldate(time.time()), raw_msg_bytes)
        imap.logout()
    except Exception as e:
        st.sidebar.warning(f"⚠️ Не вдалося зберегти копію в 'Надіслані': {e}")

# ================= 4. ЗАВАНТАЖЕННЯ ДАНИХ ТА ФАЙЛІВ =================
st.subheader("1. Джерела даних та файлів")
col_file, col_dir = st.columns(2)

with col_file:
    uploaded_table = st.file_uploader(
        "📄 Оберіть файл контактів (Excel або CSV) *Обов'язково*", 
        type=["xlsx", "xls", "csv"]
    )

with col_dir:
    uploaded_pdfs = st.file_uploader(
        "📂 Оберіть PDF-файли (Опціонально — якщо без файлів, надішлеться лише текст)", 
        type=["pdf"], 
        accept_multiple_files=True
    )

# ================= 5. ШАБЛОН ТЕМИ ТА ТІЛА ЛИСТА =================
st.markdown("---")
st.subheader("2. Шаблон листа")

subject_template = st.text_input(
    "Шаблон теми листа (змінна {company} автоматично підставить назву)", 
    value="Пропозиція партнерства: Бізнес-інкубатор Ukrainian Future та {company}"
)

default_html_body = """<p>Вітаю!</p>

<p>Нагадую вам про пропозицію партнерства по <strong>UF Бібліотеці матеріалів</strong>.</p>

<p>Ми вже створили стенд матеріалів від українських виробників і наразі <strong>отримали перші 20 зразків</strong>. Також нам дуже хотілося б бачити серед зразків ваші матеріали, щоб показувати їх студентам, старшокласникам і засновникам стартапів.</p>

<p>Ми готові максимально спростити для вас логістику й самостійно організувати доставку зразка Новою Поштою за наш рахунок.</p>

<p>Підкажіть, будь ласка, чи цікава вам ця пропозиція?</p>

<p>Гарного дня!</p>
"""

email_body_template = st.text_area(
    "HTML-шаблон тексту листа (використовуйте {company})",
    height=220,
    value=default_html_body
)

# ================= 6. ВАЛІДАЦІЯ ТА ПОПЕРЕДНІЙ ПЕРЕГЛЯД =================
st.markdown("---")
st.subheader("3. Попередній перегляд списку розсилки")

if uploaded_table:
    try:
        if uploaded_table.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_table)
        else:
            df = pd.read_excel(uploaded_table)
        
        df.columns = df.columns.astype(str).str.strip()

        company_col = next((c for c in df.columns if any(k in c.lower() for k in ['company', 'назва', 'організац', 'компані'])), None)
        email_col = next((c for c in df.columns if any(k in c.lower() for k in ['email', 'пошта', 'контакт', 'mail'])), None)

        if not company_col or not email_col:
            st.error(f"❌ Не вдалося знайти колонки компанії або email. Знайдені колонки: {list(df.columns)}")
        else:
            df_clean = df.dropna(subset=[company_col, email_col]).drop_duplicates(subset=[company_col]).copy()
            
            # Словник з PDF-файлами (якщо користувач їх завантажив)
            pdf_dict = {f.name.lower(): f for f in uploaded_pdfs} if uploaded_pdfs else {}

            matched_data = []
            for _, row in df_clean.iterrows():
                comp_name = str(row[company_col]).strip()
                comp_email = str(row[email_col]).strip()

                matched_pdf = None
                if pdf_dict:
                    for fname, fobj in pdf_dict.items():
                        if comp_name.lower() in fname:
                            matched_pdf = fobj
                            break

                matched_data.append({
                    "Компанія": comp_name,
                    "Email": comp_email,
                    "Вкладення": f"📎 {matched_pdf.name}" if matched_pdf else "Лише текст (без файлу)",
                    "_file_obj": matched_pdf
                })

            display_df = pd.DataFrame(matched_data)[["Компанія", "Email", "Вкладення"]]
            st.dataframe(display_df, use_container_width=True)

            total_count = len(matched_data)
            with_pdf_count = sum(1 for item in matched_data if item["_file_obj"] is not None)
            st.write(f"📊 Всього листів до відправки: **{total_count}** (з них із PDF-файлом: **{with_pdf_count}**)")

            # ================= 7. ВІДПРАВКА ТА ЗБЕРЕЖЕННЯ =================
            st.markdown("---")
            if st.button("🚀 Запустити розсилку", type="primary"):
                if not sender_email or not sender_password:
                    st.error("⚠️ Вкажіть Email та Пароль у лівій панелі перед відправкою!")
                elif total_count == 0:
                    st.warning("⚠️ Немає контактів для відправки.")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    log_box = st.container()

                    try:
                        status_text.text("Підключення до поштового сервера (SMTP)...")
                        with smtplib.SMTP_SSL(mail_server, int(smtp_port)) as server:
                            server.login(sender_email, sender_password)
                            st.toast("✅ Авторизація успішна!", icon="🔓")

                            sent = 0
                            for i, item in enumerate(matched_data):
                                comp = item["Компанія"]
                                target_to = sender_email if test_mode else item["Email"]
                                file_obj = item["_file_obj"]

                                msg = MIMEMultipart()
                                msg["From"] = sender_email
                                msg["To"] = target_to
                                msg["Subject"] = subject_template.format(company=comp)

                                body = email_body_template.format(company=comp)
                                msg.attach(MIMEText(body, "html", "utf-8"))

                                # Прикріплюємо файл, тільки якщо він існує
                                if file_obj is not None:
                                    pdf_bytes = file_obj.getvalue()
                                    part = MIMEApplication(pdf_bytes, _subtype="pdf")
                                    part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", file_obj.name))
                                    msg.attach(part)
                                    file_status_label = f"| Файл: `{file_obj.name}`"
                                else:
                                    file_status_label = "| Лише текст"

                                # 1. Відправка адресату через SMTP
                                server.send_message(msg)
                                
                                # 2. Збереження копії в IMAP "Надіслані"
                                save_to_sent_folder(
                                    imap_host=mail_server,
                                    imap_port=int(imap_port),
                                    user=sender_email,
                                    password=sender_password,
                                    raw_msg_bytes=msg.as_bytes()
                                )

                                sent += 1
                                log_box.success(f"✅ Надіслано ({sent}/{total_count}): **{comp}** ➔ `{target_to}` {file_status_label}")

                                progress_bar.progress((i + 1) / total_count)
                                status_text.text(f"Опрацьовано {i + 1} з {total_count}...")
                                time.sleep(delay_seconds)

                        st.balloons()
                        st.success(f"🎉 Розсилку завершено! Успішно надіслано листів: {sent}")

                    except Exception as e:
                        st.error(f"❌ Помилка під час відправки: {e}")

    except Exception as err:
        st.error(f"Помилка зчитування таблиці: {err}")
else:
    st.info("👆 Завантажте файл таблиці ліворуч, щоб переглянути контакти та розпочати.")