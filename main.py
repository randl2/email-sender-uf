import imaplib
import os
import smtplib
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
import streamlit as st

# ================= 1. НАЛАШТУВАННЯ СТОРІНКИ =================
st.set_page_config(page_title="UF Mail Automation", page_icon="📧", layout="wide")

st.title("📧 Автоматизація розсилки листів з PDF-вкладеннями")
st.caption("Інструмент для автоматичного підбору файлів за назвою компанії та персоналізованої відправки.")

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
        st.info("У тестовому режимі лист надсилатиметься на вашу адресу для перевірки теми та файлу.")

# ================= 3. ФУНКЦІЯ ЗБЕРЕЖЕННЯ В "НАДІСЛАНІ" (IMAP) =================
def save_to_sent_folder(imap_host, imap_port, user, password, raw_msg_bytes):
    """Підключається через IMAP і додає копію відправленого листа у папку 'Надіслані'."""
    try:
        imap = imaplib.IMAP4_SSL(imap_host, imap_port)
        imap.login(user, password)

        # Отримуємо список усіх доступних папок
        status, folder_list = imap.list()
        target_folder = None
        
        # Шукаємо назву папки надісланих (Sent, Sent Messages, Надіслані тощо)
        for folder_entry in folder_list:
            decoded = folder_entry.decode('utf-8', errors='ignore')
            for possible_name in ['Sent', 'INBOX.Sent', 'Sent Messages', 'Надіслані', 'Отправленные']:
                if f'"{possible_name}"' in decoded or f' {possible_name}' in decoded:
                    target_folder = possible_name
                    break
            if target_folder:
                break
        
        if not target_folder:
            target_folder = "Sent"  # За замовчуванням для більшості поштовиків

        # Зберігаємо лист з міткою \Seen (прочитане)
        imap.append(target_folder, '\\Seen', imaplib.Time2Internaldate(time.time()), raw_msg_bytes)
        imap.logout()
    except Exception as e:
        st.sidebar.warning(f"⚠️ Не вдалося зберегти копію в 'Надіслані': {e}")

# ================= 4. ЗАВАНТАЖЕННЯ ДАНИХ ТА ФАЙЛІВ =================
st.subheader("1. Джерела даних та файлів")
col_file, col_dir = st.columns(2)

with col_file:
    uploaded_table = st.file_uploader(
        "📄 Оберіть файл контактів (Excel або CSV)", 
        type=["xlsx", "xls", "csv"]
    )

with col_dir:
    uploaded_pdfs = st.file_uploader(
        "📂 Оберіть всі PDF-файли (виділіть усі разом у вікні)", 
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

<p>Я Олександр, проєктний менеджер бізнес-інкубатору Ukrainian Future НЦ “Мала академія наук України”. Минулого року ми започаткували <strong>некомерційний освітній проєкт <a href="https://ufincubator.com/ua/pro-biblioteku-materialiv" target="_blank">UF Бібліотека матеріалів</a></strong>, метою якого є познайомити студентів, старшокласників та засновників стартапів з сучасними матеріалами. Наразі в бібліотеці представлено 300 фізичних зразків, у якій відвідувачі можуть ознайомитись з матеріалами та підібрати для своїх прототипів чи наукових досліджень за фізико-хімічними властивостями.</p>

<p>Окрему увагу ми приділяємо рішенням, які підтримують екологічні принципи, зменшують негативний вплив на довкілля та відповідають стандартам циркулярної економіки.</p>

<p>Ми ознайомились з продукцією компанії <strong>{company}</strong>, і <strong>хотіли б запропонувати додати матеріали вашої компанії до нашої бібліотеки</strong> у вигляді фізичних зразків. Розміщення зразків абсолютно безкоштовне, детальну офіційну пропозицію додаю до цього листа у форматі PDF.</p>

<p>Підкажіть, будь ласка, <strong>чи цікава вам співпраця з нашим бізнес-інкубатором та розміщення ваших зразків на нашому стенді</strong>? У разі необхідності, можемо провести коротку зустріч онлайн або екскурсію по нашій бібліотеці матеріалів вживу.</p>

<p>--<br>
<strong>З повагою</strong>,<br>
Олександр, проєктний менеджер бізнес-інкубатора Ukrainian Future<br>
НЦ “Мала академія наук України”</p>

<p>🌐 <a href="https://ufincubator.com/ua" target="_blank">Сайт інкубатора</a><br>
🔹 <a href="https://www.facebook.com/UFincubator" target="_blank">Facebook</a> | <a href="https://www.instagram.com/uf_incubator" target="_blank">Instagram</a> | <a href="https://www.linkedin.com/company/ukrainian-future/posts/?feedView=all" target="_blank">LinkedIn</a></p>
"""

email_body_template = st.text_area(
    "HTML-шаблон тексту листа (використовуйте {company})",
    height=240,
    value=default_html_body
)

# ================= 6. ВАЛІДАЦІЯ, ПОШУК ТА ПРЕВ'Ю =================
st.markdown("---")
st.subheader("3. Попередній перегляд відповідностей")

if uploaded_table and uploaded_pdfs:
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
            pdf_dict = {f.name.lower(): f for f in uploaded_pdfs}

            matched_data = []
            for _, row in df_clean.iterrows():
                comp_name = str(row[company_col]).strip()
                comp_email = str(row[email_col]).strip()

                matched_pdf = None
                for fname, fobj in pdf_dict.items():
                    if comp_name.lower() in fname:
                        matched_pdf = fobj
                        break

                matched_data.append({
                    "Компанія": comp_name,
                    "Email": comp_email,
                    "Знайдений PDF": matched_pdf.name if matched_pdf else "❌ НЕ ЗНАЙДЕНО",
                    "_file_obj": matched_pdf
                })

            display_df = pd.DataFrame(matched_data)[["Компанія", "Email", "Знайдений PDF"]]
            st.dataframe(display_df, use_container_width=True)

            total_count = len(matched_data)
            ready_count = sum(1 for item in matched_data if item["_file_obj"] is not None)
            st.write(f"📊 Всього компаній: **{total_count}** | Готово до відправки: **{ready_count}**")

            # ================= 7. ВІДПРАВКА ТА ЗБЕРЕЖЕННЯ =================
            st.markdown("---")
            if st.button("🚀 Запустити розсилку", type="primary"):
                if not sender_email or not sender_password:
                    st.error("⚠️ Вкажіть Email та Пароль у лівій панелі перед відправкою!")
                elif ready_count == 0:
                    st.warning("⚠️ Не знайдено жодного відповідного PDF-файлу.")
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

                                if not file_obj:
                                    log_box.warning(f"⏩ Пропущено {comp}: файл відсутній.")
                                    progress_bar.progress((i + 1) / total_count)
                                    continue

                                msg = MIMEMultipart()
                                msg["From"] = sender_email
                                msg["To"] = target_to
                                msg["Subject"] = subject_template.format(company=comp)

                                body = email_body_template.format(company=comp)
                                msg.attach(MIMEText(body, "html", "utf-8"))

                                # Новий варіант: чіткий PDF-тип + підтримка українських літер у назві
                                pdf_bytes = file_obj.getvalue()
                                part = MIMEApplication(pdf_bytes, _subtype="pdf")
                                part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", file_obj.name))
                                msg.attach(part)

                                # 1. Відправка адресату через SMTP
                                server.send_message(msg)
                                
                                # 2. Збереження копії у папку "Надіслані" через IMAP
                                save_to_sent_folder(
                                    imap_host=mail_server,
                                    imap_port=int(imap_port),
                                    user=sender_email,
                                    password=sender_password,
                                    raw_msg_bytes=msg.as_bytes()
                                )

                                sent += 1
                                log_box.success(f"✅ Надіслано та збережено в 'Надіслані' ({sent}/{ready_count}): **{comp}** ➔ `{target_to}` | Файл: `{file_obj.name}`")

                                progress_bar.progress((i + 1) / total_count)
                                status_text.text(f"Опрацьовано {i + 1} з {total_count}...")
                                time.sleep(delay_seconds)

                        st.balloons()
                        st.success(f"🎉 Розсилку завершено! Успішно надіслано та заархівовано листів: {sent}")

                    except Exception as e:
                        st.error(f"❌ Помилка під час відправки: {e}")

    except Exception as err:
        st.error(f"Помилка зчитування файлів: {err}")
else:
    st.info("👆 Завантажте файл таблиці ліворуч та виділіть усі PDF-файли праворуч, щоб переглянути зіставлення.")