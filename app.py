import io
import smtplib
import time
from dataclasses import dataclass
from email.message import EmailMessage
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


@dataclass
class SendResult:
    success: List[Dict[str, str]]
    failed: List[Dict[str, str]]


class SafeFormatDict(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def load_contacts(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    if "Email" not in df.columns:
        raise ValueError("Excel file must include an 'Email' column.")
    return df.dropna(subset=["Email"]).copy()


def personalized_text(template: str, row: pd.Series) -> str:
    data = {k: "" if pd.isna(v) else str(v) for k, v in row.to_dict().items()}
    return template.format_map(SafeFormatDict(data))


def create_message(sender: str, recipient: str, subject: str, body: str) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    return message


def connect_smtp(server: str, port: int, email: str, password: str):
    if int(port) == 465:
        smtp = smtplib.SMTP_SSL(server, int(port), timeout=30)
        smtp.login(email, password)
        return smtp

    smtp = smtplib.SMTP(server, int(port), timeout=30)
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
    smtp.login(email, password)
    return smtp


def send_campaign(
    contacts_df: pd.DataFrame,
    sender_email: str,
    sender_password: str,
    smtp_server: str,
    smtp_port: int,
    subject_template: str,
    body_template: str,
    delay_seconds: int,
) -> SendResult:
    success: List[Dict[str, str]] = []
    failed: List[Dict[str, str]] = []

    progress_bar = st.progress(0)
    metrics = st.empty()
    live_status = st.empty()
    live_log = st.empty()

    total = len(contacts_df)
    rolling_logs: List[str] = []

    with connect_smtp(smtp_server, smtp_port, sender_email, sender_password) as smtp:
        for index, (_, row) in enumerate(contacts_df.iterrows(), start=1):
            recipient = str(row.get("Email", "")).strip()
            if not recipient:
                failed.append({"Email": "", "Reason": "Missing email address"})
                continue

            name_or_email = str(row.get("Name", recipient))
            live_status.info(f"Sending to **{recipient}**...")
            metrics.markdown(
                f"### Emails Sent: {len(success)} / {total}  \n"
                f"**Current Email:** {recipient}  \n"
                f"**Status:** Sending..."
            )

            try:
                personalized_subject = personalized_text(subject_template, row)
                personalized_body = personalized_text(body_template, row)
                message = create_message(sender_email, recipient, personalized_subject, personalized_body)
                smtp.send_message(message)
                success.append({"Name": name_or_email, "Email": recipient, "Status": "Sent"})
                rolling_logs.append(f"✅ Sent: {recipient}")
            except Exception as exc:  # noqa: BLE001
                failed.append({"Name": name_or_email, "Email": recipient, "Reason": str(exc)})
                rolling_logs.append(f"❌ Failed: {recipient} ({exc})")

            progress_bar.progress(index / total)
            metrics.markdown(
                f"### Emails Sent: {len(success)} / {total}  \n"
                f"**Current Email:** {recipient}  \n"
                f"**Status:** Completed"
            )
            live_log.code("\n".join(rolling_logs[-12:]) if rolling_logs else "No logs yet")

            if index < total:
                time.sleep(delay_seconds)

    live_status.success("Campaign finished.")
    return SendResult(success=success, failed=failed)


def failed_to_excel(failed_rows: List[Dict[str, str]]) -> bytes:
    output = io.BytesIO()
    failed_df = pd.DataFrame(failed_rows)
    failed_df.to_excel(output, index=False)
    output.seek(0)
    return output.read()


def init_page() -> None:
    st.set_page_config(page_title="Email Automation Studio", page_icon="📨", layout="wide")
    st.markdown(
        """
        <style>
            .stApp {
                background: linear-gradient(180deg, #f8fbff 0%, #f5f8fc 100%);
                color: #13233a;
            }
            .block-container {padding-top: 1.6rem; padding-bottom: 2rem;}
            .card {
                background: #ffffff;
                border: 1px solid #eaf1fb;
                border-radius: 14px;
                padding: 1rem 1.2rem;
                box-shadow: 0 8px 24px rgba(16, 39, 72, 0.06);
                margin-bottom: 1rem;
            }
            .hero-title {
                font-size: 1.8rem;
                font-weight: 700;
                margin-bottom: 0.1rem;
            }
            .hero-subtitle {color: #4b5f7e; font-size: 1rem;}
            div[data-testid="stButton"] button {
                background: linear-gradient(90deg, #2563eb, #1d4ed8);
                border: none;
                color: white;
                border-radius: 10px;
                padding: 0.62rem 1.2rem;
                font-weight: 600;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    init_page()

    st.markdown('<div class="hero-title">📨 Email Automation Studio</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="hero-subtitle">Upload contacts, personalize your message, and launch a reliable bulk email campaign.</div>',
        unsafe_allow_html=True,
    )
    st.write("")

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("1) Upload Excel Contacts")
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx", "xls"])
    contacts_df = None

    if uploaded_file:
        try:
            contacts_df = load_contacts(uploaded_file)
            st.success(f"Loaded {len(contacts_df)} contacts.")
            st.dataframe(contacts_df.head(20), use_container_width=True)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Unable to read contacts: {exc}")
    st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("2) Sender Configuration")
        sender_email = st.text_input("Sender Email")
        sender_password = st.text_input("Password / App Password", type="password")
        smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP Port", min_value=1, max_value=65535, value=587, step=1)
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("3) Email Composition")
        subject_template = st.text_input("Email Subject", value="Hello {Name}, quick update for you")
        body_template = st.text_area(
            "Email Message",
            height=220,
            value=(
                "Hello {Name},\n\n"
                "I hope you're doing well at {Company}.\n"
                "I wanted to quickly reach out regarding our latest offer.\n\n"
                "Best regards,\n"
                "Your Name"
            ),
        )
        st.caption("Use placeholders from Excel columns, e.g., {Name}, {Company}, {Email}.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("4) Delay Configuration")
    delay_seconds = st.number_input("Delay between emails (seconds)", min_value=0, max_value=3600, value=30, step=1)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("5) Launch Campaign")
    start = st.button("Start Sending Emails", use_container_width=True)

    if start:
        missing_fields: List[Tuple[str, bool]] = [
            ("Excel file", contacts_df is not None and len(contacts_df) > 0),
            ("Sender email", bool(sender_email.strip())),
            ("Password", bool(sender_password.strip())),
            ("SMTP server", bool(smtp_server.strip())),
            ("Email subject", bool(subject_template.strip())),
            ("Email message", bool(body_template.strip())),
        ]
        missing = [name for name, ok in missing_fields if not ok]

        if missing:
            st.error(f"Please complete: {', '.join(missing)}")
        else:
            st.subheader("6) Sending Progress")
            result = send_campaign(
                contacts_df=contacts_df,
                sender_email=sender_email,
                sender_password=sender_password,
                smtp_server=smtp_server,
                smtp_port=int(smtp_port),
                subject_template=subject_template,
                body_template=body_template,
                delay_seconds=int(delay_seconds),
            )

            st.subheader("7) Campaign Logs")
            sent_df = pd.DataFrame(result.success)
            failed_df = pd.DataFrame(result.failed)

            col_s, col_f = st.columns(2)
            with col_s:
                st.success(f"Successfully sent: {len(sent_df)}")
                if not sent_df.empty:
                    st.dataframe(sent_df, use_container_width=True)

            with col_f:
                st.error(f"Failed: {len(failed_df)}")
                if not failed_df.empty:
                    st.dataframe(failed_df, use_container_width=True)
                    st.download_button(
                        "Export Failed Emails to Excel",
                        data=failed_to_excel(result.failed),
                        file_name="failed_emails.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
