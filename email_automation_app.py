import io
import smtplib
import ssl
import time
from dataclasses import dataclass
from email.message import EmailMessage
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


REQUIRED_COLUMNS = {"Email"}


@dataclass
class CampaignResult:
    sent: List[Dict[str, str]]
    failed: List[Dict[str, str]]


def validate_dataframe(df: pd.DataFrame) -> Tuple[bool, str]:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        return False, f"Missing required column(s): {', '.join(missing)}"
    if df.empty:
        return False, "The uploaded file has no contacts."
    return True, ""


def personalize_text(template: str, row: pd.Series) -> str:
    safe_row = {k: "" if pd.isna(v) else str(v) for k, v in row.to_dict().items()}
    try:
        return template.format_map(safe_row)
    except KeyError:
        return template


def send_campaign(
    contacts_df: pd.DataFrame,
    sender_email: str,
    sender_password: str,
    smtp_server: str,
    smtp_port: int,
    subject_template: str,
    message_template: str,
    delay_seconds: int,
    use_tls: bool,
    progress_bar,
    status_placeholder,
    log_placeholder,
) -> CampaignResult:
    sent: List[Dict[str, str]] = []
    failed: List[Dict[str, str]] = []
    total = len(contacts_df)
    logs: List[str] = []

    def update_ui(index: int, current_email: str, current_status: str) -> None:
        progress_bar.progress(index / total, text=f"Emails sent: {index} / {total}")
        status_placeholder.markdown(
            f"""
            <div class='status-card'>
                <p><b>Emails Sent:</b> {index} / {total}</p>
                <p><b>Current Email:</b> {current_email}</p>
                <p><b>Status:</b> {current_status}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        log_placeholder.code("\n".join(logs[-12:]) if logs else "No logs yet.", language="bash")

    try:
        if use_tls:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls(context=ssl.create_default_context())
        else:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=ssl.create_default_context())

        with server:
            server.login(sender_email, sender_password)

            for idx, (_, row) in enumerate(contacts_df.iterrows(), start=1):
                recipient = str(row.get("Email", "")).strip()
                if not recipient:
                    failed.append({"Email": "", "Reason": "Missing recipient email"})
                    logs.append(f"[FAILED] row={idx} reason=Missing recipient email")
                    update_ui(idx - 1, "N/A", "Skipped")
                    continue

                message = EmailMessage()
                message["From"] = sender_email
                message["To"] = recipient
                message["Subject"] = personalize_text(subject_template, row)
                message.set_content(personalize_text(message_template, row))

                update_ui(idx - 1, recipient, "Sending...")

                try:
                    server.send_message(message)
                    sent.append({"Email": recipient, "Status": "Sent"})
                    logs.append(f"[SENT] {recipient}")
                    update_ui(idx, recipient, "Sent")
                except Exception as exc:  # noqa: BLE001
                    failed.append({"Email": recipient, "Reason": str(exc)})
                    logs.append(f"[FAILED] {recipient} reason={exc}")
                    update_ui(idx, recipient, "Failed")

                if idx < total:
                    time.sleep(delay_seconds)

    except Exception as exc:  # noqa: BLE001
        logs.append(f"[ERROR] SMTP connection failed: {exc}")
        log_placeholder.code("\n".join(logs[-12:]), language="bash")
        st.error(f"Campaign stopped due to SMTP error: {exc}")

    return CampaignResult(sent=sent, failed=failed)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="failed_emails")
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Email Automation Studio", page_icon="✉️", layout="wide")

    st.markdown(
        """
        <style>
            .stApp { background: linear-gradient(180deg, #f7f9fc 0%, #f2f5fb 100%); }
            .main-title { font-size: 2rem; font-weight: 700; color: #1f2937; margin-bottom: .2rem; }
            .subtitle { color: #4b5563; margin-bottom: 1.4rem; }
            .section-card {
                background: white; border-radius: 16px; padding: 1rem 1.2rem;
                box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08); margin-bottom: 1rem;
            }
            .status-card {
                background: #eef6ff; border-left: 4px solid #3b82f6; border-radius: 10px;
                padding: .8rem 1rem; margin-top: .6rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div class='main-title'>Email Automation Studio</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='subtitle'>Upload contacts, personalize your message, and send campaigns safely with timed intervals.</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("1) Upload Excel Contacts")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    contacts_df = None
    if uploaded_file:
        contacts_df = pd.read_excel(uploaded_file)
        valid, error_msg = validate_dataframe(contacts_df)
        if not valid:
            st.error(error_msg)
            contacts_df = None
        else:
            st.success(f"Loaded {len(contacts_df)} contacts")
            st.dataframe(contacts_df, use_container_width=True, height=220)
    st.markdown("</div>", unsafe_allow_html=True)

    left_col, right_col = st.columns(2)

    with left_col:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("2) Sender Email Configuration")
        sender_email = st.text_input("Sender Email", placeholder="you@company.com")
        sender_password = st.text_input("Password / App Password", type="password")
        smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP Port", min_value=1, max_value=65535, value=587, step=1)
        use_tls = st.checkbox("Use STARTTLS (recommended for port 587)", value=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right_col:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("3) Email Composition")
        subject_template = st.text_input("Email Subject", value="Hello {Name}")
        message_template = st.text_area(
            "Email Message",
            value="Hello {Name},\n\nWe would love to connect with {Company}.\n\nBest regards,",
            height=180,
        )
        st.caption("Use placeholders like {Name}, {Email}, {Company} based on your Excel column names.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("4) Delay Configuration")
    delay_seconds = st.slider("Delay between emails (seconds)", min_value=1, max_value=600, value=30)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("5) Start Campaign")
    start_button = st.button("Start Sending Emails", type="primary", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.subheader("6) Email Sending Progress")
    progress_bar = st.progress(0, text="Emails sent: 0 / 0")
    status_placeholder = st.empty()
    log_placeholder = st.empty()

    if start_button:
        if contacts_df is None:
            st.error("Please upload a valid Excel file with an 'Email' column before starting.")
            return

        required_values = {
            "Sender Email": sender_email,
            "Password": sender_password,
            "SMTP Server": smtp_server,
            "SMTP Port": smtp_port,
            "Email Subject": subject_template,
            "Email Message": message_template,
        }
        missing_fields = [field for field, value in required_values.items() if not str(value).strip()]
        if missing_fields:
            st.error(f"Please complete all fields: {', '.join(missing_fields)}")
            return

        result = send_campaign(
            contacts_df=contacts_df,
            sender_email=sender_email,
            sender_password=sender_password,
            smtp_server=smtp_server,
            smtp_port=int(smtp_port),
            subject_template=subject_template,
            message_template=message_template,
            delay_seconds=delay_seconds,
            use_tls=use_tls,
            progress_bar=progress_bar,
            status_placeholder=status_placeholder,
            log_placeholder=log_placeholder,
        )

        st.subheader("7) Campaign Logs")
        success_df = pd.DataFrame(result.sent)
        failed_df = pd.DataFrame(result.failed)

        success_col, failed_col = st.columns(2)
        with success_col:
            st.markdown("#### Successfully Sent Emails")
            st.metric("Total Sent", len(success_df))
            if not success_df.empty:
                st.dataframe(success_df, use_container_width=True)

        with failed_col:
            st.markdown("#### Failed Emails")
            st.metric("Total Failed", len(failed_df))
            if not failed_df.empty:
                st.dataframe(failed_df, use_container_width=True)
                st.download_button(
                    "Export Failed Emails to Excel",
                    data=to_excel_bytes(failed_df),
                    file_name="failed_emails.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()
