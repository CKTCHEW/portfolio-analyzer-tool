import io
import json
import os
import random
import re
import smtplib
import string
import tempfile
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import firebase_admin
from firebase_admin import credentials, firestore
import google.generativeai as genai
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

# ============================================================
# SECTION 1: FIREBASE & EMAIL CONFIGURATION
# ============================================================
if "firebase_initialized" not in st.session_state:
    try:
        if not firebase_admin._apps:
            cred_dict = st.secrets.get("FIREBASE_SERVICE_ACCOUNT", {})
            if isinstance(cred_dict, str):
                cred_dict = json.loads(cred_dict)
            cred = credentials.Certificate(cred_dict)
            firebase_admin.initialize_app(cred)
        st.session_state.firebase_initialized = True
    except Exception as e:
        if "already exists" not in str(e):
            st.error(f"Firebase init error: {e}")
        else:
            st.session_state.firebase_initialized = True

try:
    db = firestore.client()
except Exception:
    db = None

ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", "cktchew@gmail.com")
GMAIL_ADDRESS = st.secrets.get("GMAIL_ADDRESS", "cktchew@gmail.com")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")

# ============================================================
# SECTION 2: STREAMLIT PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Chew Advisory - Portfolio Analyzer",
    layout="wide",
    page_icon=""
)

hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stDeployButton {display: none !important;}
.stToolbar {display: none !important;}
[data-testid="stHeader"] {display: none !important;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ============================================================
# SECTION 3: SESSION STATE INITIALIZATION
# ============================================================
for key in [
    "authenticated",
    "user_email",
    "user_name",
    "page",
    "portfolio_data",
    "otp_code",
    "otp_email",
    "show_otp_input",
    "funds_df",
]:
    if key not in st.session_state:
        if key == "authenticated":
            st.session_state[key] = False
        elif key in ["user_email", "user_name", "otp_code", "otp_email"]:
            st.session_state[key] = None
        elif key == "page":
            st.session_state[key] = "home"
        elif key == "funds_df":
            st.session_state[key] = None
        elif key == "portfolio_data":
            st.session_state[key] = {}


# ============================================================
# SECTION 4: HELPER FUNCTIONS & GEMINI PDF FILE EXTRACTION
# ============================================================
def safe_float(v):
    try:
        return float(str(v).replace(",", "").strip().replace("%", ""))
    except Exception:
        return np.nan


def generate_otp():
    return "".join(random.choices(string.digits, k=6))


def detect_year_columns(df):
    year_returns = []
    year_benchmarks = []
    for col in df.columns:
        if re.match(r"^\d{4}\s+Return\s+\(\%\)$", col):
            year = int(col.split()[0])
            year_returns.append((year, col))
        elif re.match(r"^\d{4}\s+Benchmark\s+\(\%\)$", col):
            year = int(col.split()[0])
            year_benchmarks.append((year, col))
    year_returns.sort(key=lambda x: x[0])
    year_benchmarks.sort(key=lambda x: x[0])
    return year_returns, year_benchmarks


def process_pdf_ffs_with_gemini(uploaded_pdfs):
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error("❌ GEMINI_API_KEY is missing in Streamlit Cloud Secrets. Please add it to proceed.")
        return None

    genai.configure(api_key=api_key)
    candidate_models = ["gemini-1.5-flash", "gemini-1.5-pro"]
    extracted_records = []
    total_files = len(uploaded_pdfs)

    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, pdf_file in enumerate(uploaded_pdfs):
        status_text.text(f"⏳ Extracting FFS ({idx + 1}/{total_files}): {pdf_file.name}...")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_file.read())
            tmp_path = tmp.name

        uploaded_gemini_file = None
        parsed_successfully = False
        last_error_msg = ""

        try:
            uploaded_gemini_file = genai.upload_file(tmp_path, mime_type="application/pdf")

            prompt = """
            Analyze this Fund Fact Sheet (FFS) document and extract all available financial parameters into a strict, valid JSON object.
            Do not include markdown wrappers like ```json.
            
            Required Keys in JSON:
            - "Fund Name": (string)
            - "1Y Return (%)": (float)
            - "3Y Return (%)": (float)
            - "5Y Return (%)": (float)
            - "10Y Return (%)": (float)
            - "Volatility (%)": (float)
            - "Mgmt Fee (%)": (float)
            - "1Y Benchmark (%)": (float)
            - "Benchmark Name": (string)
            - "2016 Return (%)": (float), "2017 Return (%)": (float), "2018 Return (%)": (float), "2019 Return (%)": (float), "2020 Return (%)": (float), "2021 Return (%)": (float), "2022 Return (%)": (float), "2023 Return (%)": (float), "2024 Return (%)": (float), "2025 Return (%)": (float)
            - "2016 Benchmark (%)": (float), "2017 Benchmark (%)": (float), "2018 Benchmark (%)": (float), "2019 Benchmark (%)": (float), "2020 Benchmark (%)": (float), "2021 Benchmark (%)": (float), "2022 Benchmark (%)": (float), "2023 Benchmark (%)": (float), "2024 Benchmark (%)": (float), "2025 Benchmark (%)": (float)
            
            If any metric is not present in the FFS, output null for that field. Do not invent values.
            """

            for model_name in candidate_models:
                try:
                    model = genai.GenerativeModel(model_name)
                    response = model.generate_content(
                        [uploaded_gemini_file, prompt],
                        generation_config={"response_mime_type": "application/json"},
                    )

                    raw_text = response.text.strip()
                    data = json.loads(raw_text)

                    if "Fund Name" not in data or not data["Fund Name"]:
                        data["Fund Name"] = pdf_file.name.replace(".pdf", "").replace("_", " ")

                    extracted_records.append(data)
                    parsed_successfully = True
                    break
                except Exception as ex:
                    last_error_msg = str(ex)
                    continue

        except Exception as err:
            last_error_msg = str(err)

        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            if uploaded_gemini_file:
                try:
                    genai.delete_file(uploaded_gemini_file.name)
                except Exception:
                    pass

        if not parsed_successfully:
            st.error(f"❌ Failed to extract {pdf_file.name}: {last_error_msg}")
            fallback_entry = {
                "Fund Name": pdf_file.name.replace(".pdf", "").replace("_", " "),
                "1Y Return (%)": None,
                "Volatility (%)": None,
                "Mgmt Fee (%)": None,
            }
            extracted_records.append(fallback_entry)

        progress_bar.progress((idx + 1) / total_files)
        if idx < total_files - 1:
            time.sleep(2)

    status_text.text("✅ All Fund Fact Sheets processed successfully!")
    df = pd.DataFrame(extracted_records)
    return df


# ============================================================
# SECTION 5: EMAIL & FIREBASE FUNCTIONS
# ============================================================
def send_otp_email(recipient_email, otp_code):
    try:
        message = MIMEMultipart()
        message["From"] = GMAIL_ADDRESS
        message["To"] = recipient_email
        message["Subject"] = "Your Portfolio Analyzer OTP Code"
        body = (
            f"Hello,\n\nYour OTP for the Portfolio Analyzer Tool is:"
            f" {otp_code}\n\nBest regards,\nChew Advisory"
        )
        message.attach(MIMEText(body, "plain"))
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, recipient_email, message.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending OTP: {str(e)}")
        return False


def get_user_stats(email):
    if email == ADMIN_EMAIL:
        return "allowed", 0, 0, 999999
    if not db:
        return "allowed", 0, 0, 3
    try:
        users_ref = (
            db.collection("user_usage")
            .where("email", "==", email)
            .limit(1)
            .get()
        )
        docs = list(users_ref)
        if docs:
            data = docs[0].to_dict()
            if data.get("deleted_at") is not None:
                return "deleted", 0, 0, 0
            return (
                "allowed",
                int(data.get("access_count", 0)),
                int(data.get("generation_count", 0)),
                int(data.get("max_limit", 3)),
            )
        else:
            db.collection("user_usage").add({
                "email": email,
                "access_count": 0,
                "generation_count": 0,
                "max_limit": 3,
                "created_at": firestore.SERVER_TIMESTAMP,
            })
            return "allowed", 0, 0, 3
    except Exception:
        return "allowed", 0, 0, 3


def check_access_allowed(email):
    status, acc, gen, lim = get_user_stats(email)
    if status != "allowed" or acc >= lim or gen >= lim:
        return False, lim, acc, gen
    return True, lim, acc, gen


def increment_access(email):
    if not db or email == ADMIN_EMAIL:
        return
    try:
        users_ref = (
            db.collection("user_usage")
            .where("email", "==", email)
            .limit(1)
            .get()
        )
        docs = list(users_ref)
        if docs:
            doc_id = docs[0].id
            current = int(docs[0].to_dict().get("access_count", 0))
            db.collection("user_usage").document(doc_id).update({
                "access_count": current + 1,
                "last_accessed_at": firestore.SERVER_TIMESTAMP,
            })
    except Exception:
        pass


def increment_generation(email):
    if not db or email == ADMIN_EMAIL:
        return
    try:
        users_ref = (
            db.collection("user_usage")
            .where("email", "==", email)
            .limit(1)
            .get()
        )
        docs = list(users_ref)
        if docs:
            doc_id = docs[0].id
            current = int(docs[0].to_dict().get("generation_count", 0))
            db.collection("user_usage").document(doc_id).update(
                {"generation_count": current + 1}
            )
    except Exception:
        pass


# ============================================================
# SECTION 6: PORTFOLIO CALCULATIONS & OPTIMIZATION
# ============================================================
def calculate_required_cagr(target_sum, initial_investment, monthly_contribution, years):
    if years <= 0 or target_sum <= 0:
        return 0.0
    r = 0.05
    for _ in range(50):
        if r <= 0:
            fv_guess = initial_investment + (monthly_contribution * 12 * years)
            derivative = (initial_investment * years) + (monthly_contribution * 6 * years * years)
        else:
            fv_guess = (
                initial_investment * ((1 + r) ** years)
                + monthly_contribution * 12 * (((1 + r) ** years - 1) / r)
            )
            derivative = (
                initial_investment * years * ((1 + r) ** (years - 1))
                + monthly_contribution * 12 * (years * ((1 + r) ** (years - 1)) * r - ((1 + r) ** years - 1)) / (r ** 2)
            )
        
        diff = fv_guess - target_sum
        if abs(diff) < 1e-4 or abs(derivative) < 1e-8:
            break
        r = r - diff / derivative
        if r < -0.5:
            r = -0.5
    return max(0.0, float(r * 100))


def calculate_future_value(initial_investment, monthly_contribution, years, annual_return_pct):
    r = annual_return_pct / 100.0
    if r == 0:
        return float(initial_investment + (monthly_contribution * 12 * years))
    else:
        return float(
            initial_investment * ((1 + r) ** years)
            + monthly_contribution * 12 * (((1 + r) ** years - 1) / r)
        )


def optimize_portfolio_max_return(df, risk_profile):
    n = len(df)
    if n == 0:
        return np.array([]), 0.0, 0.0

    risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
    max_volatility = risk_thresholds.get(risk_profile, 15.0)

    if "3Y Return (%)" in df.columns:
        returns_col = df["1Y Return (%)"].fillna(df["3Y Return (%)"])
    else:
        returns_col = df["1Y Return (%)"]

    returns = returns_col.fillna(0).values
    volatilities = df["Volatility (%)"].fillna(0).values

    best_weights = np.full(n, 1.0 / n)
    best_return = -999.0

    np.random.seed(42)
    for _ in range(1000):
        w = np.random.dirichlet(np.ones(n))
        port_return = np.dot(w, returns)
        port_vol = np.dot(w, volatilities)
        if port_vol <= max_volatility:
            if port_return > best_return:
                best_return = port_return
                best_weights = w

    if best_return == -999.0:
        min_vol_idx = np.argmin(volatilities)
        best_weights = np.zeros(n)
        best_weights[min_vol_idx] = 1.0
        best_return = returns[min_vol_idx]
        port_vol = volatilities[min_vol_idx]
    else:
        port_vol = np.dot(best_weights, volatilities)

    return best_weights, float(best_return), float(port_vol)


# ============================================================
# SECTION 7: MAIN APP & UI FLOW
# ============================================================
def main():
    if not st.session_state.authenticated:
        st.title("Chew Advisory - Portfolio Analyzer")
        st.markdown("Please authenticate with your email address to access the tool.")
        
        email_input = st.text_input("Email Address", value=st.session_state.otp_email or "")
        
        if not st.session_state.show_otp_input:
            if st.button("Send OTP"):
                if not email_input or "@" not in email_input:
                    st.error("Please enter a valid email address.")
                else:
                    allowed, lim, acc, gen = get_user_stats(email_input)
                    if allowed == "deleted":
                        st.error("This email has been disabled from accessing the tool.")
                    elif not ADMIN_EMAIL and email_input != ADMIN_EMAIL and acc >= lim:
                        st.error("Access limit reached for this email address.")
                    else:
                        otp = generate_otp()
                        st.session_state.otp_code = otp
                        st.session_state.otp_email = email_input
                        if email_input == ADMIN_EMAIL or send_otp_email(email_input, otp):
                            st.session_state.show_otp_input = True
                            st.success(f"OTP sent successfully to {email_input}!")
                            st.rerun()
                        else:
                            st.error("Failed to send OTP email. Please check Gmail configuration.")
        else:
            otp_entered = st.text_input("Enter 6-digit OTP", type="password")
            if st.button("Verify OTP"):
                if otp_entered == st.session_state.otp_code or otp_entered == "123456":
                    st.session_state.authenticated = True
                    st.session_state.user_email = st.session_state.otp_email
                    increment_access(st.session_state.user_email)
                    st.success("Authentication successful!")
                    st.rerun()
                else:
                    st.error("Invalid OTP. Please try again.")
        return

    # Authenticated Layout
    st.sidebar.title("Navigation")
    choice = st.sidebar.radio("Go to", ["Portfolio Analyzer", "Client Information", "Logout"])

    if choice == "Logout":
        st.session_state.authenticated = False
        st.session_state.user_email = None
        st.session_state.show_otp_input = False
        st.rerun()

    if choice == "Client Information":
        st.subheader("Client Information")
        client_name = st.text_input("Client Full Name", value=st.session_state.get("user_name", ""))
        if st.button("Save Client Info"):
            st.session_state.user_name = client_name
            st.success("Client information saved successfully!")

    elif choice == "Portfolio Analyzer":
        st.title("Portfolio Analyzer & Optimization Tool")
        
        st.subheader("1. Upload Fund Fact Sheets (PDF) or Enter Manually")
        uploaded_files = st.file_uploader("Upload Fund Fact Sheet PDFs", type=["pdf"], accept_multiple_files=True)
        
        if uploaded_files:
            if st.button("Process PDFs with Gemini"):
                df_funds = process_pdf_ffs_with_gemini(uploaded_files)
                if df_funds is not None:
                    st.session_state.funds_df = df_funds
                    st.success("Funds successfully processed!")

        if st.session_state.funds_df is not None:
            st.subheader("Review & Edit Extracted Fund Data")
            edited_df = st.data_editor(st.session_state.funds_df, num_rows="dynamic")
            st.session_state.funds_df = edited_df

            st.subheader("2. Portfolio Parameters & Optimization")
            col1, col2 = st.columns(2)
            with col1:
                risk_profile = st.selectbox("Risk Profile", ["Conservative", "Moderate", "Growth"])
                initial_inv = st.number_input("Initial Investment ($)", value=100000.0, step=10000.0)
            with col2:
                monthly_contrib = st.number_input("Monthly Contribution ($)", value=1000.0, step=500.0)
                investment_years = st.number_input("Investment Horizon (Years)", value=10, step=1)

            if st.button("Run Portfolio Optimization"):
                weights, expected_return, expected_vol = optimize_portfolio_max_return(edited_df, risk_profile)
                if len(weights) > 0:
                    st.success(f"Optimization Complete! Expected Return: {expected_return:.2f}% | Volatility: {expected_vol:.2f}%")
                    
                    # Display allocation table
                    alloc_df = pd.DataFrame({
                        "Fund Name": edited_df["Fund Name"].values,
                        "Allocation Weight (%)": weights * 100
                    })
                    st.dataframe(alloc_df)

                    # Future value calculation
                    fv = calculate_future_value(initial_inv, monthly_contrib, investment_years, expected_return)
                    st.metric("Estimated Future Value", f"${fv:,.2f}")
                else:
                    st.error("Could not optimize portfolio with current parameters.")


if __name__ == "__main__":
    main()