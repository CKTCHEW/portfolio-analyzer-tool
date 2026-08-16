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
        if re.match(r"^\d{4}\s+Return\s+\(%\)$", col):
            year = int(col.split()[0])
            year_returns.append((year, col))
        elif re.match(r"^\d{4}\s+Benchmark\s+\(%\)$", col):
            year = int(col.split()[0])
            year_benchmarks.append((year, col))
    year_returns.sort(key=lambda x: x[0])
    year_benchmarks.sort(key=lambda x: x[0])
    return year_returns, year_benchmarks


def process_pdf_ffs_with_gemini(uploaded_pdfs):
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error(
            "❌ GEMINI_API_KEY is missing in Streamlit Cloud Secrets. Please add it to proceed."
        )
        return None

    genai.configure(api_key=api_key)
    candidate_models = ["gemini-1.5-flash", "gemini-1.5-pro"]
    extracted_records = []
    total_files = len(uploaded_pdfs)

    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, pdf_file in enumerate(uploaded_pdfs):
        status_text.text(
            f"⏳ Extracting FFS ({idx + 1}/{total_files}): {pdf_file.name}..."
        )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_file.read())
            tmp_path = tmp.name

        uploaded_gemini_file = None
        parsed_successfully = False
        last_error_msg = ""

        try:
            uploaded_gemini_file = genai.upload_file(
                tmp_path, mime_type="application/pdf"
            )

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

    weights = np.full(n, 0.05)
    remaining_weight = 1.0 - np.sum(weights)

    sorted_indices = np.argsort(returns)[::-1]

    for idx in sorted_indices:
        if remaining_weight <= 0.001:
            break

        max_add = 0.40 - weights[idx]
        add_weight = min(max_add, remaining_weight)

        step = add_weight
        while step > 0.001:
            test_weights = weights.copy()
            test_weights[idx] += step
            test_vol = np.sum(volatilities * test_weights)

            if test_vol <= max_volatility:
                weights[idx] += step
                remaining_weight -= step
                break
            else:
                step /= 2.0

    total_w = np.sum(weights)
    if total_w > 0:
        weights = weights / total_w
    opt_return = float(np.sum(returns * weights))
    opt_vol = float(np.sum(volatilities * weights))

    return weights, opt_return, opt_vol


def calculate_portfolio_metrics(df, weights, year_returns_cols):
    if "3Y Return (%)" in df.columns:
        returns_1y = df["1Y Return (%)"].fillna(df["3Y Return (%)"])
    else:
        returns_1y = df["1Y Return (%)"]

    volatilities = df["Volatility (%)"]
    fees = df["Mgmt Fee (%)"]

    valid_return_mask = returns_1y.notna()
    valid_vol_mask = volatilities.notna()
    valid_fee_mask = fees.notna()

    portfolio_return = float(np.sum(returns_1y[valid_return_mask] * weights[valid_return_mask])) if valid_return_mask.any() else 0.0
    portfolio_volatility = float(np.sum(volatilities[valid_vol_mask] * weights[valid_vol_mask])) if valid_vol_mask.any() else 0.0
    portfolio_fee = float(np.sum(fees[valid_fee_mask] * weights[valid_fee_mask])) if valid_fee_mask.any() else 0.0

    risk_adjusted = portfolio_return / portfolio_volatility if portfolio_volatility > 0 else 0.0

    yearly_returns = []
    for year, col_name in year_returns_cols:
        if col_name in df.columns:
            year_data = df[col_name]
            valid_mask = year_data.notna()
            if valid_mask.any():
                year_return = float(np.sum(year_data[valid_mask] * weights[valid_mask]))
                yearly_returns.append(year_return)

    best_year = float(max(yearly_returns)) if yearly_returns else None
    worst_year = float(min(yearly_returns)) if yearly_returns else None
    avg_yearly = float(np.mean(yearly_returns)) if yearly_returns else None
    positive_years = sum(1 for r in yearly_returns if r > 0)
    consistency = float((positive_years / len(yearly_returns) * 100)) if yearly_returns else 0.0

    return {
        "return": portfolio_return,
        "volatility": portfolio_volatility,
        "fee": portfolio_fee,
        "risk_adjusted": risk_adjusted,
        "best_year": best_year,
        "worst_year": worst_year,
        "avg_yearly": avg_yearly,
        "consistency": consistency,
        "yearly_returns": yearly_returns,
    }


# ============================================================
# SECTION 7: LOGIN PAGE
# ============================================================
def show_login_page():
    st.markdown(
        "<h2 style='text-align: center;'>🔐 Login to Portfolio Analyzer</h2>",
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns(