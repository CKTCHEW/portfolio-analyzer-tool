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

import firebase_admin
from firebase_admin import credentials, firestore
import google.generativeai as genai
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

# ============================================================
# SECTION: FIREBASE & EMAIL CONFIGURATION
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
# SECTION: STREAMLIT PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Chew Advisory - Portfolio Analyzer", layout="wide", page_icon=""
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
# SECTION: SESSION STATE INITIALIZATION
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
    st.session_state[key] = (
        False
        if key == "authenticated"
        else (
            None
            if key
            in ["user_email", "user_name", "otp_code", "otp_email", "funds_df"]
            else "home"
            if key == "page"
            else {}
        )
    )


# ============================================================
# SECTION: HELPER FUNCTIONS
# ============================================================
def safe_float(v):
  try:
    return float(str(v).replace(",", "").strip())
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


# ============================================================
# SECTION: EMAIL & FIREBASE FUNCTIONS
# ============================================================
def send_otp_email(recipient_email, otp_code):
  try:
    message = MIMEMultipart()
    message["From"] = GMAIL_ADDRESS
    message["To"] = recipient_email
    message["Subject"] = "Your Portfolio Analyzer OTP Code"
    body = f"Hello,\n\nYour OTP for the Portfolio Analyzer Tool is: {otp_code}\n\nBest regards,\nChew Advisory"
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
        db.collection("user_usage").where("email", "==", email).limit(1).get()
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
        db.collection("user_usage").where("email", "==", email).limit(1).get()
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
        db.collection("user_usage").where("email", "==", email).limit(1).get()
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
# SECTION: GEMINI PDF FFS EXTRACTION
# ============================================================
def process_pdf_ffs_with_gemini(uploaded_pdfs):
  api_key = st.secrets.get("GEMINI_API_KEY")
  if not api_key:
    st.error("❌ GEMINI_API_KEY missing in secrets.")
    return None
  genai.configure(api_key=api_key)
  records = []
  bar = st.progress(0)
  status = st.empty()

  for idx, pdf in enumerate(uploaded_pdfs):
    status.text(
        f"⏳ Extracting ({idx + 1}/{len(uploaded_pdfs)}): {pdf.name}..."
    )
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
      tmp.write(pdf.read())
      tmp_path = tmp.name

    g_file, success = None, False
    try:
      g_file = genai.upload_file(tmp_path, mime_type="application/pdf")
      prompt = (
          "Extract fund fact sheet data into a strict JSON object with keys:"
          ' "Fund Name", "1Y Return (%)", "3Y Return (%)", "5Y Return (%)",'
          ' "10Y Return (%)", "Volatility (%)", "Mgmt Fee (%)",'
          ' "1Y Benchmark (%)", "Benchmark Name". Output null for missing'
          " fields. Ensure numbers are floats."
      )
      model = genai.GenerativeModel("gemini-1.5-flash")
      resp = model.generate_content(
          [g_file, prompt],
          generation_config={"response_mime_type": "application/json"},
      )
      data = json.loads(resp.text.strip())
      if not data.get("Fund Name"):
        data["Fund Name"] = pdf.name.replace(".pdf", "")
      records.append(data)
      success = True
    except Exception as e:
      st.warning(f"Error processing {pdf.name}: {e}")
    finally:
      if os.path.exists(tmp_path):
        os.remove(tmp_path)
      if g_file:
        try:
          genai.delete_file(g_file.name)
        except Exception:
          pass
    if not success:
      records.append({
          "Fund Name": pdf.name.replace(".pdf", ""),
          "1Y Return (%)": None,
          "3Y Return (%)": None,
          "5Y Return (%)": None,
          "10Y Return (%)": None,
          "Volatility (%)": None,
          "Mgmt Fee (%)": None,
          "1Y Benchmark (%)": None,
          "Benchmark Name": None,
      })
    bar.progress((idx + 1) / len(uploaded_pdfs))

  status.text("✅ Gemini extraction complete!")
  time.sleep(0.5)
  status.empty()
  bar.empty()
  return pd.DataFrame(records)


# ============================================================
# SECTION: PORTFOLIO CALCULATIONS & OPTIMIZATION
# ============================================================
def calculate_required_cagr(
    target_sum, initial_investment, monthly_contribution, years
):
  if years <= 0 or target_sum <= 0:
    return 0.0
  r = 0.05
  for _ in range(50):
    fv_guess = initial_investment * ((1 + r) ** years) + (
        monthly_contribution
        * 12
        * (((1 + r) ** years - 1) / r
           if r > 0
           else monthly_contribution * 12 * years)
    )
    derivative = initial_investment * years * ((1 + r) ** (years - 1))
    if r > 0:
      derivative += (
          monthly_contribution
          * 12
          * (years * (1 + r) ** (years - 1) * r - ((1 + r) ** years - 1))
          / (r ** 2)
      )
    if abs(derivative) < 1e-8:
      break
    r = r - (fv_guess - target_sum) / derivative
    if r < -0.5:
      r = -0.5
  return max(0.0, r * 100)


def calculate_future_value(
    initial_investment, monthly_contribution, years, annual_return_pct
):
  r = annual_return_pct / 100
  if r == 0:
    fv = initial_investment + (monthly_contribution * 12 * years)
  else:
    fv = initial_investment * ((1 + r) ** years) + (
        monthly_contribution * 12 * (((1 + r) ** years - 1) / r)
    )
  return fv


def optimize_portfolio_max_return(df, risk_profile):
  n = len(df)
  if n == 0:
    return np.array([]), 0, 0

  risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
  max_volatility = risk_thresholds.get(risk_profile, 15.0)

  returns_col = df["1Y Return (%)"].fillna(df["3Y Return (%)"])
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

  weights = weights / np.sum(weights)
  opt_return = np.sum(returns * weights)
  opt_vol = np.sum(volatilities * weights)
  return weights, opt_return, opt_vol


def calculate_portfolio_metrics(df, weights, year_returns_cols):
  returns_1y = df["1Y Return (%)"].fillna(df["3Y Return (%)"])
  volatilities = df["Volatility (%)"]
  fees = df["Mgmt Fee (%)"]

  valid_return_mask = returns_1y.notna()
  valid_vol_mask = volatilities.notna()
  valid_fee_mask = fees.notna()

  portfolio_return = (
      np.sum(returns_1y[valid_return_mask] * weights[valid_return_mask])
      if valid_return_mask.any()
      else 0
  )
  portfolio_volatility = (
      np.sum(volatilities[valid_vol_mask] * weights[valid_vol_mask])
      if valid_vol_mask.any()
      else 0
  )
  portfolio_fee = (
      np.sum(fees[valid_fee_mask] * weights[valid_fee_mask])
      if valid_fee_mask.any()
      else 0
  )
  risk_adjusted = (
      portfolio_return / portfolio_volatility
      if portfolio_volatility > 0
      else 0
  )

  yearly_returns = []
  for year, col_name in year_returns_cols:
    if col_name in df.columns:
      year_data = df[col_name]
      valid_mask = year_data.notna()
      if valid_mask.any():
        year_return = np.sum(year_data[valid_mask] * weights[valid_mask])
        yearly_returns.append(year_return)

  best_year = max(yearly_returns) if yearly_returns else None
  worst_year = min(yearly_returns) if yearly_returns else None
  avg_yearly = np.mean(yearly_returns) if yearly_returns else None
  positive_years = sum(1 for r in yearly_returns if r > 0)
  consistency = (
      (positive_years / len(yearly_returns) * 100) if yearly_returns else 0
  )

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
# SECTION: LOGIN PAGE
# ============================================================
def show_login_page():
  st.markdown(
      "<h2 style='text-align: center;'>🔐 Login to Portfolio Analyzer</h2>",
      unsafe_allow_html=True,
  )
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    email = st.text_input(
        "Enter your email address:",
        placeholder="your.email@example.com",
        key="login_email",
    )
    if st.button(
        "Send OTP Code", use_container_width=True, key="send_otp_btn"
    ):
      if "@" in email and "." in email.split("@")[1]:
        otp = generate_otp()
        if send_otp_email(email, otp):
          st.session_state.otp_email = email
          st.session_state.otp_code = otp
          st.session_state.show_otp_input = True
          st.success(f"✅ OTP sent to {email}. Check your email!")
        else:
          st.error("Failed to send OTP.")
      else:
        st.error("Please enter a valid email address.")

    if st.session_state.get("show_otp_input", False):
      st.info("An OTP code has been sent to your email.")
      otp_input = st.text_input(
          "Enter 6-digit OTP:",
          placeholder="000000",
          key="otp_input",
          type="password",
      )
      if st.button("Verify OTP", use_container_width=True, key="verify_otp_btn"):
        if otp_input == st.session_state.otp_code:
          allowed, lim, acc, gen = check_access_allowed(email)
          if allowed:
            st.session_state.authenticated = True
            st.session_state.user_email = email
            increment_access(email)
            st.success("✅ Login successful!")
            st.rerun()
          else:
            st.error(f"❌ Limit reached ({lim}). Contact cktchew@gmail.com.")
        else:
          st.error("❌ Incorrect OTP.")


# ============================================================
# SECTION: MAIN APP LOGIC
# ============================================================
if not st.session_state.authenticated:
  show_login_page()
else:
  st.markdown(
      """
    <style>
    .main-header { text-align: center; color: #1f77b4; font-size: 2.5em; font-weight: bold; margin-bottom: 10px; }
    .sub-header { text-align: center; color: #666; font-size: 1.1em; margin-bottom: 20px; }
    </style>
    <div class="main-header">CHEW ADVISORY</div>
    <div class="sub-header">Portfolio Analysis Tool</div>
    """,
      unsafe_allow_html=True,
  )

  col1, col2 = st.columns([0.9, 0.1])
  with col2:
    if st.button("Logout", key="btn_logout"):
      for key in [
          "authenticated",
          "user_email",
          "user_name",
          "page",
          "portfolio_data",
          "show_otp_input",
          "otp_code",
          "otp_email",
          "funds_df",
      ]:
        st.session_state[key] = (
            False
            if key == "authenticated"
            else (
                None
                if key
                in ["user_email", "user_name", "otp_code", "otp_email", "funds_df"]
                else "home"
                if key == "page"
                else {}
            )
        )
      st.rerun()

  if st.session_state.page == "home":
    st.markdown("---")
    st.markdown("### Welcome to the Portfolio Analyzer")
    col1, col2, col3 = st.columns(3)
    with col1:
      if st.button(
          "📝 Start New Analysis", use_container_width=True, type="primary"
      ):
        st.session_state.page = "input"
        st.session_state.portfolio_data = {}
        st.session_state.funds_df = None
        st.session_state.user_name = None
        st.rerun()
    with col2:
      if st.session_state.user_email == ADMIN_EMAIL:
        st.info("⭐ Admin: Unlimited access")

  elif st.session_state.page == "input":
    st.header("📌 Step 1: Client Info, Goal & Fund Data")

    if st.session_state.portfolio_data:
      st.info("💡 You have existing analysis results from your previous run.")
      if st.button(
          "📂 Return to Previous Analysis",
          use_container_width=True,
          type="primary",
          key="btn_return_analysis",
      ):
        st.session_state.page = "analysis"
        st.rerun()
      st.markdown("---")

    st.subheader("Client Information & Parameters")
    client_name = st.text_input(
        "Client Full Name:",
        placeholder="Enter client name",
        value=st.session_state.get("user_name", "") or "",
    )
    st.session_state.user_name = client_name if client_name else None

    col1, col2, col3 = st.columns(3)
    with col1:
      goal_type = st.radio(
          "Goal Type:", ["Reach a Target Sum ($)", "Achieve Target Annual Growth (%)"]
      )
    with col2:
      years = st.number_input(
          "Time Horizon (Years)", min_value=1, max_value=50, value=10
      )
    with col3:
      risk_profile = st.selectbox(
          "Risk Profile:", ["Conservative", "Moderate", "Growth"]
      )

    if goal_type == "Reach a Target Sum ($)":
      target_value = st.number_input(
          "Target Final Sum ($)", min_value=1000, value=100000, step=1000
      )
      target_growth = None
    else:
      target_growth = st.number_input(
          "Target Annual Growth (%)", min_value=1.0, max_value=20.0, value=8.0, step=0.5
      )
      target_value = None

    col1, col2 = st.columns(2)
    with col1:
      initial_investment = st.number_input(
          "Initial Lump Sum Investment ($)", min_value=0, value=10000, step=1000
      )
    with col2:
      monthly_contribution = st.number_input(
          "Monthly Contribution ($)", min_value=0, value=500, step=100
      )

    st.markdown("---")
    st.subheader("Fund Data Source (Choose Excel or PDF FFS with Gemini)")
    tab_excel, tab_pdf = st.tabs(
        ["📊 Upload Excel File", "📄 Upload PDF Fund Fact Sheets (Gemini AI)"]
    )

    with tab_excel:
      st.info(
          "Upload an Excel file with your fund data using the standard template."
      )
      if st.button("📥 Download Excel Template", key="btn_dl_template"):
        template_df = pd.DataFrame(columns=[
            "Fund Name",
            "1Y Return (%)",
            "3Y Return (%)",
            "5Y Return (%)",
            "10Y Return (%)",
            "Volatility (%)",
            "Mgmt Fee (%)",
            "1Y Benchmark (%)",
            "Benchmark Name",
            "2016 Return (%)",
            "2017 Return (%)",
            "2018 Return (%)",
            "2019 Return (%)",
            "2020 Return (%)",
            "2021 Return (%)",
            "2022 Return (%)",
            "2023 Return (%)",
            "2024 Return (%)",
            "2025 Return (%)",
        ])
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
          template_df.to_excel(writer, sheet_name="Fund Data", index=False)
        st.download_button(
            label="Download Template Excel",
            data=buffer.getvalue(),
            file_name="fund_data_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

      uploaded_file = st.file_uploader(
          "Upload Excel File", type=["xlsx", "xls"], key="excel_upload"
      )
      if uploaded_file:
        try:
          df = pd.read_excel(uploaded_file)
          df.columns = df.columns.str.strip()
          st.session_state.funds_df = df
          st.success(f"✅ Loaded {len(df)} funds from Excel successfully!")
        except Exception as e:
          st.error(f"Error reading Excel file: {str(e)}")

    with tab_pdf:
      st.info(
          "Upload one or more Fund Fact Sheet (FFS) PDFs. Gemini AI will"
          " automatically extract returns, volatility, fees, and benchmarks."
      )
      uploaded_pdfs = st.file_uploader(
          "Upload PDF Fact Sheets",
          type=["pdf"],
          accept_multiple_files=True,
          key="pdf_upload",
      )
      if uploaded_pdfs and st.button(
          "🤖 Extract Data with Gemini AI", key="btn_gemini_extract"
      ):
        extracted_df = process_pdf_ffs_with_gemini(uploaded_pdfs)
        if extracted_df is not None:
          st.session_state.funds_df = extracted_df
          st.success(
              f"✅ Extracted data for {len(extracted_df)} funds successfully!"
          )

    if st.session_state.funds_df is not None:
      st.markdown("---")
      st.subheader("Review & Edit Extracted / Uploaded Fund Data")
      st.session_state.funds_df = st.data_editor(
          st.session_state.funds_df, num_rows="dynamic", key="fund_data_editor"
      )

      if st.button(
          "Calculate Portfolio Analysis 🚀",
          use_container_width=True,
          type="primary",
          key="btn_calc_analysis",
      ):
        df = st.session_state.funds_df
        numeric_cols = [
            "1Y Return (%)",
            "3Y Return (%)",
            "5Y Return (%)",
            "10Y Return (%)",
            "Volatility (%)",
            "Mgmt Fee (%)",
            "1Y Benchmark (%)",
        ]
        year_returns, year_benchmarks = detect_year_columns(df)
        for year, col in year_returns:
          numeric_cols.append(col)
        for year, col in year_benchmarks:
          numeric_cols.append(col)

        for col in numeric_cols:
          if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        st.session_state.portfolio_data = {
            "client_name": st.session_state.user_name,
            "goal_type": goal_type,
            "target_value": target_value,
            "target_growth": target_growth,
            "years": years,
            "initial_investment": initial_investment,
            "monthly_contribution": monthly_contribution,
            "risk_profile": risk_profile,
            "funds_df": df,
            "year_returns": year_returns,
            "year_benchmarks": year_benchmarks,
        }
        increment_generation(st.session_state.user_email)
        st.session_state.page = "analysis"
        st.rerun()

    if st.button("⬅️ Back to Home", key="btn_back_home_input"):
      st.session_state.page = "home"
      st.rerun()

  elif st.session_state.page == "analysis":
    st.header("📊 Step 2: Portfolio Analysis & Graphics")
    data = st.session_state.portfolio_data
    df = data["funds_df"].copy()
    year_returns_cols = data["year_returns"]
    year_benchmarks_cols = data["year_benchmarks"]

    if data["goal_type"] == "Reach a Target Sum ($)":
      target_return = calculate_required_cagr(
          data["target_value"],
          data["initial_investment"],
          data["monthly_contribution"],
          data["years"],
      )
      target_amount = data["target_value"]
    else:
      target_return = data["target_growth"]
      target_amount = calculate_future_value(
          data["initial_investment"],
          data["monthly_contribution"],
          data["years"],
          target_return,
      )

    n = len(df)
    equal_weights = np.ones(n) / n
    eq_metrics = calculate_portfolio_metrics(df, equal_weights, year_returns_cols)
    eq_amount = calculate_future_value(
        data["initial_investment"],
        data["monthly_contribution"],
        data["years"],
        eq_metrics["return"],
    )

    opt_weights, opt_return, opt_vol = optimize_portfolio_max_return(
        df, data["risk_profile"]
    )
    opt_metrics = calculate_portfolio_metrics(df, opt_weights, year_returns_cols)
    opt_amount = calculate_future_value(
        data["initial_investment"],
        data["monthly_contribution"],
        data["years"],
        opt_metrics["return"],
    )

    risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
    max_vol_threshold = risk_thresholds[data["risk_profile"]]

    st.subheader("Portfolio Performance Summary")
    col1, col2, col3 = st.columns(3)
    with col1:
      st.metric(
          "Equal-Weighted Return", f"{eq_metrics['return']:.2f}% p.a."
      )
    with col2:
      st.metric(
          "Optimized Return (Max)", f"{opt_metrics['return']:.2f}% p.a."
      )
    with col3:
      st.metric("Target Return", f"{target_return:.2f}% p.a.")

    st.subheader("Goal Feasibility & Risk Assessment")
    st.info(
        f"**Max Allowable Volatility for {data['risk_profile']} Profile:**"
        f" {max_vol_threshold}%"
    )

    col1, col2 = st.columns(2)
    with col1:
      st.metric(
          "Equal-Weighted Portfolio",
          (
              "✅ Achievable"
              if eq_amount >= target_amount
              else "⚠️ Shortfall"
          ),
          delta=f"${eq_amount - target_amount:,.0f}",
      )
      st.caption(
          f"Volatility: {eq_metrics['volatility']:.1f}% ("
          f"{'✅ Matches' if eq_metrics['volatility'] <= max_vol_threshold else '⚠️ Exceeds'}"
          f" {data['risk_profile']} profile)"
      )
    with col2:
      st.metric(
          "Optimized Portfolio",
          (
              "✅ Achievable"
              if opt_amount >= target_amount
              else "⚠️ Shortfall"
          ),
          delta=f"${opt_amount - target_amount:,.0f}",
      )
      st.caption(
          f"Volatility: {opt_metrics['volatility']:.1f}% ("
          f"{'✅ Matches' if opt_metrics['volatility'] <= max_vol_threshold else '⚠️ Exceeds'}"
          f" {data['risk_profile']} profile)"
      )


    # Recommendations helpers
    def find_monthly_for_target(target, initial, years, return_rate):
      r = return_rate / 100
      if r == 0:
        return max(0, (target - initial) / (years * 12))
      fv_initial = initial * ((1 + r) ** years)
      remaining = target - fv_initial
      if remaining <= 0:
        return 0
      return max(0, remaining * r / (12 * (((1 + r) ** years) - 1)))


    def find_initial_for_target(target, monthly, years, return_rate):
      r = return_rate / 100
      if r == 0:
        return max(0, target - (monthly * 12 * years))
      fv_monthly = monthly * 12 * (((1 + r) ** years) - 1) / r
      remaining = target - fv_monthly
      if remaining <= 0:
        return 0
      return max(0, remaining / ((1 + r) ** years))


    def find_years_for_target(target, initial, monthly, return_rate):
      for test_years in range(1, 100):
        if (
            calculate_future_value(initial, monthly, test_years, return_rate)
            >= target
        ):
          return test_years
      return 100


    req_monthly = find_monthly_for_target(
        target_amount,
        data["initial_investment"],
        data["years"],
        opt_metrics["return"],
    )
    req_initial_zero_monthly = find_initial_for_target(
        target_amount, 0, data["years"], opt_metrics["return"]
    )
    req_years = find_years_for_target(
        target_amount,
        data["initial_investment"],
        data["monthly_contribution"],
        opt_metrics["return"],
    )

    st.subheader("💡 Recommendations (Based on Optimized Portfolio)")
    diff_monthly = req_monthly - data["monthly_contribution"]
    if diff_monthly < -1:
      msg1 = (
          "You can **reduce** your monthly contribution by"
          f" **${abs(diff_monthly):,.0f}** (New total: ${req_monthly:,.0f}/month)"
          " and still meet your target."
      )
    elif diff_monthly > 1:
      msg1 = (
          "You need to **increase** your monthly contribution by"
          f" **${diff_monthly:,.0f}** (New total: ${req_monthly:,.0f}/month) to"
          " meet your target."
      )
    else:
      msg1 = "Your current monthly contribution is exactly on track."

    diff_initial = req_initial_zero_monthly - data["initial_investment"]
    if diff_initial < -1:
      msg2 = (
          "If monthly contribution is $0, you can **reduce** initial investment"
          f" by **${abs(diff_initial):,.0f}** (New total:"
          f" ${req_initial_zero_monthly:,.0f})."
      )
    elif diff_initial > 1:
      msg2 = (
          "If monthly contribution is $0, you need to **increase** initial"
          f" investment by **${diff_initial:,.0f}** (New total:"
          f" ${req_initial_zero_monthly:,.0f})."
      )
    else:
      msg2 = "If monthly contribution is $0, your initial investment is on track."

    diff_years = req_years - data["years"]
    if diff_years < 0:
      msg3 = (
          "You can **reduce** your time horizon by **"
          f"{abs(diff_years)} years** (New total: {req_years} years)."
      )
    elif diff_years > 0:
      msg3 = (
          "You need to **increase** your time horizon by **"
          f"{diff_years} years** (New total: {req_years} years)."
      )
    else:
      msg3 = "Your current time horizon is exactly on track."

    st.info(f"**Option 1 - Monthly Contribution:**\n{msg1}")
    st.info(f"**Option 2 - Initial Capital ($0 monthly):**\n{msg2}")
    st.info(f"**Option 3 - Time Horizon:**\n{msg3}")

    # --- CHARTS SECTION ---
    st.subheader("📊 Portfolio Allocation Charts")
    col1, col2 = st.columns(2)
    with col1:
      fig1, ax1 = plt.subplots(figsize=(7, 7))
      ax1.pie(
          equal_weights,
          labels=df["Fund Name"],
          autopct="%1.1f%%",
          startangle=90,
          colors=plt.cm.Paired.colors,
          pctdistance=0.85,
          labeldistance=1.1,
      )
      ax1.set_title(
          "Equal-Weighted Allocation", fontsize=12, fontweight="bold", pad=20
      )
      st.pyplot(fig1)

    with col2:
      fig2, ax2 = plt.subplots(figsize=(7, 7))
      ax2.pie(
          opt_weights,
          labels=df["Fund Name"],
          autopct="%1.1f%%",
          startangle=90,
          colors=plt.cm.Paired.colors,
          pctdistance=0.85,
          labeldistance=1.1,
      )
      ax2.set_title(
          "Optimized Allocation (Max Return)",
          fontsize=12,
          fontweight="bold",
          pad=20,
      )
      st.pyplot(fig2)

    st.subheader("📈 Portfolio Performance vs Benchmark")
    years_list = sorted(list(set([year for year, col in year_returns_cols])))
    fig3, ax3 = plt.subplots(figsize=(10, 5))
    has_data = False
    has_benchmark = False
    eq_port_returns = []
    opt_port_returns = []
    bench_returns = []
    benchmark_names = (
        df["Benchmark Name"].dropna().unique()
        if "Benchmark Name" in df.columns
        else []
    )
    primary_benchmark = (
        benchmark_names[0] if len(benchmark_names) > 0 else "Composite Benchmark"
    )

    for year in years_list:
      year_return_col = f"{year} Return (%)"
      year_bench_col = f"{year} Benchmark (%)"
      if year_return_col in df.columns:
        valid_mask = df[year_return_col].notna()
        if valid_mask.any():
          eq_port_returns.append(
              np.sum(df[year_return_col][valid_mask] * equal_weights[valid_mask])
          )
          opt_port_returns.append(
              np.sum(df[year_return_col][valid_mask] * opt_weights[valid_mask])
          )
          has_data = True
        else:
          eq_port_returns.append(np.nan)
          opt_port_returns.append(np.nan)
      else:
        eq_port_returns.append(np.nan)
        opt_port_returns.append(np.nan)

      if year_bench_col in df.columns:
        valid_mask = df[year_bench_col].notna()
        if valid_mask.any():
          bench_returns.append(
              np.sum(df[year_bench_col][valid_mask] * equal_weights[valid_mask])
          )
          has_benchmark = True
        else:
          bench_returns.append(np.nan)
      else:
        bench_returns.append(np.nan)

    if has_data:
      ax3.plot(
          years_list,
          eq_port_returns,
          marker="o",
          linewidth=3,
          label="Equal-Weighted",
          color="#3498db",
      )
      ax3.plot(
          years_list,
          opt_port_returns,
          marker="s",
          linewidth=3,
          label="Optimized",
          color="#2ecc71",
      )
    if has_benchmark:
      ax3.plot(
          years_list,
          bench_returns,
          marker="^",
          linewidth=3,
          label=primary_benchmark,
          color="#e74c3c",
      )

    if has_data or has_benchmark:
      ax3.set_xlabel("Year", fontweight="bold")
      ax3.set_ylabel("Return (%)", fontweight="bold")
      ax3.set_title("Portfolio Performance vs Benchmark", fontweight="bold")
      ax3.legend(loc="upper left")
      ax3.grid(True, alpha=0.3)
      st.pyplot(fig3)
    else:
      st.info("No calendar year return data found for chart plotting.")

    # PDF Download for Charts
    buf_charts = io.BytesIO()
    with PdfPages(buf_charts) as pdf:
      pdf.savefig(fig1, bbox_inches="tight")
      pdf.savefig(fig2, bbox_inches="tight")
      if has_data or has_benchmark:
        pdf.savefig(fig3, bbox_inches="tight")
    buf_charts.seek(0)
    st.download_button(
        label="📥 Download All Charts (Combined PDF)",
        data=buf_charts,
        file_name=f"Portfolio_Charts_{datetime.now().strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
      if st.button("← Back to Input", use_container_width=True):
        st.session_state.page = "input"
        st.rerun()
    with col2:
      if st.button(
          "Generate Report →",
          use_container_width=True,
          type="primary",
          key="btn_to_report",
      ):
        st.session_state.page = "report"
        st.rerun()

  elif st.session_state.page == "report":
    st.header("📄 Step 3: Word Report Generation")
    st.success("✅ Report is ready for download.")

    data = st.session_state.portfolio_data
    df = data["funds_df"]
    year_returns_cols = data["year_returns"]

    n = len(df)
    equal_weights = np.ones(n) / n
    if data["goal_type"] == "Reach a Target Sum ($)":
      target_return = calculate_required_cagr(
          data["target_value"],
          data["initial_investment"],
          data["monthly_contribution"],
          data["years"],
      )
      target_amount = data["target_value"]
    else:
      target_return = data["target_growth"]
      target_amount = calculate_future_value(
          data["initial_investment"],
          data["monthly_contribution"],
          data["years"],
          target_return,
      )

    opt_weights, opt_return, opt_vol = optimize_portfolio_max_return(
        df, data["risk_profile"]
    )
    eq_metrics = calculate_portfolio_metrics(df, equal_weights, year_returns_cols)
    opt_metrics = calculate_portfolio_metrics(df, opt_weights, year_returns_cols)
    eq_amount = calculate_future_value(
        data["initial_investment"],
        data["monthly_contribution"],
        data["years"],
        eq_metrics["return"],
    )
    opt_amount = calculate_future_value(
        data["initial_investment"],
        data["monthly_contribution"],
        data["years"],
        opt_metrics["return"],
    )

    doc = Document()
    title = doc.add_heading("Portfolio Analysis Report", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    if data.get("client_name"):
      doc.add_paragraph(f"Client Name: {data['client_name']}")
    doc.add_paragraph(f"Prepared For: {st.session_state.user_email}")
    doc.add_paragraph()

    doc.add_heading("Executive Summary", level=1)
    doc.add_paragraph(f"Goal Type: {data['goal_type']}")
    if data["goal_type"] == "Reach a Target Sum ($)":
      doc.add_paragraph(f"Target Amount: ${data['target_value']:,.0f}")
    else:
      doc.add_paragraph(f"Target Annual Growth: {data['target_growth']:.1f}%")
    doc.add_paragraph(f"Time Horizon: {data['years']} years")
    doc.add_paragraph(f"Risk Profile: {data['risk_profile']}")
    doc.add_paragraph(f"Initial Investment: ${data['initial_investment']:,.0f}")
    doc.add_paragraph(f"Monthly Contribution: ${data['monthly_contribution']:,.0f}")
    doc.add_paragraph()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    st.download_button(
        label="📄 Download Word Report (DOCX)",
        data=buffer,
        file_name=f"Portfolio_Analysis_{datetime.now().strftime('%Y%m%d')}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

    if st.button("← Back to Analysis", use_container_width=True):
      st.session_state.page = "analysis"
      st.rerun()


if __name__ == "__main__":
  main()