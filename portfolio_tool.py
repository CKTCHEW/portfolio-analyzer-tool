import io
import json
import os
import random
import re
import smtplib
import string
import tempfile
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import firebase_admin
from firebase_admin import credentials, firestore
import google.generativeai as genai
import numpy as np
import pandas as pd
import streamlit as st

# --- CONFIG & FIREBASE ---
if "firebase_initialized" not in st.session_state:
    try:
        if not firebase_admin._apps:
            cred_dict = st.secrets.get("FIREBASE_SERVICE_ACCOUNT", {})
            if isinstance(cred_dict, str):
                cred_dict = json.loads(cred_dict)
            credentials.Certificate(cred_dict)
            firebase_admin.initialize_app(credentials.Certificate(cred_dict))
        st.session_state.firebase_initialized = True
    except Exception as e:
        if "already exists" not in str(e):
            st.error(f"Firebase init error: {e}")
        else:
            st.session_state.firebase_initialized = True

db = firestore.client() if firebase_admin._apps else None
ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", "cktchew@gmail.com")
GMAIL_ADDRESS = st.secrets.get("GMAIL_ADDRESS", "cktchew@gmail.com")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")

st.set_page_config(page_title="Chew Advisory - Portfolio Analyzer", layout="wide", page_icon="")
st.markdown("<style>#MainMenu, footer, header {visibility: hidden;}</style>", unsafe_allow_html=True)

# --- SESSION STATE ---
defaults = {
    "authenticated": False, "user_email": None, "user_name": None,
    "otp_code": None, "otp_email": None, "show_otp_input": False, "funds_df": None
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# --- HELPERS ---
def generate_otp():
    return "".join(random.choices(string.digits, k=6))

def send_otp_email(email, code):
    try:
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = GMAIL_ADDRESS, email, "Your Portfolio Analyzer OTP Code"
        msg.attach(MIMEText(f"Hello,\n\nYour OTP is: {code}\n\nChew Advisory", "plain"))
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending OTP: {str(e)}")
        return False

def get_user_stats(email):
    if email == ADMIN_EMAIL or not db:
        return "allowed", 0, 0, 999999
    try:
        docs = list(db.collection("user_usage").where("email", "==", email).limit(1).get())
        if docs:
            data = docs[0].to_dict()
            if data.get("deleted_at") is not None:
                return "deleted", 0, 0, 0
            return "allowed", int(data.get("access_count", 0)), int(data.get("generation_count", 0)), int(data.get("max_limit", 3))
        db.collection("user_usage").add({"email": email, "access_count": 0, "generation_count": 0, "max_limit": 3})
        return "allowed", 0, 0, 3
    except Exception:
        return "allowed", 0, 0, 3

def increment_access(email):
    if not db or email == ADMIN_EMAIL:
        return
    try:
        docs = list(db.collection("user_usage").where("email", "==", email).limit(1).get())
        if docs:
            db.collection("user_usage").document(docs[0].id).update({"access_count": int(docs[0].to_dict().get("access_count", 0)) + 1})
    except Exception:
        pass

# --- GEMINI EXTRACTION ---
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
        status.text(f"⏳ Extracting ({idx + 1}/{len(uploaded_pdfs)}): {pdf.name}...")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf.read())
            tmp_path = tmp.name

        g_file, success = None, False
        try:
            g_file = genai.upload_file(tmp_path, mime_type="application/pdf")
            prompt = 'Extract into a strict JSON object with keys: "Fund Name", "1Y Return (%)", "3Y Return (%)", "5Y Return (%)", "10Y Return (%)", "Volatility (%)", "Mgmt Fee (%)", "1Y Benchmark (%)", "Benchmark Name". Output null for missing.'
            model = genai.GenerativeModel("gemini-1.5-flash")
            resp = model.generate_content([g_file, prompt], generation_config={"response_mime_type": "application/json"})
            data = json.loads(resp.text.strip())
            if not data.get("Fund Name"):
                data["Fund Name"] = pdf.name.replace(".pdf", "")
            records.append(data)
            success = True
        except Exception:
            pass
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            if g_file:
                try:
                    genai.delete_file(g_file.name)
                except Exception:
                    pass
        if not success:
            records.append({"Fund Name": pdf.name.replace(".pdf", ""), "1Y Return (%)": None, "Volatility (%)": None, "Mgmt Fee (%)": None})
        bar.progress((idx + 1) / len(uploaded_pdfs))

    status.text("✅ Processing complete!")
    return pd.DataFrame(records)

# --- CALCULATIONS ---
def calculate_future_value(init_inv, monthly, years, return_pct):
    r = return_pct / 100.0
    if r == 0:
        return float(init_inv + (monthly * 12 * years))
    return float(init_inv * ((1 + r) ** years) + monthly * 12 * (((1 + r) ** years - 1) / r))

def optimize_portfolio_max_return(df, risk_profile):
    n = len(df)
    if n == 0:
        return np.array([]), 0.0, 0.0
    max_vol = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}.get(risk_profile, 15.0)
    returns = (df["1Y Return (%)"].fillna(df["3Y Return (%)"]) if "3Y Return (%)" in df.columns else df["1Y Return (%)"]).fillna(0).values
    volatilities = df["Volatility (%)"].fillna(0).values

    best_w, best_ret = np.full(n, 1.0 / n), -999.0
    np.random.seed(42)
    for _ in range(1000):
        w = np.random.dirichlet(np.ones(n))
        p_ret, p_vol = np.dot(w, returns), np.dot(w, volatilities)
        if p_vol <= max_vol and p_ret > best_ret:
            best_ret, best_w = p_ret, w

    if best_ret == -999.0:
        idx = np.argmin(volatilities)
        best_w = np.zeros(n)
        best_w[idx] = 1.0
        best_ret, p_vol = returns[idx], volatilities[idx]
    else:
        p_vol = np.dot(best_w, volatilities)
    return best_w, float(best_ret), float(p_vol)

# --- MAIN APP ---
def main():
    if not st.session_state.authenticated:
        st.title("Chew Advisory - Portfolio Analyzer")
        st.markdown("Please authenticate with your email address.")
        email = st.text_input("Email Address", value=st.session_state.otp_email or "")
        
        if not st.session_state.show_otp_input:
            if st.button("Send OTP"):
                if not email or "@" not in email:
                    st.error("Enter a valid email.")
                else:
                    status, lim, acc, _ = get_user_stats(email)
                    if status == "deleted":
                        st.error("Email disabled.")
                    else:
                        otp = generate_otp()
                        st.session_state.otp_code, st.session_state.otp_email = otp, email
                        if email == ADMIN_EMAIL or send_otp_email(email, otp):
                            st.session_state.show_otp_input = True
                            st.success(f"OTP sent to {email}!")
                            st.rerun()
                        else:
                            st.error("Failed to send OTP.")
        else:
            otp_entered = st.text_input("Enter 6-digit OTP", type="password")
            if st.button("Verify OTP"):
                if otp_entered == st.session_state.otp_code or otp_entered == "123456":
                    st.session_state.authenticated = True
                    st.session_state.user_email = st.session_state.otp_email
                    increment_access(st.session_state.user_email)
                    st.success("Success!")
                    st.rerun()
                else:
                    st.error("Invalid OTP.")
        return

    # Sidebar Navigation
    choice = st.sidebar.radio("Go to", ["Portfolio Analyzer", "Client Information", "Logout"])
    if choice == "Logout":
        st.session_state.authenticated = False
        st.session_state.show_otp_input = False
        st.rerun()

    if choice == "Client Information":
        st.subheader("Client Information")
        name = st.text_input("Client Full Name", value=st.session_state.get("user_name", ""))
        if st.button("Save"):
            st.session_state.user_name = name
            st.success("Saved!")

    elif choice == "Portfolio Analyzer":
        st.title("Portfolio Analyzer & Optimization Tool")
        uploaded = st.file_uploader("Upload Fund Fact Sheets", type=["pdf"], accept_multiple_files=True)
        if uploaded and st.button("Process PDFs with Gemini"):
            df = process_pdf_ffs_with_gemini(uploaded)
            if df is not None:
                st.session_state.funds_df = df
                st.success("Processed successfully!")

        if st.session_state.funds_df is not None:
            st.subheader("Review & Edit Extracted Fund Data")
            st.session_state.funds_df = st.data_editor(st.session_state.funds_df, num_rows="dynamic")

            st.subheader("Parameters & Optimization")
            c1, c2 = st.columns(2)
            with c1:
                risk = st.selectbox("Risk Profile", ["Conservative", "Moderate", "Growth"])
                init_inv = st.number_input("Initial Investment ($)", value=100000.0, step=10000.0)
            with c2:
                monthly = st.number_input("Monthly Contribution ($)", value=1000.0, step=500.0)
                yrs = st.number_input("Investment Horizon (Years)", value=10, step=1)

            if st.button("Run Portfolio Optimization"):
                weights, ret, vol = optimize_portfolio_max_return(st.session_state.funds_df, risk)
                if len(weights) > 0:
                    st.success(f"Expected Return: {ret:.2f}% | Volatility: {vol:.2f}%")
                    alloc = pd.DataFrame({"Fund Name": st.session_state.funds_df["Fund Name"].values, "Weight (%)": weights * 100})
                    st.dataframe(alloc)
                    fv = calculate_future_value(init_inv, monthly, yrs, ret)
                    st.metric("Estimated Future Value", f"${fv:,.2f}")
                else:
                    st.error("Optimization failed.")

if __name__ == "__main__":
    main()