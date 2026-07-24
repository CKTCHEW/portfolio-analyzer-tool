import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import json
import firebase_admin
from firebase_admin import credentials, firestore
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import random
import string
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import re

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
except:
    db = None

ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", "cktchew@gmail.com")
GMAIL_ADDRESS = st.secrets.get("GMAIL_ADDRESS", "cktchew@gmail.com")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")

# ============================================================
# SECTION: STREAMLIT PAGE CONFIG
# ============================================================
st.set_page_config(page_title="Chew Advisory - Portfolio Analyzer", layout="wide", page_icon="📊")
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
for key in ['authenticated', 'user_email', 'user_name', 'page', 'portfolio_data', 'otp_code', 'otp_email', 'show_otp_input', 'funds_df']:
    if key not in st.session_state:
        st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'user_name', 'otp_code', 'otp_email'] else 'home' if key == 'page' else None if key == 'funds_df' else {})

# ============================================================
# SECTION: HELPER FUNCTIONS
# ============================================================
def safe_float(v):
    try:
        return float(str(v).replace(',', '').strip())
    except:
        return np.nan

def generate_otp():
    return ''.join(random.choices(string.digits, k=6))

def detect_year_columns(df):
    """Detect year columns dynamically using pattern matching"""
    year_returns = []
    year_benchmarks = []
    
    for col in df.columns:
        if re.match(r'^\d{4}\s+Return\s+\(%\)$', col):
            year = int(col.split()[0])
            year_returns.append((year, col))
        elif re.match(r'^\d{4}\s+Benchmark\s+\(%\)$', col):
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
    if email == ADMIN_EMAIL: return "allowed", 0, 0, 999999
    if not db: return "allowed", 0, 0, 3 
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            data = docs[0].to_dict()
            if data.get('deleted_at') is not None: return "deleted", 0, 0, 0
            return "allowed", int(data.get('access_count', 0)), int(data.get('generation_count', 0)), int(data.get('max_limit', 3))
        else:
            db.collection('user_usage').add({'email': email, 'access_count': 0, 'generation_count': 0, 'max_limit': 3, 'created_at': firestore.SERVER_TIMESTAMP})
            return "allowed", 0, 0, 3
    except: return "allowed", 0, 0, 3

def check_access_allowed(email):
    status, acc, gen, lim = get_user_stats(email)
    if status != "allowed" or acc >= lim or gen >= lim: return False, lim, acc, gen
    return True, lim, acc, gen

def increment_access(email):
    if not db or email == ADMIN_EMAIL: return
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            doc_id = docs[0].id
            current = int(docs[0].to_dict().get('access_count', 0))
            db.collection('user_usage').document(doc_id).update({'access_count': current + 1, 'last_accessed_at': firestore.SERVER_TIMESTAMP})
    except: pass

def increment_generation(email):
    if not db or email == ADMIN_EMAIL: return
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            doc_id = docs[0].id
            current = int(docs[0].to_dict().get('generation_count', 0))
            db.collection('user_usage').document(doc_id).update({'generation_count': current + 1})
    except: pass

# ============================================================
# SECTION: PORTFOLIO CALCULATIONS & OPTIMIZATION
# ============================================================
def calculate_required_cagr(target_sum, initial_investment, monthly_contribution, years):
    if years <= 0 or target_sum <= 0: return 0.0
    r = 0.05 
    for _ in range(50): 
        fv_guess = initial_investment * ((1 + r) ** years) + (monthly_contribution * 12 * (((1 + r) ** years - 1) / r) if r > 0 else monthly_contribution * 12 * years)
        derivative = initial_investment * years * ((1 + r) ** (years - 1))
        if r > 0: derivative += monthly_contribution * 12 * (years * (1 + r)**(years - 1) * r - ((1 + r)**years - 1)) / (r**2)
        if abs(derivative) < 1e-8: break
        r = r - (fv_guess - target_sum) / derivative
        if r < -0.5: r = -0.5 
    return max(0.0, r * 100)

def calculate_future_value(initial_investment, monthly_contribution, years, annual_return_pct):
    r = annual_return_pct / 100
    if r == 0:
        fv = initial_investment + (monthly_contribution * 12 * years)
    else:
        fv = initial_investment * ((1 + r) ** years) + (monthly_contribution * 12 * (((1 + r) ** years - 1) / r))
    return fv

def optimize_portfolio_risk_constrained(df, target_return, risk_profile, target_exceeded=False):
    n = len(df)
    if n == 0: return np.array([]), 0, 0
    
    risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
    max_volatility = risk_thresholds.get(risk_profile, 15.0)
    
    returns_col = df['1Y Return (%)'].fillna(df['3Y Return (%)'])
    valid_mask = returns_col.notna()
    
    if not valid_mask.any():
        return np.ones(n) / n, 0, 0
    
    returns = returns_col.fillna(0).values
    volatilities = df['Volatility (%)'].fillna(np.nan).values
    
    if target_exceeded:
        # MINIMIZE volatility while maintaining at least target return
        sorted_indices = np.argsort([volatilities[i] if not np.isnan(volatilities[i]) else 999 for i in range(n)])
        
        opt_weights = np.zeros(n)
        remaining = 1.0
        
        for idx in sorted_indices:
            if remaining <= 0: break
            
            funds_left = n - np.sum(opt_weights > 0)
            max_w = min(0.40, remaining - (0.05 * max(0, funds_left - 1)))
            w = max(0.05, max_w)
            
            test_weights = opt_weights.copy()
            test_weights[idx] = w
            test_return = np.sum(returns * test_weights)
            test_vol = np.nansum(np.array([volatilities[i] * test_weights[i] if not np.isnan(volatilities[i]) else 0 for i in range(n)]))
            
            if test_return >= target_return and test_vol <= max_volatility:
                opt_weights[idx] = w
                remaining -= w
            else:
                for test_w in np.linspace(0.05, max_w, 10):
                    test_weights = opt_weights.copy()
                    test_weights[idx] = test_w
                    test_return = np.sum(returns * test_weights)
                    test_vol = np.nansum(np.array([volatilities[i] * test_weights[i] if not np.isnan(volatilities[i]) else 0 for i in range(n)]))
                    
                    if test_return >= target_return and test_vol <= max_volatility:
                        opt_weights[idx] = test_w
                        remaining -= test_w
                        break
        
        if opt_weights.sum() > 0:
            opt_weights = opt_weights / opt_weights.sum()
        
        opt_return = np.sum(returns * opt_weights)
        opt_vol = np.nansum(np.array([volatilities[i] * opt_weights[i] if not np.isnan(volatilities[i]) else 0 for i in range(n)]))
        
    else:
        # MAXIMIZE return while staying within risk profile
        sorted_indices = np.argsort(returns)[::-1]
        
        opt_weights = np.zeros(n)
        remaining = 1.0
        
        for idx in sorted_indices:
            if remaining <= 0: break
            
            fund_vol = volatilities[idx] if not np.isnan(volatilities[idx]) else 0
            
            funds_left = n - np.sum(opt_weights > 0)
            max_w = min(0.40, remaining - (0.05 * max(0, funds_left - 1)))
            w = max(0.05, max_w)
            
            test_weights = opt_weights.copy()
            test_weights[idx] = w
            test_vol = np.nansum(np.array([volatilities[i] * test_weights[i] if not np.isnan(volatilities[i]) else 0 for i in range(n)]))
            
            if test_vol <= max_volatility:
                opt_weights[idx] = w
                remaining -= w
            else:
                if fund_vol > 0:
                    current_vol = np.nansum(np.array([volatilities[i] * opt_weights[i] if not np.isnan(volatilities[i]) else 0 for i in range(n)]))
                    max_weight_for_risk = (max_volatility - current_vol) / fund_vol
                    w = max(0.05, min(max_weight_for_risk, remaining))
                    opt_weights[idx] = w
                    remaining -= w
        
        if opt_weights.sum() > 0:
            opt_weights = opt_weights / opt_weights.sum()
        
        opt_return = np.sum(returns * opt_weights)
        opt_vol = np.nansum(np.array([volatilities[i] * opt_weights[i] if not np.isnan(volatilities[i]) else 0 for i in range(n)]))
    
    return opt_weights, opt_return, opt_vol

def calculate_portfolio_metrics(df, weights, year_returns_cols):
    returns_1y = df['1Y Return (%)'].fillna(df['3Y Return (%)'])
    volatilities = df['Volatility (%)']
    fees = df['Mgmt Fee (%)']
    
    valid_return_mask = returns_1y.notna()
    valid_vol_mask = volatilities.notna()
    valid_fee_mask = fees.notna()
    
    portfolio_return = np.sum(returns_1y[valid_return_mask] * weights[valid_return_mask]) if valid_return_mask.any() else 0
    portfolio_volatility = np.sum(volatilities[valid_vol_mask] * weights[valid_vol_mask]) if valid_vol_mask.any() else 0
    portfolio_fee = np.sum(fees[valid_fee_mask] * weights[valid_fee_mask]) if valid_fee_mask.any() else 0
    
    risk_adjusted = portfolio_return / portfolio_volatility if portfolio_volatility > 0 else 0
    
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
    consistency = (positive_years / len(yearly_returns) * 100) if yearly_returns else 0
    
    return {
        'return': portfolio_return,
        'volatility': portfolio_volatility,
        'fee': portfolio_fee,
        'risk_adjusted': risk_adjusted,
        'best_year': best_year,
        'worst_year': worst_year,
        'avg_yearly': avg_yearly,
        'consistency': consistency,
        'yearly_returns': yearly_returns
    }

# ============================================================
# SECTION: LOGIN PAGE
# ============================================================
def show_login_page():
    st.markdown("<h2 style='text-align: center;'>🔐 Login to Portfolio Analyzer</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        email = st.text_input("Enter your email address:", placeholder="your.email@example.com", key="login_email")
        if st.button("Send OTP Code", use_container_width=True, key="send_otp_btn"):
            if "@" in email and "." in email.split("@")[1]:
                otp = generate_otp()
                if send_otp_email(email, otp):
                    st.session_state.otp_email = email
                    st.session_state.otp_code = otp
                    st.session_state.show_otp_input = True
                    st.success(f"✅ OTP sent to {email}. Check your email!")
                else: st.error("Failed to send OTP.")
            else: st.error("Please enter a valid email address.")
        
        if st.session_state.get('show_otp_input', False):
            st.info("An OTP code has been sent to your email.")
            otp_input = st.text_input("Enter 6-digit OTP:", placeholder="000000", key="otp_input", type="password")
            if st.button("Verify OTP", use_container_width=True, key="verify_otp_btn"):
                if otp_input == st.session_state.otp_code:
                    allowed, lim, acc, gen = check_access_allowed(email)
                    if allowed:
                        st.session_state.authenticated = True
                        st.session_state.user_email = email
                        increment_access(email)
                        st.success("✅ Login successful!")
                        st.rerun()
                    else: st.error(f"❌ Limit reached ({lim}). Contact cktchew@gmail.com.")
                else: st.error("❌ Incorrect OTP.")

# ============================================================
# SECTION: MAIN APP LOGIC
# ============================================================
if not st.session_state.authenticated:
    show_login_page()
else:
    st.markdown("""
    <style>
    .main-header { text-align: center; color: #1f77b4; font-size: 2.5em; font-weight: bold; margin-bottom: 10px; }
    .sub-header { text-align: center; color: #666; font-size: 1.1em; margin-bottom: 20px; }
    </style>
    <div class="main-header">CHEW ADVISORY</div>
    <div class="sub-header">Portfolio Analysis Tool</div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([0.9, 0.1])
    with col2:
        if st.button("Logout"):
            for key in ['authenticated', 'user_email', 'user_name', 'page', 'portfolio_data', 'show_otp_input', 'otp_code', 'otp_email', 'funds_df']:
                st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'user_name', 'otp_code', 'otp_email'] else 'home' if key == 'page' else None if key == 'funds_df' else {})
            st.rerun()

    if st.session_state.page == 'home':
        st.markdown("---")
        st.markdown("### Welcome to the Portfolio Analyzer")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("📝 Start New Analysis", use_container_width=True, type="primary"):
                st.session_state.page = 'input'
                st.session_state.portfolio_data = {}
                st.session_state.funds_df = None
                st.session_state.user_name = None
                st.rerun()
        with col2:
            if st.session_state.user_email == ADMIN_EMAIL: st.info("🔐 Admin: Unlimited access")

    elif st.session_state.page == 'input':
        st.header("📝 Step 1: Define Goal & Upload Fund Data")
        
        st.subheader("Client Information")
        client_name = st.text_input("Client Name (Optional):", placeholder="Enter your name", value=st.session_state.get('user_name', '') or '')
        st.session_state.user_name = client_name if client_name else None
        
        st.subheader("1. Investment Goal & Risk Profile")
        col1, col2, col3 = st.columns(3)
        with col1: goal_type = st.radio("Goal Type:", ["Reach a Target Sum ($)", "Achieve Target Annual Growth (%)"])
        with col2: years = st.number_input("Time Horizon (Years)", min_value=1, max_value=50, value=10)
        with col3: risk_profile = st.selectbox("Risk Profile:", ["Conservative", "Moderate", "Growth"])
            
        if goal_type == "Reach a Target Sum ($)":
            target_value = st.number_input("Target Final Sum ($)", min_value=1000, value=100000, step=1000)
            target_growth = None
        else:
            target_growth = st.number_input("Target Annual Growth (%)", min_value=1.0, max_value=20.0, value=8.0, step=0.5)
            target_value = None

        st.subheader("2. Capital & Contributions")
        col1, col2 = st.columns(2)
        with col1: initial_investment = st.number_input("Initial Lump Sum Investment ($)", min_value=0, value=10000, step=1000)
        with col2: monthly_contribution = st.number_input("Monthly Contribution ($)", min_value=0, value=500, step=100)

        st.subheader("3. Upload Fund Data (Excel)")
        st.info("Upload an Excel file with your fund data. Use the template below as a guide.")
        
        if st.button("📥 Download Excel Template"):
            template_df = pd.DataFrame(columns=[
                'Fund Name', '1Y Return (%)', '3Y Return (%)', '5Y Return (%)', '10Y Return (%)',
                'Volatility (%)', 'Mgmt Fee (%)', '1Y Benchmark (%)', 'Benchmark Name',
                '2016 Return (%)', '2017 Return (%)', '2018 Return (%)', '2019 Return (%)', '2020 Return (%)',
                '2021 Return (%)', '2022 Return (%)', '2023 Return (%)', '2024 Return (%)', '2025 Return (%)',
                '2016 Benchmark (%)', '2017 Benchmark (%)', '2018 Benchmark (%)', '2019 Benchmark (%)', '2020 Benchmark (%)',
                '2021 Benchmark (%)', '2022 Benchmark (%)', '2023 Benchmark (%)', '2024 Benchmark (%)', '2025 Benchmark (%)'
            ])
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                template_df.to_excel(writer, sheet_name='Fund Data', index=False)
            
            st.download_button(
                label="Download Template Excel",
                data=buffer.getvalue(),
                file_name="fund_data_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                required_cols = ['Fund Name', '1Y Return (%)', 'Volatility (%)', 'Mgmt Fee (%)']
                missing_cols = [col for col in required_cols if col not in df.columns]
                
                if missing_cols:
                    st.error(f"Missing required columns: {', '.join(missing_cols)}")
                else:
                    df.columns = df.columns.str.strip()
                    numeric_cols = ['1Y Return (%)', '3Y Return (%)', '5Y Return (%)', '10Y Return (%)', 
                                    'Volatility (%)', 'Mgmt Fee (%)', '1Y Benchmark (%)']
                    
                    year_returns, year_benchmarks = detect_year_columns(df)
                    for year, col in year_returns: numeric_cols.append(col)
                    for year, col in year_benchmarks: numeric_cols.append(col)
                    
                    for col in numeric_cols:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                    
                    st.session_state.funds_df = df
                    st.success(f"✅ Loaded {len(df)} funds successfully!")
                    st.dataframe(df, use_container_width=True)
                    
                    st.markdown("---")
                    if st.button("Calculate Portfolio Analysis", use_container_width=True, type="primary"):
                        st.session_state.portfolio_data = {
                            'client_name': st.session_state.user_name,
                            'goal_type': goal_type, 'target_value': target_value, 'target_growth': target_growth,
                            'years': years, 'initial_investment': initial_investment, 'monthly_contribution': monthly_contribution,
                            'risk_profile': risk_profile, 'funds_df': df,
                            'year_returns': year_returns, 'year_benchmarks': year_benchmarks
                        }
                        increment_generation(st.session_state.user_email)
                        st.session_state.page = 'analysis'
                        st.rerun()
            except Exception as e:
                st.error(f"Error reading Excel file: {str(e)}")
        else:
            st.warning("Please upload an Excel file to proceed.")
        
        if st.button("⬅️ Back to Home"):
            st.session_state.page = 'home'
            st.rerun()

    elif st.session_state.page == 'analysis':
        st.header("📊 Step 2: Portfolio Analysis")
        data = st.session_state.portfolio_data
        df = data['funds_df'].copy()
        year_returns_cols = data['year_returns']
        year_benchmarks_cols = data['year_benchmarks']
        
        if data['goal_type'] == "Reach a Target Sum ($)":
            target_return = calculate_required_cagr(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'])
            target_amount = data['target_value']
        else:
            target_return = data['target_growth']
            target_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], target_return)

        n = len(df)
        equal_weights = np.ones(n) / n
        eq_metrics = calculate_portfolio_metrics(df, equal_weights, year_returns_cols)
        eq_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], eq_metrics['return'])
        
        target_exceeded = eq_amount >= target_amount * 1.1
        opt_weights, opt_return, opt_vol = optimize_portfolio_risk_constrained(df, target_return, data['risk_profile'], target_exceeded)
        opt_metrics = calculate_portfolio_metrics(df, opt_weights, year_returns_cols)
        opt_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], opt_metrics['return'])

        st.subheader("Portfolio Performance Analysis")
        col1, col2, col3 = st.columns(3)
        with col1: st.metric("Equal-Weighted Return", f"{eq_metrics['return']:.1f}% p.a.")
        with col2: st.metric("Optimized Return", f"{opt_metrics['return']:.1f}% p.a.")
        with col3: st.metric("Target Return", f"{target_return:.1f}% p.a.")

        st.subheader("Goal Feasibility & Risk Assessment")
        risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
        risk_thresh = risk_thresholds[data['risk_profile']]
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Equal-Weighted Portfolio", "✅ Achievable" if eq_amount >= target_amount else "⚠️ Shortfall", 
                      delta=f"${eq_amount - target_amount:,.0f}")
            st.caption(f"Volatility: {eq_metrics['volatility']:.1f}% ({'✅ Matches' if eq_metrics['volatility'] <= risk_thresh else '⚠️ Exceeds'} {data['risk_profile']} profile)")
        with col2:
            st.metric("Optimized Portfolio", "✅ Achievable" if opt_amount >= target_amount else "⚠️ Shortfall", 
                      delta=f"${opt_amount - target_amount:,.0f}")
            st.caption(f"Volatility: {opt_metrics['volatility']:.1f}% ({'✅ Matches' if opt_metrics['volatility'] <= risk_thresh else '⚠️ Exceeds'} {data['risk_profile']} profile)")

        st.subheader(" Recommendations")
        
        if opt_amount > target_amount * 1.1:
            st.success(f"✅ **Surplus Scenario:** Your optimized portfolio can achieve ${opt_amount:,.0f}, exceeding your target of ${target_amount:,.0f}.")
            
            def find_monthly_for_target(target, initial, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, (target - initial) / (years * 12))
                fv_initial = initial * ((1 + r) ** years)
                remaining = target - fv_initial
                if remaining <= 0: return 0
                return max(0, remaining * r / (12 * (((1 + r) ** years) - 1)))
            
            def find_initial_for_target(target, monthly, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, target - (monthly * 12 * years))
                fv_monthly = monthly * 12 * (((1 + r) ** years) - 1) / r
                remaining = target - fv_monthly
                if remaining <= 0: return 0
                return max(0, remaining / ((1 + r) ** years))
            
            def find_years_for_target(target, initial, monthly, return_rate):
                for test_years in range(1, 50):
                    if calculate_future_value(initial, monthly, test_years, return_rate) >= target:
                        return test_years
                return 50
            
            reduced_monthly = find_monthly_for_target(target_amount, data['initial_investment'], data['years'], opt_metrics['return'])
            st.info(f"**Option 1 - Reduce Monthly Contribution:**\nYou could reduce your monthly contribution from ${data['monthly_contribution']:,.0f} to ${reduced_monthly:,.0f} (save ${data['monthly_contribution'] - reduced_monthly:,.0f}/month)")
            
            reduced_initial = find_initial_for_target(target_amount, 0, data['years'], opt_metrics['return'])
            st.info(f"**Option 2 - Reduce Initial Investment:**\nIf you contribute $0/month, you could reduce initial investment from ${data['initial_investment']:,.0f} to ${reduced_initial:,.0f}")
            
            shortened_years = find_years_for_target(target_amount, data['initial_investment'], data['monthly_contribution'], opt_metrics['return'])
            st.info(f"**Option 3 - Shorten Time Horizon:**\nYou could achieve your target in {shortened_years} years instead of {data['years']} years")
                
        elif opt_amount < target_amount:
            st.error(f"⚠️ **Shortfall Scenario:** Your optimized portfolio can achieve ${opt_amount:,.0f}, which is ${target_amount - opt_amount:,.0f} below your target.")
            
            def find_years_for_target(target, initial, monthly, return_rate):
                for test_years in range(1, 100):
                    if calculate_future_value(initial, monthly, test_years, return_rate) >= target:
                        return test_years
                return 100
            
            def find_monthly_for_target(target, initial, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, (target - initial) / (years * 12))
                fv_initial = initial * ((1 + r) ** years)
                remaining = target - fv_initial
                if remaining <= 0: return 0
                return max(0, remaining * r / (12 * (((1 + r) ** years) - 1)))
            
            def find_initial_for_target(target, monthly, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, target - (monthly * 12 * years))
                fv_monthly = monthly * 12 * (((1 + r) ** years) - 1) / r
                remaining = target - fv_monthly
                if remaining <= 0: return 0
                return max(0, remaining / ((1 + r) ** years))
            
            increased_years = find_years_for_target(target_amount, data['initial_investment'], data['monthly_contribution'], opt_metrics['return'])
            st.info(f"**Option 1 - Extend Time Horizon:**\nYou would need {increased_years} years instead of {data['years']} years")
            
            required_monthly = find_monthly_for_target(target_amount, data['initial_investment'], data['years'], opt_metrics['return'])
            st.info(f"**Option 2 - Increase Monthly Contribution:**\nIncrease monthly contribution from ${data['monthly_contribution']:,.0f} to ${required_monthly:,.0f}/month")
            
            required_initial = find_initial_for_target(target_amount, data['monthly_contribution'], data['years'], opt_metrics['return'])
            st.info(f"**Option 3 - Increase Initial Investment:**\nIncrease initial investment from ${data['initial_investment']:,.0f} to ${required_initial:,.0f}")
        else:
            st.success(f"✅ **Target Met:** Your optimized portfolio achieves ${opt_amount:,.0f}, meeting your target of ${target_amount:,.0f}.")

        # Charts Section
        st.subheader("📊 Portfolio Allocation")
        col1, col2 = st.columns(2)
        with col1:
            fig1, ax1 = plt.subplots(figsize=(6, 6))
            ax1.pie(equal_weights, labels=df['Fund Name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax1.set_title("Equal-Weighted Allocation")
            st.pyplot(fig1)
        
        with col2:
            fig2, ax2 = plt.subplots(figsize=(6, 6))
            ax2.pie(opt_weights, labels=df['Fund Name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax2.set_title("Optimized Allocation (Risk-Constrained)")
            st.pyplot(fig2)

        st.subheader("📈 Portfolio Performance vs Benchmark")
        years_list = sorted(list(set([year for year, col in year_returns_cols])))
        fig3, ax3 = plt.subplots(figsize=(12, 7))
        
        has_data = False
        has_benchmark = False
        eq_port_returns = []
        opt_port_returns = []
        bench_returns = []
        
        benchmark_names = df['Benchmark Name'].dropna().unique() if 'Benchmark Name' in df.columns else []
        primary_benchmark = benchmark_names[0] if len(benchmark_names) > 0 else "Composite Benchmark"
        
        for year in years_list:
            year_return_col = f'{year} Return (%)'
            year_bench_col = f'{year} Benchmark (%)'
            
            if year_return_col in df.columns:
                valid_mask = df[year_return_col].notna()
                if valid_mask.any():
                    eq_port_returns.append(np.sum(df[year_return_col][valid_mask] * equal_weights[valid_mask]))
                    opt_port_returns.append(np.sum(df[year_return_col][valid_mask] * opt_weights[valid_mask]))
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
                    bench_returns.append(np.sum(df[year_bench_col][valid_mask] * equal_weights[valid_mask]))
                    has_benchmark = True
                else:
                    bench_returns.append(np.nan)
            else:
                bench_returns.append(np.nan)
        
        if has_data:
            ax3.plot(years_list, eq_port_returns, marker='o', linewidth=3, label='Equal-Weighted Portfolio', color='#3498db', markersize=10)
            ax3.plot(years_list, opt_port_returns, marker='s', linewidth=3, label='Optimized Portfolio', color='#2ecc71', markersize=10)
        
        if has_benchmark:
            ax3.plot(years_list, bench_returns, marker='^', linewidth=3, label=primary_benchmark, color='#e74c3c', markersize=10)
        
        if has_data or has_benchmark:
            ax3.set_xlabel("Year", fontsize=12, fontweight='bold')
            ax3.set_ylabel("Return (%)", fontsize=12, fontweight='bold')
            ax3.set_title("Portfolio Performance vs Benchmark", fontsize=14, fontweight='bold')
            ax3.legend(loc='upper left', fontsize=11)
            ax3.grid(True, alpha=0.3)
            ax3.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
            st.pyplot(fig3)
        else:
            st.warning("No calendar year return data found.")

        st.subheader("📉 Individual Fund Performance")
        fig4, ax4 = plt.subplots(figsize=(14, 8))
        fund_has_data = False
        colors = plt.cm.tab10.colors
        
        full_years = [y for y, c in year_returns_cols]
        
        for i, (_, row) in enumerate(df.iterrows()):
            fund_returns = []
            for year, col_name in year_returns_cols:
                if col_name in df.columns and pd.notna(row.get(col_name)):
                    fund_returns.append(row[col_name])
                else:
                    fund_returns.append(np.nan)
            
            if any(pd.notna(r) for r in fund_returns):
                ax4.plot(full_years, fund_returns, marker='o', linewidth=2, 
                        label=row['Fund Name'][:30], color=colors[i % len(colors)], markersize=8, alpha=0.8)
                fund_has_data = True
        
        if fund_has_data:
            ax4.set_xlabel("Year", fontsize=12, fontweight='bold')
            ax4.set_ylabel("Return (%)", fontsize=12, fontweight='bold')
            ax4.set_title("Individual Fund Performance Comparison", fontsize=14, fontweight='bold')
            ax4.legend(loc='upper left', fontsize=9)
            ax4.grid(True, alpha=0.3)
            ax4.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
            st.pyplot(fig4)
        else:
            st.warning("No individual fund calendar year data available.")

        # COMBINED CHARTS PDF DOWNLOAD
        st.subheader("📥 Download All Charts")
        buf_charts = io.BytesIO()
        with PdfPages(buf_charts) as pdf:
            fig_combined1 = plt.figure(figsize=(12, 6))
            ax1 = fig_combined1.add_subplot(121)
            ax1.pie(equal_weights, labels=df['Fund Name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax1.set_title("Equal-Weighted Allocation")
            ax2 = fig_combined1.add_subplot(122)
            ax2.pie(opt_weights, labels=df['Fund Name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax2.set_title("Optimized Allocation")
            pdf.savefig(fig_combined1, bbox_inches='tight')
            plt.close(fig_combined1)
            
            if has_data or has_benchmark:
                pdf.savefig(fig3, bbox_inches='tight')
            if fund_has_data:
                pdf.savefig(fig4, bbox_inches='tight')
        
        buf_charts.seek(0)
        st.download_button(
            label="📥 Download All Charts (Combined PDF)",
            data=buf_charts,
            file_name=f"Portfolio_Charts_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

        st.subheader(" Portfolio Metrics Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("Risk-Adjusted Return", f"{opt_metrics['risk_adjusted']:.2f}", help="Return per unit of volatility")
        with col2: st.metric("Best Year", f"{opt_metrics['best_year']:.1f}%" if opt_metrics['best_year'] else "N/A")
        with col3: st.metric("Worst Year", f"{opt_metrics['worst_year']:.1f}%" if opt_metrics['worst_year'] else "N/A")
        with col4: st.metric("Consistency", f"{opt_metrics['consistency']:.0f}%", help="% of years with positive returns")

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("← Back to Input", use_container_width=True):
                st.session_state.page = 'input'
                st.rerun()
        with col2:
            if st.button("Generate Report →", use_container_width=True, type="primary"):
                st.session_state.page = 'report'
                st.rerun()

    elif st.session_state.page == 'report':
        st.header(" Step 3: Portfolio Analysis Report")
        st.success("✅ Report generation is ready.")
        
        data = st.session_state.portfolio_data
        df = data['funds_df']
        year_returns_cols = data['year_returns']
        
        n = len(df)
        equal_weights = np.ones(n) / n
        
        if data['goal_type'] == "Reach a Target Sum ($)":
            target_return = calculate_required_cagr(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'])
            target_amount = data['target_value']
        else:
            target_return = data['target_growth']
            target_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], target_return)
        
        eq_metrics_temp = calculate_portfolio_metrics(df, equal_weights, year_returns_cols)
        eq_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], eq_metrics_temp['return'])
        target_exceeded = eq_amount >= target_amount * 1.1
        
        opt_weights, opt_return, opt_vol = optimize_portfolio_risk_constrained(df, target_return, data['risk_profile'], target_exceeded)
        
        eq_metrics = calculate_portfolio_metrics(df, equal_weights, year_returns_cols)
        opt_metrics = calculate_portfolio_metrics(df, opt_weights, year_returns_cols)
        
        eq_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], eq_metrics['return'])
        opt_amount = calculate_future_value(data['initial_investment'], data['monthly_contribution'], data['years'], opt_metrics['return'])
        
        doc = Document()
        title = doc.add_heading('Portfolio Analysis Report', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
        if data.get('client_name'):
            doc.add_paragraph(f"Client Name: {data['client_name']}")
        doc.add_paragraph(f"Prepared For: {st.session_state.user_email}")
        doc.add_paragraph()
        
        doc.add_heading('Executive Summary', level=1)
        doc.add_paragraph(f"Goal Type: {data['goal_type']}")
        if data['goal_type'] == "Reach a Target Sum ($)":
            doc.add_paragraph(f"Target Amount: ${data['target_value']:,.0f}")
        else:
            doc.add_paragraph(f"Target Annual Growth: {data['target_growth']:.1f}%")
        doc.add_paragraph(f"Time Horizon: {data['years']} years")
        doc.add_paragraph(f"Risk Profile: {data['risk_profile']}")
        doc.add_paragraph(f"Initial Investment: ${data['initial_investment']:,.0f}")
        doc.add_paragraph(f"Monthly Contribution: ${data['monthly_contribution']:,.0f}")
        doc.add_paragraph()
        
        doc.add_heading('Input Data Summary', level=1)
        doc.add_paragraph(f"Number of Funds Analyzed: {len(df)}")
        doc.add_paragraph()
        
        doc.add_heading('Fund Details', level=2)
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Light Grid Accent 1'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Fund Name'
        hdr_cells[1].text = '1Y Return (%)'
        hdr_cells[2].text = 'Volatility (%)'
        hdr_cells[3].text = 'Mgmt Fee (%)'
        hdr_cells[4].text = 'Equal Weight'
        hdr_cells[5].text = 'Optimized Weight'
        
        for i, (_, row) in enumerate(df.iterrows()):
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Fund Name'])
            row_cells[1].text = f"{row['1Y Return (%)']:.2f}" if pd.notna(row['1Y Return (%)']) else "N/A"
            row_cells[2].text = f"{row['Volatility (%)']:.2f}" if pd.notna(row['Volatility (%)']) else "N/A"
            row_cells[3].text = f"{row['Mgmt Fee (%)']:.2f}" if pd.notna(row['Mgmt Fee (%)']) else "N/A"
            row_cells[4].text = f"{equal_weights[i]:.1%}"
            row_cells[5].text = f"{opt_weights[i]:.1%}"
        doc.add_paragraph()
        
        doc.add_heading('Portfolio Performance Analysis', level=1)
        doc.add_paragraph(f"Equal-Weighted Return: {eq_metrics['return']:.1f}% p.a.")
        doc.add_paragraph(f"Optimized Return: {opt_metrics['return']:.1f}% p.a.")
        doc.add_paragraph(f"Target Return: {target_return:.1f}% p.a.")
        doc.add_paragraph(f"Equal-Weighted Projected Amount: ${eq_amount:,.0f}")
        doc.add_paragraph(f"Optimized Projected Amount: ${opt_amount:,.0f}")
        doc.add_paragraph(f"Target Amount: ${target_amount:,.0f}")
        doc.add_paragraph()
        
        doc.add_heading('Risk Assessment', level=1)
        risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
        risk_thresh = risk_thresholds[data['risk_profile']]
        doc.add_paragraph(f"Risk Profile Threshold: {risk_thresh}% volatility")
        doc.add_paragraph(f"Equal-Weighted Volatility: {eq_metrics['volatility']:.1f}%")
        doc.add_paragraph(f"Optimized Volatility: {opt_metrics['volatility']:.1f}%")
        doc.add_paragraph(f"Risk-Adjusted Return: {opt_metrics['risk_adjusted']:.2f}")
        doc.add_paragraph()
        
        if opt_metrics['yearly_returns']:
            doc.add_heading('Historical Performance (Calendar Year Returns)', level=1)
            doc.add_paragraph(f"Best Year: {opt_metrics['best_year']:.1f}%")
            doc.add_paragraph(f"Worst Year: {opt_metrics['worst_year']:.1f}%")
            doc.add_paragraph(f"Average Annual Return: {opt_metrics['avg_yearly']:.1f}%")
            doc.add_paragraph(f"Consistency: {opt_metrics['consistency']:.0f}% of years with positive returns")
            doc.add_paragraph()
        
        doc.add_heading('Fee Impact Analysis', level=1)
        doc.add_paragraph(f"Average Management Fee: {opt_metrics['fee']:.2f}% p.a.")
        fee_impact_10yr = data['initial_investment'] * (1 - (1 - opt_metrics['fee']/100)**10)
        doc.add_paragraph(f"Estimated Fee Impact over 10 years: ${fee_impact_10yr:,.0f}")
        doc.add_paragraph()
        
        doc.add_heading('Recommendations', level=1)
        
        if opt_amount > target_amount * 1.1:
            doc.add_paragraph("✅ SURPLUS SCENARIO: Your portfolio exceeds the target.")
            
            def find_monthly_for_target(target, initial, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, (target - initial) / (years * 12))
                fv_initial = initial * ((1 + r) ** years)
                remaining = target - fv_initial
                if remaining <= 0: return 0
                return max(0, remaining * r / (12 * (((1 + r) ** years) - 1)))
            
            def find_initial_for_target(target, monthly, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, target - (monthly * 12 * years))
                fv_monthly = monthly * 12 * (((1 + r) ** years) - 1) / r
                remaining = target - fv_monthly
                if remaining <= 0: return 0
                return max(0, remaining / ((1 + r) ** years))
            
            def find_years_for_target(target, initial, monthly, return_rate):
                for test_years in range(1, 50):
                    if calculate_future_value(initial, monthly, test_years, return_rate) >= target:
                        return test_years
                return 50
            
            reduced_monthly = find_monthly_for_target(target_amount, data['initial_investment'], data['years'], opt_metrics['return'])
            doc.add_paragraph(f"Option 1: Reduce monthly contribution from ${data['monthly_contribution']:,.0f} to ${reduced_monthly:,.0f}")
            
            reduced_initial = find_initial_for_target(target_amount, 0, data['years'], opt_metrics['return'])
            doc.add_paragraph(f"Option 2: If monthly contribution is $0, reduce initial investment to ${reduced_initial:,.0f}")
            
            shortened_years = find_years_for_target(target_amount, data['initial_investment'], data['monthly_contribution'], opt_metrics['return'])
            doc.add_paragraph(f"Option 3: Achieve target in {shortened_years} years instead of {data['years']} years")
                
        elif opt_amount < target_amount:
            doc.add_paragraph("⚠️ SHORTFALL SCENARIO: Your portfolio cannot meet the target.")
            
            def find_years_for_target(target, initial, monthly, return_rate):
                for test_years in range(1, 100):
                    if calculate_future_value(initial, monthly, test_years, return_rate) >= target:
                        return test_years
                return 100
            
            def find_monthly_for_target(target, initial, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, (target - initial) / (years * 12))
                fv_initial = initial * ((1 + r) ** years)
                remaining = target - fv_initial
                if remaining <= 0: return 0
                return max(0, remaining * r / (12 * (((1 + r) ** years) - 1)))
            
            def find_initial_for_target(target, monthly, years, return_rate):
                r = return_rate / 100
                if r == 0: return max(0, target - (monthly * 12 * years))
                fv_monthly = monthly * 12 * (((1 + r) ** years) - 1) / r
                remaining = target - fv_monthly
                if remaining <= 0: return 0
                return max(0, remaining / ((1 + r) ** years))
            
            increased_years = find_years_for_target(target_amount, data['initial_investment'], data['monthly_contribution'], opt_metrics['return'])
            doc.add_paragraph(f"Option 1: Extend time horizon to {increased_years} years")
            
            required_monthly = find_monthly_for_target(target_amount, data['initial_investment'], data['years'], opt_metrics['return'])
            doc.add_paragraph(f"Option 2: Increase monthly contribution to ${required_monthly:,.0f}")
            
            required_initial = find_initial_for_target(target_amount, data['monthly_contribution'], data['years'], opt_metrics['return'])
            doc.add_paragraph(f"Option 3: Increase initial investment to ${required_initial:,.0f}")
        else:
            doc.add_paragraph("✅ TARGET MET: Your optimized portfolio meets your investment goal.")
        doc.add_paragraph()
        
        doc.add_heading('IMPORTANT DISCLAIMER & PROFESSIONAL GUIDANCE', level=1)
        
        doc.add_heading('1. Nature of This Analysis', level=2)
        doc.add_paragraph("This Portfolio Analysis Report is generated based on the specific inputs, assumptions, and historical data provided by the user via this tool. The calculations, optimized allocations, and projections (including recommendations to adjust contributions, initial capital, or time horizons) are mathematical models intended for educational and illustrative purposes only.")
        
        doc.add_heading('2. No Guarantee of Future Performance', level=2)
        doc.add_paragraph("The analysis relies on historical performance metrics (e.g., past returns, volatility) and stated fund objectives. Past performance is not indicative of future results. Market conditions, fund management changes, and economic factors can cause actual outcomes to differ materially from the projections shown in this report.")
        
        doc.add_heading('3. Not Professional Financial Advice', level=2)
        doc.add_paragraph("This tool and its outputs do not constitute personalized financial, investment, tax, or legal advice. The recommendations provided are generic and do not take into account your complete financial picture, liquidity needs, tax status, or other personal circumstances. You should not make any investment decisions solely based on this report.")
        
        doc.add_heading('4. Professional Consultation', level=2)
        doc.add_paragraph("While this tool provides a valuable high-level feasibility assessment, building a comprehensive, concrete wealth strategy requires a holistic review of your unique financial situation. If you would like to translate this analysis into an actionable, personalized investment strategy and execute it with professional oversight, please reach out:")
        
        p = doc.add_paragraph()
        p.add_run('Christopher Chew, CFP®, CFC®\n').bold = True
        doc.add_paragraph('Certified Financial Planner | Certified Business & Financial Coach')
        doc.add_paragraph('• Email: chrischew@acaplt.com')
        doc.add_paragraph('• Mobile/WhatsApp: +6012-213 9559')
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📄 Download Word Report (DOCX)",
            data=buffer,
            file_name=f"Portfolio_Analysis_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        st.info("💡 **Note:** To convert the Word document to PDF, open it in Microsoft Word and use 'Save As' > 'PDF' format.")
        
        if st.button("← Back to Analysis", use_container_width=True):
            st.session_state.page = 'analysis'
            st.rerun()