import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import json
import firebase_admin
from firebase_admin import credentials, firestore
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import random
import string
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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
for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'otp_code', 'otp_email', 'show_otp_input', 'funds_df']:
    if key not in st.session_state:
        st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else None if key == 'funds_df' else {})

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

def optimize_portfolio_risk_constrained(df, target_return, risk_profile):
    """Optimizes portfolio to meet target return while respecting risk profile."""
    n = len(df)
    if n == 0: return np.array([]), 0, 0
    
    # Risk thresholds
    risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
    max_volatility = risk_thresholds.get(risk_profile, 15.0)
    
    returns = df['1Y Return (%)'].fillna(df['3Y Return (%)'].fillna(0)).values
    volatilities = df['Volatility (%)'].fillna(0).values
    
    # Strategy: Start with equal weights, then adjust
    weights = np.ones(n) / n
    
    # If equal-weighted already meets target and risk, return it
    eq_return = np.sum(returns * weights)
    eq_vol = np.sum(volatilities * weights)
    
    if eq_return >= target_return and eq_vol <= max_volatility:
        return weights, eq_return, eq_vol
    
    # If target is not met, shift weight to higher-return funds (up to 40% each)
    if eq_return < target_return:
        sorted_indices = np.argsort(returns)[::-1]
        weights = np.zeros(n)
        remaining = 1.0
        
        for idx in sorted_indices:
            funds_left = n - np.sum(weights > 0)
            max_w = min(0.40, remaining - (0.05 * max(0, funds_left - 1)))
            w = max(0.05, max_w)
            weights[idx] = w
            remaining -= w
            if remaining <= 0.05:
                unallocated = np.where(weights == 0)[0]
                if len(unallocated) > 0:
                    weights[unallocated] = remaining / len(unallocated)
                break
        
        weights = weights / weights.sum()
    
    # Check if we're within risk profile
    opt_return = np.sum(returns * weights)
    opt_vol = np.sum(volatilities * weights)
    
    # If volatility exceeds risk profile, shift to lower-volatility funds
    if opt_vol > max_volatility:
        # Sort by return/volatility ratio (best risk-adjusted)
        ratios = np.where(volatilities > 0, returns / volatilities, 0)
        sorted_indices = np.argsort(ratios)[::-1]
        
        weights = np.zeros(n)
        remaining = 1.0
        
        for idx in sorted_indices:
            funds_left = n - np.sum(weights > 0)
            max_w = min(0.40, remaining - (0.05 * max(0, funds_left - 1)))
            w = max(0.05, max_w)
            weights[idx] = w
            remaining -= w
            if remaining <= 0.05:
                unallocated = np.where(weights == 0)[0]
                if len(unallocated) > 0:
                    weights[unallocated] = remaining / len(unallocated)
                break
        
        weights = weights / weights.sum()
        opt_return = np.sum(returns * weights)
        opt_vol = np.sum(volatilities * weights)
    
    return weights, opt_return, opt_vol

def calculate_extra_contribution(target_sum, initial_investment, current_monthly, years, annual_return_pct):
    r = annual_return_pct / 100
    if r == 0:
        required_monthly = (target_sum - initial_investment) / (years * 12)
    else:
        required_monthly = (target_sum - initial_investment * ((1 + r) ** years)) * r / (((1 + r) ** years) - 1) / 12
    return max(0, required_monthly - current_monthly)

def calculate_reduced_initial_investment(target_sum, monthly_contribution, years, annual_return_pct):
    """Calculate how much initial investment can be reduced while still meeting target."""
    r = annual_return_pct / 100
    if r == 0:
        max_initial = target_sum - (monthly_contribution * 12 * years)
    else:
        fv_contributions = monthly_contribution * 12 * (((1 + r) ** years - 1) / r)
        max_initial = (target_sum - fv_contributions) / ((1 + r) ** years)
    return max(0, max_initial)

def calculate_portfolio_metrics(df, weights):
    """Calculate comprehensive portfolio metrics."""
    returns_1y = df['1Y Return (%)'].fillna(df['3Y Return (%)'].fillna(0))
    volatilities = df['Volatility (%)'].fillna(0)
    fees = df['Mgmt Fee (%)'].fillna(0)
    
    portfolio_return = np.sum(returns_1y * weights)
    portfolio_volatility = np.sum(volatilities * weights)
    portfolio_fee = np.sum(fees * weights)
    
    # Risk-adjusted return (return per unit of volatility)
    risk_adjusted = portfolio_return / portfolio_volatility if portfolio_volatility > 0 else 0
    
    # Calendar year metrics (if available)
    years = [2016, 2017, 2018, 2019, 2020]
    yearly_returns = []
    
    for year in years:
        col = f'{year} Return (%)'
        if col in df.columns:
            year_return = np.sum(df[col].fillna(0) * weights)
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
    st.markdown("<h2 style='text-align: center;'> Login to Portfolio Analyzer</h2>", unsafe_allow_html=True)
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
            for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'show_otp_input', 'otp_code', 'otp_email', 'funds_df']:
                st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else None if key == 'funds_df' else {})
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
                st.rerun()
        with col2:
            if st.session_state.user_email == ADMIN_EMAIL: st.info("🔐 Admin: Unlimited access")

    elif st.session_state.page == 'input':
        st.header("📝 Step 1: Define Goal & Upload Fund Data")
        
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
        
        # Provide template download
        if st.button("📥 Download Excel Template"):
            template_df = pd.DataFrame({
                'Fund Name': ['AM Income and Growth Fund', 'Manulife US Equity Fund', 'EastSpring Small Cap'],
                '1Y Return (%)': [16.28, 34.01, 31.72],
                '3Y Return (%)': [np.nan, 74.49, 49.90],
                '5Y Return (%)': [np.nan, 127.68, 1137.30],
                'Volatility (%)': [np.nan, 21.0, 19.0],
                'Mgmt Fee (%)': [1.80, 1.80, 1.50],
                '1Y Benchmark (%)': [30.82, 28.94, 10.33],
                '2016 Return (%)': [np.nan, 9.75, 0.28],
                '2017 Return (%)': [np.nan, 7.74, 21.72],
                '2018 Return (%)': [np.nan, -4.34, -18.97],
                '2019 Return (%)': [np.nan, 29.65, 18.03],
                '2020 Return (%)': [np.nan, 17.28, 19.36],
                '2016 Benchmark (%)': [np.nan, 14.45, -7.71],
                '2017 Benchmark (%)': [np.nan, 7.73, 15.87],
                '2018 Benchmark (%)': [np.nan, -4.26, -29.14],
                '2019 Benchmark (%)': [np.nan, 27.57, 25.36],
                '2020 Benchmark (%)': [np.nan, 14.33, 9.89]
            })
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                template_df.to_excel(writer, sheet_name='Fund Data', index=False)
            
            st.download_button(
                label="Download Template Excel",
                data=buffer.getvalue(),
                file_name="fund_data_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # File uploader
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                
                # Validate required columns
                required_cols = ['Fund Name', '1Y Return (%)', 'Volatility (%)', 'Mgmt Fee (%)']
                missing_cols = [col for col in required_cols if col not in df.columns]
                
                if missing_cols:
                    st.error(f"Missing required columns: {', '.join(missing_cols)}")
                else:
                    # Clean and standardize column names
                    df.columns = df.columns.str.strip()
                    
                    # Convert numeric columns
                    numeric_cols = ['1Y Return (%)', '3Y Return (%)', '5Y Return (%)', 'Volatility (%)', 'Mgmt Fee (%)', 
                                    '1Y Benchmark (%)', '2016 Return (%)', '2017 Return (%)', '2018 Return (%)', 
                                    '2019 Return (%)', '2020 Return (%)', '2016 Benchmark (%)', '2017 Benchmark (%)', 
                                    '2018 Benchmark (%)', '2019 Benchmark (%)', '2020 Benchmark (%)']
                    for col in numeric_cols:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                    
                    st.session_state.funds_df = df
                    st.success(f"✅ Loaded {len(df)} funds successfully!")
                    st.dataframe(df, use_container_width=True)
                    
                    st.markdown("---")
                    if st.button("Calculate Portfolio Analysis", use_container_width=True, type="primary"):
                        st.session_state.portfolio_data = {
                            'goal_type': goal_type, 'target_value': target_value, 'target_growth': target_growth,
                            'years': years, 'initial_investment': initial_investment, 'monthly_contribution': monthly_contribution,
                            'risk_profile': risk_profile, 'funds_df': df
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
        st.header(" Step 2: Portfolio Analysis")
        data = st.session_state.portfolio_data
        df = data['funds_df'].copy()
        
        # Calculate Target Return
        if data['goal_type'] == "Reach a Target Sum ($)":
            target_return = calculate_required_cagr(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'])
        else:
            target_return = data['target_growth']

        # 1. Equal Weighted Portfolio
        n = len(df)
        equal_weights = np.ones(n) / n
        eq_metrics = calculate_portfolio_metrics(df, equal_weights)

        # 2. Risk-Constrained Optimized Portfolio
        opt_weights, opt_return, opt_vol = optimize_portfolio_risk_constrained(df, target_return, data['risk_profile'])
        opt_metrics = calculate_portfolio_metrics(df, opt_weights)

        # Display Metrics
        st.subheader("Portfolio Performance Analysis")
        col1, col2, col3 = st.columns(3)
        with col1: st.metric("Equal-Weighted Return", f"{eq_metrics['return']:.1f}% p.a.")
        with col2: st.metric("Optimized Return", f"{opt_metrics['return']:.1f}% p.a.")
        with col3: st.metric("Target Return", f"{target_return:.1f}% p.a.")

        # Feasibility & Risk
        st.subheader("Goal Feasibility & Risk Assessment")
        risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
        risk_thresh = risk_thresholds[data['risk_profile']]
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Equal-Weighted Portfolio", "✅ Achievable" if eq_metrics['return'] >= target_return else "️ Shortfall", 
                      delta=f"{eq_metrics['return'] - target_return:.1f}%")
            st.caption(f"Volatility: {eq_metrics['volatility']:.1f}% ({'✅ Matches' if eq_metrics['volatility'] <= risk_thresh else '⚠️ Exceeds'} {data['risk_profile']} profile)")
        with col2:
            st.metric("Optimized Portfolio", "✅ Achievable" if opt_metrics['return'] >= target_return else "⚠️ Shortfall", 
                      delta=f"{opt_metrics['return'] - target_return:.1f}%")
            st.caption(f"Volatility: {opt_metrics['volatility']:.1f}% ({'✅ Matches' if opt_metrics['volatility'] <= risk_thresh else '⚠️ Exceeds'} {data['risk_profile']} profile)")

        # Recommendations
        st.subheader(" Recommendations")
        
        if opt_metrics['return'] < target_return:
            st.error("⚠️ **Target Not Met:** Even with optimization, the maximum achievable return is {:.1f}%, which is below your target of {:.1f}%.".format(opt_metrics['return'], target_return))
            extra_monthly = calculate_extra_contribution(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'], opt_metrics['return'])
            st.info(f"💡 **Option 1:** Increase monthly contribution by **${extra_monthly:,.0f}** (New total: ${data['monthly_contribution'] + extra_monthly:,.0f}/month)")
        elif opt_metrics['return'] > target_return * 1.5:  # If return is 50%+ above target
            st.success("✅ **Target Exceeded:** Your portfolio can achieve {:.1f}% return, well above your target of {:.1f}%.".format(opt_metrics['return'], target_return))
            
            # Option 1: Reduce monthly contribution
            reduced_monthly = data['monthly_contribution'] - calculate_extra_contribution(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'], opt_metrics['return'])
            st.info(f"💡 **Option 1:** You could reduce your monthly contribution to **${max(0, reduced_monthly):,.0f}** (save ${data['monthly_contribution'] - max(0, reduced_monthly):,.0f}/month)")
            
            # Option 2: Reduce initial investment
            reduced_initial = calculate_reduced_initial_investment(data['target_value'], data['monthly_contribution'], data['years'], opt_metrics['return'])
            if reduced_initial < data['initial_investment']:
                st.info(f"💡 **Option 2:** You could reduce your initial investment to **${reduced_initial:,.0f}** (save ${data['initial_investment'] - reduced_initial:,.0f})")
            
            # Option 3: Shorten time horizon
            st.info(f"💡 **Option 3:** You could achieve your target in fewer years with this portfolio performance")
        else:
            st.success("✅ **Target Met:** Your optimized portfolio achieves {:.1f}% return, meeting your target of {:.1f}%.".format(opt_metrics['return'], target_return))
            if opt_metrics['volatility'] > risk_thresh:
                st.warning(f"⚠️ However, volatility ({opt_metrics['volatility']:.1f}%) exceeds your {data['risk_profile']} risk profile ({risk_thresh}%). Consider a more conservative allocation.")

        # Charts
        st.subheader("Portfolio Allocation")
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

        # Historical Performance vs Benchmark Line Chart
        st.subheader("Historical Performance vs Benchmark (Calendar Year Returns)")
        years_list = [2016, 2017, 2018, 2019, 2020]
        fig_hist, ax_hist = plt.subplots(figsize=(12, 7))
        
        has_data = False
        has_benchmark = False
        
        # Plot portfolio returns
        port_returns = []
        bench_returns = []
        
        for year in years_list:
            col_name = f'{year} Return (%)'
            bench_col = f'{year} Benchmark (%)'
            
            if col_name in df.columns:
                port_return = np.sum(df[col_name].fillna(0) * equal_weights)
                port_returns.append(port_return)
                has_data = True
            else:
                port_returns.append(np.nan)
            
            if bench_col in df.columns:
                bench_return = np.sum(df[bench_col].fillna(0) * equal_weights)
                bench_returns.append(bench_return)
                has_benchmark = True
            else:
                bench_returns.append(np.nan)
        
        if has_data:
            ax_hist.plot(years_list, port_returns, marker='o', linewidth=3, label='Portfolio (Equal-Weighted)', color='#2ecc71', markersize=10)
        
        if has_benchmark:
            ax_hist.plot(years_list, bench_returns, marker='s', linewidth=3, label='Composite Benchmark', color='#e74c3c', markersize=10)
        
        if has_data or has_benchmark:
            ax_hist.set_xlabel("Year", fontsize=12, fontweight='bold')
            ax_hist.set_ylabel("Return (%)", fontsize=12, fontweight='bold')
            ax_hist.set_title("Portfolio vs Benchmark Performance (2016-2020)", fontsize=14, fontweight='bold')
            ax_hist.legend(loc='upper left', fontsize=11)
            ax_hist.grid(True, alpha=0.3)
            ax_hist.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
            
            # Add value labels
            for i, (year, port, bench) in enumerate(zip(years_list, port_returns, bench_returns)):
                if not np.isnan(port):
                    ax_hist.annotate(f'{port:.1f}%', (year, port), textcoords="offset points", xytext=(0,10), ha='center', fontsize=9, color='#2ecc71', fontweight='bold')
                if not np.isnan(bench):
                    ax_hist.annotate(f'{bench:.1f}%', (year, bench), textcoords="offset points", xytext=(0,-15), ha='center', fontsize=9, color='#e74c3c', fontweight='bold')
            
            st.pyplot(fig_hist)
        else:
            st.warning("No calendar year return data (2016-2020) found in the uploaded Excel file.")

        # Portfolio Metrics Summary
        st.subheader("📈 Portfolio Metrics Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Risk-Adjusted Return", f"{opt_metrics['risk_adjusted']:.2f}", help="Return per unit of volatility")
        with col2:
            st.metric("Best Year", f"{opt_metrics['best_year']:.1f}%" if opt_metrics['best_year'] else "N/A")
        with col3:
            st.metric("Worst Year", f"{opt_metrics['worst_year']:.1f}%" if opt_metrics['worst_year'] else "N/A")
        with col4:
            st.metric("Consistency", f"{opt_metrics['consistency']:.0f}%", help="% of years with positive returns")

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
        st.header("📄 Step 3: Portfolio Analysis Report")
        st.success("✅ Report generation is ready.")
        
        data = st.session_state.portfolio_data
        df = data['funds_df']
        
        # Calculate metrics
        n = len(df)
        equal_weights = np.ones(n) / n
        opt_weights, opt_return, opt_vol = optimize_portfolio_risk_constrained(df, 
            calculate_required_cagr(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years']) if data['goal_type'] == "Reach a Target Sum ($)" else data['target_growth'],
            data['risk_profile'])
        
        eq_metrics = calculate_portfolio_metrics(df, equal_weights)
        opt_metrics = calculate_portfolio_metrics(df, opt_weights)
        
        # Generate Word report
        doc = Document()
        
        # Title
        title = doc.add_heading('Portfolio Analysis Report', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
        doc.add_paragraph(f"Client: {st.session_state.user_email}")
        doc.add_paragraph()
        
        # Executive Summary
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
        
        # Portfolio Performance
        doc.add_heading('Portfolio Performance Analysis', level=1)
        doc.add_paragraph(f"Equal-Weighted Return: {eq_metrics['return']:.1f}% p.a.")
        doc.add_paragraph(f"Optimized Return: {opt_metrics['return']:.1f}% p.a.")
        doc.add_paragraph(f"Target Return: {calculate_required_cagr(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years']) if data['goal_type'] == 'Reach a Target Sum ($)' else data['target_growth']:.1f}% p.a.")
        doc.add_paragraph()
        
        # Risk Metrics
        doc.add_heading('Risk Assessment', level=1)
        doc.add_paragraph(f"Equal-Weighted Volatility: {eq_metrics['volatility']:.1f}%")
        doc.add_paragraph(f"Optimized Volatility: {opt_metrics['volatility']:.1f}%")
        doc.add_paragraph(f"Risk-Adjusted Return: {opt_metrics['risk_adjusted']:.2f}")
        doc.add_paragraph()
        
        # Historical Performance
        if opt_metrics['yearly_returns']:
            doc.add_heading('Historical Performance (Calendar Year Returns)', level=1)
            doc.add_paragraph(f"Best Year: {opt_metrics['best_year']:.1f}%")
            doc.add_paragraph(f"Worst Year: {opt_metrics['worst_year']:.1f}%")
            doc.add_paragraph(f"Average Annual Return: {opt_metrics['avg_yearly']:.1f}%")
            doc.add_paragraph(f"Consistency: {opt_metrics['consistency']:.0f}% of years with positive returns")
            doc.add_paragraph()
        
        # Fee Impact Analysis
        doc.add_heading('Fee Impact Analysis', level=1)
        doc.add_paragraph(f"Average Management Fee: {opt_metrics['fee']:.2f}% p.a.")
        fee_impact_10yr = data['initial_investment'] * (1 - (1 - opt_metrics['fee']/100)**10)
        doc.add_paragraph(f"Estimated Fee Impact over 10 years: ${fee_impact_10yr:,.0f}")
        doc.add_paragraph()
        
        # Funds Analyzed
        doc.add_heading('Funds Analyzed', level=1)
        for _, row in df.iterrows():
            doc.add_paragraph(f"• {row['Fund Name']}")
        doc.add_paragraph()
        
        # Recommendations
        doc.add_heading('Recommendations', level=1)
        target_return = calculate_required_cagr(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years']) if data['goal_type'] == "Reach a Target Sum ($)" else data['target_growth']
        
        if opt_metrics['return'] < target_return:
            doc.add_paragraph("⚠️ Target Not Met: The portfolio cannot achieve your target return with the current funds.")
            extra_monthly = calculate_extra_contribution(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'], opt_metrics['return'])
            doc.add_paragraph(f"Recommendation: Increase monthly contribution by ${extra_monthly:,.0f}")
        elif opt_metrics['return'] > target_return * 1.5:
            doc.add_paragraph("✅ Target Exceeded: Your portfolio significantly outperforms your target.")
            reduced_monthly = data['monthly_contribution'] - calculate_extra_contribution(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'], opt_metrics['return'])
            doc.add_paragraph(f"Recommendation: You could reduce monthly contribution to ${max(0, reduced_monthly):,.0f}")
        else:
            doc.add_paragraph("✅ Target Met: Your optimized portfolio meets your investment goal.")
        
        doc.add_paragraph()
        doc.add_paragraph("Disclaimer: This report is for informational purposes only and does not constitute financial advice. Past performance is not indicative of future results.")
        
        # Save to buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📄 Download Word Report",
            data=buffer,
            file_name=f"Portfolio_Analysis_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        if st.button("← Back to Analysis", use_container_width=True):
            st.session_state.page = 'analysis'
            st.rerun()