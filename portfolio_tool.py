import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import random
import string
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import matplotlib.pyplot as plt
import json
import firebase_admin
from firebase_admin import credentials, firestore

# ============================================================
# SECTION: FIREBASE & EMAIL CONFIGURATION
# ============================================================
# Initialize Firebase (only once per session)
if "firebase_initialized" not in st.session_state:
    try:
        # Check if Firebase app already exists
        if not firebase_admin._apps:
            cred_dict = st.secrets.get("FIREBASE_SERVICE_ACCOUNT", {})
            if isinstance(cred_dict, str):
                cred_dict = json.loads(cred_dict)
            cred = credentials.Certificate(cred_dict)
            firebase_admin.initialize_app(cred)
            st.session_state.firebase_initialized = True
        else:
            # Firebase already initialized
            st.session_state.firebase_initialized = True
    except Exception as e:
        if "already exists" not in str(e):
            st.error(f"Firebase init error: {e}")
        else:
            st.session_state.firebase_initialized = True

# Initialize Firestore client
try:
    db = firestore.client()
except:
    db = None

db = firestore.client() if st.session_state.get("firebase_initialized") else None

ADMIN_EMAIL = st.secrets.get("ADMIN_EMAIL", "cktchew@gmail.com")
GMAIL_ADDRESS = st.secrets.get("GMAIL_ADDRESS", "cktchew@gmail.com")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")

# ============================================================
# SECTION: STREAMLIT PAGE CONFIG
# ============================================================
st.set_page_config(page_title="Chew Advisory - Portfolio Analyzer", layout="wide")
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
for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'otp_code', 'otp_email', 'show_otp_input']:
    if key not in st.session_state:
        st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else {})

# ============================================================
# SECTION: HELPER FUNCTIONS
# ============================================================
def safe_float(v):
    try:
        return float(str(v).replace(',', '').strip() or 0)
    except:
        return 0.0

def safe_int(value, default=0):
    try:
        return int(float(value)) if value is not None and str(value).strip() != '' else default
    except (ValueError, TypeError):
        return default

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
        body = f"Hello,\n\nYour One-Time Password (OTP) for the Portfolio Analyzer Tool is:\n{otp_code}\n\nThis code will expire in 10 minutes.\n\nBest regards,\nChristopher Chew\nChew Advisory"
        message.attach(MIMEText(body, "plain"))
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, recipient_email, message.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending OTP email: {str(e)}")
        return False

def get_user_stats(email):
    if email == ADMIN_EMAIL:
        return "allowed", 0, 0, 999999
    if not db:
        return "allowed", 0, 0, 3 # Fallback if Firebase fails
    
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            data = docs[0].to_dict()
            if data.get('deleted_at') is not None:
                return "deleted", 0, 0, 0
            return "allowed", safe_int(data.get('access_count'), 0), safe_int(data.get('generation_count'), 0), safe_int(data.get('max_limit'), 3)
        else:
            # Create new user record
            db.collection('user_usage').add({
                'email': email, 'access_count': 0, 'generation_count': 0, 'max_limit': 3, 'created_at': firestore.SERVER_TIMESTAMP
            })
            return "allowed", 0, 0, 3
    except Exception as e:
        st.error(f"Database error: {str(e)}")
        return "allowed", 0, 0, 3

def check_access_allowed(email):
    status, acc, gen, lim = get_user_stats(email)
    if status != "allowed" or acc >= lim or gen >= lim:
        return False, lim, acc, gen
    return True, lim, acc, gen

def increment_access(email):
    if not db or email == ADMIN_EMAIL: return
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            doc_id = docs[0].id
            current = safe_int(docs[0].to_dict().get('access_count'), 0)
            db.collection('user_usage').document(doc_id).update({'access_count': current + 1})
    except Exception as e:
        print(f"Error incrementing access: {e}")

def increment_generation(email):
    if not db or email == ADMIN_EMAIL: return
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            doc_id = docs[0].id
            current = safe_int(docs[0].to_dict().get('generation_count'), 0)
            db.collection('user_usage').document(doc_id).update({'generation_count': current + 1})
    except Exception as e:
        print(f"Error incrementing generation: {e}")

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
                else:
                    st.error("Failed to send OTP. Please try again.")
            else:
                st.error("Please enter a valid email address.")
        
        if st.session_state.get('show_otp_input', False):
            st.info("An OTP code has been sent to your email. Please check your inbox.")
            st.markdown("---")
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
                    else:
                        st.error(f"❌ Limit reached ({lim} accesses/generations). Contact cktchew@gmail.com.")
                else:
                    st.error("❌ Incorrect OTP. Please try again.")

# ============================================================
# SECTION: PORTFOLIO CALCULATIONS
# ============================================================
def calculate_required_cagr(target_sum, initial_investment, monthly_contribution, years):
    """Calculates the required CAGR to reach a target sum."""
    if years <= 0 or target_sum <= 0:
        return 0.0
    
    r = 0.05 # Initial guess 5%
    for _ in range(50): # 50 iterations for convergence
        fv_guess = initial_investment * ((1 + r) ** years)
        if r > 0:
            fv_guess += monthly_contribution * 12 * (((1 + r) ** years - 1) / r)
        else:
            fv_guess += monthly_contribution * 12 * years
            
        derivative = initial_investment * years * ((1 + r) ** (years - 1))
        if r > 0:
            derivative += monthly_contribution * 12 * (years * (1 + r)**(years - 1) * r - ((1 + r)**years - 1)) / (r**2)
            
        if abs(derivative) < 1e-8:
            break
            
        r = r - (fv_guess - target_sum) / derivative
        if r < -0.5: r = -0.5 # Prevent negative infinity
        
    return max(0.0, r * 100) # Return as percentage

def optimize_portfolio(funds_df):
    """Simple equal weight vs. basic optimization mock for V1."""
    weights = np.ones(len(funds_df)) / len(funds_df)
    port_return = np.sum(funds_df['expected_return'] * weights)
    port_volatility = np.sqrt(np.sum((funds_df['volatility'] * weights) ** 2)) 
    
    return weights, port_return, port_volatility

# ============================================================
# SECTION: MAIN APP LOGIC
# ============================================================
if not st.session_state.authenticated:
    show_login_page()
else:
    # Header
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
            for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'show_otp_input', 'otp_code', 'otp_email']:
                st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else {})
            st.rerun()

    # Navigation
    if st.session_state.page == 'home':
        st.markdown("---")
        st.markdown("### Welcome to the Portfolio Analyzer")
        st.markdown("This tool helps you determine if your selected funds can achieve your financial goals through historical analysis and projection.")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("📝 Start New Analysis", use_container_width=True, type="primary"):
                st.session_state.page = 'input'
                st.session_state.portfolio_data = {}
                st.rerun()
        with col2:
            if st.session_state.user_email == ADMIN_EMAIL:
                st.info("🔐 Admin: Unlimited access")

    elif st.session_state.page == 'input':
        st.header("📝 Step 1: Define Goal & Upload Funds")
        
        st.subheader("1. Investment Goal")
        col1, col2 = st.columns(2)
        with col1:
            goal_type = st.radio("Goal Type:", ["Reach a Target Sum ($)", "Achieve Target Annual Growth (%)"])
        with col2:
            years = st.number_input("Time Horizon (Years)", min_value=1, max_value=50, value=10)
            
        if goal_type == "Reach a Target Sum ($)":
            target_value = st.number_input("Target Final Sum ($)", min_value=1000, value=100000, step=1000)
            target_growth = None
        else:
            target_growth = st.number_input("Target Annual Growth (%)", min_value=1.0, max_value=20.0, value=8.0, step=0.5)
            target_value = None

        st.subheader("2. Capital & Contributions")
        col1, col2 = st.columns(2)
        with col1:
            initial_investment = st.number_input("Initial Lump Sum Investment ($)", min_value=0, value=10000, step=1000)
        with col2:
            monthly_contribution = st.number_input("Monthly Contribution ($)", min_value=0, value=500, step=100)

        st.subheader("3. Fund Data Input (FFS)")
        st.info("Upload a CSV file with columns: `Fund Name`, `Expected Return (%)`, `Volatility (%)`, `Expense Ratio (%)`. \n*(A template will be provided if you don't have one).*")
        
        uploaded_file = st.file_uploader("Upload Fund Data (CSV)", type=['csv'])
        
        # Provide a template download
        if st.button(" Download CSV Template"):
            template = pd.DataFrame({
                "Fund Name": ["Fund A (e.g., S&P 500 ETF)", "Fund B (e.g., Global Bond Fund)"],
                "Expected Return (%)": [8.5, 4.0],
                "Volatility (%)": [15.0, 5.0],
                "Expense Ratio (%)": [0.15, 0.25]
            })
            csv = template.to_csv(index=False).encode('utf-8')
            st.download_button("Download Template CSV", csv, "fund_template.csv", "text/csv")

        # Data Editing Table
        if uploaded_file is not None:
            try:
                df = pd.read_csv(uploaded_file)
                st.session_state.portfolio_data['funds_df'] = df
                st.success("✅ File uploaded successfully! Please review/edit the data below before calculating.")
                
                # Allow manual editing
                st.markdown("**Review & Edit Fund Data:**")
                edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
                st.session_state.portfolio_data['funds_df'] = edited_df
                
                st.markdown("---")
                if st.button("🚀 Calculate Portfolio Analysis", use_container_width=True, type="primary"):
                    st.session_state.portfolio_data['goal_type'] = goal_type
                    st.session_state.portfolio_data['target_value'] = target_value
                    st.session_state.portfolio_data['target_growth'] = target_growth
                    st.session_state.portfolio_data['years'] = years
                    st.session_state.portfolio_data['initial_investment'] = initial_investment
                    st.session_state.portfolio_data['monthly_contribution'] = monthly_contribution
                    st.session_state.page = 'results'
                    st.rerun()
            except Exception as e:
                st.error(f"Error reading file. Please ensure it matches the template. Error: {e}")
        else:
            st.warning("Please upload a CSV file to proceed.")
            
        if st.button("⬅️ Back to Home"):
            st.session_state.page = 'home'
            st.rerun()

    elif st.session_state.page == 'results':
        st.header("📊 Step 2: Portfolio Analysis Results")
        data = st.session_state.portfolio_data
        df = data['funds_df']
        
        # Calculations
        weights, port_return, port_vol = optimize_portfolio(df)
        df['Allocation (%)'] = (weights * 100).round(1)
        
        if data['goal_type'] == "Reach a Target Sum ($)":
            required_cagr = calculate_required_cagr(
                data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years']
            )
            projected_value = data['initial_investment'] * ((1 + port_return/100) ** data['years'])
            if port_return > 0:
                projected_value += data['monthly_contribution'] * 12 * (((1 + port_return/100) ** data['years'] - 1) / (port_return/100))
            else:
                projected_value += data['monthly_contribution'] * 12 * data['years']
                
            goal_met = port_return >= required_cagr
        else:
            required_cagr = data['target_growth']
            projected_value = data['initial_investment'] * ((1 + port_return/100) ** data['years'])
            if port_return > 0:
                projected_value += data['monthly_contribution'] * 12 * (((1 + port_return/100) ** data['years'] - 1) / (port_return/100))
            else:
                projected_value += data['monthly_contribution'] * 12 * data['years']
            goal_met = port_return >= required_cagr

        # Metrics
        st.subheader("Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Projected Portfolio Return", f"{port_return:.1f}% p.a.")
        with col2:
            st.metric("Required Return for Goal", f"{required_cagr:.1f}% p.a.")
        with col3:
            st.metric("Goal Feasibility", "✅ Achievable" if goal_met else "⚠️ Shortfall", 
                      delta=f"{port_return - required_cagr:.1f}%" if goal_met else f"{port_return - required_cagr:.1f}%")
        with col4:
            st.metric("Projected Final Value", f"${int(projected_value):,}")

        st.markdown("---")
        
        # Charts
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Portfolio Allocation")
            fig1, ax1 = plt.subplots(figsize=(6, 6))
            ax1.pie(df['Allocation (%)'], labels=df['Fund Name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax1.set_title("Equal Weight Allocation")
            st.pyplot(fig1)
            
        with col2:
            st.subheader("Return vs Required")
            fig2, ax2 = plt.subplots(figsize=(6, 6))
            bars = ax2.bar(['Projected Return', 'Required Return'], [port_return, required_cagr], color=['#2ecc71' if goal_met else '#e74c3c', '#3498db'])
            ax2.set_ylabel("Annual Return (%)")
            ax2.set_title("Can the portfolio meet the goal?")
            for bar in bars:
                yval = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2, yval + 0.2, f"{yval:.1f}%", ha='center', va='bottom')
            st.pyplot(fig2)

        st.markdown("---")
        st.subheader("Fund Details")
        st.dataframe(df, use_container_width=True)

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📄 Generate Word Report", use_container_width=True):
                doc = Document()
                doc.add_heading('Portfolio Analysis Report', 0)
                doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
                doc.add_paragraph(f"Goal: {data['goal_type']}")
                doc.add_paragraph(f"Projected Return: {port_return:.1f}% | Required: {required_cagr:.1f}%")
                doc.add_heading('Fund Allocation', level=1)
                table = doc.add_table(rows=1, cols=4)
                hdr = table.rows[0].cells
                hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = 'Fund Name', 'Return (%)', 'Volatility (%)', 'Allocation (%)'
                for _, row in df.iterrows():
                    cells = table.add_row().cells
                    cells[0].text = str(row['Fund Name'])
                    cells[1].text = f"{row['Expected Return (%)']:.1f}"
                    cells[2].text = f"{row['Volatility (%)']:.1f}"
                    cells[3].text = f"{row['Allocation (%)']:.1f}"
                
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                st.download_button("📥 Download Word Report", buffer, file_name=f"Portfolio_Analysis_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        
        with col2:
            if st.button("📊 Generate Excel Report", use_container_width=True):
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Fund Analysis', index=False)
                    summary = pd.DataFrame({
                        'Metric': ['Initial Investment', 'Monthly Contribution', 'Years', 'Projected Value', 'Required CAGR', 'Projected CAGR'],
                        'Value': [data['initial_investment'], data['monthly_contribution'], data['years'], int(projected_value), f"{required_cagr:.1f}%", f"{port_return:.1f}%"]
                    })
                    summary.to_excel(writer, sheet_name='Summary', index=False)
                st.download_button(" Download Excel Report", excel_buffer.getvalue(), file_name=f"Portfolio_Data_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if st.button("⬅️ Back to Input"):
            st.session_state.page = 'input'
            st.rerun()