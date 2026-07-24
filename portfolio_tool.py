import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import json
import firebase_admin
from firebase_admin import credentials, firestore
import pdfplumber
import matplotlib.pyplot as plt
from docx import Document
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
for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'otp_code', 'otp_email', 'show_otp_input', 'ffs_files', 'ffs_data', 'edited_funds_df']:
    if key not in st.session_state:
        st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else [] if key in ['ffs_files', 'ffs_data'] else None if key == 'edited_funds_df' else {})

# ============================================================
# SECTION: HELPER FUNCTIONS
# ============================================================
def safe_float(v):
    try:
        if isinstance(v, str):
            v = v.replace(',', '').replace('%', '').strip()
            if '(' in v and ')' in v:
                v = '-' + v.replace('(', '').replace(')', '')
        return float(v)
    except:
        return np.nan

def generate_otp():
    return ''.join(random.choices(string.digits, k=6))

def clean_text(text):
    """Clean and normalize text for searching."""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_number_from_text(text, context=None):
    """Extract a number from text with context awareness."""
    if not text:
        return np.nan
    
    # If context is provided, look for numbers near the context
    if context:
        # Look for context followed by a number
        pattern = rf'{re.escape(context)}.*?([\d\.]+)'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = safe_float(match.group(1))
            if not np.isnan(val):
                return val
    
    # Look for any number in the text
    numbers = re.findall(r'([\d\.]+)', text)
    for num in numbers:
        val = safe_float(num)
        if not np.isnan(val) and 0 < val < 1000:
            return val
    
    return np.nan

def extract_fund_data(pdf_file):
    """
    Smart extraction using keyword-based search.
    This approach searches for specific keywords and extracts nearby data.
    """
    data = {
        'fund_name': 'Unknown Fund',
        '1y_return': np.nan,
        '3y_return': np.nan,
        '5y_return': np.nan,
        'volatility': np.nan,
        'management_fee': np.nan,
        'ret_2016': np.nan, 'ret_2017': np.nan, 'ret_2018': np.nan, 'ret_2019': np.nan, 'ret_2020': np.nan,
        'ret_2021': np.nan
    }
    
    try:
        if hasattr(pdf_file, 'seek'):
            pdf_file.seek(0)
        
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            # Extract text from all pages
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text += text + "\n"
            
            # Normalize text
            full_text = clean_text(full_text)
            
            # ============================================================
            # 1. EXTRACT FUND NAME
            # ============================================================
            lines = full_text.split('\n')
            fund_name_found = False
            
            # Look for fund name patterns in the first 30 lines
            for i, line in enumerate(lines[:30]):
                line = line.strip()
                if not line or len(line) < 5 or len(line) > 80:
                    continue
                
                # Common fund name indicators
                fund_indicators = ['fund', 'income', 'growth', 'equity', 'dividend', 'titans', 'asia', 
                                 'pacific', 'global', 'select', 'opportunity', 'small cap', 'mid cap',
                                 'large cap', 'balanced', 'dynamic', 'strategic', 'advantage', 'plus']
                
                # Skip lines that are clearly not fund names
                skip_patterns = ['source:', 'as at', 'page', 'disclaimer', 'growwithus', 'annual', 
                               'management', 'volatility', 'performance', 'distribution', 'benchmark',
                               'investment objective', 'risk', 'portfolio']
                
                if any(skip in line.lower() for skip in skip_patterns):
                    continue
                
                # Check if line contains fund indicators
                if any(indicator in line.lower() for indicator in fund_indicators):
                    # Clean up the line
                    fund_name = re.sub(r'\s+', ' ', line).strip()
                    # Remove common prefixes
                    fund_name = re.sub(r'^(An|A)\s+', '', fund_name)
                    # Remove trailing "Fund" if it exists to get cleaner name
                    if len(fund_name) > 15:
                        data['fund_name'] = fund_name
                        fund_name_found = True
                        break
            
            # If no fund name found, try a more aggressive approach
            if not fund_name_found:
                for line in lines[:20]:
                    line = line.strip()
                    if 'fund' in line.lower() and len(line) > 10 and len(line) < 60:
                        data['fund_name'] = clean_text(line)
                        break
            
            # ============================================================
            # 2. EXTRACT ANNUAL RETURNS
            # ============================================================
            # Look for "1 Year", "3 Year", "5 Year" patterns
            return_keywords = {
                '1y_return': ['1 year', '1yr', '1y', '1-yr', '1-year', '1 year annualised'],
                '3y_return': ['3 year', '3yr', '3y', '3-yr', '3-year', '3 year annualised'],
                '5y_return': ['5 year', '5yr', '5y', '5-yr', '5-year', '5 year annualised']
            }
            
            # Find the Annualised Return section
            annual_section = None
            if 'annualised' in full_text.lower():
                # Find the section containing "Annualised Return"
                lines_split = full_text.split('\n')
                for i, line in enumerate(lines_split):
                    if 'annualised' in line.lower() or 'annual return' in line.lower():
                        # Get a window of text around this section
                        start = max(0, i - 2)
                        end = min(len(lines_split), i + 20)
                        annual_section = '\n'.join(lines_split[start:end])
                        break
            
            if annual_section:
                # Search for returns in the annual section
                for key, keywords in return_keywords.items():
                    for keyword in keywords:
                        # Look for keyword followed by a number
                        pattern = rf'{re.escape(keyword)}.*?([\d\.]+)'
                        match = re.search(pattern, annual_section, re.IGNORECASE)
                        if match:
                            val = safe_float(match.group(1))
                            if 0 < val < 2000:
                                data[key] = val
                                break
            
            # If not found, search the entire text
            if np.isnan(data['1y_return']) or np.isnan(data['3y_return']) or np.isnan(data['5y_return']):
                for key, keywords in return_keywords.items():
                    if np.isnan(data[key]):
                        for keyword in keywords:
                            pattern = rf'{re.escape(keyword)}.*?([\d\.]+)'
                            match = re.search(pattern, full_text, re.IGNORECASE)
                            if match:
                                val = safe_float(match.group(1))
                                if 0 < val < 2000:
                                    data[key] = val
                                    break
            
            # ============================================================
            # 3. EXTRACT MANAGEMENT FEE
            # ============================================================
            fee_keywords = ['annual management fee', 'management fee', 'management fee (% p.a.)', 'annual fee']
            for keyword in fee_keywords:
                pattern = rf'{re.escape(keyword)}.*?([\d\.]+)'
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    val = safe_float(match.group(1))
                    if 0 < val < 10:
                        data['management_fee'] = val
                        break
            
            # ============================================================
            # 4. EXTRACT VOLATILITY
            # ============================================================
            vol_keywords = ['volatility factor', 'vf', 'volatility (vf)', 'volatility']
            for keyword in vol_keywords:
                pattern = rf'{re.escape(keyword)}.*?([\d\.]+)'
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    val = safe_float(match.group(1))
                    if 0 < val < 100:
                        data['volatility'] = val
                        break
            
            # ============================================================
            # 5. EXTRACT CALENDAR YEAR RETURNS
            # ============================================================
            years_to_find = [2021, 2020, 2019, 2018, 2017, 2016]
            
            # Find the Calendar Year Returns section
            calendar_section = None
            lines_split = full_text.split('\n')
            for i, line in enumerate(lines_split):
                if 'calendar year' in line.lower() or 'year to date' in line.lower() or 'calendar return' in line.lower():
                    start = max(0, i - 1)
                    end = min(len(lines_split), i + 30)
                    calendar_section = '\n'.join(lines_split[start:end])
                    break
            
            # If no specific section, use the whole text
            search_text = calendar_section if calendar_section else full_text
            
            for year in years_to_find:
                # Try multiple patterns
                patterns = [
                    rf'{year}.*?([\d\.]+)\s*%',
                    rf'{year}\s+([\d\.]+)',
                    rf'{year}\s*[\(\)]*([\d\.]+)',
                    rf'[\(\)]*{year}[\)]*\s+([\d\.]+)',
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, search_text)
                    if match:
                        val = safe_float(match.group(1))
                        if -100 < val < 200:
                            data[f'ret_{year}'] = val
                            break
                
                # If still not found, try looking for year in context of returns
                if np.isnan(data[f'ret_{year}']):
                    # Look for "Year 2021: 12.3%" style
                    pattern = rf'[Yy]ear\s+{year}[:\s]+([\d\.]+)'
                    match = re.search(pattern, full_text)
                    if match:
                        val = safe_float(match.group(1))
                        if -100 < val < 200:
                            data[f'ret_{year}'] = val
            
            # ============================================================
            # 6. FALLBACK: Try to find any numbers in the annualized section
            # ============================================================
            if np.isnan(data['1y_return']) and np.isnan(data['3y_return']) and np.isnan(data['5y_return']):
                # Look for a table-like structure with returns
                table_patterns = [
                    r'(\d+\.?\d*)\s*(\d+\.?\d*)\s*(\d+\.?\d*)\s*[0-9]',  # Three numbers in a row
                    r'Return.*?(\d+\.?\d*).*?(\d+\.?\d*).*?(\d+\.?\d*)',  # Return followed by numbers
                ]
                
                for pattern in table_patterns:
                    matches = re.findall(pattern, full_text)
                    if matches:
                        for match in matches:
                            vals = [safe_float(v) for v in match if not np.isnan(safe_float(v))]
                            if len(vals) >= 3:
                                # Assign the first three numbers as 1Y, 3Y, 5Y
                                if np.isnan(data['1y_return']):
                                    data['1y_return'] = vals[0]
                                if np.isnan(data['3y_return']):
                                    data['3y_return'] = vals[1]
                                if np.isnan(data['5y_return']):
                                    data['5y_return'] = vals[2]
                                break
            
    except Exception as e:
        st.warning(f"Error processing PDF: {str(e)}")
    
    return data

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

def optimize_portfolio_weights(df, target_return):
    n = len(df)
    if n == 0: return np.array([])
    
    returns = df['1y_return'].fillna(df['3y_return']).fillna(0).values
    sorted_indices = np.argsort(returns)[::-1]
    
    opt_weights = np.zeros(n)
    remaining_weight = 1.0
    
    for idx in sorted_indices:
        if remaining_weight <= 0.05:
            unallocated = np.where(opt_weights == 0)[0]
            if len(unallocated) > 0:
                opt_weights[unallocated] = remaining_weight / len(unallocated)
            break
        
        funds_left = n - np.sum(opt_weights > 0)
        max_allowed = min(0.40, remaining_weight - (0.05 * (funds_left - 1)))
        w = max(0.05, max_allowed)
        
        opt_weights[idx] = w
        remaining_weight -= w
        
    return opt_weights / opt_weights.sum()

def calculate_extra_contribution(target_sum, initial_investment, current_monthly, years, annual_return_pct):
    r = annual_return_pct / 100
    if r == 0:
        required_monthly = (target_sum - initial_investment) / (years * 12)
    else:
        required_monthly = (target_sum - initial_investment * ((1 + r) ** years)) * r / (((1 + r) ** years) - 1) / 12
    return max(0, required_monthly - current_monthly)

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
            for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'show_otp_input', 'otp_code', 'otp_email', 'ffs_files', 'ffs_data', 'edited_funds_df']:
                st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else [] if key in ['ffs_files', 'ffs_data'] else None if key == 'edited_funds_df' else {})
            st.rerun()

    if st.session_state.page == 'home':
        st.markdown("---")
        st.markdown("### Welcome to the Portfolio Analyzer")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("📝 Start New Analysis", use_container_width=True, type="primary"):
                st.session_state.page = 'input'
                st.session_state.portfolio_data = {}
                st.session_state.ffs_files = []
                st.session_state.ffs_data = []
                st.session_state.edited_funds_df = None
                st.rerun()
        with col2:
            if st.session_state.user_email == ADMIN_EMAIL: st.info("🔐 Admin: Unlimited access")

    elif st.session_state.page == 'input':
        st.header("📝 Step 1: Define Goal & Upload Fund Factsheets")
        
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

        st.subheader("3. Upload Fund Factsheets (FFS)")
        st.info("Upload PDF files. The tool will automatically search for and extract key data.")
        uploaded_ffs = st.file_uploader("Upload FFS (PDF files)", type=['pdf'], accept_multiple_files=True, key="ffs_uploader")
        
        if uploaded_ffs:
            st.session_state.ffs_files = uploaded_ffs
            with st.spinner("Processing FFS documents..."):
                st.session_state.ffs_data = []
                for f in uploaded_ffs:
                    f.seek(0)
                    extracted = extract_fund_data(f)
                    st.session_state.ffs_data.append(extracted)
                
                # Create DataFrame for editing
                cols = ['fund_name', '1y_return', '3y_return', '5y_return', 'volatility', 'management_fee', 
                        'ret_2016', 'ret_2017', 'ret_2018', 'ret_2019', 'ret_2020', 'ret_2021']
                st.session_state.edited_funds_df = pd.DataFrame(st.session_state.ffs_data)[cols]
            st.success(f"✅ Processed {len(uploaded_ffs)} FFS documents!")

        if st.session_state.edited_funds_df is not None and not st.session_state.edited_funds_df.empty:
            st.subheader("4. Review & Edit Extracted Data")
            st.warning("⚠️ Please review the extracted data below. **You can click on any cell to manually correct fund names, returns, or volatility if the PDF parser missed it.**")
            
            # Configure columns for data editor
            column_config = {
                "fund_name": st.column_config.TextColumn("Fund Name", width="large"),
                "1y_return": st.column_config.NumberColumn("1Y Return (%)", format="%.2f"),
                "3y_return": st.column_config.NumberColumn("3Y Return (%)", format="%.2f"),
                "5y_return": st.column_config.NumberColumn("5Y Return (%)", format="%.2f"),
                "volatility": st.column_config.NumberColumn("Volatility (%)", format="%.2f"),
                "management_fee": st.column_config.NumberColumn("Mgmt Fee (%)", format="%.2f"),
                "ret_2016": st.column_config.NumberColumn("2016 Return (%)", format="%.2f"),
                "ret_2017": st.column_config.NumberColumn("2017 Return (%)", format="%.2f"),
                "ret_2018": st.column_config.NumberColumn("2018 Return (%)", format="%.2f"),
                "ret_2019": st.column_config.NumberColumn("2019 Return (%)", format="%.2f"),
                "ret_2020": st.column_config.NumberColumn("2020 Return (%)", format="%.2f"),
                "ret_2021": st.column_config.NumberColumn("2021 Return (%)", format="%.2f"),
            }
            
            edited_df = st.data_editor(st.session_state.edited_funds_df, column_config=column_config, use_container_width=True, num_rows="dynamic")
            st.session_state.edited_funds_df = edited_df
            
            st.markdown("---")
            if st.button("Calculate Portfolio Analysis", use_container_width=True, type="primary"):
                st.session_state.portfolio_data = {
                    'goal_type': goal_type, 'target_value': target_value, 'target_growth': target_growth,
                    'years': years, 'initial_investment': initial_investment, 'monthly_contribution': monthly_contribution,
                    'risk_profile': risk_profile, 'funds_df': edited_df
                }
                increment_generation(st.session_state.user_email)
                st.session_state.page = 'analysis'
                st.rerun()
        else:
            st.warning("Please upload at least one FFS document to proceed.")
        
        if st.button("⬅️ Back to Home"):
            st.session_state.page = 'home'
            st.rerun()

    elif st.session_state.page == 'analysis':
        st.header("📊 Step 2: Portfolio Analysis")
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
        eq_return = np.sum(df['1y_return'].fillna(df['3y_return']).fillna(0) * equal_weights)
        eq_vol = np.sum(df['volatility'].fillna(0) * equal_weights)

        # 2. Optimized Portfolio
        opt_weights = optimize_portfolio_weights(df, target_return)
        opt_return = np.sum(df['1y_return'].fillna(df['3y_return']).fillna(0) * opt_weights)
        opt_vol = np.sum(df['volatility'].fillna(0) * opt_weights)

        # Display Metrics
        st.subheader("Portfolio Performance Analysis")
        col1, col2, col3 = st.columns(3)
        with col1: st.metric("Equal-Weighted Return", f"{eq_return:.1f}% p.a.")
        with col2: st.metric("Optimized Return", f"{opt_return:.1f}% p.a.")
        with col3: st.metric("Target Return", f"{target_return:.1f}% p.a.")

        # Feasibility & Risk
        st.subheader("Goal Feasibility & Risk Assessment")
        risk_thresholds = {"Conservative": 10.0, "Moderate": 15.0, "Growth": 20.0}
        risk_thresh = risk_thresholds[data['risk_profile']]
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Equal-Weighted Portfolio", "✅ Achievable" if eq_return >= target_return else "️ Shortfall", delta=f"{eq_return - target_return:.1f}%")
            st.caption(f"Volatility: {eq_vol:.1f}% ({'✅ Matches' if eq_vol <= risk_thresh else '⚠️ Exceeds'} {data['risk_profile']} profile)")
        with col2:
            st.metric("Optimized Portfolio", "✅ Achievable" if opt_return >= target_return else "⚠️ Shortfall", delta=f"{opt_return - target_return:.1f}%")
            st.caption(f"Volatility: {opt_vol:.1f}% ({'✅ Matches' if opt_vol <= risk_thresh else '⚠️ Exceeds'} {data['risk_profile']} profile)")

        # Shortfall Recommendation
        if opt_return < target_return:
            st.error("⚠️ **Target Not Met:** Even with the best possible mix of these funds, the maximum return is {:.1f}%. To reach your target of {:.1f}%, you need to increase your monthly contribution.".format(opt_return, target_return))
            extra_monthly = calculate_extra_contribution(data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years'], opt_return)
            st.info(f" **Recommendation:** Increase your monthly contribution by **${extra_monthly:,.0f}** (New total: ${data['monthly_contribution'] + extra_monthly:,.0f}/month) to reach your goal.")

        # Charts
        st.subheader("Portfolio Allocation")
        col1, col2 = st.columns(2)
        with col1:
            fig1, ax1 = plt.subplots(figsize=(6, 6))
            ax1.pie(equal_weights, labels=df['fund_name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax1.set_title("Equal-Weighted Allocation")
            st.pyplot(fig1)
        with col2:
            fig2, ax2 = plt.subplots(figsize=(6, 6))
            ax2.pie(opt_weights, labels=df['fund_name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax2.set_title("Optimized Allocation")
            st.pyplot(fig2)

        # Historical Performance Line Chart (Calendar Year Returns)
        st.subheader("Historical Performance (Calendar Year Returns)")
        years_list = [2016, 2017, 2018, 2019, 2020, 2021]
        fig_hist, ax_hist = plt.subplots(figsize=(12, 7))
        
        has_data = False
        for i, row in df.iterrows():
            returns = [row[f'ret_{y}'] for y in years_list]
            if not all(pd.isna(r) for r in returns):
                ax_hist.plot(years_list, returns, marker='o', linewidth=2, label=row['fund_name'][:25])
                has_data = True
        
        if has_data:
            # Calculate equal-weighted portfolio return for each year
            port_returns = []
            for y in years_list:
                col_returns = df[f'ret_{y}'].fillna(0)
                port_returns.append(np.sum(col_returns * equal_weights))
            ax_hist.plot(years_list, port_returns, marker='s', linewidth=3, label='Equal-Weighted Portfolio', color='black', linestyle='--')
            
            ax_hist.set_xlabel("Year", fontsize=12)
            ax_hist.set_ylabel("Return (%)", fontsize=12)
            ax_hist.set_title("Calendar Year Returns Comparison", fontsize=14, fontweight='bold')
            ax_hist.legend(loc='upper left', fontsize=9)
            ax_hist.grid(True, alpha=0.3)
            ax_hist.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
            st.pyplot(fig_hist)
        else:
            st.warning("No calendar year return data (2016-2021) was extracted from the FFS documents to plot the historical chart.")

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
        st.success("✅ Report generation is ready. The tool has analyzed your portfolio based on the uploaded FFS documents.")
        st.info("📥 You can download the analysis results as a Word document or Excel file.")
        
        # Generate Word report
        doc = Document()
        doc.add_heading('Portfolio Analysis Report', 0)
        doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
        doc.add_paragraph(f"Client: {st.session_state.user_email}")
        
        data = st.session_state.portfolio_data
        df = data['funds_df']
        
        # Add goal details
        doc.add_heading('Investment Goal', level=1)
        doc.add_paragraph(f"Goal Type: {data['goal_type']}")
        if data['goal_type'] == "Reach a Target Sum ($)":
            doc.add_paragraph(f"Target Amount: ${data['target_value']:,.0f}")
        else:
            doc.add_paragraph(f"Target Annual Growth: {data['target_growth']:.1f}%")
        doc.add_paragraph(f"Time Horizon: {data['years']} years")
        doc.add_paragraph(f"Risk Profile: {data['risk_profile']}")
        
        # Add capital details
        doc.add_heading('Capital & Contributions', level=1)
        doc.add_paragraph(f"Initial Investment: ${data['initial_investment']:,.0f}")
        doc.add_paragraph(f"Monthly Contribution: ${data['monthly_contribution']:,.0f}")
        
        # Add fund list
        doc.add_heading('Funds Analyzed', level=1)
        for _, row in df.iterrows():
            doc.add_paragraph(f"• {row['fund_name']}")
        
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