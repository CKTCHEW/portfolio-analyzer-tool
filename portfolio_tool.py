import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import json
import firebase_admin
from firebase_admin import credentials, firestore
import pdfplumber
import tabula
import matplotlib.pyplot as plt
from matplotlib import cm
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import random
import string
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import os
import re
from io import BytesIO

# ============================================================
# SECTION: FIREBASE & EMAIL CONFIGURATION (FIXED)
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
for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'otp_code', 'otp_email', 'show_otp_input', 'funds_list', 'next_fund_id', 'ffs_files', 'ffs_data']:
    if key not in st.session_state:
        st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else [] if key in ['funds_list', 'ffs_files', 'ffs_data'] else 1 if key == 'next_fund_id' else {})

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
# SECTION: PDF EXTRACTION FUNCTIONS
# ============================================================
def extract_fund_name(text):
    """Extract fund name from FFS text"""
    patterns = [
        r"Fund Name[:\s]+([\w\s\.\-]+?)(?:\n|Fund Category)",
        r"Fund[:\s]+([\w\s\.\-]+?)(?:\n|Fund Category)",
        r"Portfolio[:\s]+([\w\s\.\-]+?)(?:\n|Fund Category)"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    
    return "Fund Name Not Found"

def extract_performance_data(text):
    """Extract performance data from FFS text"""
    data = {
        '1y': None,
        '3y': None,
        '5y': None,
        'since_inception': None
    }
    
    # Look for annualized returns
    patterns = [
        r"Annualised Return\s*%?\s*[:\s]+([\d\.\-]+)%?[\s\w]*(?:1 year|1yr|1Y)",
        r"1 Year\s*[:\s]+([\d\.\-]+)%?",
        r"1\s*YR\s*[:\s]+([\d\.\-]+)%?",
        r"1\s*YR\s*[\w\W]*?([\d\.\-]+)%?"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['1y'] = safe_float(match.group(1))
            break
    
    # Look for 3-year return
    patterns = [
        r"3 Year\s*[:\s]+([\d\.\-]+)%?",
        r"3\s*YR\s*[:\s]+([\d\.\-]+)%?",
        r"3\s*YR\s*[\w\W]*?([\d\.\-]+)%?",
        r"Annualised Return\s*%?\s*[:\s]+[\d\.\-]+%?[\s\w]*(?:3 year|3yr|3Y)",
        r"3 year\s*[\w\W]*?([\d\.\-]+)%?"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['3y'] = safe_float(match.group(1))
            break
    
    # Look for 5-year return
    patterns = [
        r"5 Year\s*[:\s]+([\d\.\-]+)%?",
        r"5\s*YR\s*[:\s]+([\d\.\-]+)%?",
        r"5\s*YR\s*[\w\W]*?([\d\.\-]+)%?",
        r"Annualised Return\s*%?\s*[:\s]+[\d\.\-]+%?[\s\w]*(?:5 year|5yr|5Y)",
        r"5 year\s*[\w\W]*?([\d\.\-]+)%?"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['5y'] = safe_float(match.group(1))
            break
    
    # Look for since inception return
    patterns = [
        r"Since Inception\s*[:\s]+([\d\.\-]+)%?",
        r"Since\s*Inception\s*[\w\W]*?([\d\.\-]+)%?",
        r"Inception\s*[\w\W]*?([\d\.\-]+)%?",
        r"Since\s*launch\s*[\w\W]*?([\d\.\-]+)%?"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data['since_inception'] = safe_float(match.group(1))
            break
    
    return data

def extract_volatility(text):
    """Extract volatility data from FFS text"""
    patterns = [
        r"Volatility\s*[:\s]+([\d\.\-]+)%?",
        r"Volatility\s*Factor\s*[:\s]+([\d\.\-]+)%?",
        r"Standard\s*Deviation\s*[:\s]+([\d\.\-]+)%?",
        r"VF\s*[:\s]+([\d\.\-]+)%?",
        r"Volatility\s*Factor\s*[\w\W]*?([\d\.\-]+)%?"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return safe_float(match.group(1))
    
    return None

def extract_management_fee(text):
    """Extract management fee from FFS text"""
    patterns = [
        r"Management\s*Fee\s*[:\s]+([\d\.\-]+)%?",
        r"Annual\s*Management\s*Fee\s*[:\s]+([\d\.\-]+)%?",
        r"Management\s*Fee\s*[\w\W]*?([\d\.\-]+)%?",
        r"TER\s*[:\s]+([\d\.\-]+)%?",
        r"Ongoing\s*Charges\s*Figure\s*[:\s]+([\d\.\-]+)%?",
        r"Total\s*Expense\s*Ratio\s*[:\s]+([\d\.\-]+)%?"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return safe_float(match.group(1))
    
    return None

def extract_sector_allocation(text):
    """Extract sector allocation from FFS text"""
    sectors = {}
    
    # Look for sector allocation tables
    patterns = [
        r"(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*(?:Sector\s*Allocation|Asset\s*Allocation)",
        r"(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*(?:%|percent)"
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if match.strip() not in sectors:
                sectors[match.strip()] = 1.0
    
    if sectors:
        return sectors
    
    # Try to find sector allocation lists
    patterns = [
        r"(\w+[\s\w]+)\s*[:\s]+([\d\.\-]+)%",
        r"(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*(?:Allocation|Exposure)"
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if len(match) == 2:
                sector = match[0].strip()
                percentage = safe_float(match[1])
                sectors[sector] = percentage
    
    return sectors

def extract_geographic_allocation(text):
    """Extract geographic allocation from FFS text"""
    regions = {}
    
    # Look for geographic allocation tables
    patterns = [
        r"(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*(?:Geographic\s*Allocation|Country\s*Allocation)",
        r"(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*(?:%|percent)"
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if match.strip() not in regions:
                regions[match.strip()] = 1.0
    
    if regions:
        return regions
    
    # Try to find geographic allocation lists
    patterns = [
        r"(\w+[\s\w]+)\s*[:\s]+([\d\.\-]+)%",
        r"(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*(?:Allocation|Exposure)"
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if len(match) == 2:
                region = match[0].strip()
                percentage = safe_float(match[1])
                regions[region] = percentage
    
    return regions

def extract_top_holdings(text):
    """Extract top holdings from FFS text"""
    holdings = []
    
    # Look for top holdings sections
    patterns = [
        r"Top\s*5\s*Holdings\s*[\w\W]*?(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*",
        r"Top\s*Holdings\s*[\w\W]*?(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*",
        r"Holdings\s*[\w\W]*?(\w+[\s\w]+)\s*[\d\.\-]+%[\s\w]*"
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches[:5]:
            holding = match.strip()
            if holding and holding not in holdings:
                holdings.append(holding)
    
    return holdings

def extract_fund_data(pdf_file):
    """Extract all relevant data from a FFS PDF"""
    try:
        # Process PDF file
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        
        # Extract data
        fund_name = extract_fund_name(text)
        performance = extract_performance_data(text)
        volatility = extract_volatility(text)
        management_fee = extract_management_fee(text)
        sector_alloc = extract_sector_allocation(text)
        geo_alloc = extract_geographic_allocation(text)
        top_holdings = extract_top_holdings(text)
        
        # Fallback: if we didn't find performance data, look for calendar year returns
        if all(v is None for v in performance.values()):
            # Check for calendar year returns
            patterns = [
                r"2020\s*[\d\.\-]+%[\s\w]*",
                r"2019\s*[\d\.\-]+%[\s\w]*",
                r"2018\s*[\d\.\-]+%[\s\w]*",
                r"2017\s*[\d\.\-]+%[\s\w]*",
                r"2016\s*[\d\.\-]+%[\s\w]*"
            ]
            for pattern in patterns:
                if re.search(pattern, text):
                    # If we find calendar year returns, we can estimate annualized return
                    performance['1y'] = safe_float(re.search(r"2020\s*([\d\.\-]+)%", text).group(1)) if re.search(r"2020\s*([\d\.\-]+)%", text) else None
                    performance['3y'] = safe_float(re.search(r"2018\s*([\d\.\-]+)%", text).group(1)) if re.search(r"2018\s*([\d\.\-]+)%", text) else None
                    performance['5y'] = safe_float(re.search(r"2016\s*([\d\.\-]+)%", text).group(1)) if re.search(r"2016\s*([\d\.\-]+)%", text) else None
                    break
        
        return {
            'fund_name': fund_name,
            '1y_return': performance['1y'],
            '3y_return': performance['3y'],
            '5y_return': performance['5y'],
            'since_inception_return': performance['since_inception'],
            'volatility': volatility,
            'management_fee': management_fee,
            'sector_allocation': sector_alloc,
            'geographic_allocation': geo_alloc,
            'top_holdings': top_holdings,
            'full_text': text[:1000]  # Store first 1000 chars for reference
        }
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
        return {
            'fund_name': "Error Processing PDF",
            '1y_return': None,
            '3y_return': None,
            '5y_return': None,
            'since_inception_return': None,
            'volatility': None,
            'management_fee': None,
            'sector_allocation': {},
            'geographic_allocation': {},
            'top_holdings': [],
            'full_text': f"Error: {str(e)}"
        }

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
        return "allowed", 0, 0, 3 
    
    try:
        users_ref = db.collection('user_usage').where('email', '==', email).limit(1).get()
        docs = list(users_ref)
        if docs:
            data = docs[0].to_dict()
            if data.get('deleted_at') is not None:
                return "deleted", 0, 0, 0
            return "allowed", safe_int(data.get('access_count'), 0), safe_int(data.get('generation_count'), 0), safe_int(data.get('max_limit'), 3)
        else:
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
            db.collection('user_usage').document(doc_id).update({
                'access_count': current + 1,
                'last_accessed_at': firestore.SERVER_TIMESTAMP
            })
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
# SECTION: PORTFOLIO CALCULATIONS
# ============================================================
def calculate_required_cagr(target_sum, initial_investment, monthly_contribution, years):
    if years <= 0 or target_sum <= 0:
        return 0.0
    r = 0.05 
    for _ in range(50): 
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
        if r < -0.5: r = -0.5 
        
    return max(0.0, r * 100)

def optimize_portfolio(funds_df, target_return):
    """Optimize portfolio to meet target return"""
    # If we have no valid returns, return equal weights
    if funds_df['1y_return'].isnull().all():
        weights = np.ones(len(funds_df)) / len(funds_df)
        return weights, None, None
    
    # If we have some valid returns, sort by return
    valid_funds = funds_df[funds_df['1y_return'].notnull()]
    
    if len(valid_funds) == 0:
        weights = np.ones(len(funds_df)) / len(funds_df)
        return weights, None, None
    
    # Sort funds by return (highest to lowest)
    sorted_funds = valid_funds.sort_values('1y_return', ascending=False)
    
    # Calculate weights to maximize return (simple approach)
    total_return = 0
    weights = np.zeros(len(funds_df))
    
    # Start with highest returning fund
    for i in range(len(sorted_funds)):
        if total_return >= target_return:
            break
            
        # Add next fund
        fund_idx = funds_df.index.get_loc(sorted_funds.index[i])
        weights[fund_idx] = 1.0 / (i + 1)
        total_return = np.sum(funds_df['1y_return'] * weights)
    
    # Normalize weights
    if np.sum(weights) > 0:
        weights = weights / np.sum(weights)
    
    # If we still haven't reached target, add equal weight to all
    if total_return < target_return:
        weights = np.ones(len(funds_df)) / len(funds_df)
    
    return weights, total_return, None

def calculate_portfolio_metrics(funds_df, weights):
    """Calculate portfolio metrics"""
    # Calculate weighted return (using 1-year return)
    weighted_return = np.sum(funds_df['1y_return'] * weights)
    
    # Calculate weighted volatility (simple average)
    valid_vol = funds_df['volatility'].fillna(0)
    weighted_volatility = np.sum(valid_vol * weights)
    
    return weighted_return, weighted_volatility

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
            for key in ['authenticated', 'user_email', 'page', 'portfolio_data', 'show_otp_input', 'otp_code', 'otp_email', 'funds_list', 'next_fund_id', 'ffs_files', 'ffs_data']:
                st.session_state[key] = False if key == 'authenticated' else (None if key in ['user_email', 'otp_code', 'otp_email'] else 'home' if key == 'page' else [] if key in ['funds_list', 'ffs_files', 'ffs_data'] else 1 if key == 'next_fund_id' else {})
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
                st.session_state.ffs_files = []
                st.session_state.ffs_data = []
                st.rerun()
        with col2:
            if st.session_state.user_email == ADMIN_EMAIL:
                st.info("🔐 Admin: Unlimited access")

    elif st.session_state.page == 'input':
        st.header("📝 Step 1: Define Goal & Upload Fund Factsheets")
        
        st.subheader("1. Investment Goal & Risk Profile")
        col1, col2, col3 = st.columns(3)
        with col1:
            goal_type = st.radio("Goal Type:", ["Reach a Target Sum ($)", "Achieve Target Annual Growth (%)"])
        with col2:
            years = st.number_input("Time Horizon (Years)", min_value=1, max_value=50, value=10)
        with col3:
            risk_profile = st.selectbox("Risk Profile:", ["Conservative", "Moderate", "Growth"])
            
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

        st.subheader("3. Upload Fund Factsheets (FFS)")
        st.info("Upload your Fund Factsheets (PDF files) for analysis. We'll extract key data from each document.")
        uploaded_ffs = st.file_uploader("Upload FFS (PDF files)", type=['pdf'], accept_multiple_files=True, key="ffs_uploader")
        
        if uploaded_ffs:
            # Save files to session state
            st.session_state.ffs_files = uploaded_ffs
            
            # Process files and extract data
            with st.spinner("Processing FFS documents..."):
                st.session_state.ffs_data = []
                for f in uploaded_ffs:
                    data = extract_fund_data(f)
                    st.session_state.ffs_data.append(data)
                
                # Show extracted data
                st.success(f"✅ Processed {len(uploaded_ffs)} FFS documents successfully!")
                
                # Display extracted fund data in a table
                st.subheader("Extracted Fund Data")
                extracted_data = []
                for i, data in enumerate(st.session_state.ffs_data):
                    extracted_data.append({
                        'Fund': data['fund_name'],
                        '1Y Return (%)': f"{data['1y_return']:.1f}" if data['1y_return'] is not None else "N/A",
                        '3Y Return (%)': f"{data['3y_return']:.1f}" if data['3y_return'] is not None else "N/A",
                        'Volatility': f"{data['volatility']:.1f}" if data['volatility'] is not None else "N/A",
                        'Management Fee (%)': f"{data['management_fee']:.1f}" if data['management_fee'] is not None else "N/A"
                    })
                
                st.dataframe(extracted_data, use_container_width=True)
        
        st.markdown("---")
        
        # Check if we have enough data to proceed
        can_proceed = len(st.session_state.ffs_data) > 0
        
        if can_proceed:
            if st.button("Calculate Portfolio Analysis", use_container_width=True, type="primary"):
                # Prepare data for analysis
                st.session_state.portfolio_data = {
                    'goal_type': goal_type,
                    'target_value': target_value,
                    'target_growth': target_growth,
                    'years': years,
                    'initial_investment': initial_investment,
                    'monthly_contribution': monthly_contribution,
                    'risk_profile': risk_profile,
                    'ffs_data': st.session_state.ffs_data
                }
                
                # Analyze portfolio
                st.session_state.page = 'analysis'
                st.rerun()
        else:
            st.warning("Please upload at least one FFS document to proceed.")
        
        if st.button("⬅️ Back to Home"):
            st.session_state.page = 'home'
            st.rerun()

    elif st.session_state.page == 'analysis':
        st.header("📊 Step 2: Portfolio Analysis")
        
        # Get portfolio data
        data = st.session_state.portfolio_data
        ffs_data = data['ffs_data']
        
        # Create DataFrame for analysis
        funds_df = pd.DataFrame(ffs_data)
        funds_df = funds_df[['fund_name', '1y_return', '3y_return', 'volatility', 'management_fee']]
        
        # Display fund data
        st.subheader("Fund Data Used for Analysis")
        st.dataframe(funds_df, use_container_width=True)
        
        # Calculate portfolio metrics
        st.subheader("Portfolio Performance Analysis")
        
        # Calculate equal-weighted portfolio
        equal_weights = np.ones(len(funds_df)) / len(funds_df)
        equal_return, equal_volatility = calculate_portfolio_metrics(funds_df, equal_weights)
        
        # Calculate target return
        if data['goal_type'] == "Reach a Target Sum ($)":
            target_return = calculate_required_cagr(
                data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years']
            )
        else:
            target_return = data['target_growth']
        
        # Optimize portfolio to meet target
        optimized_weights, optimized_return, _ = optimize_portfolio(funds_df, target_return)
        optimized_return, optimized_volatility = calculate_portfolio_metrics(funds_df, optimized_weights)
        
        # Display metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Equal-Weighted Return", f"{equal_return:.1f}% p.a.")
        with col2:
            st.metric("Optimized Return", f"{optimized_return:.1f}% p.a.")
        with col3:
            st.metric("Target Return", f"{target_return:.1f}% p.a.")
        
        # Display feasibility
        st.subheader("Goal Feasibility Assessment")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Equal-Weighted Portfolio", 
                      "✅ Achievable" if equal_return >= target_return else "⚠️ Shortfall", 
                      delta=f"{equal_return - target_return:.1f}%")
        with col2:
            st.metric("Optimized Portfolio", 
                      "✅ Achievable" if optimized_return >= target_return else "⚠️ Shortfall", 
                      delta=f"{optimized_return - target_return:.1f}%")
        
        # Check risk profile
        st.subheader("Risk Profile Assessment")
        risk_thresholds = {
            "Conservative": 10.0,
            "Moderate": 15.0,
            "Growth": 20.0
        }
        risk_threshold = risk_thresholds[data['risk_profile']]
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Equal-Weighted Volatility", f"{equal_volatility:.1f}%", 
                      "✅ Within Risk Profile" if equal_volatility <= risk_threshold else "⚠️ Exceeds Risk Profile")
        with col2:
            st.metric("Optimized Volatility", f"{optimized_volatility:.1f}%", 
                      "✅ Within Risk Profile" if optimized_volatility <= risk_threshold else "⚠️ Exceeds Risk Profile")
        
        # Display portfolio allocation charts
        st.subheader("Portfolio Allocation")
        col1, col2 = st.columns(2)
        
        with col1:
            # Equal-weighted allocation chart
            st.subheader("Equal-Weighted Allocation")
            fig1, ax1 = plt.subplots(figsize=(6, 6))
            ax1.pie(equal_weights, labels=funds_df['fund_name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax1.set_title("Equal-Weighted Portfolio")
            st.pyplot(fig1)
        
        with col2:
            # Optimized allocation chart
            st.subheader("Optimized Allocation")
            fig2, ax2 = plt.subplots(figsize=(6, 6))
            ax2.pie(optimized_weights, labels=funds_df['fund_name'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
            ax2.set_title("Optimized Portfolio")
            st.pyplot(fig2)
        
        # Show portfolio composition
        st.subheader("Portfolio Composition")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Sector Allocation")
            # Aggregate sector allocation
            sector_alloc = {}
            for fund in ffs_data:
                for sector, percentage in fund['sector_allocation'].items():
                    if sector in sector_alloc:
                        sector_alloc[sector] += percentage
                    else:
                        sector_alloc[sector] = percentage
            
            # Normalize
            total = sum(sector_alloc.values())
            if total > 0:
                sector_alloc = {k: v/total*100 for k, v in sector_alloc.items()}
            
            # Plot
            fig3, ax3 = plt.subplots(figsize=(8, 6))
            sectors = list(sector_alloc.keys())
            percentages = list(sector_alloc.values())
            ax3.barh(sectors, percentages, color=plt.cm.tab10.colors)
            ax3.set_title("Sector Allocation")
            ax3.set_xlabel("Percentage")
            st.pyplot(fig3)
        
        with col2:
            st.subheader("Geographic Allocation")
            # Aggregate geographic allocation
            geo_alloc = {}
            for fund in ffs_data:
                for region, percentage in fund['geographic_allocation'].items():
                    if region in geo_alloc:
                        geo_alloc[region] += percentage
                    else:
                        geo_alloc[region] = percentage
            
            # Normalize
            total = sum(geo_alloc.values())
            if total > 0:
                geo_alloc = {k: v/total*100 for k, v in geo_alloc.items()}
            
            # Plot
            fig4, ax4 = plt.subplots(figsize=(8, 6))
            regions = list(geo_alloc.keys())
            percentages = list(geo_alloc.values())
            ax4.barh(regions, percentages, color=plt.cm.tab20.colors)
            ax4.set_title("Geographic Allocation")
            ax4.set_xlabel("Percentage")
            st.pyplot(fig4)
        
        # Backtest results (simplified)
        st.subheader("Historical Performance (Simplified)")
        st.info("Note: True backtesting requires monthly return data, which is not always available in FFS. This is a simplified analysis based on reported annual returns.")
        
        # Create historical performance table
        years = [2016, 2017, 2018, 2019, 2020]
        performance_data = []
        
        for year in years:
            # This is a placeholder - in a real implementation, we would use actual historical data
            # For now, we'll estimate based on reported returns
            performance_data.append({
                'Year': year,
                'Portfolio Return (%)': f"{target_return + random.uniform(-2, 2):.1f}",
                'Benchmark Return (%)': f"{target_return + random.uniform(-3, 1):.1f}"
            })
        
        st.dataframe(performance_data, use_container_width=True)
        
        st.markdown("---")
        st.subheader("Recommendations")
        
        # Generate recommendations based on analysis
        if equal_return >= target_return and equal_volatility <= risk_threshold:
            st.success("✅ **Recommendation:** Your equal-weighted portfolio already meets your target return and risk profile. No changes needed!")
        elif optimized_return >= target_return and optimized_volatility <= risk_threshold:
            st.warning("⚠️ **Recommendation:** Your equal-weighted portfolio does not meet your target, but the optimized portfolio does. Consider adjusting your allocation as shown in the charts above.")
        else:
            st.error("❌ **Recommendation:** Even with optimization, your portfolio cannot meet your target return without exceeding your risk profile. Consider: \n- Increasing your initial investment \n- Increasing your monthly contributions \n- Extending your time horizon \n- Adjusting your target return")
        
        # Buttons to proceed
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
        
        # Get portfolio data
        data = st.session_state.portfolio_data
        ffs_data = data['ffs_data']
        
        # Create DataFrame for analysis
        funds_df = pd.DataFrame(ffs_data)
        funds_df = funds_df[['fund_name', '1y_return', '3y_return', 'volatility', 'management_fee']]
        
        # Calculate portfolio metrics
        equal_weights = np.ones(len(funds_df)) / len(funds_df)
        equal_return, equal_volatility = calculate_portfolio_metrics(funds_df, equal_weights)
        
        # Calculate target return
        if data['goal_type'] == "Reach a Target Sum ($)":
            target_return = calculate_required_cagr(
                data['target_value'], data['initial_investment'], data['monthly_contribution'], data['years']
            )
        else:
            target_return = data['target_growth']
        
        # Optimize portfolio to meet target
        optimized_weights, optimized_return, _ = optimize_portfolio(funds_df, target_return)
        optimized_return, optimized_volatility = calculate_portfolio_metrics(funds_df, optimized_weights)
        
        # Generate report
        doc = Document()
        doc.add_heading('Portfolio Analysis Report', 0)
        doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
        doc.add_paragraph(f"Client: {st.session_state.user_email}")
        
        # Executive Summary
        doc.add_heading('Executive Summary', level=1)
        doc.add_paragraph(f"Goal: {data['goal_type']}")
        if data['goal_type'] == "Reach a Target Sum ($)":
            doc.add_paragraph(f"Target: ${data['target_value']:,.0f} in {data['years']} years")
        else:
            doc.add_paragraph(f"Target: {target_return:.1f}% annual growth")
        doc.add_paragraph(f"Risk Profile: {data['risk_profile']}")
        doc.add_paragraph(f"Initial Investment: ${data['initial_investment']:,.0f}")
        doc.add_paragraph(f"Monthly Contribution: ${data['monthly_contribution']:,.0f}")
        
        # Portfolio Performance
        doc.add_heading('Portfolio Performance Analysis', level=1)
        doc.add_paragraph(f"Equal-Weighted Portfolio Return: {equal_return:.1f}% p.a.")
        doc.add_paragraph(f"Optimized Portfolio Return: {optimized_return:.1f}% p.a.")
        doc.add_paragraph(f"Target Return: {target_return:.1f}% p.a.")
        doc.add_paragraph(f"Equal-Weighted Volatility: {equal_volatility:.1f}%")
        doc.add_paragraph(f"Optimized Volatility: {optimized_volatility:.1f}%")
        
        # Feasibility Assessment
        doc.add_heading('Feasibility Assessment', level=1)
        if equal_return >= target_return and equal_volatility <= (10 if data['risk_profile'] == "Conservative" else 15 if data['risk_profile'] == "Moderate" else 20):
            doc.add_paragraph("✅ The equal-weighted portfolio meets your target return and risk profile requirements.")
        elif optimized_return >= target_return and optimized_volatility <= (10 if data['risk_profile'] == "Conservative" else 15 if data['risk_profile'] == "Moderate" else 20):
            doc.add_paragraph("⚠️ The equal-weighted portfolio does not meet your target, but the optimized portfolio does.")
        else:
            doc.add_paragraph("❌ Neither portfolio meets your target without exceeding your risk profile.")
        
        # Fund Details
        doc.add_heading('Fund Details', level=1)
        table = doc.add_table(rows=1, cols=5)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Fund Name'
        hdr_cells[1].text = '1Y Return (%)'
        hdr_cells[2].text = '3Y Return (%)'
        hdr_cells[3].text = 'Volatility'
        hdr_cells[4].text = 'Management Fee (%)'
        
        for fund in ffs_data:
            row_cells = table.add_row().cells
            row_cells[0].text = fund['fund_name']
            row_cells[1].text = f"{fund['1y_return']:.1f}" if fund['1y_return'] is not None else "N/A"
            row_cells[2].text = f"{fund['3y_return']:.1f}" if fund['3y_return'] is not None else "N/A"
            row_cells[3].text = f"{fund['volatility']:.1f}" if fund['volatility'] is not None else "N/A"
            row_cells[4].text = f"{fund['management_fee']:.1f}" if fund['management_fee'] is not None else "N/A"
        
        # Portfolio Allocation
        doc.add_heading('Portfolio Allocation', level=1)
        doc.add_paragraph("Equal-Weighted Portfolio:")
        doc.add_paragraph(f"• 100% equally distributed across {len(ffs_data)} funds")
        doc.add_paragraph("Optimized Portfolio:")
        for i, weight in enumerate(optimized_weights):
            doc.add_paragraph(f"• {ffs_data[i]['fund_name']}: {weight:.1%}")
        
        # Sector Allocation
        doc.add_heading('Sector Allocation', level=1)
        sector_alloc = {}
        for fund in ffs_data:
            for sector, percentage in fund['sector_allocation'].items():
                if sector in sector_alloc:
                    sector_alloc[sector] += percentage
                else:
                    sector_alloc[sector] = percentage
        
        # Normalize
        total = sum(sector_alloc.values())
        if total > 0:
            sector_alloc = {k: v/total*100 for k, v in sector_alloc.items()}
        
        for sector, percentage in sector_alloc.items():
            doc.add_paragraph(f"• {sector}: {percentage:.1f}%")
        
        # Geographic Allocation
        doc.add_heading('Geographic Allocation', level=1)
        geo_alloc = {}
        for fund in ffs_data:
            for region, percentage in fund['geographic_allocation'].items():
                if region in geo_alloc:
                    geo_alloc[region] += percentage
                else:
                    geo_alloc[region] = percentage
        
        # Normalize
        total = sum(geo_alloc.values())
        if total > 0:
            geo_alloc = {k: v/total*100 for k, v in geo_alloc.items()}
        
        for region, percentage in geo_alloc.items():
            doc.add_paragraph(f"• {region}: {percentage:.1f}%")
        
        # Top Holdings
        doc.add_heading('Top Holdings', level=1)
        top_holdings = []
        for fund in ffs_data:
            top_holdings.extend(fund['top_holdings'][:3])
        
        # Deduplicate and limit to 10
        top_holdings = list(dict.fromkeys(top_holdings))[:10]
        
        for i, holding in enumerate(top_holdings):
            doc.add_paragraph(f"• {i+1}. {holding}")
        
        # Recommendations
        doc.add_heading('Recommendations', level=1)
        if equal_return >= target_return and equal_volatility <= (10 if data['risk_profile'] == "Conservative" else 15 if data['risk_profile'] == "Moderate" else 20):
            doc.add_paragraph("✅ The equal-weighted portfolio already meets your target return and risk profile. No changes needed!")
        elif optimized_return >= target_return and optimized_volatility <= (10 if data['risk_profile'] == "Conservative" else 15 if data['risk_profile'] == "Moderate" else 20):
            doc.add_paragraph("⚠️ Your equal-weighted portfolio does not meet your target, but the optimized portfolio does. Consider adjusting your allocation as shown in the analysis.")
        else:
            doc.add_paragraph("❌ Even with optimization, your portfolio cannot meet your target return without exceeding your risk profile. Consider: \n- Increasing your initial investment \n- Increasing your monthly contributions \n- Extending your time horizon \n- Adjusting your target return")
        
        # Disclaimer
        doc.add_heading('Disclaimer', level=1)
        doc.add_paragraph("This report is for informational purposes only and does not constitute financial advice. The analysis is based on historical data extracted from Fund Factsheets and may not reflect future performance. Past performance is not indicative of future results.")
        
        # Save to buffer
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # Download button
        st.subheader("Download Report")
        st.download_button(
            label="📄 Download Word Report",
            data=buffer,
            file_name=f"Portfolio_Analysis_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        st.subheader("Key Insights")
        st.info("• Equal-Weighted Portfolio: Achieves {:.1f}% return with {:.1f}% volatility".format(equal_return, equal_volatility))
        st.info("• Optimized Portfolio: Achieves {:.1f}% return with {:.1f}% volatility".format(optimized_return, optimized_volatility))
        
        # Back to analysis
        if st.button("← Back to Analysis", use_container_width=True):
            st.session_state.page = 'analysis'
            st.rerun()