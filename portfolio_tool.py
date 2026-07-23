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
# SECTION: HELPER & PDF EXTRACTION FUNCTIONS
# ============================================================
def safe_float(v):
    try:
        return float(str(v).replace(',', '').strip())
    except:
        return np.nan

def generate_otp():
    return ''.join(random.choices(string.digits, k=6))

def extract_fund_data(text):
    """Extracts data from FFS text based on common patterns."""
    data = {
        'fund_name': 'Unknown Fund',
        '1y_return': np.nan,
        '3y_return': np.nan,
        '5y_return': np.nan,
        'volatility': np.nan,
        'management_fee': np.nan,
        'sector_allocation': {},
        'geographic_allocation': {},
        'ret_2016': np.nan, 'ret_2017': np.nan, 'ret_2018': np.nan, 'ret_2019': np.nan, 'ret_2020': np.nan
    }
    
    # 1. Fund Name (Look for title-like text in first 15 lines)
    lines = text.split('\n')
    for line in lines[:15]:
        line = line.strip()
        if 'fund' in line.lower() and 10 < len(line) < 60 and not line.startswith('Source'):
            data['fund_name'] = line
            break
            
    # 2. Returns (1Y, 3Y, 5Y)
    patterns_1y = [r'1[\s-]*[Yy]ear[:\s]*([\d\.]+)', r'1[\s-]*[Yy][:\s]*([\d\.]+)', r'1 Year[^\n]*?([\d\.]+)']
    patterns_3y = [r'3[\s-]*[Yy]ear[:\s]*([\d\.]+)', r'3[\s-]*[Yy][:\s]*([\d\.]+)', r'3 Year[^\n]*?([\d\.]+)']
    patterns_5y = [r'5[\s-]*[Yy]ear[:\s]*([\d\.]+)', r'5[\s-]*[Yy][:\s]*([\d\.]+)', r'5 Year[^\n]*?([\d\.]+)']
    
    for p in patterns_1y:
        m = re.search(p, text)
        if m: data['1y_return'] = safe_float(m.group(1)); break
    for p in patterns_3y:
        m = re.search(p, text)
        if m: data['3y_return'] = safe_float(m.group(1)); break
    for p in patterns_5y:
        m = re.search(p, text)
        if m: data['5y_return'] = safe_float(m.group(1)); break

    # 3. Volatility
    vol_patterns = [r'(?:Volatility Factor|VF|Volatility)[:\s]*(?:Up to |Max )?([\d\.]+)', r'Volatility[^\n]*?([\d\.]+)']
    for p in vol_patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m: data['volatility'] = safe_float(m.group(1)); break

    # 4. Management Fee
    fee_patterns = [r'(?:Annual )?Management Fee[:\s]*(?:Up to |Max )?([\d\.]+)', r'Management Fee[^\n]*?([\d\.]+)']
    for p in fee_patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m: data['management_fee'] = safe_float(m.group(1)); break

    # 5. Calendar Year Returns (2016-2020)
    for year in [2016, 2017, 2018, 2019, 2020]:
        # Look for the year followed by a number (return %)
        pattern = rf'{year}[^\n]*?(-?[\d\.]+)'
        m = re.search(pattern, text)
        if m:
            val = safe_float(m.group(1))
            if val is not np.nan and -50 < val < 100: # Sanity check
                data[f'ret_{year}'] = val

    # 6. Sector & Geographic Allocation (Simplified extraction)
    sectors = ['Technology', 'Financials', 'Healthcare', 'Consumer Discretionary', 'Consumer Staples', 'Industrials', 'Communication Services', 'Energy', 'Materials', 'Real Estate', 'Utilities', 'Manufacturing']
    countries = ['United States', 'Malaysia', 'China', 'Japan', 'Europe', 'United Kingdom', 'Taiwan', 'India', 'Australia', 'Hong Kong']
    
    for s in sectors:
        m = re.search(rf'{s}[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if m: data['sector_allocation'][s] = safe_float(m.group(1))
    for c in countries:
        m = re.search(rf'{c}[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if m: data['geographic_allocation'][c] =