import os
import io
import json
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
import openpyxl
import firebase_admin
from firebase_admin import credentials, firestore
from google import genai

# --- Page Configuration ---
st.set_page_config(
    page_title="Chew Advisory Portfolio Analyzer",
    page_icon="📈",
    layout="wide"
)

# --- Firebase Initialization ---
@st.cache_resource
def init_firebase():
    if not firebase_admin._apps:
        try:
            if "firebase" in st.secrets:
                cred_dict = dict(st.secrets["firebase"])
                cred = credentials.Certificate(cred_dict)
                firebase_admin.initialize_app(cred)
        except Exception as e:
            st.warning(f"Firebase initialization skipped or failed: {e}")

init_firebase()

# --- Gemini Client Initialization ---
@st.cache_resource
def get_gemini_client():
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key and "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    
    if api_key:
        return genai.Client(api_key=api_key)
    else:
        return genai.Client()

client = get_gemini_client()

# --- App Header ---
st.title("📈 Chew Advisory Portfolio Analyzer & Optimizer")
st.markdown("Welcome! Upload your Fund Fact Sheet PDFs or review your portfolio allocations below.")

# --- Sidebar Controls ---
st.sidebar.header("Configuration")
risk_profile = st.sidebar.selectbox("Select Risk Profile", ["Conservative", "Moderate", "Growth", "Aggressive"])
investment_horizon = st.sidebar.slider("Investment Horizon (Years)", 1, 30, 10)

# --- Main Tabs ---
tab1, tab2, tab3 = st.tabs(["📂 PDF Fact Sheet Processor", "📊 Portfolio Analysis & Optimization", "📑 Report Generator"])

with tab1:
    st.header("Fund Fact Sheet PDF Extraction")
    uploaded_pdfs = st.file_uploader(
        "Upload Fund Fact Sheet PDFs", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if uploaded_pdfs:
        st.info(f"Processing {len(uploaded_pdfs)} fund fact sheet(s) using Gemini...")
        
        extracted_results = []
        
        for pdf_file in uploaded_pdfs:
            try:
                pdf_bytes = pdf_file.read()
                
                prompt = (
                    "Extract the following information from this fund fact sheet into a clean JSON object: "
                    "fund_name, fund_category, nav, currency, benchmark, "
                    "asset_allocation (breakdown of equities, bonds, cash in percentages), "
                    "and top_holdings (list of top holdings)."
                )
                
                response = client.models.generate_content(
                    model='gemini-2.0-flash',
                    contents=[
                        {
                            "mime_type": "application/pdf",
                            "data": pdf_bytes
                        },
                        prompt
                    ]
                )
                
                text_output = response.text
                if "```json" in text_output:
                    text_output = text_output.split("```json")[1].split("```")[0].strip()
                elif "```" in text_output:
                    text_output = text_output.split("```")[1].split("```")[0].strip()
                
                fund_data = json.loads(text_output)
                fund_data["file_name"] = pdf_file.name
                extracted_results.append(fund_data)
                
                st.success(f"Successfully processed: {pdf_file.name}")
            except Exception as e:
                st.error(f"Error processing {pdf_file.name}: {e}")
        
        if extracted_results:
            st.subheader("Extracted Fund Summary")
            df_funds = pd.DataFrame(extracted_results)
            st.dataframe(df_funds)
            st.session_state["extracted_funds"] = df_funds

with tab2:
    st.header("Portfolio Asset Allocation & Simulation")
    
    if "extracted_funds" in st.session_state:
        st.write("Using data from processed fund fact sheets.")
        df = st.session_state["extracted_funds"]
        
        fig, ax = plt.subplots(figsize=(8, 4))
        categories = ['Equities', 'Fixed Income', 'Cash / Other']
        
        if risk_profile == "Conservative":
            allocations = [20, 70, 10]
        elif risk_profile == "Moderate":
            allocations = [50, 45, 5]
        elif risk_profile == "Growth":
            allocations = [75, 20, 5]
        else:
            allocations = [90, 5, 5]
            
        ax.bar(categories, allocations, color=['#1f77b4', '#ff7f0e', '#2ca02c'])
        ax.set_ylabel("Allocation Percentage (%)")
        ax.set_title(f"Recommended Asset Allocation for {risk_profile} Profile")
        st.pyplot(fig)
    else:
        st.info("Please upload and process PDF fact sheets in Tab 1 to view portfolio recommendations.")

with tab3:
    st.header("Professional Word Report Generator")
    client_name = st.text_input("Client Name", "Valued Client")
    
    if st.button("Generate Word Report"):
        doc = Document()
        doc.add_heading("Chew Advisory Wealth Portfolio Report", 0)
        doc.add_paragraph(f"Prepared for: {client_name}")
        doc.add_paragraph(f"Risk Profile: {risk_profile}")
        doc.add_paragraph(f"Investment Horizon: {investment_horizon} Years")
        
        doc.add_heading("Executive Summary", level=1)
        doc.add_paragraph("This report outlines your personalized asset allocation strategy designed to align with your long-term financial objectives.")
        
        if "extracted_funds" in st.session_state:
            doc.add_heading("Analyzed Funds", level=1)
            for _, row in st.session_state["extracted_funds"].iterrows():
                doc.add_paragraph(f"• {row.get('fund_name', row['file_name'])} ({row.get('fund_category', 'N/A')})")
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="Download Client Report (.docx)",
            data=buffer,
            file_name=f"Portfolio_Report_{client_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )