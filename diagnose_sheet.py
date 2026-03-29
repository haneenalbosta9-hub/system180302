"""
Google Sheets Header Diagnostic Tool
This script checks if your Google Sheet headers match the expected columns.
"""

import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SPREADSHEET_ID = "1EXiXsOQ0VsfIbZlUpN3r6g0-aRNUUEKZDVZHh_xZnEY"
SHEET_SAMPLES = "Samples"

EXPECTED_COLUMNS = [
    "Received Date", "Sample ID", "Unit No.", "Sample Type", "Sample Batch No.",
    "Customer Name", "Reference No.", "Type of Test",
    "Test Performing Date", "Test Status", "Product Name",
    "Customer Name (AR)", "Customer Name (EN)"
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def main():
    st.title("🔍 Google Sheets Diagnostic Tool")
    
    try:
        # Connect to Google Sheets
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(SPREADSHEET_ID)
        ws = ss.worksheet(SHEET_SAMPLES)
        
        # Get current headers
        current_headers = ws.row_values(1)
        
        st.subheader("📋 Current Headers in Google Sheet:")
        st.code(str(current_headers))
        
        st.subheader("✅ Expected Headers:")
        st.code(str(EXPECTED_COLUMNS))
        
        # Compare
        st.subheader("🔍 Comparison:")
        
        if current_headers == EXPECTED_COLUMNS:
            st.success("✅ Headers match perfectly! Your sheet is configured correctly.")
        else:
            st.error("❌ Headers DO NOT match! This is causing data loss.")
            
            st.subheader("Differences:")
            for i, (current, expected) in enumerate(zip(current_headers, EXPECTED_COLUMNS)):
                if current != expected:
                    st.error(f"Column {i+1} (Column {chr(65+i)}): Got '{current}' but expected '{expected}'")
            
            # Show fix instructions
            st.subheader("🔧 How to Fix:")
            st.markdown("""
            1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1EXiXsOQ0VsfIbZlUpN3r6g0-aRNUUEKZDVZHh_xZnEY/edit
            2. Look at the first row (headers)
            3. Change the headers to match the expected values above
            4. **Most important**: Change Column A from "10" to "Received Date"
            5. Run this diagnostic again to verify the fix
            """)
            
            if st.button("🔄 Auto-Fix Headers (USE WITH CAUTION)"):
                st.warning("⚠️ This will update the header row in your Google Sheet!")
                if st.checkbox("I understand this will modify my Google Sheet"):
                    ws.update('A1:M1', [EXPECTED_COLUMNS])
                    st.success("✅ Headers have been updated! Please refresh the page.")
                    st.balloons()
        
        # Show sample data
        st.subheader("📊 Sample Data (first 5 rows):")
        sample_data = ws.get_all_values()[:6]  # Header + 5 data rows
        st.table(sample_data)
        
    except Exception as e:
        st.error(f"❌ Error connecting to Google Sheets: {e}")
        st.info("Make sure your .streamlit/secrets.toml file is configured correctly.")

if __name__ == "__main__":
    main()
