import streamlit as st  # type: ignore
import pandas as pd  # type: ignore
import pyodbc
import subprocess
import io
import time
import os
import sys

# Optional: Import win32com.client only if on Windows
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# Streamlit Page Config
st.set_page_config(page_title="REPORTING WEBSITE", layout="wide")

# Database Connection Config
DB_HOST = "192.168.15.197"
DB_USER = "jborromeo"
DB_PASSWORD = "$PMadrid1234jb"
DB_NAME = "bcrm"
DSN_NAME = "data"  # ODBC Data Source Name

# ODBC Connection String (For SQL Server)
CONN_STRING = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={DB_HOST};DATABASE={DB_NAME};UID={DB_USER};PWD={DB_PASSWORD}"

# Function to run the Excel macro
def run_excel_macro():
    if not WIN32_AVAILABLE:
        st.error("Excel automation is only supported on Windows.")
        return False
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Run in the background
        
        wb = excel.Workbooks.Open(r"\\192.168.15.241\admin\ACTIVE\jlborromeo\CBS HOME LOAN\CBS HEADER MAPPING V2.xlsm")
        excel.Application.Run("AlignDataBasedOnMappingWithMissingHeaders")
        wb.Save()
        wb.Close()
        excel.Quit()
        return True
    except Exception as e:
        st.error(f"Failed to run macro: {e}")
        return False

# Function to run an external Python script
def run_python_script():
    try:
        script_path = r"importing/import.py"
        subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=True)
        st.success("Python Import Script Executed Successfully! ✅")
        return True
    except subprocess.CalledProcessError as e:
        st.error(f"Importing Error! ❌\n{e.stderr}")
        return False

# Define query paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
QUERIES_PATH = os.path.join(BASE_DIR, "queries")

REPORT_QUERIES = {
    "MASTERLIST": os.path.join(QUERIES_PATH, "masterlist.sql"),
    "SKIPS AND COLLECT REPORT": os.path.join(QUERIES_PATH, "skips_and_collect_report.sql"),
    "COLLECT REPORT": os.path.join(QUERIES_PATH, "collect_report.sql"),
}

# Function to load SQL query
def load_query(report_type):
    file_path = REPORT_QUERIES.get(report_type)
    if not file_path or not os.path.exists(file_path):
        st.error(f"⚠ SQL query file not found: {file_path}")
        return None
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            return file.read()
    except Exception as e:
        st.error(f"❌ Error loading SQL query file: {file_path}\nError: {e}")
        return None

# Function to fetch data from SQL Server
def load_data(report_type):
    query = load_query(report_type)
    if not query:
        return pd.DataFrame()
    try:
        conn = pyodbc.connect(CONN_STRING)
        df = pd.read_sql(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"❌ Database connection error: {e}")
        return pd.DataFrame()

# Function to convert dataframe to Excel
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    processed_data = output.getvalue()
    return processed_data

# Sidebar Navigation
campaigns = ["CBS HOMELOAN", "PNB HOMELOAN", "SBF HOMELOAN", "BDO HOMELOAN", "OPTION 5"]
st.session_state["selected_campaign"] = st.sidebar.selectbox("Choose a campaign:", campaigns, index=0)

st.title(st.session_state["selected_campaign"])

# CBS HOMELOAN - Masterlist
if st.session_state["selected_campaign"] == "CBS HOMELOAN":
    df_masterlist = load_data("MASTERLIST")
    if not df_masterlist.empty:
        st.dataframe(df_masterlist)
        
        col1, col2 = st.columns([0.8, 0.2])
        with col1:
            if st.button("PROCESS ENDORSEMENT"):
                if run_python_script():
                    st.info("Processing completed!")
        with col2:
            st.download_button("📥 DOWNLOAD MASTERLIST", convert_df_to_excel(df_masterlist), "Masterlist.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# BDO HOMELOAN - Reports
elif st.session_state["selected_campaign"] == "BDO HOMELOAN":
    report_type = st.radio("Select Report", ["SKIPS AND COLLECT REPORT", "COLLECT REPORT"])
    df_report = load_data(report_type)
    st.dataframe(df_report.head(30))
    if not df_report.empty:
        st.download_button("📥 Download Full Report", convert_df_to_excel(df_report), f"{report_type.replace(' ', '_')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
