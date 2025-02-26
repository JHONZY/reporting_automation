import streamlit as st
import pandas as pd
import mysql.connector
import subprocess
import io
import time
import os
import sys

# Streamlit Page Config
st.set_page_config(page_title="REPORTING WEBSITE", layout="wide")

# Load Database Credentials
try:
    DB_HOST = st.secrets["DB"]["DB_HOST"]
    DB_USER = st.secrets["DB"]["DB_USER"]
    DB_PASSWORD = st.secrets["DB"]["DB_PASSWORD"]
    DB_NAME = st.secrets["DB"]["DB_NAME"]
except KeyError:
    st.error("❌ Database credentials missing! Set them in Streamlit Secrets.")
    st.stop()

# Base Directory for Queries
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
QUERIES_PATH = os.path.join(BASE_DIR, "queries")

# Report Query Paths
REPORT_QUERIES = {
    "MASTERLIST": os.path.join(QUERIES_PATH, "masterlist.sql"),
    "SKIPS AND COLLECT REPORT": os.path.join(QUERIES_PATH, "skips_and_collect_report.sql"),
    "COLLECT REPORT": os.path.join(QUERIES_PATH, "collect_report.sql"),
}

# Function to Load SQL Query
def load_query(report_type):
    file_path = REPORT_QUERIES.get(report_type)
    if not file_path or not os.path.exists(file_path):
        st.error(f"⚠ SQL query file not found: {file_path}")
        return None

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            return file.read()
    except Exception as e:
        st.error(f"❌ Error loading SQL query file: {e}")
        return None

# Function to Connect to MySQL & Fetch Data
def load_data(report_type):
    query = load_query(report_type)
    if not query:
        return pd.DataFrame()

    try:
        conn = mysql.connector.connect(
            host=DB_HOST,    # Use the actual IP, not "localhost"
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME,
            port=3306,       # Explicitly set MySQL port
            use_pure=True
        )
        cursor = conn.cursor(dictionary=True)
        cursor.execute(query)
        data = cursor.fetchall()
        df = pd.DataFrame(data)
        cursor.close()
        conn.close()
        return df
    except mysql.connector.Error as e:
        st.error(f"❌ Database connection error: {e}")
        return pd.DataFrame()


# Function to Convert DataFrame to Excel File
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

# Function to Run External Python Script
def run_python_script():
    try:
        script_path = os.path.join(BASE_DIR, "importing", "import.py")
        subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=True)
        st.success("Python Import Script Executed Successfully! ✅")
        return True
    except subprocess.CalledProcessError as e:
        st.error(f"Importing Error! ❌\n{e.stderr}")
        return False

# Sidebar Navigation
campaigns = ["CBS HOMELOAN", "PNB HOMELOAN", "SBF HOMELOAN", "BDO HOMELOAN", "OPTION 5"]

if "selected_campaign" not in st.session_state:
    st.session_state["selected_campaign"] = campaigns[0]

selected_campaign = st.sidebar.selectbox("Choose a campaign:", campaigns, index=campaigns.index(st.session_state["selected_campaign"]))

if selected_campaign != st.session_state["selected_campaign"]:
    st.session_state["selected_campaign"] = selected_campaign
    st.rerun()

st.title(f"{selected_campaign}" if selected_campaign else "REPORTING WEBSITE")

# CBS HOMELOAN - SHOW MASTERLIST + PROCESS ENDORSEMENT
if selected_campaign == "CBS HOMELOAN":
    df_masterlist = load_data("MASTERLIST")
    if not df_masterlist.empty:
        st.dataframe(df_masterlist)
        col1, col2 = st.columns([0.79, 0.15])

        with col1:
            if st.button("PROCESS ENDORSEMENT", use_container_width=False):
                status_placeholder = st.empty()
                status_placeholder.info("Running Import Python Script... Please wait.")
                time.sleep(5)
                status_placeholder.empty()

                if run_python_script():
                    status_placeholder.info("Processing complete.")
                    time.sleep(2)
                    status_placeholder.empty()
                else:
                    status_placeholder.error("Importing Error! ❌")
                    time.sleep(5)
                    status_placeholder.empty()

        with col2:
            st.download_button(
                label="📥 DOWNLOAD MASTERLIST",
                data=convert_df_to_excel(df_masterlist),
                file_name="Masterlist.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# BDO HOMELOAN - Report Selection
elif selected_campaign == "BDO HOMELOAN":
    if "report_type" not in st.session_state:
        st.session_state["report_type"] = "SKIPS AND COLLECT REPORT"

    col1, col2 = st.columns(2)
    if col1.button("SKIPS AND COLLECT REPORT", use_container_width=True):
        st.session_state["report_type"] = "SKIPS AND COLLECT REPORT"
        st.rerun()
    if col2.button("COLLECT REPORT", use_container_width=True):
        st.session_state["report_type"] = "COLLECT REPORT"
        st.rerun()

    report_type = st.session_state["report_type"]
    st.title(f"BDO HOMELOAN - {report_type}")

    df_option1 = load_data(report_type)
    st.dataframe(df_option1.head(30))

    if not df_option1.empty:
        st.download_button(
            label="📥 Download Full Report",
            data=convert_df_to_excel(df_option1),
            file_name=f"{report_type.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
