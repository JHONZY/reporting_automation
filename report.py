import streamlit as st  # type: ignore
import pandas as pd  # type: ignore
import pyodbc
import subprocess
import io
import time
import os
import sys

# Streamlit Page Config
st.set_page_config(page_title="REPORTING WEBSITE", layout="wide")

# Database Connection Config
DB_SERVER = "192.168.15.197"
DB_USER = "jborromeo"
DB_PASSWORD = "$PMadrid1234jb"
DB_NAME = "bcrm"
DSN_NAME = "data"  # ODBC Data Source Name

# Function to run the Excel macro
def run_excel_macro():
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

# Function to run the external Python script
def run_python_script():
    try:
        script_path = r"importing\import.py"

        # Run the script
        subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=True)

        # Show success message
        success_message = st.empty()
        success_message.success("Python Import Script Executed Successfully! ✅")

        # Hide message after 3 seconds
        time.sleep(30)
        success_message.empty()

        return True
    except subprocess.CalledProcessError as e:
        st.error(f"Importing Error! ❌\n{e.stderr}")  # Show error details
        return False

# File paths to SQL queries
QUERIES_PATH = r"C:\Users\SPM.SPMWNDT0659\Documents\Python\streamlit\queries"
REPORT_QUERIES = {
    "MASTERLIST": os.path.join(QUERIES_PATH, "masterlist.sql"),
    "SKIPS AND COLLECT REPORT": os.path.join(QUERIES_PATH, "skips_and_collect_report.sql"),
    "COLLECT REPORT": os.path.join(QUERIES_PATH, "collect_report.sql"),
}

# Function to read an SQL query from a file
def load_query(report_type):
    file_path = REPORT_QUERIES.get(report_type)
    if not file_path:
        st.error(f"Invalid report type: {report_type}")
        return None

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            return file.read()
    except Exception as e:
        st.error(f"Error loading SQL query file: {file_path}\nError: {e}")
        return None

# Function to fetch data from ODBC database (Masterlist + Reports)
def load_data(report_type):
    query = load_query(report_type)
    if not query:
        return pd.DataFrame()

    try:
        # Create a single database connection
        conn = pyodbc.connect(f"DSN={DSN_NAME};UID={DB_USER};PWD={DB_PASSWORD}", autocommit=True)
        with conn:
            df = pd.read_sql(query, conn)
        return df
    except Exception as e:
        st.error(f"Database connection error: {e}")
        return pd.DataFrame()

        
# Function to convert DataFrame to Excel
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    return output.getvalue()

# Sidebar Navigation

campaigns = ["CBS HOMELOAN", "PNB HOMELOAN", "SBF HOMELOAN", "BDO HOMELOAN", "OPTION 5"]

# Check if there's a campaign stored in session state, else use the first option
if "selected_campaign" not in st.session_state:
    st.session_state["selected_campaign"] = campaigns[0]  # Default to first campaign

# Create a selectbox for campaign selection
selected_campaign = st.sidebar.selectbox("Choose a campaign:", campaigns, index=campaigns.index(st.session_state["selected_campaign"]))

# Update session state when selection changes
if selected_campaign != st.session_state["selected_campaign"]:
    st.session_state["selected_campaign"] = selected_campaign
    st.rerun()  # Refresh page with new selection

# ✅ Update Page Title Based on Selected Campaign
if selected_campaign:
    st.title(f"{selected_campaign}")
else:
    st.title("REPORTING WEBSITE")

# CBS HOMELOAN - SHOW MASTERLIST + PROCESS ENDORSEMENT
if selected_campaign == "CBS HOMELOAN":

    # Display Masterlist if the session state flag is True
    if st.session_state.get("show_masterlist", True):
        df_masterlist = load_data("MASTERLIST")
        if not df_masterlist.empty:
            st.dataframe(df_masterlist)
            # Add download button
            col1, col2 = st.columns([0.79, 0.15])  # 85% width for col1, 15% for col2


            with col1:  # Left side: Process Endorsement button
                if st.button("PROCESS ENDORSEMENT", use_container_width=False):
                    #st.info("Running Excel Macro... Please wait.")
                    #if run_excel_macro():
                        #st.success("Excel Macro Completed Successfully! ✅")

                    status_placeholder = st.empty()  # Create a placeholder for dynamic updates

                    status_placeholder.info("Running Import Python Script... Please wait.")
                    time.sleep(5)
                    status_placeholder.empty()

                    if run_python_script():  
                        #status_placeholder.info("Running Import Python Script... Please wait.")
                        status_placeholder.info("Please wait.")
                        #time.sleep(2)  # Wait for 5 seconds
                        status_placeholder.empty()  # Clear the message
                    else:
                        status_placeholder.error("Importing Error! ❌ File not found!")
                        time.sleep(5)  # Wait for 5 seconds
                        status_placeholder.empty()  # Clear the message

            with col2:  # Right side: Download button
                st.download_button(
                    label="📥 DOWNLOAD MASTERLIST",
                    data=convert_df_to_excel(df_masterlist),
                    file_name="Masterlist.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            

# BDO HOMELOAN - Report Selection
elif selected_campaign == "BDO HOMELOAN":
    if "report_type" not in st.session_state:
        st.session_state["report_type"] = "SKIPS AND COLLECT REPORT"  # Default view

    col1, col2 = st.columns(2)
    if col1.button("SKIPS AND COLLECT REPORT", use_container_width=True):
        st.session_state["report_type"] = "SKIPS AND COLLECT REPORT"
        st.rerun()
    if col2.button("COLLECT REPORT", use_container_width=True):
        st.session_state["report_type"] = "COLLECT REPORT"
        st.rerun()

    # Load and display selected report
    report_type = st.session_state["report_type"]
    st.title(f"BDO HOMELOAN - {report_type}")

    df_option1 = load_data(report_type)

    # ✅ Display only the first 30 rows in the table
    st.dataframe(df_option1.head(30))

    # ✅ Add a "Download Report" button to export full data
    if not df_option1.empty:
        st.download_button(
            label="📥 Download Full Report",
            data=convert_df_to_excel(df_option1),
            file_name=f"{report_type.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
