import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import io
import os
import json
import base64
from datetime import datetime, timedelta
from fpdf import FPDF
import math

# ==========================================
# CONFIGURATION & SETUP
# ==========================================
st.set_page_config(page_title="Mobitel Material Tracker", layout="wide", page_icon="📡")

# ලංකාවේ වෙලාව ලබා ගැනීමේ ශ්‍රිතය (Sri Lanka Time: UTC + 5:30)
def get_sl_time():
    return datetime.utcnow() + timedelta(hours=5, minutes=30)

# --- BACKGROUND IMAGE FUNCTION ---
def set_bg_hack(main_bg):
    main_bg_ext = "jpg"
    try:
        with open(main_bg, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
            
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url(data:image/{main_bg_ext};base64,{encoded_string});
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }}
            .stApp::before {{
                content: "";
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(14, 17, 23, 0.95); 
                z-index: -1;
            }}
            [data-testid="block-container"] {{
                background-color: rgba(14, 17, 23, 0.85);
                padding: 2rem;
                border-radius: 15px;
                box-shadow: 0px 0px 20px rgba(0, 0, 0, 0.5);
            }}
            [data-testid="stMetric"] {{
                background-color: rgba(255, 255, 255, 0.05);
                backdrop-filter: blur(10px);
                padding: 20px;
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 0.1);
            }}
            </style>
            """,
            unsafe_allow_html=True
        )
    except FileNotFoundError:
        pass 

set_bg_hack("bg.jpg")
# ---------------------------------

# Ensure Backup Directories Exist (For local execution)
BACKUP_DIR = "Delivery_Notes_Backup"
EVIDENCE_DIR = "Evidence_Backup"
for d in [BACKUP_DIR, EVIDENCE_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

# ==========================================
# GOOGLE SHEETS CONNECTION
# ==========================================
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

try:
    if "gcp_json" in st.secrets:
        creds_dict = json.loads(st.secrets["gcp_json"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    
    # Connect Sheets
    client = gspread.authorize(creds)
    
except Exception as e:
    st.error(f"Error connecting to Google Sheets: {e}")
    st.stop()

MAIN_SHEET_ID = "1pGS-qmg5MifIneWINIFJ9-TZVxQ7pfMe6SHRCUIcr54"
REMOVAL_SHEET_ID = "1tCll8K66UbsMF7nI7CF-tczr6WidSJpIcXdKHE4O5UI"

@st.cache_resource
def get_sheets():
    try:
        main_sh = client.open_by_key(MAIN_SHEET_ID).sheet1
        removal_sh = client.open_by_key(REMOVAL_SHEET_ID).sheet1
        return main_sh, removal_sh
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets. (Error: {e})")
        st.stop()

main_sheet, removal_sheet = get_sheets()

# ==========================================
# AUTHENTICATION SYSTEM
# ==========================================
USERS_FILE = "users.json"

def load_users():
    if not os.path.exists(USERS_FILE):
        default_users = {
            "admin": {"password": "admin9499", "role": "admin"},
            "membermgm": {"password": "mgm123", "role": "member"}
        }
        save_users(default_users)
        return default_users
    with open(USERS_FILE, "r") as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=4)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "role" not in st.session_state:
    st.session_state.role = ""
    
if "main_up_key" not in st.session_state:
    st.session_state.main_up_key = 0
if "rem_up_key" not in st.session_state:
    st.session_state.rem_up_key = 0

users_db = load_users()

# Login Screen
if not st.session_state.logged_in:
    st.title("🔐 Login to Mobitel Material Tracker")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("User Login")
        input_user = st.text_input("Username")
        input_pass = st.text_input("Password", type="password")
        
        if st.button("Login", type="primary", use_container_width=True):
            if input_user in users_db and users_db[input_user]["password"] == input_pass:
                st.session_state.logged_in = True
                st.session_state.username = input_user
                st.session_state.role = users_db[input_user]["role"]
                st.rerun()
            else:
                st.error("Invalid Username or Password!")
    st.stop()

st.sidebar.markdown(f"👤 **Logged in as:** {st.session_state.username.upper()} ({st.session_state.role.capitalize()})")
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.rerun()

# ==========================================
# DATA FETCHING 
# ==========================================
@st.cache_data(ttl=60)
def fetch_main_data():
    return main_sheet.get_all_records()

@st.cache_data(ttl=60)
def fetch_removal_data():
    return removal_sheet.get_all_records()

# ==========================================
# HELPER FUNCTIONS (EXCEL & PDF)
# ==========================================
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Exported Data')
    return output.getvalue()

def generate_delivery_note_pdf(dn_number, date, issued_by, issued_to, mapped_from, items_df, rec_company, rec_name, rec_nic, rec_mobile, rec_vehicle):
    pdf = FPDF()
    pdf.add_page()
    if os.path.exists("logo.png"):
        pdf.image("logo.png", x=10, y=8, w=45)
    pdf.set_font("Arial", "B", 18)
    pdf.set_text_color(144, 12, 63)
    pdf.cell(0, 8, "ONNEXTA HOLDINGS PVT LTD", ln=True, align="R")
    pdf.set_font("Arial", "B", 12)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 8, f"DELIVERY NOTE #{dn_number}", ln=True, align="R")
    pdf.ln(15)

    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(245, 245, 245)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(50, 8, f" Date: {date}", border=1, fill=True)
    pdf.cell(140, 8, f" Issued By: {issued_by}", border=1, ln=True)
    pdf.cell(50, 8, f" Issued To: {issued_to}", border=1, fill=True)
    pdf.cell(140, 8, " Project: Mobitel", border=1, ln=True)
    pdf.cell(190, 8, f" Mapped From (Site): {mapped_from}", border=1, ln=True, fill=True)
    pdf.ln(8)

    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(44, 62, 80)
    pdf.set_text_color(255, 255, 255)
    cols = ["INDEX", "ITEM", "DESCRIPTION", "SERIAL", "QTY", "REMARKS"]
    widths = [15, 35, 65, 35, 15, 25]
    for i in range(len(cols)):
        pdf.cell(widths[i], 8, cols[i], border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_font("Arial", "", 8)
    pdf.set_text_color(0, 0, 0)
    for idx, row in enumerate(items_df.iterrows(), 1):
        r = row[1]
        fill_row = True if idx % 2 == 0 else False
        if fill_row: pdf.set_fill_color(248, 248, 248)
        else: pdf.set_fill_color(255, 255, 255)
        pdf.cell(widths[0], 8, f"{idx:02d}", border=1, align="C", fill=fill_row)
        
        item_name = str(r.get("Generic Name", r.get("Removed Item", "")))[:20]
        desc = str(r.get("Item Description", r.get("Removed Item Description", "")))[:40]
        serial = str(r.get("SN", "N/A"))
        qty = str(r.get("Handed Over Qty", r.get("Removal Qty", "")))
        remarks = str(r.get("Remarks", ""))[:15]
        
        pdf.cell(widths[1], 8, f" {item_name}", border=1, fill=fill_row)
        pdf.cell(widths[2], 8, f" {desc}", border=1, fill=fill_row)
        pdf.cell(widths[3], 8, serial, border=1, align="C", fill=fill_row)
        pdf.cell(widths[4], 8, qty, border=1, align="C", fill=fill_row)
        pdf.cell(widths[5], 8, f" {remarks}", border=1, fill=fill_row)
        pdf.ln()

    pdf.ln(10)
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(44, 62, 80)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(190, 8, " RECEIVER DETAILS", ln=True, fill=True, border=1)
    
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(95, 8, f" Company: {rec_company}", border=1)
    pdf.cell(95, 8, f" Receiver Name: {rec_name}", border=1, ln=True)
    pdf.cell(95, 8, f" NIC: {rec_nic}", border=1)
    pdf.cell(95, 8, f" Mobile: {rec_mobile}", border=1, ln=True)
    pdf.cell(190, 8, f" Vehicle: {rec_vehicle}", border=1, ln=True)
    pdf.ln(20)
    pdf.cell(95, 8, "________________________", align="C")
    pdf.cell(95, 8, "________________________", align="C", ln=True)
    pdf.cell(95, 5, "Authorized Signature", align="C")
    pdf.cell(95, 5, "Receiver's Signature", align="C", ln=True)
    return bytes(pdf.output())

def generate_table_export_pdf(df, title):
    pdf = FPDF(orientation='L')
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.set_text_color(44, 62, 80)
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.set_font("Arial", "I", 10)
    pdf.set_text_color(100, 100, 100)
    
    sl_time_str = get_sl_time().strftime('%Y-%m-%d %H:%M')
    pdf.cell(0, 8, f"Generated on: {sl_time_str} (SLST)", ln=True, align="C")
    
    pdf.ln(5)
    
    cols = list(df.columns)
    
    # SN එකට 45 ක් දීලා තියෙනවා එක පේළියට එන්න. අනිත් ඒවත් ගැලපුවා.
    width_map = {
        "Index": 10, "Site ID": 18, "Site Name": 25, "Removed Item": 30, 
        "Removed Item Description": 50, "UOM": 12, "Removal Qty": 20, 
        "SN": 45, "Return Status": 22, "Returned Qty": 20, "Remarks": 45,
        "Status": 18, "Required Qty": 20, "Materials From": 25,
        "Request Type": 22, "IR/MO": 20, "Item Code_INV": 25,
        "SE": 18, "Subcon": 18, "Generic Name": 35, "Item Description": 55,
        "ERP Site ID": 30, "Mat Req Ref": 30, "HQ/TaskID": 35,
        "Handed Over Qty": 20
    }
    
    widths = [width_map.get(col, 25) for col in cols]
    
    # --- අලුත් කරපු Dynamic Scaling සහ Centering ---
    total_table_width = sum(widths)
    max_page_width = 285 # A4 කොළේ පළල 297mm. දෙපැත්තෙන් 6mm ගානේ ඉතුරු කරන්න 285mm.
    
    # ටේබල් එක කොළේට වඩා ලොකු නම්, ඔටෝම Scale කරලා පොඩි කරනවා
    if total_table_width > max_page_width:
        scale_factor = max_page_width / total_table_width
        widths = [w * scale_factor for w in widths]
        total_table_width = sum(widths)
        
    page_width = 297 
    left_margin = (page_width - total_table_width) / 2
    if left_margin < 5: 
        left_margin = 5 
        
    pdf.set_left_margin(left_margin)
    pdf.set_x(left_margin)
    # -----------------------------------------------
    
    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(44, 62, 80)
    pdf.set_text_color(255, 255, 255)
    
    for i in range(len(cols)):
        pdf.cell(widths[i], 8, str(cols[i]), border=1, align="C", fill=True)
    pdf.ln()
    
    pdf.set_font("Arial", "", 8)
    pdf.set_text_color(0, 0, 0)
    line_height = 5
    
    for idx, row in df.iterrows():
        fill_row = True if idx % 2 == 0 else False
        
        max_lines = 1
        for i, col in enumerate(cols):
            text = str(row.get(col, ""))
            if text == "nan" or text == "None": text = ""
            
            safe_width = widths[i] - 2 if widths[i] > 2 else 1
            
            lines = 0
            for p in text.split('\n'):
                words = p.split(' ')
                line_w = 0
                for w in words:
                    w_w = pdf.get_string_width(w + " ")
                    if line_w + w_w > safe_width and line_w > 0:
                        lines += 1
                        line_w = w_w
                    else:
                        line_w += w_w
                lines += 1
                
            math_lines = math.ceil(pdf.get_string_width(text) / safe_width)
            final_lines = max(lines, math_lines)
            
            if final_lines > max_lines:
                max_lines = final_lines
                
        row_height = max_lines * line_height
        
        if pdf.get_y() + row_height > 190:
            pdf.add_page()
            pdf.set_x(left_margin) 
            pdf.set_font("Arial", "B", 9)
            pdf.set_fill_color(44, 62, 80)
            pdf.set_text_color(255, 255, 255)
            for i in range(len(cols)):
                pdf.cell(widths[i], 8, str(cols[i]), border=1, align="C", fill=True)
            pdf.ln()
            pdf.set_font("Arial", "", 8)
            pdf.set_text_color(0, 0, 0)

        start_x = pdf.get_x()
        start_y = pdf.get_y()
        
        for i, col in enumerate(cols):
            text = str(row.get(col, ""))
            if text == "nan" or text == "None": text = ""
            
            if fill_row: 
                pdf.set_fill_color(248, 248, 248)
            else: 
                pdf.set_fill_color(255, 255, 255)
                
            pdf.rect(start_x, start_y, widths[i], row_height, style='DF')
            
            pdf.set_xy(start_x, start_y)
            align = "C" if widths[i] < 25 else "L"
            
            pdf.multi_cell(widths[i], line_height, text, border=0, align=align)
            
            start_x += widths[i]
            
        pdf.set_xy(left_margin, start_y + row_height)
        
    return bytes(pdf.output())

# ==========================================
# MAIN UI
# ==========================================
st.title("📡 Mobitel Site Material Tracker")
st.markdown("---")

if st.session_state.role == "admin":
    tab_dash, tab1, tab2, tab_settings = st.tabs(["📊 Dashboard", "📦 Main Materials", "♻️ Site Removal Materials", "⚙️ Settings"])
else:
    tab_dash, tab1, tab2 = st.tabs(["📊 Dashboard", "📦 Main Materials", "♻️ Site Removal Materials"])

main_data = fetch_main_data()
removal_data = fetch_removal_data()
df_main = pd.DataFrame(main_data) if main_data else pd.DataFrame()
df_removal = pd.DataFrame(removal_data) if removal_data else pd.DataFrame()

# DASHBOARD TAB
with tab_dash:
    st.subheader("Project Overview")
    if not df_main.empty:
        total_sites = df_main['Site ID'].nunique()
        total_items = len(df_main)
        
        col_metric1, col_metric2, col_metric3 = st.columns(3)
        col_metric1.metric("Total Active Sites", total_sites)
        col_metric2.metric("Total Material Entries", total_items)
        if not df_removal.empty:
            col_metric3.metric("Total Removal Entries", len(df_removal))
        
        st.markdown("---")
        st.write("**Material Status Breakdown**")
        status_counts = df_main['Status'].value_counts()
        st.bar_chart(status_counts)
    else:
        st.info("Not enough data to display dashboard.")

# TAB 1: MAIN MATERIALS
with tab1:
    st.subheader("Update Main and Surplus Materials")
    if not df_main.empty:
        site_list = df_main['Site ID'].unique().tolist()
        selected_site = st.selectbox("Select Site ID to Update:", ["-- Select a Site --"] + site_list, key="main_site_select")
        
        if selected_site != "-- Select a Site --":
            df_filtered = df_main[df_main['Site ID'] == selected_site].copy()
            df_filtered.insert(0, "Index", range(1, len(df_filtered) + 1))
            
            disabled_cols = ["Index", "ERP Site ID", "Mat Req Ref", "Site ID", "Site Name", "HQ/TaskID", "Generic Name", "Item Description", "UOM", "Required Qty", "Materials From", "Request Type", "IR/MO", "Item Code_INV", "SE", "Subcon"]
            
            edited_df = st.data_editor(
                df_filtered, 
                disabled=disabled_cols, 
                use_container_width=True, 
                hide_index=True, 
                key="main_data_editor"
            )
            
            if st.button("Save Updates to Database", type="primary", key="save_main"):
                with st.spinner("Saving data..."):
                    try:
                        save_df = edited_df.drop(columns=["Index"])
                        df_main.update(save_df)
                        updated_data = [df_main.columns.values.tolist()] + df_main.fillna("").values.tolist()
                        main_sheet.update(updated_data)
                        st.cache_data.clear()
                        
                        st.session_state.main_up_key += 1
                        st.session_state.main_success_msg = "Data successfully updated in Google Sheets!"
                        st.rerun() 
                        
                    except Exception as e:
                        st.error(f"Failed to save data. Error: {e}")
            
            if "main_success_msg" in st.session_state:
                st.success(st.session_state.main_success_msg)
                del st.session_state.main_success_msg
        
        st.markdown("---")
        st.subheader("📥 Select & Export Main Materials Data")
        st.write("Filter by Site and Status, then select exactly which columns to export:")
        
        col_export_m1, col_export_m2 = st.columns(2)
        with col_export_m1:
            main_export_site_filter = st.selectbox("Filter Export by Site ID:", ["All Sites"] + site_list, key="main_export_site_filter")
            status_filter = st.selectbox("Filter Export by Status:", ["All", "Installed", "Surplus", "HO"], key="main_export_status")
            
        export_df = df_main.copy()
        
        if main_export_site_filter != "All Sites":
            export_df = export_df[export_df['Site ID'] == main_export_site_filter]
        if status_filter != "All":
            export_df = export_df[export_df['Status'] == status_filter]
            
        export_df.insert(0, "Index", range(1, len(export_df) + 1))
        
        all_cols_main = export_df.columns.tolist()
        default_cols_main = ["Index", "Site ID", "Site Name", "Generic Name", "Item Description", "UOM", "Required Qty", "Status", "Remarks"]
        default_cols_main = [c for c in default_cols_main if c in all_cols_main]
        
        with col_export_m2:
            selected_cols_main = st.multiselect("Select columns for Export Table:", all_cols_main, default=default_cols_main, key="main_col_select")
            
        if not selected_cols_main:
            st.warning("Please select at least one column to export.")
        else:
            df_export_display_main = export_df[selected_cols_main].copy()
            df_export_display_main.insert(0, "Select for Export", True)
            
            edited_export_main = st.data_editor(df_export_display_main, use_container_width=True, hide_index=True, key="main_export_editor")
            selected_export_data_main = edited_export_main[edited_export_main["Select for Export"] == True].drop(columns=["Select for Export"])
            
            ex_col_m1, ex_col_m2 = st.columns(2)
            with ex_col_m1:
                excel_data = to_excel(selected_export_data_main)
                ex_name = f"Main_Materials_{main_export_site_filter}_{status_filter}.xlsx" if main_export_site_filter != "All Sites" else f"Main_Materials_{status_filter}.xlsx"
                st.download_button(label="💾 Download Selected Data (Excel)", data=excel_data, file_name=ex_name, mime="application/vnd.ms-excel")
            with ex_col_m2:
                pdf_name = f"Main_Materials_{main_export_site_filter}_{status_filter}.pdf" if main_export_site_filter != "All Sites" else f"Main_Materials_{status_filter}.pdf"
                pdf_title = f"Main Materials - {status_filter} Data" if main_export_site_filter == "All Sites" else f"Main Materials - {status_filter} Data ({main_export_site_filter})"
                
                pdf_main_data = generate_table_export_pdf(selected_export_data_main, pdf_title)
                st.download_button(label="📄 Download Selected Data (PDF)", data=pdf_main_data, file_name=pdf_name, mime="application/pdf")
    else:
        st.info("No data available.")

# TAB 2: REMOVAL MATERIALS
with tab2:
    st.subheader("Update Site Removal Materials")
    if not df_removal.empty:
        site_list_rem = df_removal['Site ID'].unique().tolist()
        selected_site_rem = st.selectbox("Select Site ID:", ["-- Select a Site --"] + site_list_rem, key="rem_site_select")
        
        if selected_site_rem != "-- Select a Site --":
            df_filtered_rem = df_removal[df_removal['Site ID'] == selected_site_rem].copy()
            df_filtered_rem.insert(0, "Index", range(1, len(df_filtered_rem) + 1))
            df_filtered_rem.insert(0, "Select for Delivery Note", False)
            
            disabled_cols_rem = ["Index", "Site ID", "Site Name", "Removed Item Description", "UOM"]
            
            st.markdown("**Check the boxes to select items for Handover / Delivery Note:**")
            
            edited_df_rem = st.data_editor(
                df_filtered_rem, 
                disabled=disabled_cols_rem, 
                use_container_width=True, 
                hide_index=True, 
                key="rem_data_editor"
            )
            
            if st.button("Save Removal Updates", type="primary", key="save_rem"):
                with st.spinner("Saving data..."):
                    try:
                        save_df = edited_df_rem.drop(columns=["Select for Delivery Note", "Index"])
                        df_removal.update(save_df)
                        updated_data_rem = [df_removal.columns.values.tolist()] + df_removal.fillna("").values.tolist()
                        removal_sheet.update(updated_data_rem)
                        st.cache_data.clear()
                        
                        st.session_state.rem_up_key += 1
                        st.session_state.rem_success_msg = "Removal data successfully updated in Google Sheets!"
                        st.rerun() 
                        
                    except Exception as e:
                        st.error(f"Error: {e}")
                        
            if "rem_success_msg" in st.session_state:
                st.success(st.session_state.rem_success_msg)
                del st.session_state.rem_success_msg
            
            st.markdown("---")
            st.subheader("📄 Generate Delivery Note (PDF)")
            selected_items_df = edited_df_rem[edited_df_rem["Select for Delivery Note"] == True]
            if not selected_items_df.empty:
                with st.form("delivery_note_form"):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        dn_number = st.text_input("Delivery Note #", value=get_sl_time().strftime("%Y%m%d%H%M"))
                        issued_by = st.text_input("Issued By", value="Warehouse Officer")
                        issued_to = st.text_input("Issued To (e.g., DAR)")
                        rec_company = st.text_input("Receiver Company")
                    with col_b:
                        rec_name = st.text_input("Receiver Name")
                        rec_nic = st.text_input("NIC")
                        rec_mobile = st.text_input("Mobile")
                        rec_vehicle = st.text_input("Vehicle Number")
                    generate_btn = st.form_submit_button("Generate Delivery Note")
                    
                if generate_btn:
                    with st.spinner("Generating PDF..."):
                        try:
                            pdf_bytes = generate_delivery_note_pdf(
                                dn_number, get_sl_time().strftime("%Y/%m/%d"), issued_by, issued_to, 
                                selected_site_rem, selected_items_df, rec_company, rec_name, rec_nic, rec_mobile, rec_vehicle
                            )
                            file_name = f"DeliveryNote_{dn_number}.pdf"
                            st.success(f"Delivery Note '{file_name}' Generated!")
                            st.download_button("⬇️ Download Delivery Note to PC", data=pdf_bytes, file_name=file_name, mime="application/pdf", type="primary")
                        except Exception as e:
                            st.error(f"Error generating PDF: {e}")
            else:
                st.warning("Select items using checkboxes above to generate Delivery Note.")
                
        st.markdown("---")
        st.subheader("📥 Select & Export Removal Data")
        st.write("Filter by Site and select exactly which columns to export:")
        
        col_export1, col_export2 = st.columns(2)
        with col_export1:
            export_site_filter = st.selectbox("Filter Export by Site ID:", ["All Sites"] + site_list_rem, key="rem_export_filter")
            
        df_export_rem = df_removal.copy()
        if export_site_filter != "All Sites":
            df_export_rem = df_export_rem[df_export_rem['Site ID'] == export_site_filter]
            
        df_export_rem.insert(0, "Index", range(1, len(df_export_rem) + 1))

        all_cols = df_export_rem.columns.tolist()
        default_cols = ["Index", "Site ID", "Removed Item Description", "SN", "UOM", "Removal Qty", "Return Status", "Returned Qty", "Remarks"]
        default_cols = [c for c in default_cols if c in all_cols]

        with col_export2:
            selected_cols = st.multiselect("Select columns for Export Table:", all_cols, default=default_cols, key="rem_col_select")
            
        if not selected_cols:
            st.warning("Please select at least one column to export.")
        else:
            df_export_display = df_export_rem[selected_cols].copy()
            df_export_display.insert(0, "Select for Export", True)
            
            edited_export_rem = st.data_editor(df_export_display, use_container_width=True, hide_index=True, key="rem_export_editor")
            selected_export_data = edited_export_rem[edited_export_rem["Select for Export"] == True].drop(columns=["Select for Export"])
            
            ex_col1, ex_col2 = st.columns(2)
            with ex_col1:
                ex_rem_excel = to_excel(selected_export_data)
                ex_rem_name = f"Site_Removal_Materials_{export_site_filter}.xlsx" if export_site_filter != "All Sites" else "Site_Removal_Materials.xlsx"
                st.download_button(label="💾 Download Selected as Excel", data=ex_rem_excel, file_name=ex_rem_name, mime="application/vnd.ms-excel")
            with ex_col2:
                pdf_rem_title = f"Site Removal Materials Export ({export_site_filter})" if export_site_filter != "All Sites" else "Site Removal Materials Export"
                pdf_rem_name = f"Site_Removal_Materials_{export_site_filter}.pdf" if export_site_filter != "All Sites" else "Site_Removal_Materials.pdf"
                ex_rem_pdf = generate_table_export_pdf(selected_export_data, pdf_rem_title)
                st.download_button(label="📄 Download Selected as PDF", data=ex_rem_pdf, file_name=pdf_rem_name, mime="application/pdf")

# ==========================================
# TAB: SETTINGS (ADMIN ONLY)
# ==========================================
if st.session_state.role == "admin":
    with tab_settings:
        st.subheader("User Management")
        
        st.write("**Current Users:**")
        for user, data in users_db.items():
            st.write(f"- 👤 **{user}** (Role: {data['role']})")
            
        st.markdown("---")
        
        col_add, col_edit = st.columns(2)
        
        with col_add:
            st.markdown("### ➕ Add New User")
            with st.form("add_user_form"):
                new_user = st.text_input("New Username")
                new_pass = st.text_input("Password", type="password")
                new_role = st.selectbox("Role", ["member", "admin"])
                submit_add = st.form_submit_button("Add User", type="primary")
                
                if submit_add:
                    if not new_user or not new_pass:
                        st.error("Please fill all fields.")
                    elif new_user in users_db:
                        st.error(f"User '{new_user}' already exists! Use the Edit section to change their password.")
                    else:
                        users_db[new_user] = {"password": new_pass, "role": new_role}
                        save_users(users_db)
                        st.success(f"User '{new_user}' added successfully!")
                        st.rerun()
        
        with col_edit:
            st.markdown("### ✏️ Edit Existing User")
            with st.form("edit_user_form"):
                existing_user = st.selectbox("Select User to Edit", list(users_db.keys()))
                edit_pass = st.text_input("New Password (leave blank to keep current)", type="password")
                edit_role = st.selectbox("Update Role", ["member", "admin"])
                submit_edit = st.form_submit_button("Update User", type="primary")
                
                if submit_edit:
                    if edit_pass: 
                        users_db[existing_user]["password"] = edit_pass
                    users_db[existing_user]["role"] = edit_role
                    save_users(users_db)
                    st.success(f"User '{existing_user}' updated successfully!")
                    st.rerun()
