import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import altair as alt

# --- Smart Quarterly Reporting Assistant ---
# Author: Raghad Alqarni
# Date: July 2026

st.set_page_config(
    page_title="Smart Quarterly Reporting Assistant",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Hardcoded Streamlit theme settings for corporate high-contrast Light Theme
STREAMLIT_THEME_CONFIG = {
    "theme.base": "light",
    "theme.primaryColor": "#1F4E78",
    "theme.backgroundColor": "#FFFFFF",
    "theme.secondaryBackgroundColor": "#F8FAFC",
    "theme.textColor": "#000000",
    "theme.font": "serif"
}
for key, val in STREAMLIT_THEME_CONFIG.items():
    try:
        st.config.set_option(key, val)
    except Exception:
        pass

# --- Times New Roman & High Contrast Custom CSS Injection ---
st.markdown("""
<style>
    /* Force Times New Roman font strictly everywhere */
    * {
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }
    
    /* Force Pure Light Theme background on all outer views */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"], .main, [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
    }
    
    /* High Contrast Titles & Headers */
    .executive-title {
        color: #1F4E78 !important; /* Deep Royal Navy Blue */
        font-size: 2.8rem !important;
        font-weight: bold !important;
        text-align: center !important;
        margin-top: 10px !important;
        margin-bottom: 5px !important;
    }
    
    .gold-divider {
        height: 3px;
        background-color: #D4AF37 !important; /* Premium Gold */
        margin-bottom: 30px;
        border: none;
        border-radius: 2px;
    }
    
    /* Section Headers */
    .section-header {
        color: #1F4E78 !important;
        font-size: 1.6rem !important;
        font-weight: bold !important;
        margin-top: 30px !important;
        margin-bottom: 20px !important;
        border-left: 6px solid #D4AF37 !important;
        padding-left: 12px !important;
    }
    
    /* Dataset Configuration Cards */
    .config-card {
        background-color: #FFFFFF !important;
        padding: 22px !important;
        border-radius: 6px !important;
        border: 1.5px solid #CBD5E1 !important;
        border-top: 5px solid #1F4E78 !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.03) !important;
        margin-bottom: 15px !important;
    }
    
    .config-card-title {
        color: #1F4E78 !important;
        font-size: 1.25rem !important;
        font-weight: bold !important;
        margin-bottom: 15px !important;
        border-bottom: 1px solid #E2E8F0 !important;
        padding-bottom: 8px !important;
    }
    
    /* Selectbox Dropdown Box and Popover Menu Styling */
    div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        color: #1e293b !important;
        border-radius: 4px !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }
    
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #1e293b !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }

    div[data-baseweb="select"] * {
        background-color: #FFFFFF !important;
        color: #1e293b !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }

    [data-baseweb="popover"], [data-baseweb="popover"] * {
        background-color: #FFFFFF !important;
        color: #1e293b !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }
    
    [data-baseweb="menu"], [data-baseweb="menu"] * {
        background-color: #FFFFFF !important;
        color: #1e293b !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }

    div[role="listbox"], ul[role="listbox"], div[role="listbox"] *, ul[role="listbox"] * {
        background-color: #FFFFFF !important;
        color: #1e293b !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }
    
    div[role="option"]:hover, li[role="option"]:hover,
    div[role="option"][aria-selected="true"], li[role="option"][aria-selected="true"],
    div[role="option"]:hover *, li[role="option"]:hover *,
    div[role="option"][aria-selected="true"] *, li[role="option"][aria-selected="true"] * {
        background-color: #D4AF37 !important;
        color: #1e293b !important;
    }

    /* Input/Textarea elements general fallback styling */
    input, select, textarea {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1.5px solid #CBD5E1 !important;
        border-radius: 4px !important;
    }

    /* File upload container styling */
    div[data-testid="stFileUploader"],
    div[data-testid="stFileUploader"] > section,
    div[data-testid="stFileUploader"] > section > div,
    div[data-testid="stFileUploader"] [data-testid="stFileDropzone"],
    div[data-testid="stFileDropzone"] {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        border-radius: 6px !important;
    }
    
    div[data-testid="stFileUploader"] {
        border: 1.5px dashed #CBD5E1 !important;
        padding: 16px !important;
        text-align: center !important;
    }
    
    div[data-testid="stFileUploader"] *,
    div[data-testid="stFileUploader"] label,
    div[data-testid="stFileUploader"] span,
    div[data-testid="stFileUploader"] small,
    div[data-testid="stFileUploader"] p,
    div[data-testid="stFileUploader"] div {
        color: #000000 !important;
    }

    /* Clean File Uploader Button styling */
    div[data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]):not([aria-label*="Delete"]):not([aria-label*="Remove"]):not([class*="delete"]) {
        background-color: #EFEFEF !important;
        border: 1.5px solid #CBD5E1 !important;
        border-radius: 4px !important;
        padding: 8px 16px !important;
        transition: background-color 0.25s ease-in-out, border-color 0.25s ease-in-out !important;
        font-size: 0px !important;
        color: transparent !important;
        line-height: 0 !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    
    div[data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]):not([aria-label*="Delete"]):not([aria-label*="Remove"]):not([class*="delete"]) * {
        display: none !important;
        font-size: 0px !important;
        color: transparent !important;
    }
    
    div[data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]):not([aria-label*="Delete"]):not([aria-label*="Remove"]):not([class*="delete"])::after {
        content: "Upload" !important;
        display: block !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
        font-weight: bold !important;
        font-size: 1rem !important;
        color: #000000 !important;
    }
    
    div[data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]):not([aria-label*="Delete"]):not([aria-label*="Remove"]):not([class*="delete"]):hover {
        background-color: #D0D0D0 !important;
        border-color: #CBD5E1 !important;
    }
    
    div[data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]):not([aria-label*="Delete"]):not([aria-label*="Remove"]):not([class*="delete"]):hover::after {
        color: #000000 !important;
    }

    /* Distinct styling for the File Deletion button */
    div[data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"],
    div[data-testid="stFileUploader"] button[aria-label*="Delete"],
    div[data-testid="stFileUploader"] button[aria-label*="Remove"],
    div[data-testid="stFileUploader"] button[class*="delete"] {
        background-color: #F8D7DA !important;
        border: 1.5px solid #F5C2C7 !important;
        border-radius: 4px !important;
        padding: 8px 16px !important;
        transition: background-color 0.25s ease-in-out, border-color 0.25s ease-in-out !important;
        font-size: 0px !important;
        color: transparent !important;
        line-height: 0 !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        cursor: pointer !important;
    }
    
    div[data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"] *,
    div[data-testid="stFileUploader"] button[aria-label*="Delete"] *,
    div[data-testid="stFileUploader"] button[aria-label*="Remove"] *,
    div[data-testid="stFileUploader"] button[class*="delete"] * {
        display: none !important;
        font-size: 0px !important;
        color: transparent !important;
    }
    
    div[data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"]::after,
    div[data-testid="stFileUploader"] button[aria-label*="Delete"]::after,
    div[data-testid="stFileUploader"] button[aria-label*="Remove"]::after,
    div[data-testid="stFileUploader"] button[class*="delete"]::after {
        content: "🗑️ Remove" !important;
        display: block !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
        font-weight: bold !important;
        font-size: 1rem !important;
        color: #842029 !important;
    }
    
    div[data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"]:hover,
    div[data-testid="stFileUploader"] button[aria-label*="Delete"]:hover,
    div[data-testid="stFileUploader"] button[aria-label*="Remove"]:hover,
    div[data-testid="stFileUploader"] button[class*="delete"]:hover {
        background-color: #F1AEB5 !important;
        border-color: #EA868F !important;
    }
    
    div[data-testid="stFileUploader"] button[data-testid="stFileUploaderDeleteBtn"]:hover::after,
    div[data-testid="stFileUploader"] button[aria-label*="Delete"]:hover::after,
    div[data-testid="stFileUploader"] button[aria-label*="Remove"]:hover::after,
    div[data-testid="stFileUploader"] button[class*="delete"]:hover::after {
        color: #842029 !important;
    }

    /* Buttons styling */
    .stButton > button, .stDownloadButton > button, .feedback-submit {
        background-color: #1F4E78 !important;
        color: #FFFFFF !important;
        border: 1.5px solid #D4AF37 !important;
        font-weight: bold !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
        font-size: 1.15rem !important;
        padding: 12px 24px !important;
        border-radius: 4px !important;
        width: 100% !important;
        transition: all 0.25s ease-in-out !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important;
        cursor: pointer !important;
        text-align: center !important;
    }
    
    .stButton > button:hover, .stDownloadButton > button:hover, .feedback-submit:hover {
        background-color: #D4AF37 !important;
        color: #1F4E78 !important;
        border-color: #1F4E78 !important;
        box-shadow: 0 6px 12px rgba(212, 175, 55, 0.4) !important;
    }
    
    /* Widget labels color correction */
    label[data-testid="stWidgetLabel"] {
        color: #1F4E78 !important;
        font-weight: bold !important;
        font-size: 1.05rem !important;
    }
    
    /* Style st.tabs headers */
    .stTabs [data-baseweb="tab"], .stTabs button, .stTabs [role="tab"], .stTabs [data-baseweb="tab"] *, .stTabs button *, .stTabs [role="tab"] * {
        color: #1e293b !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }
    
    /* High contrast table styles */
    table, .stTable, div[data-testid="stTable"] {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        font-family: 'Times New Roman', Times, Georgia, serif !important;
    }
    
    table, th, td {
        border: 1px solid #CCCCCC !important;
    }
    
    table th, .stTable th, div[data-testid="stTable"] th {
        background-color: #F1F5F9 !important;
        color: #000000 !important;
        font-weight: bold !important;
        border: 1px solid #CCCCCC !important;
        padding: 8px 12px !important;
        text-align: left !important;
    }
    
    table td, .stTable td, div[data-testid="stTable"] td {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #CCCCCC !important;
        padding: 8px 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Clinics Meta-Data Definitions ---
CLINICS_METADATA = [
    {
        "id": 1,
        "short_name": "Al-Hejaz",
        "full_name": "Al Osrah Medical Center in Hijaz of Royal commission in Jubail Industrial",
        "status_rows": (4, 21),
        "rejection_rows": (4, 21)
    },
    {
        "id": 2,
        "short_name": "Al-Dafi",
        "full_name": "Al Dafi first aid Medical Center of Royal commission in Jubail Industrial",
        "status_rows": (26, 43),
        "rejection_rows": (26, 43)
    },
    {
        "id": 3,
        "short_name": "Al-Huwaylat",
        "full_name": "Al Howilat first aid Medical Center of Royal commission in Jubail Industrial",
        "status_rows": (48, 65),
        "rejection_rows": (48, 65)
    },
    {
        "id": 4,
        "short_name": "Al-Farouq",
        "full_name": "Al Farouq first aid Medical Center of Royal commission in Jubail Industrial",
        "status_rows": (69, 86),
        "rejection_rows": (69, 86)
    },
    {
        "id": 5,
        "short_name": "Jalmudah",
        "full_name": "Jalmoud first aid Medical Center of Royal commission in Jubail Industrial",
        "status_rows": (91, 108),
        "rejection_rows": (91, 108)
    },
    {
        "id": 6,
        "short_name": "RCH-Jubail",
        "full_name": "Royal Commission Hospital in Jubail",
        "status_rows": (113, 130),
        "rejection_rows": (113, 130)
    },
    {
        "id": 7,
        "short_name": "Ras-Al-Khair",
        "full_name": "Ras Al Khair first aid Medical Center of Royal commission in Jubail Industrial",
        "status_rows": (135, 152),
        "rejection_rows": (135, 152)
    }
]

INSURANCE_COMPANIES = [
    "Bupa", "Tawuniya", "Medgulf", "Malath", "SAICO", "Gulf Union", 
    "Alrajhi Takaful", "Gulf Insurance Group – GIG", "TCS", "GLOBMED", 
    "NEXTCARE", "ARABIAN SHIELD"
]

STATUS_CATEGORIES = ["Approved/Paid", "Rejected", "Pending", "In Process", "Billed", "Disputed"]
REJECTION_REASONS = [
    "Missing Authorization", "Service Not Covered", "Duplicate Claim",
    "Incorrect Member ID", "Diagnosis/Procedure Mismatch", "Invalid Code",
    "Provider Not in Network", "Timely Filing Limit Exceeded"
]

# --- Initialize session states ---
if "months" not in st.session_state:
    st.session_state["months"] = [
        {
            "is_uploaded": False,
            "file_name": "",
            "file_bytes": None,
            "month_date": datetime.date(2026, 4, 1),
            "insurance": "Bupa",
            "clinics_data": {c["short_name"]: {"status": [], "rejection": []} for c in CLINICS_METADATA}
        },
        {
            "is_uploaded": False,
            "file_name": "",
            "file_bytes": None,
            "month_date": datetime.date(2026, 5, 1),
            "insurance": "Tawuniya",
            "clinics_data": {c["short_name"]: {"status": [], "rejection": []} for c in CLINICS_METADATA}
        },
        {
            "is_uploaded": False,
            "file_name": "",
            "file_bytes": None,
            "month_date": datetime.date(2026, 6, 1),
            "insurance": "Medgulf",
            "clinics_data": {c["short_name"]: {"status": [], "rejection": []} for c in CLINICS_METADATA}
        }
    ]

if "data_processed" not in st.session_state:
    st.session_state["data_processed"] = False


# --- HEADER PANEL ---
st.markdown("<div class='executive-title'>Smart Quarterly Reporting Assistant</div>", unsafe_allow_html=True)
st.markdown("<div class='gold-divider'></div>", unsafe_allow_html=True)


# --- EXCEL PARSING CORE LOGIC ---
def parse_monthly_file(file_bytes, month_idx):
    try:
        # Reset BytesIO pointer for complete server memory parsing
        buffer = BytesIO(file_bytes)
        buffer.seek(0)
        
        wb = openpyxl.load_workbook(buffer, data_only=True)
        if "Dashboard" not in wb.sheetnames:
            st.error(f"Error in Month {month_idx+1}: Sheet 'Dashboard' not found in uploaded file.")
            return False
            
        ws = wb["Dashboard"]
        month_state = st.session_state["months"][month_idx]
        
        for clinic in CLINICS_METADATA:
            c_name = clinic["short_name"]
            
            # Status Table (Columns A to C)
            status_list = []
            s_start, s_end = clinic["status_rows"]
            for r in range(s_start, s_end + 1):
                category = ws.cell(row=r, column=1).value  # Col A
                cases = ws.cell(row=r, column=2).value     # Col B
                amount = ws.cell(row=r, column=3).value    # Col C
                
                if category is not None:
                    cat_str = str(category).strip()
                    if cat_str and not cat_str.lower().startswith("total") and cat_str.lower() != "status":
                        try:
                            cases_val = int(cases) if cases is not None else 0
                        except (ValueError, TypeError):
                            cases_val = 0
                        try:
                            amount_val = float(amount) if amount is not None else 0.0
                        except (ValueError, TypeError):
                            amount_val = 0.0
                        
                        status_list.append({
                            "category": cat_str,
                            "cases": cases_val,
                            "amount": amount_val
                        })
            
            # Rejection Table (Columns E to G)
            rejection_list = []
            r_start, r_end = clinic["rejection_rows"]
            for r in range(r_start, r_end + 1):
                reason = ws.cell(row=r, column=5).value    # Col E
                cases = ws.cell(row=r, column=6).value     # Col F
                amount = ws.cell(row=r, column=7).value    # Col G
                
                if reason is not None:
                    reason_str = str(reason).strip()
                    if reason_str and not reason_str.lower().startswith("total") and reason_str.lower() != "rejection reason":
                        try:
                            cases_val = int(cases) if cases is not None else 0
                        except (ValueError, TypeError):
                            cases_val = 0
                        try:
                            amount_val = float(amount) if amount is not None else 0.0
                        except (ValueError, TypeError):
                            amount_val = 0.0
                        
                        rejection_list.append({
                            "category": reason_str,
                            "cases": cases_val,
                            "amount": amount_val
                        })
                        
            month_state["clinics_data"][c_name]["status"] = status_list
            month_state["clinics_data"][c_name]["rejection"] = rejection_list
            
        return True
    except Exception as e:
        st.error(f"Failed to extract metrics from Month {month_idx+1}: {str(e)}")
        return False


# --- AGGREGATION & BALANCING ENGINE ---
def compute_consolidated_aggregates():
    aggregated_status = {}
    aggregated_rejection = {}
    
    for m in st.session_state["months"]:
        if not m["is_uploaded"]:
            continue
        for c_name, c_data in m["clinics_data"].items():
            for item in c_data["status"]:
                cat = item["category"]
                cases = item["cases"]
                amount = item["amount"]
                if cat not in aggregated_status:
                    aggregated_status[cat] = {"cases": 0, "amount": 0.0}
                aggregated_status[cat]["cases"] += cases
                aggregated_status[cat]["amount"] += amount
                
            for item in c_data["rejection"]:
                reason = item["category"]
                cases = item["cases"]
                amount = item["amount"]
                if reason not in aggregated_rejection:
                    aggregated_rejection[reason] = {"cases": 0, "amount": 0.0}
                aggregated_rejection[reason]["cases"] += cases
                aggregated_rejection[reason]["amount"] += amount
                
    status_summary = [
        {"category": k, "cases": v["cases"], "amount": round(v["amount"], 2)}
        for k, v in aggregated_status.items()
    ]
    rejection_summary = [
        {"category": k, "cases": v["cases"], "amount": round(v["amount"], 2)}
        for k, v in aggregated_rejection.items()
    ]
    
    # Strictly reconcile the sums to enforce absolute equality
    sum_status_cases = sum(x["cases"] for x in status_summary)
    sum_status_amount = sum(x["amount"] for x in status_summary)
    
    sum_rejection_cases = sum(x["cases"] for x in rejection_summary)
    sum_rejection_amount = sum(x["amount"] for x in rejection_summary)
    
    diff_cases = sum_status_cases - sum_rejection_cases
    diff_amount = sum_status_amount - sum_rejection_amount
    
    if diff_cases > 0 or abs(diff_amount) > 0.01:
        rejection_summary.append({
            "category": "Approved/Paid Claims (No Rejection)",
            "cases": max(0, diff_cases),
            "amount": round(max(0.0, diff_amount), 2)
        })
        
    return status_summary, rejection_summary


# --- 1. CONFIGURATION LAYOUT ---
st.markdown("<div class='section-header'>1. Monthly Datasets Configuration</div>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

for idx in range(3):
    m_data = st.session_state["months"][idx]
    m_name = f"Month {idx + 1}"
    
    with [col1, col2, col3][idx]:
        st.markdown(f"""
        <div class="config-card">
            <div class="config-card-title">📅 {m_name} Settings</div>
        """, unsafe_allow_html=True)
        
        m_data["insurance"] = st.selectbox(
            f"Select Insurance - {m_name}",
            options=INSURANCE_COMPANIES,
            index=INSURANCE_COMPANIES.index(m_data["insurance"]),
            key=f"ins_sel_{idx}"
        )
        
        available_dates = []
        for y in range(2020, 2031):
            for m in range(1, 13):
                available_dates.append(datetime.date(y, m, 1))
        
        date_options_labels = [d.strftime("%B %Y") for d in available_dates]
        
        current_date_val = m_data["month_date"]
        if not isinstance(current_date_val, datetime.date):
            current_date_val = datetime.date(2026, 4 + idx, 1)
        else:
            current_date_val = datetime.date(current_date_val.year, current_date_val.month, 1)
            
        try:
            default_opt_idx = available_dates.index(current_date_val)
        except ValueError:
            default_opt_idx = available_dates.index(datetime.date(2026, 4 + idx, 1))
            
        selected_label = st.selectbox(
            f"Month Reference Date - {m_name}",
            options=date_options_labels,
            index=default_opt_idx,
            key=f"month_year_sel_{idx}"
        )
        
        m_data["month_date"] = available_dates[date_options_labels.index(selected_label)]
        
        file_val = st.file_uploader(
            f"Drop Month {idx+1} Report (.xlsx)",
            type=["xlsx"],
            key=f"file_up_{idx}"
        )
        
        if file_val is not None:
            uploaded_bytes = file_val.getvalue()
            if m_data["file_name"] != file_val.name or m_data["file_bytes"] != uploaded_bytes:
                m_data["file_name"] = file_val.name
                m_data["file_bytes"] = uploaded_bytes
                m_data["is_uploaded"] = True
                
                success = parse_monthly_file(m_data["file_bytes"], idx)
                if success:
                    st.session_state["data_processed"] = True
                
        if m_data["is_uploaded"]:
            st.markdown(f"""
                <div style="margin-top: 15px; padding: 10px; background-color: #E2F0D9; border: 1.5px solid #385723; border-radius: 4px; text-align: center;">
                    <b style="color: #385723;">✓ File Connected:</b><br/>
                    <span style="font-size: 0.9rem; color: #111111;">{m_data['file_name']}</span>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div style="margin-top: 15px; padding: 10px; background-color: #FCE4D6; border: 1.5px solid #C65911; border-radius: 4px; text-align: center;">
                    <b style="color: #C65911;">⚠ Awaiting Dataset File</b>
                </div>
            """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)


# --- 2. EXECUTIVE DASHBOARD & SUMMARY ---
if st.session_state["data_processed"]:
    st.markdown("<div class='section-header'>2. Consolidated Quarterly Performance Dashboard</div>", unsafe_allow_html=True)
    
    status_summary, rejection_summary = compute_consolidated_aggregates()
    
    df_status = pd.DataFrame(status_summary)
    df_rejection = pd.DataFrame(rejection_summary)
    
    tab1, tab2 = st.tabs(["📊 Claims Status Breakdown", "⚠️ Rejection Reasons Analysis"])
    
    with tab1:
        c1, c2 = st.columns([1, 1])
        with c1:
            st.dataframe(df_status, use_container_width=True)
        with c2:
            if not df_status.empty:
                chart = alt.Chart(df_status).mark_bar(color='#1F4E78').encode(
                    x=alt.X('category:N', title='Status Category', sort=None),
                    y=alt.Y('amount:Q', title='Total Amount (SAR)'),
                    tooltip=['category', 'cases', 'amount']
                ).properties(height=350)
                st.altair_chart(chart, use_container_width=True)
                
    with tab2:
        c1, c2 = st.columns([1, 1])
        with c1:
            st.dataframe(df_rejection, use_container_width=True)
        with c2:
            if not df_rejection.empty:
                chart_rej = alt.Chart(df_rejection).mark_bar(color='#D4AF37').encode(
                    x=alt.X('category:N', title='Rejection Reason', sort=None),
                    y=alt.Y('amount:Q', title='Total Amount (SAR)'),
                    tooltip=['category', 'cases', 'amount']
                ).properties(height=350)
                st.altair_chart(chart_rej, use_container_width=True)

# Footer
st.markdown("""
<div class="executive-footer">
    Smart Quarterly Reporting Assistant &copy; 2026 | Royal Commission Health Services
</div>
""", unsafe_allow_html=True)
