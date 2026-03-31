import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Premium Luxury Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-title { 
        color: #1a3a5f; font-family: 'Sarabun', sans-serif; 
        font-weight: 800; text-align: center; margin-bottom: 30px;
        letter-spacing: 1px;
    }
    .kpi-box {
        background: white; padding: 25px; border-radius: 20px; text-align: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        border-top: 5px solid #1a3a5f;
        transition: 0.3s;
    }
    .kpi-box:hover { transform: translateY(-5px); box-shadow: 0 15px 35px rgba(0,0,0,0.1); }
    .stButton>button {
        border-radius: 15px; background: linear-gradient(135deg, #1a3a5f 0%, #2c5282 100%);
        color: white; border: none; height: 3.5em; font-weight: bold; width: 100%;
        box-shadow: 0 4px 15px rgba(26, 58, 95, 0.2);
    }
    </style>
    """, unsafe_allow_html=True)

# --- Keywords Configuration ---
keywords = ["Ventilator", "Foley", "Central line", "Port A Cath"]

def get_safe_total(df_in):
    """ฟังก์ชันช่วยหายอดรวมอุปกรณ์จาก Keyword"""
    d_cols = [c for c in df_in.columns if any(k.lower() in str(c).lower() for k in keywords)]
    if not d_cols: return 0, []
    df_c = df_in.copy()
    for c in d_cols: 
        df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
    return df_c[d_cols].values.sum(), d_cols

def process_file_summary(uploaded_file):
    """ฟังก์ชันอ่านไฟล์และสรุปยอดราย Sheet"""
    excel_file = pd.ExcelFile(uploaded_file)
    data = []
    for s in excel_file.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
        total_sum, _ = get_safe_total(df)
        data.append({'Ward': s, 'Total_Days': total_sum})
    return pd.DataFrame(data)

# --- Sidebar ---
st.sidebar.markdown("<h2 style='text-align:center; color:#1a3a5f;'>🏥 Device Comparison</h2>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 ไฟล์ที่ 1 (เดือนก่อนหน้า)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 ไฟล์ที่ 2 (เดือนปัจจุบัน)", type=["xlsx"], key="f2")

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.radio("Navigation:", ["📊 Comparison Analytics", "📄 Data Editor (File 2)"])

    # --- คำนวณข้อมูลเปรียบเทียบ ---
    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    
    # Merge ข้อมูลเข้าด้วยกัน
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_File1', '_File2')).fillna(0)
    df_compare['Diff'] = df_compare['Total_Days_File2'] - df_compare['Total_Days_File1']
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_File1']) * 100).fillna(0)

    # --- หน้าเปรียบเทียบ ---
    if page == "📊 Comparison Analytics":
        st.markdown("<h1 class='main-title'>📊 Executive Comparison Dashboard</h1>", unsafe_allow_html=True)

        # KPI Metrics
        total_1 = df_compare['Total_Days_File1'].sum()
        total_2 = df_compare['Total_Days_File2'].sum()
        total_diff = total_2 - total_1
        total_growth = (total_diff / total_1 * 100) if total_1 != 0 else 0

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi-box'><p style='color:#6c757d; font-weight:bold;'>PREVIOUS TOTAL</p><h1 style='color:#1a3a5f;'>{int(total_1):,}</h1></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-box'><p style='color:#6c757d; font-weight:bold;'>CURRENT TOTAL</p><h1 style='color:#1a3a5f;'>{int(total_2):,}</h1></div>", unsafe_allow_html=True)
        with c3:
            color = "#27ae60" if total_diff >= 0 else "#e74c3c"
            st.markdown(f"<div class='kpi-box' style='border-top-color:{color};'><p style='color:#6c757d; font-weight:bold;'>VARIANCE</p><h1 style='color:{color};'>{int(total_diff):+,}</h1><p style='color:{color}; font-weight:bold;'>({total_growth:+.1f}%)</p></div>", unsafe_allow_html=True)

        st.markdown("---")
        
        # กราฟเปรียบเทียบ
        st.subheader("📈 Ward Comparison: File 1 vs File 2")
        fig_compare = px.bar(
            df_compare, x='Ward', y=['Total_Days_File1', 'Total_Days_File2'],
            barmode='group',
            labels={'value': 'Total Days', 'variable': 'File Source'},
            title="เปรียบเทียบภาระงานรายวอร์ด",
            color_discrete_sequence=['#adb5bd', '#1a3a5f'] # เทา vs น้ำเงินเข้ม
        )
        st.plotly_chart(fig_compare, use_container_width=True)

        # ตารางสรุปผลต่าง
        st.subheader("📋 Comparison Summary Table")
        st.dataframe(df_compare.style.format({'% Growth': '{:.2f}%', 'Total_Days_File1': '{:,.0f}', 'Total_Days_File2': '{:,.0f}', 'Diff': '{:+,.0f}'}), use_container_width=True)

        # ปุ่ม Export Comparison
        buf_comp = io.BytesIO()
        df_compare.to_excel(buf_comp, index=False)
        st.download_button("📥 ดาวน์โหลดรายงานเปรียบเทียบ (Excel)", data=buf_comp.getvalue(), file_name="Device_Comparison_Report.xlsx")

    # --- หน้า Data Editor (แสดงข้อมูลไฟล์ล่าสุด) ---
    elif page == "📄 Data Editor (File 2)":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("เลือกแผนก (Ward):", excel_2.sheet_names)
        st.markdown(f"<h1 class='main-title'>📄 แผนก: {selected_sheet} (Current File)</h1>", unsafe_allow_html=True)
        
        df = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df, use_container_width=True, hide_index=True)
        
        total_val, device_cols = get_safe_total(edited_df)
        st.subheader("📊 สรุปยอดอุปกรณ์รายแผนก")
        if device_cols:
            cols_grid = st.columns(len(device_cols))
            sum_vals = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum()
            for i, col_name in enumerate(device_cols):
                with cols_grid[i % len(device_cols)]:
                    st.metric(label=col_name, value=f"{int(sum_vals[col_name]):,}")

else:
    st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏦 Device Comparative Analysis</h1><p>Luxury Web-based Data Management System</p><p style='color:#6c757d;'>กรุณาอัปโหลดไฟล์ Excel 2 ชุดเพื่อเปรียบเทียบข้อมูล</p></div>", unsafe_allow_html=True)
