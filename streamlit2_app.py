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
    }
    .kpi-box {
        background: white; padding: 25px; border-radius: 20px; text-align: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05); border-top: 5px solid #1a3a5f;
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
st.sidebar.markdown("<h2 style='text-align:center; color:#1a3a5f;'>🏥 Device Management</h2>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 ไฟล์ที่ 1 (เดือนก่อนหน้า)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 ไฟล์ที่ 2 (เดือนปัจจุบัน)", type=["xlsx"], key="f2")

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.radio("Navigation:", ["📊 Comparison Analytics", "📄 Data Editor (File 2)"])

    # ประมวลผลข้อมูล
    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_File1', '_File2')).fillna(0)
    df_compare['Diff'] = df_compare['Total_Days_File2'] - df_compare['Total_Days_File1']
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_File1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0)

    if page == "📊 Comparison Analytics":
        st.markdown("<h1 class='main-title'>📊 Executive Comparison Dashboard</h1>", unsafe_allow_html=True)

        # KPI Metrics
        t1, t2 = df_compare['Total_Days_File1'].sum(), df_compare['Total_Days_File2'].sum()
        diff = t2 - t1
        growth = (diff / t1 * 100) if t1 != 0 else 0

        c1, c2, c3 = st.columns(3)
        c1.markdown(f"<div class='kpi-box'><b>PREVIOUS</b><h2>{int(t1):,}</h2></div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='kpi-box'><b>CURRENT</b><h2>{int(t2):,}</h2></div>", unsafe_allow_html=True)
        color = "#27ae60" if diff >= 0 else "#e74c3c"
        c3.markdown(f"<div class='kpi-box' style='border-top-color:{color};'><b>VARIANCE</b><h2 style='color:{color};'>{int(diff):+,} ({growth:+.1f}%)</h2></div>", unsafe_allow_html=True)

        st.markdown("---")
        
        # --- แยกกราฟเป็น 2 ส่วน ---
        col_chart1, col_chart2 = st.columns([6, 4])
        
        with col_chart1:
            st.subheader("📈 Total Days Comparison")
            fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_File1', 'Total_Days_File2'],
                             barmode='group', color_discrete_sequence=['#adb5bd', '#1a3a5f'])
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_chart2:
            st.subheader("📉 % Growth by Ward")
            # สร้างกราฟ % แยกออกมา
            df_compare['Color'] = df_compare['% Growth'].apply(lambda x: 'Positive' if x >= 0 else 'Negative')
            fig_growth = px.bar(df_compare, x='% Growth', y='Ward', orientation='h',
                                color='Color', color_discrete_map={'Positive': '#27ae60', 'Negative': '#e74c3c'},
                                text_auto='.1f')
            st.plotly_chart(fig_growth, use_container_width=True)

        st.subheader("📋 Detailed Comparison Table")
        st.dataframe(df_compare.drop(columns=['Color']).style.background_gradient(subset=['Diff'], cmap='RdYlGn'), use_container_width=True)

    elif page == "📄 Data Editor (File 2)":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("เลือกแผนก (Ward):", excel_2.sheet_names)
        st.markdown(f"<h1 class='main-title'>📄 Editor: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        # --- ส่วนคำนวณรวมในหน้า Editor ---
        st.markdown("---")
        st.subheader("🧮 Live Summary (Current Sheet)")
        
        total_val, device_cols = get_safe_total(edited_df)
        
        if device_cols:
            # สร้าง DataFrame สรุปยอดรายอุปกรณ์
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum()
            summary_df = pd.DataFrame(summary_row).T
            summary_df.index = ["TOTAL SUM"]
            
            # แสดงตารางผลรวมอุปกรณ์ทั้งหมด
            st.write("ตารางคำนวณรวมทุกอุปกรณ์ในแผนก:")
            st.dataframe(summary_df.style.highlight_max(axis=1, color='#d4edda'), use_container_width=True)
            
            # Metric Cards
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{int(summary_row[col]):,}")
            m_cols[-1].metric("GRAND TOTAL", f"{int(total_val):,}", delta_color="normal")
        else:
            st.warning("ไม่พบคอลัมน์อุปกรณ์ (Ventilator, Foley, etc.) ใน Sheet นี้")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือนเพื่อเริ่มต้นการวิเคราะห์เปรียบเทียบ")
