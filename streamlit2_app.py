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
    d_cols = [c for c in df_in.columns if any(k.lower() in str(c).lower() for k in keywords)]
    if not d_cols: return 0, []
    df_c = df_in.copy()
    for c in d_cols: 
        df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
    return int(df_c[d_cols].values.sum()), d_cols

def process_file_summary(uploaded_file):
    excel_file = pd.ExcelFile(uploaded_file)
    data = []
    for s in excel_file.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
        total_sum, _ = get_safe_total(df)
        data.append({'Ward': s, 'Total_Days': int(total_sum)})
    return pd.DataFrame(data)

# --- Sidebar ---
st.sidebar.markdown("<h2 style='text-align:center; color:#1a3a5f;'>🏥 Device Management</h2>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 ไฟล์ที่ 1 (เดือนก่อนหน้า)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 ไฟล์ที่ 2 (เดือนปัจจุบัน)", type=["xlsx"], key="f2")

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.radio("Navigation:", ["📊 Comparison Analytics", "📄 Data Editor (File 2)"])

    # ประมวลผลข้อมูลเปรียบเทียบ
    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    
    # คำนวณส่วนต่างเป็นจำนวนเต็ม
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    if page == "📊 Comparison Analytics":
        st.markdown("<h1 class='main-title'>📊 Executive Comparison Dashboard</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff = t2 - t1
        growth = int((diff / t1 * 100)) if t1 != 0 else 0

        c1, c2, c3 = st.columns(3)
        c1.markdown(f"<div class='kpi-box'><b>PREVIOUS (M1)</b><h2>{t1:,}</h2></div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='kpi-box'><b>CURRENT (M2)</b><h2>{t2:,}</h2></div>", unsafe_allow_html=True)
        color = "#27ae60" if diff >= 0 else "#e74c3c"
        c3.markdown(f"<div class='kpi-box' style='border-top-color:{color};'><b>VARIANCE</b><h2 style='color:{color};'>{diff:+,} ({growth:+,}%)</h2></div>", unsafe_allow_html=True)

        st.markdown("---")
        
        # กราฟเปรียบเทียบจำนวนวัน (ใส่ตัวเลขบนกราฟ)
        st.subheader("📈 Total Days Comparison (M1 vs M2)")
        fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                         barmode='group', text_auto=',.0f', # ใส่ตัวเลขจำนวนเต็ม
                         color_discrete_sequence=['#adb5bd', '#1a3a5f'],
                         labels={'value': 'Total Days', 'variable': 'Period'})
        fig_bar.update_traces(textposition='outside')
        st.plotly_chart(fig_bar, use_container_width=True)

        # กราฟ % Growth (ใส่ตัวเลขบนกราฟ)
        st.subheader("📉 % Growth by Ward")
        df_compare['Color_Status'] = df_compare['% Growth'].apply(lambda x: 'Positive' if x >= 0 else 'Negative')
        fig_growth = px.bar(df_compare, x='Ward', y='% Growth', color='Color_Status',
                            text_auto='d', # แสดงเป็นเลขจำนวนเต็ม
                            color_discrete_map={'Positive': '#27ae60', 'Negative': '#e74c3c'})
        fig_growth.update_traces(textposition='outside')
        st.plotly_chart(fig_growth, use_container_width=True)

        # ตารางข้อมูลและปุ่ม Export
        st.subheader("📋 Comparison Summary")
        st.dataframe(df_compare.drop(columns=['Color_Status']).style.format(precision=0), use_container_width=True)
        
        buf_comp = io.BytesIO()
        df_compare.drop(columns=['Color_Status']).to_excel(buf_comp, index=False)
        st.download_button("📥 Export Comparison Report", data=buf_comp.getvalue(), file_name="Comparison_Report.xlsx")

    elif page == "📄 Data Editor (File 2)":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("เลือกแผนก (Ward):", excel_2.sheet_names)
        st.markdown(f"<h1 class='main-title'>📄 Editor: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        total_val, device_cols = get_safe_total(edited_df)
        
        if device_cols:
            st.subheader("🧮 Ward Summary (Live Calculation)")
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            
            # Metric Display
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{summary_row[col]:,}")
            m_cols[-1].metric("GRAND TOTAL", f"{total_val:,}", delta_color="normal")
            
            # ปุ่ม Export เฉพาะ Ward นี้
            st.markdown("---")
            buf_ward = io.BytesIO()
            # เพิ่มบรรทัด Grand Total ลงในไฟล์ที่ดาวน์โหลด
            export_df = edited_df.copy()
            for c in device_cols: export_df[c] = pd.to_numeric(export_df[c], errors='coerce').fillna(0)
            summary_df = pd.DataFrame(summary_row).T
            summary_df.index = [len(export_df)]
            final_export = pd.concat([export_df, summary_df], axis=0)
            final_export.iloc[-1, 0] = "GRAND TOTAL"
            
            final_export.to_excel(buf_ward, index=False)
            st.download_button(f"📥 Export {selected_sheet} Data", data=buf_ward.getvalue(), file_name=f"Report_{selected_sheet}.xlsx")
        else:
            st.warning("ไม่พบคอลัมน์อุปกรณ์ใน Sheet นี้")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือนเพื่อเริ่มการวิเคราะห์เปรียบเทียบ")
