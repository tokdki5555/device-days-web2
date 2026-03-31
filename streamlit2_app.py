import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Ultra-Modern Executive Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700;800&display=swap');
    
    /* บังคับตัวหนังสือดำเข้มและชัดเจนที่สุด */
    html, body, [class*="css"] { 
        font-family: 'Sarabun', sans-serif; 
        color: #000000 !important; 
    }
    
    .stApp { background-color: #ffffff; }
    
    .main-title { 
        color: #0f172a; font-size: 3.5rem; font-weight: 800; 
        text-align: center; margin-bottom: 50px;
    }

    /* MEGA KPI Card - ปรับปรุงให้ใหญ่ยักษ์และชัดเจนตามรูป */
    .mega-kpi-container {
        background: #f8fafc;
        padding: 4.5rem 1rem;
        border-radius: 40px;
        text-align: center;
        box-shadow: 0 15px 30px rgba(0,0,0,0.06);
        border: 2px solid #e2e8f0;
        min-height: 380px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .mega-label { 
        color: #475569; font-size: 2rem; font-weight: 700; 
        text-transform: uppercase; margin-bottom: 20px;
        letter-spacing: 2px;
    }
    
    .mega-value { 
        color: #000000; font-size: 7.5rem; font-weight: 800; 
        margin: 0; line-height: 0.9; letter-spacing: -3px;
    }
    
    .mega-delta-tag { 
        margin-top: 2.5rem; padding: 15px 40px; border-radius: 100px; 
        display: inline-block; font-weight: 800; font-size: 2.5rem; 
    }

    /* ปุ่มกดตัวหนาชัดเจน */
    .stButton>button {
        border-radius: 20px; background: #0f172a; color: white !important;
        font-weight: 800; border: none; padding: 1.5rem 2rem; font-size: 1.5rem;
        width: 100%; box-shadow: 0 10px 15px rgba(15, 23, 42, 0.2);
    }
    </style>
    """, unsafe_allow_html=True)

# ฟังก์ชันคำนวณและดึงข้อมูล
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
st.sidebar.markdown("<h1 style='text-align:center;'>🏥 SYSTEM</h1>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 เดือนที่ 1 (Previous Month)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 เดือนที่ 2 (Current Month)", type=["xlsx"], key="f2")

if file_1 and file_2:
    page = st.sidebar.selectbox("🎯 เลือกดูข้อมูล", ["📊 Utilization Analytics", "📄 Data Editor"])

    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    if page == "📊 Utilization Analytics":
        st.markdown("<h1 class='main-title'>Device Utilization Analytics</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff, growth = t2 - t1, int((t2-t1)/t1*100) if t1 != 0 else 0
        bg_color = "#d1fae5" if diff >= 0 else "#fee2e2"
        text_color = "#047857" if diff >= 0 else "#b91c1c"

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='mega-kpi-container'><p class='mega-label'>Previous Total</p><p class='mega-value'>{t1:,}</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='mega-kpi-container'><p class='mega-label'>Current Total</p><p class='mega-value' style='color:#2563eb;'>{t2:,}</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""<div class='mega-kpi-container'>
                <p class='mega-label'>Variance</p>
                <p class='mega-value' style='color:{text_color};'>{diff:+,}</p>
                <div class='mega-delta-tag' style='background:{bg_color}; color:{text_color};'>{growth:+,}% Change</div>
            </div>""", unsafe_allow_html=True)

        # ปุ่มดาวน์โหลดรายงาน (แก้ไข Syntax Error แล้ว)
        st.markdown("---")
        if st.button("📥 Export Full Hospital Report (รวบรวมทุกแผนก)"):
            with st.spinner('กำลังรวบรวมข้อมูล...'):
                excel_2 = pd.ExcelFile(file_2)
                bulk_buf = io.BytesIO()
                with pd.ExcelWriter(bulk_buf, engine='xlsxwriter') as writer:
                    for s in excel_2.sheet_names:
                        s_df = pd.read_excel(file_2, sheet_name=s).dropna(how='all')
                        _, s_cols = get_safe_total(s_df)
                        if s_cols:
                            for c in s_cols: s_df[c] = pd.to_numeric(s_df[c], errors='coerce').fillna(0).astype(int)
                            s_sum = s_df[s_cols].sum().to_frame().T
                            s_sum.index = [len(s_df)]
                            s_df = pd.concat([s_df, s_sum])
                            s_df.iloc[-1, 0] = "GRAND TOTAL"
                        s_df.to_excel(writer, sheet_name=s, index=False)
                st.download_button("✅ คลิกที่นี่เพื่อดาวน์โหลดไฟล์สมบูรณ์", data=bulk_buf.getvalue(), file_name="Hospital_Report.xlsx")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือน")
