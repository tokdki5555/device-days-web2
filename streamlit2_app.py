import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
    .delta-pos { color: #28a745; font-size: 0.9em; font-weight: bold; }
    .delta-neg { color: #dc3545; font-size: 0.9em; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- Shared Functions ---
keywords = ["Ventilator", "Foley", "Central line", "Port A Cath"]

def get_safe_total(df_in):
    d_cols = [c for c in df_in.columns if any(k.lower() in str(c).lower() for k in keywords)]
    if not d_cols: return 0, []
    df_c = df_in.copy()
    for c in d_cols: df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
    return df_c[d_cols].values.sum(), d_cols

def process_file_summary(uploaded_file):
    excel_file = pd.ExcelFile(uploaded_file)
    ward_data = []
    for s in excel_file.sheet_names:
        df_t = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
        total_sum, _ = get_safe_total(df_t)
        ward_data.append({'Ward': s, 'Total_Days': total_sum})
    return pd.DataFrame(ward_data)

# --- Sidebar ---
st.sidebar.markdown("<h2 style='text-align:center; color:#1a3a5f;'>🏥 Device Analytics Pro</h2>", unsafe_allow_html=True)

page = st.sidebar.radio("Navigation:", ["📄 Data Editor", "📊 Executive Analytics", "🔄 Data Comparison"])

# --- Page 1 & 2 logic ---
if page in ["📄 Data Editor", "📊 Executive Analytics"]:
    uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx"], key="single_upload")
    
    if uploaded_file:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        selected_sheet = st.sidebar.selectbox("เลือกแผนก (Ward):", sheet_names)

        if page == "📄 Data Editor":
            st.markdown(f"<h1 class='main-title'>📄 แผนก: {selected_sheet}</h1>", unsafe_allow_html=True)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet).dropna(how='all')
            for col in df.columns:
                if 'date' in col.lower():
                    try: df[col] = pd.to_datetime(df[col]).dt.date
                    except: pass
            
            edited_df = st.data_editor(df, use_container_width=True, hide_index=True)
            total_val, device_cols = get_safe_total(edited_df)
            
            st.subheader("📊 สรุปยอดอุปกรณ์รายแผนก")
            if device_cols:
                cols_grid = st.columns(len(device_cols))
                sum_vals = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum()
                for i, col_name in enumerate(device_cols):
                    with cols_grid[i % len(device_cols)]:
                        st.metric(label=col_name, value=f"{int(sum_vals[col_name]):,}")

        elif page == "📊 Executive Analytics":
            st.markdown("<h1 class='main-title'>📊 Executive Summary Dashboard</h1>", unsafe_allow_html=True)
            df_stats = process_file_summary(uploaded_file)
            grand_total = df_stats['Total_Days'].sum()
            
            if grand_total > 0:
                c1, c2, c3 = st.columns(3)
                c1.metric("Grand Total", f"{int(grand_total):,}")
                c2.metric("Average / Ward", f"{df_stats['Total_Days'].mean():.1f}")
                
                fig_bar = px.bar(df_stats, x='Ward', y='Total_Days', color='Total_Days', title="Total Days by Ward")
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("ไม่พบข้อมูลตัวเลขในไฟล์")

# --- Page 3: Comparison Mode ---
elif page == "🔄 Data Comparison":
    st.markdown("<h1 class='main-title'>🔄 Comparative Performance Analysis</h1>", unsafe_allow_html=True)
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        file1 = st.file_uploader("📂 ไฟล์ฐาน (Base File - e.g. Month 1)", type=["xlsx"])
    with col_f2:
        file2 = st.file_uploader("📂 ไฟล์เปรียบเทียบ (Target File - e.g. Month 2)", type=["xlsx"])

    if file1 and file2:
        df1 = process_file_summary(file1)
        df2 = process_file_summary(file2)
        
        # Merge data for comparison
        df_comp = pd.merge(df1, df2, on="Ward", suffixes=('_Base', '_Target'), how='outer').fillna(0)
        df_comp['Diff'] = df_comp['Total_Days_Target'] - df_comp['Total_Days_Base']
        df_comp['Growth_%'] = (df_comp['Diff'] / df_comp['Total_Days_Base'] * 100).fillna(0)

        # KPI Metrics
        t1 = df_comp['Total_Days_Base'].sum()
        t2 = df_comp['Total_Days_Target'].sum()
        diff_total = t2 - t1
        pct_total = (diff_total / t1 * 100) if t1 != 0 else 0

        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"<div class='kpi-box'><h3>Base Total</h3><h1>{int(t1):,}</h1></div>", unsafe_allow_html=True)
        with m2:
            st.markdown(f"<div class='kpi-box'><h3>Target Total</h3><h1>{int(t2):,}</h1></div>", unsafe_allow_html=True)
        with m3:
            color = "#28a745" if diff_total >= 0 else "#dc3545"
            arrow = "↑" if diff_total >= 0 else "↓"
            st.markdown(f"<div class='kpi-box' style='border-top-color:{color}'><h3>Variance</h3><h1 style='color:{color}'>{arrow} {abs(diff_total):,.0f}</h1><p class='{'delta-pos' if diff_total >= 0 else 'delta-neg'}'>{pct_total:.1f}% Change</p></div>", unsafe_allow_html=True)

        # Comparative Chart
        st.markdown("### 📊 Side-by-Side Comparison by Ward")
        fig_compare = go.Figure(data=[
            go.Bar(name='Base Period', x=df_comp['Ward'], y=df_comp['Total_Days_Base'], marker_color='#1a3a5f'),
            go.Bar(name='Target Period', x=df_comp['Ward'], y=df_comp['Total_Days_Target'], marker_color='#4a90e2')
        ])
        fig_compare.update_layout(barmode='group', height=500)
        st.plotly_chart(fig_compare, use_container_width=True)

        # Summary Table
        st.markdown("### 📋 Comparison Data Table")
        st.dataframe(df_comp.style.format({'Growth_%': '{:.1f}%', 'Total_Days_Base': '{:,.0f}', 'Total_Days_Target': '{:,.0f}', 'Diff': '{:+.0f}'}), use_container_width=True)

    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ทั้ง 2 เดือนเพื่อดูการเปรียบเทียบข้อมูล")

# --- Welcome Screen ---
if not st.sidebar.file_uploader: # Simplified logic for welcome
     st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏦 Smart Device Analytics</h1><p>Luxury Web-based Data Management System</p></div>", unsafe_allow_html=True)
