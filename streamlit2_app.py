import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Executive Luxury Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
    .stApp { background-color: #f0f4f8; }
    
    /* Header Style */
    .main-title { 
        color: #0f172a; font-size: 2.5rem; font-weight: 800; 
        text-align: center; margin-bottom: 30px;
        background: linear-gradient(90deg, #1e293b, #334155);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }

    /* Executive Card Style */
    .kpi-card {
        background: white; padding: 2rem; border-radius: 24px; text-align: center;
        box-shadow: 0 10px 30px rgba(0,0,0,0.04);
        border: 1px solid rgba(255,255,255,0.8);
        transition: all 0.3s ease;
    }
    .kpi-card:hover { transform: translateY(-5px); box-shadow: 0 20px 40px rgba(0,0,0,0.08); }
    .kpi-label { color: #64748b; font-size: 0.9rem; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; }
    .kpi-value { color: #1e293b; font-size: 2.2rem; font-weight: 800; margin: 10px 0; }
    .kpi-delta { font-size: 1rem; font-weight: 700; }

    /* Button Style */
    .stButton>button {
        border-radius: 12px; background: #1e293b; color: white;
        font-weight: 700; letter-spacing: 0.5px; border: none; padding: 0.5rem 2rem;
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
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2413/2413110.png", width=100)
st.sidebar.markdown("<h2 style='text-align:center; color:#1e293b;'>MEDICAL SYSTEM</h2>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 Month 1 (Previous)", type=["xlsx"])
file_2 = st.sidebar.file_uploader("📂 Month 2 (Current)", type=["xlsx"])

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.selectbox("Navigation", ["📊 Comparison Analytics", "📄 Data Editor (File 2)"])

    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    # --- PAGE 1: COMPARISON ---
    if page == "📊 Comparison Analytics":
        st.markdown("<h1 class='main-title'>Executive Comparison Insights</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff = t2 - t1
        growth = int((diff / t1 * 100)) if t1 != 0 else 0
        color_code = "#10b981" if diff >= 0 else "#ef4444"

        # KPI Luxury Row
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Previous Month</p><p class='kpi-value'>{t1:,}</p><p style='color:#94a3b8;'>Total Days</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Current Month</p><p class='kpi-value' style='color:#2563eb;'>{t2:,}</p><p style='color:#94a3b8;'>Total Days</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Variance</p><p class='kpi-value' style='color:{color_code};'>{diff:+,}</p><p class='kpi-delta' style='color:{color_code};'>{growth:+,}% from last month</p></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Charts Row
        col_l, col_r = st.columns([7, 3])
        with col_l:
            st.subheader("🏥 Device Usage by Department")
            fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                             barmode='group', text_auto=',.0f',
                             color_discrete_sequence=['#e2e8f0', '#1e293b'])
            fig_bar.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', legend_title="")
            fig_bar.update_traces(textposition='outside', marker_line_width=0)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_r:
            st.subheader("📈 Growth Trend (%)")
            df_compare['Status'] = df_compare['% Growth'].apply(lambda x: 'Up' if x >= 0 else 'Down')
            fig_growth = px.bar(df_compare, x='% Growth', y='Ward', orientation='h',
                                color='Status', color_discrete_map={'Up': '#10b981', 'Down': '#ef4444'},
                                text_auto='d')
            fig_growth.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_growth, use_container_width=True)

        # Export Comparison
        st.markdown("---")
        st.subheader("📋 Comparison Data Table")
        st.dataframe(df_compare.drop(columns=['Status']).style.format(precision=0), use_container_width=True)
        
        buf_comp = io.BytesIO()
        df_compare.drop(columns=['Status']).to_excel(buf_comp, index=False)
        st.download_button("📥 Export Comparison Summary", data=buf_comp.getvalue(), file_name="Executive_Summary.xlsx")

    # --- PAGE 2: EDITOR & BULK EXPORT ---
    elif page == "📄 Data Editor (File 2)":
        excel_2 = pd.ExcelFile(file_2)
        sheet_names = excel_2.sheet_names
        selected_sheet = st.sidebar.selectbox("Select Ward:", sheet_names)
        
        st.markdown(f"<h1 class='main-title'>Data Management: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        # Live Stats for Selected Ward
        st.markdown("---")
        total_val, device_cols = get_safe_total(edited_df)
        if device_cols:
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{summary_row[col]:,}")
            m_cols[-1].metric("TOTAL DEVICE DAYS", f"{total_val:,}", delta_color="normal")

        # --- BULK EXPORT ALL SHEETS ---
        st.markdown("---")
        st.subheader("📥 Bulk Export Options (Current File)")
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("📥 Prepare All Wards (รวบรวมทุกแผนก)"):
                with st.spinner('กำลังประมวลผลทุก Sheet...'):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        for s_name in sheet_names:
                            temp_df = pd.read_excel(file_2, sheet_name=s_name).dropna(how='all')
                            t_val, d_cols = get_safe_total(temp_df)
                            if d_cols:
                                # เพิ่มแถวสรุปยอดท้ายแต่ละ Sheet
                                for c in d_cols: temp_df[c] = pd.to_numeric(temp_df[c], errors='coerce').fillna(0)
                                sum_line = temp_df[d_cols].sum().to_frame().T
                                sum_line.index = [len(temp_df)]
                                temp_df = pd.concat([temp_df, sum_line])
                                temp_df.iloc[-1, 0] = "GRAND TOTAL"
                            temp_df.to_excel(writer, sheet_name=s_name, index=False)
                    st.download_button("📥 คลิกเพื่อดาวน์โหลดทุกแผนก (Single File)", 
                                       data=output.getvalue(), 
                                       file_name=f"Full_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")
        
        with col_btn2:
            st.info("💡 ปุ่มด้านบนจะรวมข้อมูลจากทุก Sheet (ICU, CCU, Ward ต่างๆ) เข้ามาเป็นไฟล์ Excel เดียวกันพร้อมยอดรวมท้ายตาราง")

else:
    st.markdown("""
        <div style='text-align:center; margin-top:100px;'>
            <h1 style='color:#1e293b; font-size:3rem;'>🏥 Smart Executive Analytics</h1>
            <p style='color:#64748b; font-size:1.2rem;'>กรุณาอัปโหลดไฟล์ Excel 2 ช่วงเวลาเพื่อเปรียบเทียบประสิทธิภาพ</p>
        </div>
        """, unsafe_allow_html=True)
