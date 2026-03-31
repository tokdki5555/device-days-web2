import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Ultra-Clear Executive Styling
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
        color: #0f172a; font-size: 3.8rem; font-weight: 800; 
        text-align: center; margin-bottom: 50px;
    }

    /* Mega KPI Card - ใหญ่ ชัดเจน อ่านง่ายจากระยะไกล */
    .kpi-container {
        background: #f8fafc;
        padding: 3.5rem 2rem;
        border-radius: 35px;
        text-align: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        border: 2px solid #e2e8f0;
    }
    
    .kpi-label { 
        color: #475569; font-size: 1.5rem; font-weight: 700; 
        text-transform: uppercase; margin-bottom: 10px;
    }
    
    .kpi-value { 
        color: #000000; font-size: 5.8rem; font-weight: 800; 
        margin: 0; line-height: 1;
    }
    
    .kpi-delta-tag { 
        margin-top: 1.5rem; padding: 10px 25px; border-radius: 100px; 
        display: inline-block; font-weight: 800; font-size: 2rem; 
    }

    /* ปุ่มกดตัวหนาชัดเจน */
    .stButton>button {
        border-radius: 15px; background: #0f172a; color: white !important;
        font-weight: 800; border: none; padding: 1.2rem 2rem; font-size: 1.4rem;
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

def color_growth(val):
    color = '#059669' if val > 0 else '#dc2626' if val < 0 else '#64748b'
    return f'color: {color}; font-weight: 800; font-size: 1.3rem;'

# --- Sidebar ---
st.sidebar.markdown("<h1 style='text-align:center;'>🏥 SYSTEM</h1>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 เดือนที่ 1 (Previous Month)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 เดือนที่ 2 (Current Month)", type=["xlsx"], key="f2")

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.selectbox("🎯 เลือกดูข้อมูล", ["📊 Utilization Analytics", "📄 Data Editor"])

    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    if page == "📊 Utilization Analytics":
        st.markdown("<h1 class='main-title'>Utilization Analytics</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff, growth = t2 - t1, int((t2-t1)/t1*100) if t1 != 0 else 0
        bg_color = "#d1fae5" if diff >= 0 else "#fee2e2"
        text_color = "#047857" if diff >= 0 else "#b91c1c"

        # --- KPI SECTION ---
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi-container'><p class='kpi-label'>Previous Month</p><p class='kpi-value'>{t1:,}</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-container'><p class='kpi-label'>Current Month</p><p class='kpi-value' style='color:#2563eb;'>{t2:,}</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""<div class='kpi-container'>
                <p class='kpi-label'>Variance</p>
                <p class='kpi-value' style='color:{text_color};'>{diff:+,}</p>
                <div class='kpi-delta-tag' style='background:{bg_color}; color:{text_color};'>{growth:+,}% Change</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # --- กราฟแท่งเปรียบเทียบ (M1 vs M2) ---
        st.subheader("📊 เปรียบเทียบจำนวนวันการใช้งานรายแผนก")
        fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                         barmode='group', text_auto=',.0f',
                         color_discrete_sequence=['#cbd5e1', '#0f172a'])
        fig_bar.update_layout(
            font=dict(size=18, family="Sarabun", color="#000"),
            plot_bgcolor='rgba(0,0,0,0)', 
            xaxis=dict(tickfont=dict(size=20, color="#000", weight='bold')),
            legend=dict(font=dict(size=18), orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        fig_bar.update_traces(textposition='outside', textfont_size=22, textfont_weight="bold", textfont_color="#000")
        st.plotly_chart(fig_bar, use_container_width=True)

        # --- กราฟอัตราการเติบโต (Growth %) ---
        st.subheader("📈 อัตราการเปลี่ยนแปลงรายแผนก (%)")
        df_compare['Status'] = df_compare['% Growth'].apply(lambda x: 'Up' if x >= 0 else 'Down')
        fig_growth = px.bar(df_compare, x='Ward', y='% Growth', color='Status',
                            text_auto='d', color_discrete_map={'Up': '#10b981', 'Down': '#ef4444'})
        fig_growth.update_layout(
            font=dict(size=18, family="Sarabun", color="#000"),
            plot_bgcolor='rgba(0,0,0,0)', showlegend=False,
            xaxis=dict(tickfont=dict(size=20, color="#000", weight='bold'))
        )
        fig_growth.update_traces(textposition='outside', textfont_size=22, textfont_weight="bold", textfont_color="#000")
        st.plotly_chart(fig_growth, use_container_width=True)

        # --- ตารางข้อมูลและปุ่ม Export (หน้า Dashboard) ---
        st.markdown("---")
        st.subheader("📋 รายละเอียดข้อมูลเปรียบเทียบ")
        styled_df = df_compare[['Ward', 'Total_Days_M1', 'Total_Days_M2', 'Diff', '% Growth']].style.format(precision=0).applymap(color_growth, subset=['% Growth'])
        st.dataframe(styled_df, use_container_width=True)
        
        # ปุ่ม Export รายงาน Comparison
        buf_comp = io.BytesIO()
        df_compare.drop(columns=['Status']).to_excel(buf_comp, index=False)
        st.download_button("📥 ดาวน์โหลดรายงานเปรียบเทียบ (Excel)", data=buf_comp.getvalue(), file_name="Utilization_Comparison.xlsx")

    elif page == "📄 Data Editor":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("เลือกวอร์ด:", excel_2.sheet_names)
        st.markdown(f"<h1 class='main-title'>Editor: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        total_val, device_cols = get_safe_total(edited_df)
        if device_cols:
            st.markdown("### 🧮 Live สรุปยอด")
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{summary_row[col]:,}")
            m_cols[-1].metric("GRAND TOTAL", f"{total_val:,}")
            
            # --- Bulk Export (รวบรวมทุกวอร์ด) ---
            st.markdown("---")
            if st.button("📥 รวบรวมข้อมูลทุกวอร์ด (Bulk Export)"):
                try:
                    bulk_buf = io.BytesIO()
                    with pd.ExcelWriter(bulk_buf, engine='xlsxwriter') as writer:
                        for s in excel_2.sheet_names:
                            s_df = pd.read_excel(file_2, sheet_name=s).dropna(how='all')
                            _, s_cols = get_safe_total(s_df)
                            if s_cols:
                                for c in s_cols: s_df[c] = pd.to_numeric(s_df[c], errors='coerce').fillna(0).astype(int)
                                s_sum = s_df[s_cols].sum().to_frame().T
                                s_sum.index = [len(s_df)]; s_df = pd.concat([s_df, s_sum])
                                s_df.iloc[-1, 0] = "GRAND TOTAL"
                            s_df.to_excel(writer, sheet_name=s, index=False)
                    st.download_button("📥 คลิกเพื่อดาวน์โหลดรายงานโรงพยาบาล", data=bulk_buf.getvalue(), file_name="Full_Device_Report.xlsx")
                except Exception as e:
                    st.error(f"Error: กรุณาตรวจสอบไฟล์ requirements.txt ว่ามี 'xlsxwriter' หรือไม่? ({e})")

else:
    st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏦 EXECUTIVE SYSTEM</h1><p>กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มวิเคราะห์</p></div>", unsafe_allow_html=True)
