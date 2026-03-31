import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Ultra Modern Executive Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
    
    .stApp { background-color: #fcfcfc; }
    
    /* Header Style */
    .main-title { 
        color: #0f172a; font-size: 3rem; font-weight: 800; 
        text-align: center; margin-bottom: 40px;
        letter-spacing: -1px;
    }

    /* Modern KPI Card */
    .kpi-card {
        background: white; padding: 2rem; border-radius: 24px; text-align: center;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.04);
        border: 2px solid #f1f5f9;
        transition: all 0.3s ease;
    }
    .kpi-label { color: #64748b; font-size: 1.1rem; font-weight: 700; text-transform: uppercase; }
    .kpi-value { color: #1e293b; font-size: 3.8rem; font-weight: 800; margin: 0.5rem 0; }
    .kpi-delta { font-size: 1.4rem; font-weight: 800; padding: 5px 15px; border-radius: 50px; display: inline-block; }

    /* Custom Button */
    .stButton>button {
        border-radius: 12px; background: #0f172a; color: white;
        font-weight: 700; border: none; padding: 0.8rem 2rem; font-size: 1rem;
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
st.sidebar.markdown("<h2 style='text-align:center;'>🏥 MANAGEMENT</h2>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 Month 1 (Previous)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 Month 2 (Current)", type=["xlsx"], key="f2")

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.selectbox("เมนูนำทาง", ["📊 Comparison Dashboard", "📄 Interactive Editor"])

    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    
    # คำนวณความแม่นยำสูง
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    # เตรียมข้อมูล Label พิเศษสำหรับกราฟ (โชว์ทั้งตัวเลขและ %)
    df_compare['M2_Label'] = df_compare.apply(lambda x: f"{int(x['Total_Days_M2']):,} ({int(x['% Growth']):+d}%)", axis=1)

    if page == "📊 Comparison Dashboard":
        st.markdown("<h1 class='main-title'>Device Usage Comparison</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff = t2 - t1
        growth = int((diff / t1 * 100)) if t1 != 0 else 0
        bg_color = "#d1fae5" if diff >= 0 else "#fee2e2"
        text_color = "#065f46" if diff >= 0 else "#991b1b"

        # KPI Row
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Previous Month</p><p class='kpi-value'>{t1:,}</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Current Month</p><p class='kpi-value' style='color:#4f46e5;'>{t2:,}</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
                <div class='kpi-card'>
                    <p class='kpi-label'>Variance</p>
                    <p class='kpi-value' style='color:{text_color};'>{diff:+,}</p>
                    <span class='kpi-delta' style='background:{bg_color}; color:{text_color};'>{growth:+,}%</span>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # --- กราฟแท่งสีใหม่ ดีไซน์ทันสมัย ---
        st.subheader("📊 Department Comparison with Growth %")
        
        # ใช้ Plotly แบบปรับแต่งพิเศษ
        fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                         barmode='group',
                         color_discrete_sequence=['#cbd5e1', '#4f46e5'], # สีเทาอ่อน vs สีน้ำเงิน Indigo
                         labels={'value': 'Total Days', 'variable': 'Month'})
        
        # ใส่ตัวเลขบนแท่งกราฟ (แท่ง M1 โชว์เลข / แท่ง M2 โชว์เลข + %)
        fig_bar.data[0].text = df_compare['Total_Days_M1'].apply(lambda x: f"{int(x):,}")
        fig_bar.data[1].text = df_compare['M2_Label']
        
        fig_bar.update_traces(textposition='outside', textfont_size=14, textfont_color="#1e293b", textfont_family="Sarabun")
        
        fig_bar.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=14)),
            xaxis=dict(title="", tickfont=dict(size=16, color="#0f172a", family="Sarabun")),
            yaxis=dict(title="Total Days", showgrid=True, gridcolor="#f1f5f9", tickfont=dict(size=14)),
            margin=dict(t=50, b=50, l=0, r=0)
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # ตารางข้อมูลและปุ่ม Export
        st.markdown("---")
        st.subheader("📋 Comparison Summary Report")
        st.dataframe(df_compare[['Ward', 'Total_Days_M1', 'Total_Days_M2', 'Diff', '% Growth']].style.format(precision=0), use_container_width=True)
        
        buf_comp = io.BytesIO()
        df_compare.drop(columns=['M2_Label']).to_excel(buf_comp, index=False)
        st.download_button("📥 Export Detailed Report (Excel)", data=buf_comp.getvalue(), file_name="Executive_Comparison.xlsx")

    elif page == "📄 Interactive Editor":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("Select Ward:", excel_2.sheet_names)
        st.markdown(f"<h1 class='main-title'>Management: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        total_val, device_cols = get_safe_total(edited_df)
        if device_cols:
            st.markdown("### 🧮 Live Ward Summary")
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{summary_row[col]:,}")
            m_cols[-1].metric("GRAND TOTAL", f"{total_val:,}", delta_color="normal")
            
            # Bulk Export Option
            st.markdown("---")
            if st.button("📥 Prepare Bulk Hospital Export"):
                with st.spinner('ประมวลผลทุกแผนก...'):
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
                    st.download_button("📥 Click to Download Full Report", data=bulk_buf.getvalue(), file_name="Full_Hospital_Report.xlsx")
        else:
            st.warning("ไม่พบคอลัมน์อุปกรณ์ใน Sheet นี้")

else:
    st.markdown("""
        <div style='text-align:center; margin-top:100px;'>
            <h1 style='color:#0f172a; font-size:4rem; font-weight:800;'>Device Analytics Hub</h1>
            <p style='color:#64748b; font-size:1.5rem;'>กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการวิเคราะห์เปรียบเทียบ</p>
        </div>
    """, unsafe_allow_html=True)
