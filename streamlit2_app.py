import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Ultra Modern Luxury Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
    
    .stApp { background-color: #fcfcfc; }
    
    /* Header Style */
    .main-title { 
        color: #0f172a; font-size: 2.8rem; font-weight: 800; 
        text-align: center; margin-bottom: 40px;
    }

    /* KPI Card แบบพรีเมียมตัวเลขใหญ่ชัดเจน */
    .kpi-card {
        background: white; padding: 2.2rem; border-radius: 24px; text-align: center;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.04);
        border: 1px solid #f1f5f9;
        transition: all 0.3s ease;
    }
    .kpi-label { color: #64748b; font-size: 1.1rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { color: #1e293b; font-size: 3.8rem; font-weight: 800; margin: 0.5rem 0; }
    .kpi-delta-box { padding: 8px 20px; border-radius: 12px; display: inline-block; font-weight: 800; font-size: 1.4rem; }

    /* Button Style */
    .stButton>button {
        border-radius: 12px; background: #0f172a; color: white;
        font-weight: 700; border: none; padding: 0.8rem 2rem; width: 100%;
    }
    
    /* ปรับแต่งตารางให้ตัวหนังสือชัด */
    .stDataFrame { border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; }
    </style>
    """, unsafe_allow_html=True)

# --- Keywords Configuration ---
keywords = ["Ventilator", "Foley", "Central line", "Port A Cath"]

def get_safe_total(df_in):
    """คำนวณผลรวมโดยดึงเฉพาะคอลัมน์อุปกรณ์และลบทศนิยม"""
    d_cols = [c for c in df_in.columns if any(k.lower() in str(c).lower() for k in keywords)]
    if not d_cols: return 0, []
    df_c = df_in.copy()
    for c in d_cols: 
        df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
    return int(df_c[d_cols].values.sum()), d_cols

def process_file_summary(uploaded_file):
    """สรุปยอดรวมแต่ละแผนก (Sheet)"""
    excel_file = pd.ExcelFile(uploaded_file)
    data = []
    for s in excel_file.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
        total_sum, _ = get_safe_total(df)
        data.append({'Ward': s, 'Total_Days': int(total_sum)})
    return pd.DataFrame(data)

# --- Sidebar ---
st.sidebar.markdown("<h2 style='text-align:center; color:#0f172a;'>🏥 DEVICE SYSTEM</h2>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 เดือนที่ 1 (Previous)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 เดือนที่ 2 (Current)", type=["xlsx"], key="f2")

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.selectbox("เมนูนำทาง", ["📊 Comparison Dashboard", "📄 Interactive Editor"])

    # ประมวลผลข้อมูลเปรียบเทียบ
    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    
    # คำนวณส่วนต่างและ % (Integer Only)
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    if page == "📊 Comparison Dashboard":
        st.markdown("<h1 class='main-title'>Executive Device Analytics</h1>", unsafe_allow_html=True)

        # คำนวณ Grand Total
        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff = t2 - t1
        growth = int((diff / t1 * 100)) if t1 != 0 else 0
        bg_color = "#dcfce7" if diff >= 0 else "#fee2e2"
        text_color = "#15803d" if diff >= 0 else "#b91c1c"

        # KPI Row
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Previous Total</p><p class='kpi-value'>{t1:,}</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>Current Total</p><p class='kpi-value' style='color:#4f46e5;'>{t2:,}</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
                <div class='kpi-card'>
                    <p class='kpi-label'>Variance</p>
                    <p class='kpi-value' style='color:{text_color};'>{diff:+,}</p>
                    <div class='kpi-delta-box' style='background:{bg_color}; color:{text_color};'>{growth:+,}% Change</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- กราฟแยกข้อมูล: กราฟแท่งเปรียบเทียบ ---
        st.subheader("📊 จำนวนวันการใช้งานแยกตามแผนก (M1 vs M2)")
        fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                         barmode='group', text_auto=',.0f',
                         color_discrete_sequence=['#cbd5e1', '#1e293b']) # เทาอ่อน vs น้ำเงินเข้ม
        fig_bar.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                              legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        fig_bar.update_traces(textposition='outside', textfont_size=16, textfont_weight="bold")
        st.plotly_chart(fig_bar, use_container_width=True)

        # --- กราฟแยกข้อมูล: กราฟ % Growth ---
        st.subheader("📈 อัตราการเติบโตรายแผนก (%)")
        df_compare['Status'] = df_compare['% Growth'].apply(lambda x: 'Positive' if x >= 0 else 'Negative')
        fig_growth = px.bar(df_compare, x='% Growth', y='Ward', orientation='h',
                            color='Status', color_discrete_map={'Positive': '#10b981', 'Negative': '#ef4444'},
                            text_auto='d')
        fig_growth.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False)
        fig_growth.update_traces(textposition='outside', textfont_size=16, textfont_weight='bold')
        st.plotly_chart(fig_growth, use_container_width=True)

        # ตารางสรุปข้อมูล
        st.markdown("---")
        st.subheader("📋 ตารางข้อมูลเปรียบเทียบฉบับสมบูรณ์")
        st.dataframe(df_compare[['Ward', 'Total_Days_M1', 'Total_Days_M2', 'Diff', '% Growth']].style.format(precision=0), use_container_width=True)
        
        buf_comp = io.BytesIO()
        df_compare.drop(columns=['Status']).to_excel(buf_comp, index=False)
        st.download_button("📥 Export Comparison Report (Excel)", data=buf_comp.getvalue(), file_name="Executive_Summary.xlsx")

    elif page == "📄 Interactive Editor":
        excel_2 = pd.ExcelFile(file_2)
        sheet_names = excel_2.sheet_names
        selected_sheet = st.sidebar.selectbox("เลือกแผนกที่ต้องการแก้ไข:", sheet_names)
        st.markdown(f"<h1 class='main-title'>Data Editor: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        # ส่วนแสดงผล Metric รายแผนก
        total_val, device_cols = get_safe_total(edited_df)
        if device_cols:
            st.markdown("### 🧮 Live Ward Totals")
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{summary_row[col]:,}")
            m_cols[-1].metric("GRAND TOTAL", f"{total_val:,}")
            
            # Export Options
            st.markdown("---")
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                if st.button("📥 Prepare All Wards (รวบรวมทุกแผนก)"):
                    with st.spinner('กำลังประมวลผล...'):
                        bulk_buf = io.BytesIO()
                        with pd.ExcelWriter(bulk_buf, engine='xlsxwriter') as writer:
                            for s in sheet_names:
                                s_df = pd.read_excel(file_2, sheet_name=s).dropna(how='all')
                                _, s_cols = get_safe_total(s_df)
                                if s_cols:
                                    for c in s_cols: s_df[c] = pd.to_numeric(s_df[c], errors='coerce').fillna(0).astype(int)
                                    s_sum = s_df[s_cols].sum().to_frame().T
                                    s_sum.index = [len(s_df)]; s_df = pd.concat([s_df, s_sum])
                                    s_df.iloc[-1, 0] = "GRAND TOTAL"
                                s_df.to_excel(writer, sheet_name=s, index=False)
                        st.download_button("📥 คลิกเพื่อดาวน์โหลด (ทุกแผนก)", data=bulk_buf.getvalue(), file_name="Full_Hospital_Report.xlsx")
        else:
            st.warning("ไม่พบคอลัมน์อุปกรณ์ใน Sheet นี้")

else:
    st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏥 Smart Device Analytics</h1><p>Luxury Web-based Data Management System</p><p>กรุณาอัปโหลดไฟล์ Excel 2 เดือนเพื่อเริ่มการวิเคราะห์</p></div>", unsafe_allow_html=True)
