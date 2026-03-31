import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Ultra Clear Executive Styling
st.set_page_config(page_title="Device Analytics Pro", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
    html, body, [class*="css"] { font-family: 'Sarabun', sans-serif; }
    
    .stApp { background-color: #ffffff; }
    
    /* หัวข้อใหญ่ชัดเจน */
    .main-title { 
        color: #0f172a; font-size: 3rem; font-weight: 800; 
        text-align: center; margin-bottom: 40px;
    }

    /* KPI Card แบบเน้นตัวเลขใหญ่พิเศษ */
    .kpi-card {
        background: #f8fafc; padding: 30px; border-radius: 15px; text-align: center;
        border: 2px solid #e2e8f0;
    }
    .kpi-label { color: #475569; font-size: 1.2rem; font-weight: 700; text-transform: uppercase; }
    .kpi-value { color: #1e293b; font-size: 3.5rem; font-weight: 800; margin: 10px 0; line-height: 1; }
    .kpi-delta { font-size: 1.5rem; font-weight: 700; }

    /* ปรับตัวเลขใน Table ให้ชัดเจน */
    .stDataFrame { font-size: 1.1rem; }
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
st.sidebar.markdown("<h1 style='text-align:center;'>🏥 SYSTEM</h1>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 Month 1 (Previous)", type=["xlsx"])
file_2 = st.sidebar.file_uploader("📂 Month 2 (Current)", type=["xlsx"])

if file_1 and file_2:
    st.sidebar.markdown("---")
    page = st.sidebar.selectbox("เมนูนำทาง", ["📊 Comparison Analytics", "📄 Data Editor (File 2)"])

    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    # --- หน้า 1: COMPARISON (เน้นตัวหนา เลขใหญ่ สีกราฟใหม่) ---
    if page == "📊 Comparison Analytics":
        st.markdown("<h1 class='main-title'>Dashboard เปรียบเทียบผลงาน</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff = t2 - t1
        growth = int((diff / t1 * 100)) if t1 != 0 else 0
        color_code = "#059669" if diff >= 0 else "#dc2626" # เขียวเข้ม / แดงเข้ม

        # KPI Row (เน้นความชัดเจน)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>เดือนก่อนหน้า</p><p class='kpi-value'>{t1:,}</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>เดือนปัจจุบัน</p><p class='kpi-value' style='color:#2563eb;'>{t2:,}</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi-card'><p class='kpi-label'>ผลต่าง (Variance)</p><p class='kpi-value' style='color:{color_code};'>{diff:+,}</p><p class='kpi-delta' style='color:{color_code};'>{growth:+,}% Change</p></div>", unsafe_allow_html=True)

        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # กราฟเปรียบเทียบ (เปลี่ยนสีกราฟให้ชัดเจนขึ้น)
        st.subheader("📈 ยอดการใช้งานอุปกรณ์แยกตามแผนก")
        fig_bar = px.bar(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                         barmode='group', text_auto=',.0f',
                         color_discrete_sequence=['#94a3b8', '#1e293b'], # เทาเข้ม กับ น้ำเงินดำ Contrast สูง
                         labels={'value': 'จำนวนวัน (Days)', 'variable': 'เดือน'})
        
        fig_bar.update_layout(
            font=dict(size=16, color="#000000"), # ปรับตัวอักษรกราฟให้ใหญ่และดำสนิท
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        fig_bar.update_traces(textfont_size=16, textfont_color="white", textposition='inside')
        st.plotly_chart(fig_bar, use_container_width=True)

        # Export ส่วนกลาง
        st.markdown("---")
        st.subheader("📋 ตารางข้อมูลเปรียบเทียบ")
        st.dataframe(df_compare.style.format(precision=0), use_container_width=True)
        
        buf_comp = io.BytesIO()
        df_compare.to_excel(buf_comp, index=False)
        st.download_button("📥 ดาวน์โหลดตารางเปรียบเทียบ", data=buf_comp.getvalue(), file_name="Comparison_Summary.xlsx")

    # --- หน้า 2: EDITOR (แยกรายการกลับมาเหมือนเดิม) ---
    elif page == "📄 Data Editor (File 2)":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("เลือกแผนก (Ward):", excel_2.sheet_names)
        
        st.markdown(f"<h1 class='main-title'>แก้ไขข้อมูล: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        # อ่านข้อมูล
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        
        # ส่วนแสดงผล Metric แยกรายการ (นำกลับมาไว้ด้านบน)
        total_val, device_cols = get_safe_total(df_raw)
        
        if device_cols:
            st.subheader("📊 ยอดรวมแยกตามอุปกรณ์ (Current Data)")
            # คำนวณยอดดิบก่อน Edit
            current_sum = df_raw[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(label=col, value=f"{current_sum[col]:,}")
            m_cols[-1].metric(label="ยอดรวมทั้งหมด", value=f"{total_val:,}", delta_color="off")
            st.markdown("---")

        # ส่วน Editor
        st.subheader("📝 ตารางแก้ไขข้อมูล")
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        # ปุ่ม Export ทั้งไฟล์ (Bulk Export)
        st.markdown("---")
        st.subheader("📥 ดาวน์โหลดข้อมูลทั้งหมด (File 2)")
        if st.button("📥 เตรียมไฟล์ Export ทุกแผนก"):
            with st.spinner('กำลังรวบรวมทุก Sheet...'):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for s_name in excel_2.sheet_names:
                        temp_df = pd.read_excel(file_2, sheet_name=s_name).dropna(how='all')
                        t_val, d_cols = get_safe_total(temp_df)
                        if d_cols:
                            for c in d_cols: temp_df[c] = pd.to_numeric(temp_df[c], errors='coerce').fillna(0)
                            sum_line = temp_df[d_cols].sum().to_frame().T
                            sum_line.index = [len(temp_df)]
                            temp_df = pd.concat([temp_df, sum_line])
                            temp_df.iloc[-1, 0] = "GRAND TOTAL"
                        temp_df.to_excel(writer, sheet_name=s_name, index=False)
                st.download_button("📥 คลิกเพื่อดาวน์โหลด (Single Excel File)", 
                                   data=output.getvalue(), 
                                   file_name=f"Full_Device_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")

else:
    st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏦 Device Analytics System</h1><p>กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือนเพื่อเริ่มการวิเคราะห์</p></div>", unsafe_allow_html=True)
