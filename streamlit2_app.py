import streamlit as st
import pandas as pd
import plotly.express as px

# --- 1. Settings & Style ---
st.set_page_config(page_title="Medical Device Days Dashboard", layout="wide")
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stMetric { background: white; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 5px solid #2563eb; }
    </style>
    """, unsafe_allow_html=True)

def clean_and_sum(df):
    """ฟังก์ชันทำความสะอาดข้อมูลและหายอดรวมที่ถูกต้อง"""
    # 1. ลบบรรทัดที่ว่างทั้งหมดออก
    df = df.dropna(how='all').reset_index(drop=True)
    
    # 2. ค้นหาคอลัมน์ตัวเลข (คอลัมน์ที่มีคำว่า 'Total' หรือคอลัมน์ที่ 2 เป็นต้นไป)
    # เราจะแปลงทุกคอลัมน์เป็นตัวเลข ถ้าแปลงไม่ได้ให้เป็น NaN
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # 3. เลือกคอลัมน์ที่มีตัวเลขมากที่สุด (มักจะเป็นคอลัมน์ Total Device Days)
    numeric_df = df.select_dtypes(include=['number'])
    if not numeric_df.empty:
        # หายอดรวมของคอลัมน์ที่มีค่ารวมสูงสุด (เพื่อกันไปหยิบเอาคอลัมน์ลำดับที่มา)
        sums = numeric_df.sum()
        return sums.max(), df # คืนค่าผลรวมสูงสุด
    return 0, df

def main():
    st.title("🏥 Executive Device Days Dashboard")
    st.divider()

    with st.sidebar:
        st.header("📂 Data Import")
        f1 = st.file_uploader("เดือนที่ 1 (Base)", type=['xlsx'], key="u1")
        f2 = st.file_uploader("เดือนที่ 2 (Compare)", type=['xlsx'], key="u2")

    if f1 and f2:
        try:
            xls1 = pd.ExcelFile(f1)
            xls2 = pd.ExcelFile(f2)
            all_sheets = xls1.sheet_names

            summary_data = []
            g_total1, g_total2 = 0, 0

            for sheet in all_sheets:
                # อ่านไฟล์โดยเริ่มเช็คตั้งแต่บรรทัดแรก
                d1_raw = pd.read_excel(f1, sheet_name=sheet)
                d2_raw = pd.read_excel(f2, sheet_name=sheet)

                # คำนวณยอดรวมที่ถูกต้อง
                val1, _ = clean_and_sum(d1_raw)
                val2, _ = clean_and_sum(d2_raw)

                g_total1 += val1
                g_total2 += val2
                
                summary_data.append({"แผนก": sheet, "M1": val1, "M2": val2})

            df_summary = pd.DataFrame(summary_data)

            # --- ส่วนแสดงผล KPI ---
            st.subheader("💰 สรุปภาพรวม")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ยอดรวม M1", f"{g_total1:,.0f}")
            m2.metric("ยอดรวม M2", f"{g_total2:,.0f}")
            diff = g_total2 - g_total1
            m3.metric("ผลต่าง", f"{diff:+,.0f}")
            m4.metric("Growth (%)", f"{(diff/g_total1*100 if g_total1 != 0 else 0):+.2f}%")

            # --- กราฟเปรียบเทียบรายแผนก ---
            st.divider()
            st.subheader("📊 เปรียบเทียบรายแผนก")
            fig = px.bar(df_summary, x='แผนก', y=['M1', 'M2'], barmode='group',
                         color_discrete_sequence=['#94a3b8', '#2563eb'], template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

            # --- เจาะลึกรายรายการ ---
            st.divider()
            st.subheader("🔍 รายละเอียดรายการในแผนก")
            selected_dept = st.selectbox("เลือกแผนก:", all_sheets)
            
            # ดึงข้อมูลมาแสดงเป็นตาราง (ใช้ฟังก์ชันเดิมที่เคยแก้เรื่อง Error ไว้)
            temp1 = pd.read_excel(f1, sheet_name=selected_dept)
            temp2 = pd.read_excel(f2, sheet_name=selected_dept)
            
            # แสดงตารางเปรียบเทียบแบบง่าย
            st.write(f"ข้อมูลดิบจากแผนก: {selected_dept}")
            col_a, col_b = st.columns(2)
            with col_a: st.write("เดือนที่ 1"); st.dataframe(temp1.dropna(how='all').iloc[:10, :2])
            with col_b: st.write("เดือนที่ 2"); st.dataframe(temp2.dropna(how='all').iloc[:10, :2])

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือน")

if __name__ == "__main__":
    main()
