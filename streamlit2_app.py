import streamlit as st
import pandas as pd
import plotly.express as px

# --- 1. SETTINGS ---
st.set_page_config(page_title="Medical Device Days Dashboard", layout="wide")

def get_accurate_sum(df):
    """ฟังก์ชันดึงค่าตัวเลขที่ถูกต้องรายบรรทัด (ป้องกันการบวกเลข Total ซ้ำ)"""
    # ลบบรรทัดว่าง
    df = df.dropna(how='all').reset_index(drop=True)
    
    # โดยปกติ: Column 0 คือชื่ออุปกรณ์, Column 1 คือตัวเลข
    # เราจะกรองเอาเฉพาะบรรทัดที่มีชื่ออุปกรณ์และมีตัวเลขจริงๆ
    # และตัดบรรทัดที่มีคำว่า 'Total', 'รวม', 'Sum' ออกเพื่อไม่ให้เลขเบิ้ล
    
    temp_df = df.copy()
    col_name = temp_df.columns[0]
    col_val = temp_df.columns[1]
    
    # แปลงคอลัมน์ตัวเลขให้เป็น numeric (ค่าไหนไม่ใช่เลขจะเป็น NaN)
    temp_df[col_val] = pd.to_numeric(temp_df[col_val], errors='coerce')
    
    # กรองเอาเฉพาะบรรทัดที่:
    # 1. คอลัมน์แรกไม่ใช่คำว่า Total/รวม
    # 2. คอลัมน์สองเป็นตัวเลขและมากกว่า 0
    mask = (
        temp_df[col_name].astype(str).str.contains('Total|รวม|Sum|Grand', case=False, na=False) == False
    ) & (temp_df[col_val] > 0)
    
    final_df = temp_df[mask]
    return final_df[col_val].sum(), final_df

def main():
    st.title("🏥 Executive Device Days Dashboard (Corrected)")
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
                d1_raw = pd.read_excel(f1, sheet_name=sheet)
                d2_raw = pd.read_excel(f2, sheet_name=sheet)

                # คำนวณยอดแบบแม่นยำรายบรรทัด
                val1, _ = get_accurate_sum(d1_raw)
                val2, _ = get_accurate_sum(d2_raw)

                g_total1 += val1
                g_total2 += val2
                
                summary_data.append({"แผนก": sheet, "M1": val1, "M2": val2})

            df_summary = pd.DataFrame(summary_data)

            # --- KPI Cards ---
            st.subheader("💰 สรุปยอดรวม (ตรวจสอบความถูกต้อง)")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ยอดรวม M1", f"{g_total1:,.0f}")
            m2.metric("ยอดรวม M2", f"{g_total2:,.0f}")
            m3.metric("ผลต่าง", f"{g_total2 - g_total1:+,.0f}")
            m4.metric("Growth (%)", f"{((g_total2-g_total1)/g_total1*100 if g_total1 != 0 else 0):+.2f}%")

            # --- กราฟเปรียบเทียบรายแผนก ---
            st.divider()
            fig = px.bar(df_summary, x='แผนก', y=['M1', 'M2'], barmode='group',
                         title="เปรียบเทียบ Device Days รายแผนก (ยอดสุทธิ)",
                         color_discrete_sequence=['#94a3b8', '#2563eb'], template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

            # --- ตารางตรวจสอบความถูกต้องรายแผนก ---
            st.divider()
            st.subheader("🔍 ตารางตรวจสอบรายรายการ (ป้องกันเลขเบิ้ล)")
            selected_dept = st.selectbox("เลือกแผนกเพื่อดูรายการที่นำมาคำนวณ:", all_sheets)
            
            _, d1_detailed = get_accurate_sum(pd.read_excel(f1, sheet_name=selected_dept))
            _, d2_detailed = get_accurate_sum(pd.read_excel(f2, sheet_name=selected_dept))

            col_a, col_b = st.columns(2)
            with col_a:
                st.write(f"รายการที่ระบบดึงจาก M1 ({selected_dept})")
                st.dataframe(d1_detailed.iloc[:, :2], use_container_width=True)
            with col_b:
                st.write(f"รายการที่ระบบดึงจาก M2 ({selected_dept})")
                st.dataframe(d2_detailed.iloc[:, :2], use_container_width=True)

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือน")

if __name__ == "__main__":
    main()
