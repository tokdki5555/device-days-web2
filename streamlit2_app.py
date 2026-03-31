import streamlit as st
import pandas as pd
import plotly.express as px

# ตั้งค่าหน้าจอ
st.set_page_config(page_title="Unit Performance Analysis", layout="wide")

def main():
    st.title("🏥 Unit Performance Analysis (CCU / ICU / Ward)")
    st.markdown("---")

    with st.sidebar:
        st.header("📂 Data Upload")
        file1 = st.file_uploader("Upload ไฟล์เดือนที่ 1", type=['xlsx'])
        file2 = st.file_uploader("Upload ไฟล์เดือนที่ 2", type=['xlsx'])
        st.divider()
        st.info("ระบบจะอ่านรายชื่อ Sheet จากไฟล์ที่คุณอัปโหลด")

    if file1 and file2:
        # 1. อ่านรายชื่อ Sheet ทั้งหมดจากไฟล์
        xls1 = pd.ExcelFile(file1)
        xls2 = pd.ExcelFile(file2)
        sheet_names = xls1.sheet_names

        # 2. ให้ผู้ใช้เลือก Sheet และคอลัมน์
        st.subheader("🎯 Selection")
        c1, c2, c3 = st.columns(3)
        with c1:
            selected_sheet = st.selectbox("เลือกแผนก/Sheet ที่ต้องการดู:", sheet_names)
        
        # อ่านข้อมูลจาก Sheet ที่เลือก
        df1 = pd.read_excel(file1, sheet_name=selected_sheet)
        df2 = pd.read_excel(file2, sheet_name=selected_sheet)

        all_columns = df1.columns.tolist()
        
        with c2:
            cat_col = st.selectbox("คอลัมน์รายการ (เช่น รายการยา):", all_columns)
        with c3:
            val_col = st.selectbox("คอลัมน์ตัวเลข (เช่น จำนวน):", all_columns)

        st.divider()

        try:
            # คำนวณสรุปยอด
            total1 = df1[val_col].sum()
            total2 = df2[val_col].sum()
            diff_val = total2 - total1
            diff_pct = (diff_val / total1) * 100 if total1 != 0 else 0

            # --- ส่วนแสดงผล KPI ---
            st.subheader(f"📊 ผลการวิเคราะห์แผนก: {selected_sheet}")
            m1, m2, m3 = st.columns(3)
            m1.metric(f"ยอดรวม (เดือน 1)", f"{total1:,.2f}")
            m2.metric(f"ยอดรวม (เดือน 2)", f"{total2:,.2f}")
            m3.metric("อัตราการเปลี่ยนแปลง", f"{diff_pct:+.2f}%", delta=f"{diff_val:,.2f}")

            # --- ส่วนกราฟเปรียบเทียบ ---
            df1_temp = df1[[cat_col, val_col]].copy()
            df1_temp['Period'] = 'เดือน 1'
            df2_temp = df2[[cat_col, val_col]].copy()
            df2_temp['Period'] = 'เดือน 2'
            combined_df = pd.concat([df1_temp, df2_temp])

            fig_bar = px.bar(
                combined_df, x=cat_col, y=val_col, color='Period',
                barmode='group', text_auto='.2s',
                title=f"เปรียบเทียบรายหมวดหมู่ใน {selected_sheet}",
                color_discrete_sequence=['#1f77b4', '#ff7f0e']
            )
            st.plotly_chart(fig_bar, use_container_width=True)

            # --- ตารางสรุปข้อมูล (ส่วนที่แก้ Error วงเล็บ) ---
            st.subheader("📋 ตารางข้อมูลรายละเอียด")
            summary = combined_df.groupby([cat_col, 'Period'])[val_col].sum().unstack().reset_index()
            summary = summary.fillna(0)
            
            # ตรวจสอบว่ามีคอลัมน์ครบไหมก่อนคำนวณ
            if 'เดือน 1' in summary.columns and 'เดือน 2' in summary.columns:
                summary['ผลต่าง'] = summary['เดือน 2'] - summary['เดือน 1']
                summary['% การเปลี่ยนแปลง'] = (summary['ผลต่าง'] / summary['เดือน 1']) * 100
            
            # แก้ไขจุดที่ Syntax Error (ปิดวงเล็บให้ครบ)
            st.dataframe(
                summary.style.format({
                    'เดือน 1': '{:,.2f}', 
                    'เดือน 2': '{:,.2f}',
                    'ผลต่าง': '{:,.2f}', 
                    '% การเปลี่ยนแปลง': '{:+.2f}%'
                }).background_gradient(subset=['% การเปลี่ยนแปลง'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการคำนวณ: {e}")
    else:
        st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการวิเคราะห์")

if __name__ == "__main__":
    main()
