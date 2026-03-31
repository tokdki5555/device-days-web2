import streamlit as st
import pandas as pd
import plotly.express as px

# --- 1. การตั้งค่าหน้าจอ (Theme & Layout) ---
st.set_page_config(page_title="Medical Device Days Dashboard", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stMetric { 
        background: white; padding: 20px; border-radius: 15px; 
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #2563eb;
    }
    h1, h2, h3 { color: #1e293b; font-family: 'Sarabun', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🏥 Executive Device Days Dashboard")
    st.caption("วิเคราะห์และเปรียบเทียบข้อมูลการใช้อุปกรณ์รายแผนก (CCU / ICU / Ward)")
    st.divider()

    # --- 2. ส่วนการ Upload ไฟล์ ---
    with st.sidebar:
        st.header("📂 Data Import")
        f1 = st.file_uploader("ไฟล์เดือนที่ 1 (Base Period)", type=['xlsx'], key="u1")
        f2 = st.file_uploader("ไฟล์เดือนที่ 2 (Comparison Period)", type=['xlsx'], key="u2")
        st.divider()
        if f1 and f2:
            st.success("✅ อัปโหลดข้อมูลสำเร็จ")
        else:
            st.info("💡 กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือน")

    if f1 and f2:
        try:
            # อ่านชื่อ Sheet ทั้งหมด (CCU, ICU, Ward...)
            xls1 = pd.ExcelFile(f1)
            xls2 = pd.ExcelFile(f2)
            all_sheets = xls1.sheet_names

            # --- 3. ประมวลผลภาพรวม (Grand Total) ---
            summary_list = []
            grand_total1 = 0
            grand_total2 = 0

            for sheet in all_sheets:
                # อ่านข้อมูลแต่ละแผนก
                d1 = pd.read_excel(f1, sheet_name=sheet)
                d2 = pd.read_excel(f2, sheet_name=sheet)

                # บังคับคอลัมน์ที่ 2 เป็นตัวเลข (กัน Error str+int)
                col_val1 = d1.columns[1]
                col_val2 = d2.columns[1]
                
                v1 = pd.to_numeric(d1[col_val1], errors='coerce').sum()
                v2 = pd.to_numeric(d2[col_val2], errors='coerce').sum()

                grand_total1 += v1
                grand_total2 += v2
                
                summary_list.append({
                    "แผนก": sheet,
                    "เดือนที่ 1": v1,
                    "เดือนที่ 2": v2
                })

            df_summary = pd.DataFrame(summary_list)
            diff_total = grand_total2 - grand_total1
            pct_total = (diff_total / grand_total1 * 100) if grand_total1 != 0 else 0

            # --- 4. แสดงผล KPI Cards ---
            st.subheader("💰 สรุปภาพรวมทุกแผนก (Grand Total)")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ยอดรวม M1", f"{grand_total1:,.2f}")
            m2.metric("ยอดรวม M2", f"{grand_total2:,.2f}")
            m3.metric("ผลต่างรวม", f"{diff_total:+,.2f}")
            m4.metric("Growth (%)", f"{pct_total:+.2f}%", delta=f"{pct_total:+.2f}%")
            
            st.divider()

            # --- 5. กราฟเปรียบเทียบรายแผนก ---
            st.subheader("📊 เปรียบเทียบผลงานแยกตามแผนก")
            fig = px.bar(
                df_summary, x='แผนก', y=['เดือนที่ 1', 'เดือนที่ 2'],
                barmode='group',
                color_discrete_sequence=['#94a3b8', '#2563eb'],
                template="plotly_white",
                height=500
            )
            fig.update_layout(legend_title_text='ช่วงเวลา', yaxis_title="Device Days")
            st.plotly_chart(fig, use_container_width=True)

            # --- 6. เจาะลึกรายแผนก (Section แก้ Error duplicated) ---
            st.divider()
            st.subheader("🔍 เจาะลึกรายละเอียดรายการในแผนก")
            
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการตรวจสอบ:", all_sheets)

            # อ่านข้อมูลรายรายการ
            deep1 = pd.read_excel(f1, sheet_name=selected_dept)
            deep2 = pd.read_excel(f2, sheet_name=selected_dept)

            # หาชื่อคอลัมน์
            name_col = deep1.columns[0] # รายการ
            val_col = deep1.columns[1]  # ยอดรวม

            # รวมกลุ่มข้อมูลป้องกันชื่อรายการซ้ำในแผ่นเดียว
            t1 = deep1.groupby(name_col)[val_col].sum()
            t2 = deep2.groupby(deep2.columns[0])[deep2.columns[1]].sum()

            # สร้างตารางเปรียบเทียบใหม่ (เลี่ยงการใช้ merge ชื่อเดิม)
            final_df = pd.DataFrame({
                "รายการอุปกรณ์": t1.index,
                "ยอดเดือนที่ 1": t1.values,
                "ยอดเดือนที่ 2": t2.reindex(t1.index, fill_value=0).values
            })
            final_df["ผลต่าง"] = final_df["ยอดเดือนที่ 2"] - final_df["ยอดเดือนที่ 1"]

            st.info(f"📍 กำลังแสดงข้อมูล: **{selected_dept}**")
            
            # แสดงตารางพร้อมสี Gradient (ต้องมี matplotlib ใน requirements.txt)
            st.dataframe(
                final_df.style.format({
                    "ยอดเดือนที่ 1": "{:,.2f}", 
                    "ยอดเดือนที่ 2": "{:,.2f}", 
                    "ผลต่าง": "{:+,.2f}"
                }).background_gradient(subset=['ผลต่าง'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ ระบบขัดข้อง: {e}")
            st.info("💡 คำแนะนำ: ตรวจสอบว่าคอลัมน์แรกคือชื่อรายการ และคอลัมน์ที่สองคือตัวเลข")
    else:
        # หน้า Welcome เมื่อยังไม่ลงไฟล์
        st.markdown("""
            <div style="text-align: center; padding: 100px;">
                <h1 style="font-size: 60px;">📊</h1>
                <h2>ยินดีต้อนรับสู่ระบบวิเคราะห์ Device Days</h2>
                <p>กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือนที่ Sidebar เพื่อเริ่มต้น</p>
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
