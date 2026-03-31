import streamlit as st
import pandas as pd
import plotly.express as px

# --- 1. การตั้งค่าหน้าจอและสไตล์ (Custom CSS) ---
st.set_page_config(page_title="Executive Medical Dashboard", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stMetric { 
        background: white; padding: 20px; border-radius: 15px; 
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #3b82f6;
    }
    div[data-testid="stExpander"] { background-color: white; border-radius: 10px; }
    h1, h2, h3 { color: #1e293b; font-family: 'Sarabun', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🏥 Executive Medical Performance Dashboard")
    st.caption("ระบบวิเคราะห์และเปรียบเทียบข้อมูลรายแผนก (CCU / ICU / Ward)")
    st.divider()

    # --- 2. ส่วนการ Upload ไฟล์ด้านข้าง ---
    with st.sidebar:
        st.header("📂 Data Import")
        f1 = st.file_uploader("เดือนที่ 1 (Base Period)", type=['xlsx'], key="u1")
        f2 = st.file_uploader("เดือนที่ 2 (Comparison Period)", type=['xlsx'], key="u2")
        st.divider()
        if f1 and f2:
            st.success("✅ อัปโหลดข้อมูลสำเร็จ")
        else:
            st.info("💡 กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือน")

    if f1 and f2:
        try:
            # อ่านชื่อ Sheet ทั้งหมด (ต้องตรงกันทั้ง 2 ไฟล์)
            xls1 = pd.ExcelFile(f1)
            xls2 = pd.ExcelFile(f2)
            all_sheets = xls1.sheet_names

            # --- 3. การประมวลผลข้อมูลภาพรวม (Grand Total) ---
            summary_list = []
            grand_total1 = 0
            grand_total2 = 0

            for sheet in all_sheets:
                # อ่านข้อมูลแต่ละ Sheet
                df1 = pd.read_excel(f1, sheet_name=sheet)
                df2 = pd.read_excel(f2, sheet_name=sheet)

                # บังคับคอลัมน์ที่ 2 ให้เป็นตัวเลข (ป้องกัน Error str+int)
                val_col1 = df1.columns[1]
                val_col2 = df2.columns[1]
                
                v1 = pd.to_numeric(df1[val_col1], errors='coerce').sum()
                v2 = pd.to_numeric(df2[val_col2], errors='coerce').sum()

                grand_total1 += v1
                grand_total2 += v2
                
                summary_list.append({
                    "แผนก": sheet,
                    "เดือนที่ 1": v1,
                    "เดือนที่ 2": v2
                })

            df_summary = pd.DataFrame(summary_list)
            diff_total = grand_total2 - grand_total1
            growth_pct = (diff_total / grand_total1 * 100) if grand_total1 != 0 else 0

            # --- 4. แสดง KPI Cards ---
            st.subheader("💰 สรุปภาพรวมทุกแผนก")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("ยอดรวมรวม M1", f"{grand_total1:,.2f}")
            k2.metric("ยอดรวมรวม M2", f"{grand_total2:,.2f}")
            k3.metric("ผลต่างรวม", f"{diff_total:+,.2f}")
            k4.metric("Growth (%)", f"{growth_pct:+.2f}%", delta=f"{growth_pct:+.2f}%")
            
            st.divider()

            # --- 5. กราฟแท่งเปรียบเทียบรายแผนก ---
            st.subheader("📊 เปรียบเทียบผลงานรายแผนก")
            fig = px.bar(
                df_summary, x='แผนก', y=['เดือนที่ 1', 'เดือนที่ 2'],
                barmode='group',
                color_discrete_sequence=['#94a3b8', '#2563eb'],
                template="plotly_white",
                height=450
            )
            fig.update_layout(legend_title_text='ช่วงเวลา', yaxis_title="จำนวน/ยอดรวม")
            st.plotly_chart(fig, use_container_width=True)

            # --- 6. เจาะลึกรายละเอียดรายแผนก (Section ที่แก้ Error) ---
            st.divider()
            st.subheader("🔍 เจาะลึกรายละเอียดรายแผนก")
            
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการตรวจสอบ:", all_sheets)

            # อ่านข้อมูลเฉพาะแผนกที่เลือก
            deep1 = pd.read_excel(f1, sheet_name=selected_dept)
            deep2 = pd.read_excel(f2, sheet_name=selected_dept)

            # เตรียมข้อมูลสำหรับตาราง (ใช้ index ในการจับคู่รายการเพื่อเลี่ยงชื่อคอลัมน์ซ้ำ)
            col_name = deep1.columns[0] # ชื่อรายการ
            col_val = deep1.columns[1]  # ยอดเงิน/จำนวน
            
            # บังคับข้อมูลเป็นตัวเลขก่อน Group
            deep1[col_val] = pd.to_numeric(deep1[col_val], errors='coerce').fillna(0)
            deep2[col_val] = pd.to_numeric(deep2[val_col2], errors='coerce').fillna(0)

            t1 = deep1.groupby(col_name)[col_val].sum()
            t2 = deep2.groupby(deep2.columns[0])[deep2.columns[1]].sum()

            # สร้าง DataFrame ผลลัพธ์
            final_df = pd.DataFrame({
                "รายการ": t1.index,
                "ยอดเดือนที่ 1": t1.values,
                "ยอดเดือนที่ 2": t2.reindex(t1.index, fill_value=0).values
            })
            final_df["ผลต่าง"] = final_df["ยอดเดือนที่ 2"] - final_df["ยอดเดือนที่ 1"]

            st.write(f"กำลังแสดงข้อมูลแผนก: **{selected_dept}**")
            
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
            st.info("💡 ข้อแนะนำ: ตรวจสอบว่าในไฟล์ Excel คอลัมน์แรกคือชื่อรายการ และคอลัมน์ที่สองคือตัวเลข")
    else:
        # หน้า Welcome
        st.markdown("""
            <div style="text-align: center; padding: 100px;">
                <h1 style="font-size: 50px;">📊</h1>
                <h2>ยินดีต้อนรับสู่ระบบวิเคราะห์ข้อมูล</h2>
                <p>กรุณาอัปโหลดไฟล์ Excel เดือนที่ 1 และ 2 ที่แถบด้านซ้ายเพื่อเริ่มการทำงาน</p>
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
