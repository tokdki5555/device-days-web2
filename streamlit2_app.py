import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration
st.set_page_config(page_title="Medical Analytics Pro", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f7f6; }
    .stMetric { 
        background: white; padding: 20px; border-radius: 12px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.03); border: 1px solid #e1e4e8;
    }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🏥 Executive Medical Dashboard")
    st.divider()

    with st.sidebar:
        st.header("📂 Data Import")
        f1 = st.file_uploader("เดือนที่ 1 (Base)", type=['xlsx'], key="file1")
        f2 = st.file_uploader("เดือนที่ 2 (Compare)", type=['xlsx'], key="file2")

    if f1 and f2:
        try:
            # อ่านชื่อ Sheet ทั้งหมด
            xls1 = pd.ExcelFile(f1)
            xls2 = pd.ExcelFile(f2)
            all_sheets = xls1.sheet_names

            # --- ส่วนที่ 1: คำนวณสรุปยอดรายแผนก (เลี่ยงการใช้ Merge ชื่อคอลัมน์) ---
            summary_data = []
            grand_total1 = 0
            grand_total2 = 0

            for sheet in all_sheets:
                # อ่านข้อมูล
                d1 = pd.read_excel(f1, sheet_name=sheet)
                d2 = pd.read_excel(f2, sheet_name=sheet)

                # แปลงข้อมูลเป็นตัวเลขอัตโนมัติ ป้องกัน Error
                for d in [d1, d2]:
                    for col in d.columns:
                        d[col] = pd.to_numeric(d[col], errors='coerce')
                
                # หาคอลัมน์ตัวเลขแรกที่เจอ
                num_col1 = d1.select_dtypes(include=['number']).columns[0]
                num_col2 = d2.select_dtypes(include=['number']).columns[0]

                val1 = d1[num_col1].sum()
                val2 = d2[num_col2].sum()

                grand_total1 += val1
                grand_total2 += val2
                
                summary_data.append({
                    "แผนก": sheet,
                    "เดือนที่ 1": val1,
                    "เดือนที่ 2": val2
                })

            df_summary = pd.DataFrame(summary_data)
            diff_all = grand_total2 - grand_total1
            pct_all = (diff_all / grand_total1 * 100) if grand_total1 != 0 else 0

            # --- แสดง KPI Cards ---
            st.markdown("### 💰 สรุปภาพรวม")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ยอดรวม M1", f"{grand_total1:,.2f}")
            m2.metric("ยอดรวม M2", f"{grand_total2:,.2f}")
            m3.metric("ผลต่าง", f"{diff_all:+,.2f}")
            m4.metric("Growth (%)", f"{pct_all:+.2f}%", delta=f"{pct_all:+.2f}%")
            
            st.divider()

            # --- ส่วนที่ 2: กราฟแท่ง ---
            st.markdown("### 📊 เปรียบเทียบผลงานรายแผนก")
            fig = px.bar(df_summary, x='แผนก', y=['เดือนที่ 1', 'เดือนที่ 2'],
                         barmode='group', color_discrete_sequence=['#94a3b8', '#2563eb'],
                         template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

            # --- ส่วนที่ 3: เจาะลึกรายแผนก (Fixed Error จุดนี้) ---
            st.divider()
            st.markdown("### 🔍 เจาะลึกรายละเอียดรายแผนก")
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการตรวจสอบ:", all_sheets)

            # อ่านข้อมูลแผนกที่เลือกแบบใหม่ ไม่มีการใช้ชื่อคอลัมน์เดิมจากไฟล์มาสร้างตัวแปร
            deep1 = pd.read_excel(f1, sheet_name=selected_dept)
            deep2 = pd.read_excel(f2, sheet_name=selected_dept)

            cat_col = deep1.columns[0] # ชื่อรายการ (คอลัมน์แรก)
            # หาคอลัมน์ตัวเลขแรก
            val_col_name = deep1.select_dtypes(include=['number', 'object']).columns
            # บังคับเป็นตัวเลข
            deep1[deep1.columns[1]] = pd.to_numeric(deep1[deep1.columns[1]], errors='coerce').fillna(0)
            deep2[deep2.columns[1]] = pd.to_numeric(deep2[deep2.columns[1]], errors='coerce').fillna(0)

            # สรุปรายรายการ
            t1 = deep1.groupby(deep1.columns[0])[deep1.columns[1]].sum()
            t2 = deep2.groupby(deep2.columns[0])[deep2.columns[1]].sum()

            # สร้าง DataFrame ใหม่จาก Index เพื่อเลี่ยงชื่อคอลัมน์เดิมชนกัน
            final_df = pd.DataFrame({
                "รายการ": t1.index,
                "เดือนที่ 1": t1.values,
                "เดือนที่ 2": t2.reindex(t1.index, fill_value=0).values
            })
            final_df["ผลต่าง"] = final_df["เดือนที่ 2"] - final_df["เดือนที่ 1"]

            st.write(f"วิเคราะห์แผนก: **{selected_dept}**")
            st.dataframe(
                final_df.style.format({"เดือนที่ 1": "{:,.2f}", "เดือนที่ 2": "{:,.2f}", "ผลต่าง": "{:+,.2f}"})
                .background_gradient(subset=['ผลต่าง'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ ระบบขัดข้อง: {e}")
            st.info("💡 ข้อแนะนำ: ตรวจสอบว่าคอลัมน์แรกคือชื่อรายการ และคอลัมน์ที่สองคือตัวเลข")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้น")

if __name__ == "__main__":
    main()
