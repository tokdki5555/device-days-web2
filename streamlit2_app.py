import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration & Styling
st.set_page_config(page_title="Hospital Performance Dashboard", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #eef2f7; }
    .css-10trblm { color: #1e3a8a; } /* Header color */
    </style>
    """, unsafe_allow_html=True)

def process_all_sheets(file):
    """ฟังก์ชันอ่านทุก Sheet และรวมข้อมูล"""
    xls = pd.ExcelFile(file)
    all_data = []
    total_value = 0
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        # ค้นหาคอลัมน์ที่เป็นตัวเลขคอลัมน์แรก
        num_cols = df.select_dtypes(include=['number']).columns
        if not num_cols.empty:
            val_col = num_cols[0]
            sheet_sum = df[val_col].sum()
            total_value += sheet_sum
            all_data.append({'Department': sheet, 'Total': sheet_sum})
            
    return total_value, pd.DataFrame(all_data)

def main():
    st.title("🏥 Executive Performance Dashboard")
    st.subheader("สรุปภาพรวมและแยกรายแผนก (CCU / ICU / Ward)")
    st.divider()

    # Sidebar สำหรับจัดการไฟล์
    with st.sidebar:
        st.header("⚙️ Data Management")
        file1 = st.file_uploader("📂 ไฟล์เดือนที่ 1 (Base Period)", type=['xlsx'])
        file2 = st.file_uploader("📂 ไฟล์เดือนที่ 2 (Comparison Period)", type=['xlsx'])
        st.divider()
        if file1 and file2:
            st.success("โหลดข้อมูลเรียบร้อย!")

    if file1 and file2:
        try:
            # --- ประมวลผลข้อมูลระดับภาพรวม ---
            grand_total1, df_dept1 = process_all_sheets(file1)
            grand_total2, df_dept2 = process_all_sheets(file2)
            
            diff_total = grand_total2 - grand_total1
            pct_total = (diff_total / grand_total1 * 100) if grand_total1 != 0 else 0

            # --- SECTION 1: รวมทั้งหมด (Grand Total) ---
            st.markdown("### 💰 สรุปยอดรวมทุกแผนก (Grand Total)")
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("ยอดรวมรวมทุกแผนก (เดือน 1)", f"{grand_total1:,.2f}")
            kpi2.metric("ยอดรวมรวมทุกแผนก (เดือน 2)", f"{grand_total2:,.2f}")
            kpi3.metric("ภาพรวมการเปลี่ยนแปลง", f"{pct_total:+.2f}%", delta=f"{diff_total:,.2f}")
            
            st.divider()

            # --- SECTION 2: แยกแต่ละแผนก (Department Analysis) ---
            st.markdown("### 📊 การเปรียบเทียบรายแผนก (Department Breakdown)")
            
            # เตรียมข้อมูลกราฟเปรียบเทียบราย Sheet
            df_dept1['Period'] = 'เดือน 1'
            df_dept2['Period'] = 'เดือน 2'
            combined_dept = pd.concat([df_dept1, df_dept2])

            col_chart, col_table = st.columns([2, 1])
            
            with col_chart:
                fig_dept = px.bar(
                    combined_dept, x='Department', y='Total', color='Period',
                    barmode='group', text_auto='.3s',
                    color_discrete_sequence=['#94a3b8', '#1e40af'],
                    title="ยอดรวมเปรียบเทียบแยกตามแผนก"
                )
                st.plotly_chart(fig_dept, use_container_width=True)

            with col_table:
                # ตารางสรุปสั้นๆ รายแผนก
                summary_table = df_dept1.merge(df_dept2, on='Department', suffixes=(' (M1)', ' (M2)'))
                summary_table = summary_table[['Department', 'Total (M1)', 'Total (M2)']]
                st.write("**ตารางสรุปรายแผนก**")
                st.dataframe(summary_table.style.format({'Total (M1)': '{:,.0f}', 'Total (M2)': '{:,.0f}'}), use_container_width=True)

            # --- SECTION 3: เจาะลึกรายรายการ (Itemized Deep Dive) ---
            st.divider()
            st.markdown("### 🔍 เจาะลึกข้อมูลรายแผนก")
            
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการดูรายละเอียดรายการ:", df_dept1['Department'].tolist())
            
            # อ่านข้อมูลรายแผนกที่เลือก
            df_item1 = pd.read_excel(file1, sheet_name=selected_dept)
            df_item2 = pd.read_excel(file2, sheet_name=selected_dept)
            
            # ค้นหาคอลัม
