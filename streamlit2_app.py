import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration
st.set_page_config(page_title="Executive Performance Dashboard", layout="wide")

# Custom CSS
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    h1, h2, h3 { color: #2c3e50; }
    </style>
    """, unsafe_allow_html=True)

def get_all_sheets_total(file):
    """ฟังก์ชันรวมยอดจากทุก Sheet ในไฟล์เดียว"""
    xls = pd.ExcelFile(file)
    grand_total = 0
    sheet_data = []
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        # กรองเฉพาะคอลัมน์ที่เป็นตัวเลข
        num_df = df.select_dtypes(include=['number'])
        if not num_df.empty:
            sheet_sum = num_df.iloc[:, 0].sum() # เอายอดจากคอลัมน์ตัวเลขแรกที่เจอ
            grand_total += sheet_sum
            sheet_data.append({'Sheet': sheet, 'Total': sheet_sum})
            
    return grand_total, pd.DataFrame(sheet_data)

def main():
    st.title("🏥 Executive Performance Dashboard")
    st.caption("สรุปยอดรวมทุกแผนกและเปรียบเทียบรายหน่วยงาน")
    st.divider()

    with st.sidebar:
        st.header("📂 Data Upload")
        file1 = st.file_uploader("ไฟล์เดือนที่ 1 (Base)", type=['xlsx'])
        file2 = st.file_uploader("ไฟล์เดือนที่ 2 (Compare)", type=['xlsx'])
        
    if file1 and file2:
        try:
            # --- ประมวลผลยอดรวมทั้งหมด (Grand Total) ---
            total_v1, df_sheets_v1 = get_all_sheets_total(file1)
            total_v2, df_sheets_v2 = get_all_sheets_total(file2)
            
            diff_all = total_v2 - total_v1
            pct_all = (diff_all / total_v1 * 100) if total_v1 != 0 else 0

            # --- 1. แสดงยอดรวมสูงสุด (Grand Total Section) ---
            st.subheader("💰 สรุปยอดรวมทุกแผนก (Grand Total)")
            g1, g2, g3 = st.columns(3)
            g1.metric("ยอดรวมรวมทุก Sheet (เดือน 1)", f"{total_v1:,.2f}")
            g2.metric("ยอดรวมรวมทุก Sheet (เดือน 2)", f"{total_v2:,.2f}")
            g3.metric("ภาพรวมการเปลี่ยนแปลง", f"{pct_all:+.2f}%", delta=f"{diff_all:,.2f}")
            
            st.divider()

            # --- 2. กราฟเปรียบเทียบราย Sheet (Unit Comparison) ---
            st.subheader("📊 เปรีย bเทียบรายแผนก (Unit Comparison)")
            
            df_sheets_v1['Period'] = 'เดือน 1'
            df_sheets_v2['Period'] = 'เดือน 2'
            combined_sheets = pd.concat([df_sheets_v1, df_sheets_v2])

            fig_unit = px.bar(
                combined_sheets, x='Sheet', y='Total', color='Period',
                barmode='group', text_auto='.2s',
                title="ยอดรวมแยกตามชื่อ Sheet",
                color_discrete_sequence=['#94a3b8', '#3b82f6']
            )
            st.plotly_chart(fig_unit, use_container_width=True)

            # --- 3. ส่วนเลือกเจาะลึกราย Sheet (Deep Dive) ---
            st.markdown("---")
            st.subheader("🔍 เจาะลึกรายละเอียดรายแผนก")
            selected_unit = st.selectbox("เลือกแผนกที่ต้องการดูรายละเอียด:", df_sheets_v1['Sheet'].tolist())
            
            df_u1 = pd.read_excel(file1, sheet_name=selected_unit)
            df_u2 = pd.read_excel(file2, sheet_name=selected_unit)

            # ค้นหาคอลัมน์รายการ (คอลัมน์แรก) และตัวเลข (คอลัมน์ตัวเลขแรก)
            cat_col = df_u1.columns[0] 
            val_col = df_u1.select_dtypes(include=['number']).columns[0]

            u_total1 = df_u1[val_col].sum()
            u_total2 = df_u2[val_col].sum()
            
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"แผนก: **{selected_unit}** | รายการที่วิเคราะห์: **{cat_col}**")
                st.write(f"ยอดรวมเฉพาะแผนกเดือน 1: **{u_total1:,.2f}**")
                st.write(f"ยอดรวมเฉพาะแผนกเดือน 2: **{u_total2:,.2f}**")
            
            # ตารางรายละเอียดใน Sheet นั้น
            df_u1_sub = df_u1.groupby(cat_col)[val_col].sum().reset_index()
            df_u2_sub = df_u2.groupby(cat_col)[val_col].sum().reset_index()
            
            detail_table = df_u1_sub.merge(df_u2_sub, on=cat_col, how='outer', suffixes=(' (M1)', ' (M2)')).fillna(0)
            
            st.dataframe(
                detail_table.style.format({
                    detail_table.columns[1]: '{:,.2f}',
                    detail_table.columns[2]: '{:,.2f}'
                }), use_container_width=True
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.info("โปรดตรวจสอบว่าคอลัมน์แรกในทุก Sheet เป็นรายการ และมีคอลัมน์ตัวเลขอย่างน้อย 1 คอลัมน์")
    else:
        st.info("👋 ยินดีต้อนรับ! กรุณาอัปโหลดไฟล์ Excel เพื่อดูยอดรวมทั้งหมด")

if __name__ == "__main__":
    main()
