import streamlit as st
import pandas as pd
import plotly.express as px

# ตั้งค่าหน้ากระดาษ
st.set_page_config(page_title="Performance Analysis", layout="wide")

def process_data(file):
    """ฟังก์ชันสำหรับอ่านไฟล์ Excel"""
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def main():
    st.title("📊 Comparative Performance Analysis")
    st.markdown("---")

    # ส่วนการ Upload ไฟล์
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        file1 = st.file_uploader("Upload File 1 (Base Period)", type=['xlsx'])
    with col_up2:
        file2 = st.file_uploader("Upload File 2 (Comparison Period)", type=['xlsx'])

    if file1 and file2:
        df1 = process_data(file1)
        df2 = process_data(file2)

        if df1 is not None and df2 is not None:
            # --- 1. ส่วนสรุปยอดรวม (Summary Metrics) ---
            # สมมติว่ามีคอลัมน์ชื่อ 'Value' สำหรับคำนวณ
            total1 = df1['Value'].sum()
            total2 = df2['Value'].sum()
            diff_val = total2 - total1
            diff_pct = (diff_val / total1) * 100 if total1 != 0 else 0

            st.subheader("📌 Key Performance Summary")
            m1, m2, m3 = st.columns(3)
            m1.metric("ยอดรวมไฟล์ที่ 1", f"{total1:,.2f}")
            m2.metric("ยอดรวมไฟล์ที่ 2", f"{total2:,.2f}")
            m3.metric("การเปลี่ยนแปลง (%)", f"{diff_pct:+.2f}%", delta=f"{diff_val:,.2f}")

            st.divider()

            # --- 2. ส่วนกราฟเปรียบเทียบ (Visualizations) ---
            st.subheader("📈 Comparative Graphs")
            
            # เตรียมข้อมูลสำหรับกราฟ
            df1['Source'] = 'File 1'
            df2['Source'] = 'File 2'
            combined_df = pd.concat([df1, df2])

            c1, c2 = st.columns(2)
            
            with c1:
                # กราฟแท่งเปรียบเทียบตาม Category
                fig_bar = px.bar(
                    combined_df, 
                    x='Category', 
                    y='Value', 
                    color='Source', 
                    barmode='group',
                    title="ยอดขาย/ผลงาน แยกตามหมวดหมู่"
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            with c2:
                # กราฟเส้นแสดงแนวโน้ม (ถ้ามีคอลัมน์วันที่ หรือ ลำดับ)
                fig_pie = px.pie(
                    combined_df, 
                    values='Value', 
                    names='Source', 
                    hole=0.4,
                    title="สัดส่วนยอดรวมทั้งหมด"
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            # --- 3. ส่วนตารางวิเคราะห์เชิงลึก (Analysis Table) ---
            st.subheader("📋 Deep Dive Analysis")
            
            # สร้าง Pivot Table เพื่อดู % Change รายหมวดหมู่
            pivot_df = combined_df.pivot_table(index='Category', columns='Source', values='Value', aggfunc='sum')
            pivot_df['Difference'] = pivot_df['File 2'] - pivot_df['File 1']
            pivot_df['% Growth'] = (pivot_df['Difference'] / pivot_df['File 1']) * 100
            
            # แสดงตารางพร้อมใส่สี Highlight
            st.dataframe(
                pivot_df.style.format("{:,.2f}")
                .highlight_max(subset=['File 1', 'File 2'], color='#d1f2eb')
                .bar(subset=['% Growth'], color='#fb8500', align='mid')
            )

    else:
        st.info("กรุณา Upload ไฟล์ Excel ทั้ง 2 ไฟล์เพื่อเริ่มการวิเคราะห์")

if __name__ == "__main__":
    main()
