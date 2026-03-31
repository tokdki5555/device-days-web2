import streamlit as st
import pandas as pd
import plotly.express as px

# ตั้งค่าหน้ากระดาษให้ดูทันสมัย
st.set_page_config(page_title="Performance Dashboard", layout="wide", initial_sidebar_state="expanded")

# ปรับแต่ง CSS เพื่อความสวยงาม
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🚀 Professional Performance Analysis")
    
    with st.sidebar:
        st.header("📂 Data Source")
        file1 = st.file_uploader("Upload Base Period (Excel)", type=['xlsx'])
        file2 = st.file_uploader("Upload Comparison Period (Excel)", type=['xlsx'])
        st.divider()
        st.info("คำแนะนำ: ตรวจสอบให้มั่นใจว่าทั้ง 2 ไฟล์มีโครงสร้างคอลัมน์ที่เหมือนกัน")

    if file1 and file2:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # แก้ปัญหา KeyError โดยการให้ผู้ใช้เลือกคอลัมน์เอง
        st.subheader("⚙️ Settings")
        col_set1, col_set2 = st.columns(2)
        
        all_columns = df1.columns.tolist()
        
        with col_set1:
            cat_col = st.selectbox("เลือกคอลัมน์ 'หมวดหมู่' (เช่น Category, Product):", all_columns)
        with col_set2:
            val_col = st.selectbox("เลือกคอลัมน์ 'ตัวเลข' (เช่น Value, Sales, Amount):", all_columns)

        st.divider()

        try:
            # คำนวณสรุปยอด
            total1 = df1[val_col].sum()
            total2 = df2[val_col].sum()
            diff_val = total2 - total1
            diff_pct = (diff_val / total1) * 100 if total1 != 0 else 0

            # --- 1. สรุปยอดรวมแบบสวยงาม (KPI Cards) ---
            m1, m2, m3 = st.columns(3)
            m1.metric(f"Total {val_col} (File 1)", f"{total1:,.2f}")
            m2.metric(f"Total {val_col} (File 2)", f"{total2:,.2f}")
            m3.metric("Growth Rate", f"{diff_pct:+.2f}%", delta=f"{diff_val:,.2f}")

            # --- 2. กราฟเปรียบเทียบ ---
            st.subheader("📊 Comparative Visualization")
            
            # รวมข้อมูลเพื่อพล็อตกราฟ
            df1_temp = df1[[cat_col, val_col]].copy()
            df1_temp['Period'] = 'File 1'
            df2_temp = df2[[cat_col, val_col]].copy()
            df2_temp['Period'] = 'File 2'
            combined_df = pd.concat([df1_temp, df2_temp])

            c1, c2 = st.columns([2, 1])
            
            with c1:
                fig_bar = px.bar(
                    combined_df, x=cat_col, y=val_col, color='Period',
                    barmode='group', text_auto='.2s',
                    template="plotly_white",
                    color_discrete_sequence=['#636EFA', '#EF553B']
                )
                fig_bar.update_layout(title_font_size=20)
                st.plotly_chart(fig_bar, use_container_width=True)

            with c2:
                # กราฟ Donut สัดส่วนยอดรวม
                pie_data = pd.DataFrame({
                    'Period': ['File 1', 'File 2'],
                    'Total': [total1, total2]
                })
                fig_donut = px.pie(pie_data, values='Total', names='Period', hole=0.5,
                                  color_discrete_sequence=['#636EFA', '#EF553B'])
                st.plotly_chart(fig_donut, use_container_width=True)

            # --- 3. ตารางวิเคราะห์เชิงลึก ---
            st.subheader("🔍 Deep Dive Table")
            summary = combined_df.groupby([cat_col, 'Period'])[val_col].sum().unstack().reset_index()
            summary['Diff'] = summary['File 2'] - summary['File 1']
            summary['% Change'] = (summary['Diff'] / summary['File 1']) * 100

            st.dataframe(
                summary.style.format({
                    'File 1': '{:,.2f}', 'File 2': '{:,.2f}',
                    'Diff': '{:,.2f}', '% Change': '{:+.2f}%'
                }).background_gradient(subset=['% Change'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลข้อมูล: {e}")
    else:
        # หน้า Welcome เมื่อยังไม่ได้อัปโหลดไฟล์
        st.info("👋 ยินดีต้อนรับ! กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 ไฟล์ที่แถบด้านซ้ายเพื่อเริ่มการวิเคราะห์")
        st.image("https://img.freepik.com/free-vector/data-report-concept-illustration_114360-883.jpg", width=400)

if __name__ == "__main__":
    main()
