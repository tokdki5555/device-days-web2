import streamlit as st
import pandas as pd
import plotly.express as px

# 1. การตั้งค่าหน้าจอและสไตล์
st.set_page_config(page_title="Unit Performance Pro", layout="wide")

st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; color: #1E88E5; }
    .main { background-color: #f0f2f6; }
    div.stButton > button:first-child { background-color: #00b894; color:white; }
    </style>
    """, unsafe_allow_html=True)

def main():
    st.title("🏥 Unit Performance Analysis")
    st.caption("ระบบวิเคราะห์เปรียบเทียบผลการดำเนินงานรายแผนก (CCU / ICU / Ward)")
    st.divider()

    # 2. Sidebar สำหรับ Upload
    with st.sidebar:
        st.header("📂 Data Management")
        file1 = st.file_uploader("ไฟล์เดือนที่ 1 (Base)", type=['xlsx'])
        file2 = st.file_uploader("ไฟล์เดือนที่ 2 (Compare)", type=['xlsx'])
        
        if file1 and file2:
            st.success("อัปโหลดสำเร็จ!")
            if st.button("🔄 ล้างข้อมูลทั้งหมด"):
                st.rerun()

    if file1 and file2:
        # อ่าน Sheet Names
        xls1 = pd.ExcelFile(file1)
        xls2 = pd.ExcelFile(file2)
        sheet_names = xls1.sheet_names

        # 3. ส่วนการเลือกข้อมูล (จัดวางใหม่ให้สวย)
        st.subheader("🎯 การตั้งค่าการเปรียบเทียบ")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            selected_sheet = st.selectbox("🏥 เลือกแผนก/หน่วยงาน:", sheet_names)
        
        # อ่านข้อมูลจาก Sheet
        df1 = pd.read_excel(file1, sheet_name=selected_sheet)
        df2 = pd.read_excel(file2, sheet_name=selected_sheet)
        all_cols = df1.columns.tolist()

        with c2:
            cat_col = st.selectbox("🔍 เลือกรายการวิเคราะห์ (Category):", all_cols)
        with c3:
            # กรองเฉพาะคอลัมน์ที่เป็นตัวเลขเพื่อลด Error
            num_cols = df1.select_dtypes(include=['number']).columns.tolist()
            if not num_cols: num_cols = all_cols
            val_col = st.selectbox("💰 เลือกยอดเงิน/จำนวน (Value):", num_cols)

        st.divider()

        try:
            # ประมวลผลข้อมูล (จัดการเรื่อง Data Type เพื่อแก้ Error 'str' + 'int')
            df1[val_col] = pd.to_numeric(df1[val_col], errors='coerce').fillna(0)
            df2[val_col] = pd.to_numeric(df2[val_col], errors='coerce').fillna(0)

            total1 = df1[val_col].sum()
            total2 = df2[val_col].sum()
            diff_val = total2 - total1
            diff_pct = (diff_val / total1 * 100) if total1 != 0 else 0

            # --- 4. สรุปยอด (Executive Summary) ---
            st.subheader(f"📌 สรุปภาพรวม: {selected_sheet}")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("เดือนก่อนหน้า", f"{total1:,.2f}")
            m2.metric("เดือนปัจจุบัน", f"{total2:,.2f}")
            m3.metric("ผลต่าง", f"{diff_val:+,.2f}", delta_color="normal")
            m4.metric("การเติบโต (%)", f"{diff_pct:+.2f}%", delta=f"{diff_pct:+.2f}%")

            # --- 5. กราฟ (Visual Insight) ---
            st.subheader("📊 การวิเคราะห์รายหมวดหมู่")
            
            df1_sub = df1.groupby(cat_col)[val_col].sum().reset_index()
            df1_sub['เดือน'] = 'เดือนที่ 1'
            df2_sub = df2.groupby(cat_col)[val_col].sum().reset_index()
            df2_sub['เดือน'] = 'เดือนที่ 2'
            
            combined = pd.concat([df1_sub, df2_sub])

            g1, g2 = st.columns([2, 1])
            with g1:
                fig_bar = px.bar(
                    combined, x=cat_col, y=val_col, color='เดือน',
                    barmode='group', text_auto='.2s',
                    color_discrete_map={'เดือนที่ 1': '#AABBC3', 'เดือนที่ 2': '#1E88E5'},
                    template="simple_white"
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            
            with g2:
                # กราฟวงกลมดูสัดส่วนเดือนปัจจุบัน
                fig_pie = px.pie(df2_sub, values=val_col, names=cat_col, 
                                 title="สัดส่วนเดือนปัจจุบัน", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)

            # --- 6. ตารางรายละเอียด ---
            st.subheader("📋 รายละเอียดเชิงลึก (Deep Dive)")
            final_table = df1_sub.merge(df2_sub, on=cat_col, how='outer', suffixes=('_1', '_2'))
            final_table = final_table.fillna(0)
            final_table = final_table[[cat_col, f'{val_col}_1', f'{val_col}_2']]
            final_table.columns = ['รายการ', 'ยอดเดือน 1', 'ยอดเดือน 2']
            
            final_table['ผลต่าง'] = final_table['ยอดเดือน 2'] - final_table['ยอดเดือน 1']
            final_table['% Growth'] = (final_table['ผลต่าง'] / final_table['ยอดเดือน 1'] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)

            st.dataframe(
                final_table.style.format({
                    'ยอดเดือน 1': '{:,.2f}', 'ยอดเดือน 2': '{:,.2f}',
                    'ผลต่าง': '{:,.2f}', '% Growth': '{:+.2f}%'
                }).background_gradient(subset=['% Growth'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ ระบบขัดข้อง: {e}")
            st.info("💡 ข้อแนะนำ: ตรวจสอบว่าเลือกคอลัมน์ที่เป็น 'ตัวเลข' ในช่องขวาสุดแล้วหรือยัง")

    else:
        st.markdown("""
            <div style="text-align: center; padding: 50px;">
                <h2 style="color: #636e72;">กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการวิเคราะห์</h2>
                <p>เลือกไฟล์จาก Sidebar ด้านซ้ายมือ</p>
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
