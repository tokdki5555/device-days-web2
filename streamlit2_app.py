import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration
st.set_page_config(page_title="Medical Analytics Pro", layout="wide")

# Custom Styling
st.markdown("""
    <style>
    .main { background-color: #f4f7f6; }
    .stMetric { 
        background: white; padding: 20px; border-radius: 12px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.03); border: 1px solid #e1e4e8;
    }
    h1, h2 { color: #1a365d; font-weight: 700; }
    </style>
    """, unsafe_allow_html=True)

def get_sheet_summary(file):
    xls = pd.ExcelFile(file)
    summary_list = []
    grand_total = 0
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        
        # --- จุดแก้ไขสำคัญ: ป้องกันชื่อคอลัมน์ซ้ำ ---
        # ถ้ามีคอลัมน์ชื่อ 'Department' หรือชื่อที่ระบบจะใช้ ให้ลบออกก่อน
        df = df.drop(columns=['Dept_Name', 'Total_Val'], errors='ignore')
        
        # จัดการข้อมูลตัวเลข ป้องกัน Error str+int
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        df = df.fillna(0)
        
        # เลือกคอลัมน์ตัวเลขแรกที่พบเพื่อหายอดรวม
        num_cols = df.select_dtypes(include=['number']).columns
        if not num_cols.empty:
            val_col = num_cols[0]
            s_total = df[val_col].sum()
            grand_total += s_total
            # เก็บข้อมูลโดยใช้ชื่อที่ระบบกำหนดเอง ไม่ปนกับในไฟล์
            summary_list.append({'Dept_Name': str(sheet), 'Total_Val': s_total, 'Raw_Data': df})
            
    return grand_total, summary_list

def main():
    st.title("🏥 Executive Medical Dashboard")
    st.caption("ระบบวิเคราะห์ข้อมูลเปรียบเทียบผลการดำเนินงานรายแผนก")
    st.divider()

    with st.sidebar:
        st.header("📂 Data Import")
        f1 = st.file_uploader("เดือนที่ 1 (Base)", type=['xlsx'])
        f2 = st.file_uploader("เดือนที่ 2 (Compare)", type=['xlsx'])
        if f1 and f2:
            st.success("✅ เชื่อมต่อข้อมูลสำเร็จ")

    if f1 and f2:
        try:
            # ดึงข้อมูลและสรุปยอด
            total1, list1 = get_sheet_summary(f1)
            total2, list2 = get_sheet_summary(f2)
            
            # สร้าง DataFrame เปรียบเทียบภาพรวมแผนก
            df_comp1 = pd.DataFrame([{'Dept_Name': x['Dept_Name'], 'M1': x['Total_Val']} for x in list1])
            df_comp2 = pd.DataFrame([{'Dept_Name': x['Dept_Name'], 'M2': x['Total_Val']} for x in list2])
            df_comp = pd.merge(df_comp1, df_comp2, on='Dept_Name', how='outer').fillna(0)
            
            diff_all = total2 - total1
            pct_all = (diff_all / total1 * 100) if total1 != 0 else 0

            # --- ส่วนที่ 1: KPI Cards ---
            st.markdown("### 💰 สรุปภาพรวมทุกแผนก")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ยอดรวม M1", f"{total1:,.2f}")
            m2.metric("ยอดรวม M2", f"{total2:,.2f}")
            m3.metric("ผลต่าง", f"{diff_all:+,.2f}")
            m4.metric("Growth (%)", f"{pct_all:+.2f}%", delta=f"{pct_all:+.2f}%")
            
            st.divider()

            # --- ส่วนที่ 2: กราฟแท่ง ---
            st.markdown("### 📊 เปรียบเทียบผลงานรายแผนก")
            fig = px.bar(df_comp, x='Dept_Name', y=['M1', 'M2'],
                         barmode='group', 
                         labels={'value': 'Amount', 'variable': 'Period', 'Dept_Name': 'แผนก'},
                         color_discrete_sequence=['#94a3b8', '#2563eb'],
                         template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

            # --- ส่วนที่ 3: เจาะลึกรายแผนก (จุดที่เคย Error) ---
            st.divider()
            st.markdown("### 🔍 เจาะลึกรายละเอียดรายแผนก")
            
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการตรวจสอบ:", df_comp['Dept_Name'].tolist())
            
            # ดึงข้อมูล Raw Data ของแผนกนั้นมาวิเคราะห์รายการย่อย
            d1_data = next(item['Raw_Data'] for item in list1 if item['Dept_Name'] == selected_dept)
            d2_data = next(item['Raw_Data'] for item in list2 if item['Dept_Name'] == selected_dept)
            
            cat_col = d1_data.columns[0] # คอลัมน์แรก (ชื่อรายการ)
            val_col = d1_data.select_dtypes(include=['number']).columns[0] # คอลัมน์ตัวเลขแรก
            
            t1 = d1_data.groupby(cat_col)[val_col].sum().reset_index()
            t2 = d2_data.groupby(cat_col)[val_col].sum().reset_index()
            
            # ใช้ merge แบบระบุชื่อคอลัมน์ชัดเจน ป้องกันการเอา 'Department' มาแทรกซ้ำ
            item_table = pd.merge(t1, t2, on=cat_col, how='outer', suffixes=(' (M1)', ' (M2)')).fillna(0)
            item_table['ผลต่าง'] = item_table.iloc[:, 2] - item_table.iloc[:, 1]
            
            st.write(f"แสดงข้อมูลรายการของแผนก: **{selected_dept}**")
            st.dataframe(
                item_table.style.format({
                    item_table.columns[1]: '{:,.2f}',
                    item_table.columns[2]: '{:,.2f}',
                    'ผลต่าง': '{:+,.2f}'
                }).background_gradient(subset=['ผลต่าง'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ ระบบขัดข้อง: {e}")
            st.info("แนะนำ: ตรวจสอบว่าในไฟล์มีคอลัมน์ตัวเลขอย่างน้อย 1 คอลัมน์")
    else:
        st.info("💡 อัปโหลดไฟล์ Excel เดือนที่ 1 และ 2 เพื่อเริ่มต้น")

if __name__ == "__main__":
    main()
