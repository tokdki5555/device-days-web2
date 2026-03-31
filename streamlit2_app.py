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
        
        # จัดการข้อมูลตัวเลข ป้องกัน Error str+int
        for col in df.columns:
            if df[col].dtype == 'object':
                # พยายามเปลี่ยนเป็นตัวเลข ถ้าไม่ใช่ให้เป็น NaN แล้วเติม 0
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        df = df.fillna(0)
        
        # เลือกคอลัมน์ตัวเลขแรกที่พบเพื่อหายอดรวม
        num_cols = df.select_dtypes(include=['number']).columns
        if not num_cols.empty:
            val_col = num_cols[0]
            s_total = df[val_col].sum()
            grand_total += s_total
            summary_list.append({'Dept_Name': sheet, 'Total_Val': s_total, 'Raw_Data': df})
            
    return grand_total, summary_list

def main():
    st.title("🏥 Executive Medical Dashboard")
    st.caption("ระบบวิเคราะห์ข้อมูลเปรียบเทียบผลการดำเนินงานรายแผนก")
    st.divider()

    with st.sidebar:
        st.header("📂 Data Import")
        f1 = st.file_uploader("เดือนที่ 1 (Base)", type=['xlsx'])
        f2 = st.file_uploader("เดือนที่ 2 (Compare)", type=['xlsx'])
        st.divider()
        if f1 and f2:
            st.success("✅ เชื่อมต่อข้อมูลสำเร็จ")

    if f1 and f2:
        try:
            # ดึงข้อมูล
            total1, list1 = get_sheet_summary(f1)
            total2, list2 = get_sheet_summary(f2)
            
            diff = total2 - total1
            pct = (diff / total1 * 100) if total1 != 0 else 0

            # --- ส่วนที่ 1: สรุปยอดรวมสูงสุด ---
            st.markdown("### 💰 สรุปภาพรวมทุกแผนก (Grand Total)")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ยอดรวมรวม M1", f"{total1:,.2f}")
            m2.metric("ยอดรวมรวม M2", f"{total2:,.2f}")
            m3.metric("ผลต่างรวม", f"{diff:+,.2f}")
            m4.metric("Growth (%)", f"{pct:+.2f}%", delta=f"{pct:+.2f}%")
            
            st.divider()

            # --- ส่วนที่ 2: เปรียบเทียบแผนก (Department Comparison) ---
            st.markdown("### 📊 เปรียบเทียบผลงานรายแผนก")
            
            # สร้าง DataFrame เปรียบเทียบ
            df_comp = pd.DataFrame(list1)[['Dept_Name', 'Total_Val']].merge(
                pd.DataFrame(list2)[['Dept_Name', 'Total_Val']], 
                on='Dept_Name', suffixes=('_M1', '_M2')
            )
            
            # กราฟแท่ง
            fig = px.bar(df_comp, x='Dept_Name', y=['Total_Val_M1', 'Total_Val_M2'],
                         barmode='group', 
                         labels={'value': 'Amount', 'variable': 'Period', 'Dept_Name': 'แผนก'},
                         color_discrete_sequence=['#94a3b8', '#2563eb'],
                         template="plotly_white")
            fig.update_layout(legend_title_text='')
            st.plotly_chart(fig, use_container_width=True)

            # --- ส่วนที่ 3: เจาะลึกรายแผนก ---
            st.divider()
            st.markdown("### 🔍 เจาะลึกรายละเอียดรายแผนก")
            
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการตรวจสอบ:", df_comp['Dept_Name'])
            
            # ดึงข้อมูลแผนกที่เลือก
            d1_data = next(item['Raw_Data'] for item in list1 if item['Dept_Name'] == selected_dept)
            d2_data = next(item['Raw_Data'] for item in list2 if item['Dept_Name'] == selected_dept)
            
            cat_col = d1_data.columns[0] # คอลัมน์รายการ
            val_col = d1_data.select_dtypes(include=['number']).columns[0] # คอลัมน์ตัวเลข
            
            # รวมยอดรายรายการ
            t1 = d1_data.groupby(cat_col)[val_col].sum().reset_index()
            t2 = d2_data.groupby(cat_col)[val_col].sum().reset_index()
            
            # สร้างตารางเปรียบเทียบ (แก้ปัญหาชื่อซ้ำ)
            final_table = t1.merge(t2, on=cat_col, how='outer', suffixes=(' (M1)', ' (M2)')).fillna(0)
            final_table['ผลต่าง'] = final_table.iloc[:, 2] - final_table.iloc[:, 1]
            
            st.write(f"แสดงข้อมูลรายการของแผนก: **{selected_dept}**")
            st.dataframe(
                final_table.style.format({
                    final_table.columns[1]: '{:,.2f}',
                    final_table.columns[2]: '{:,.2f}',
                    'ผลต่าง': '{:+,.2f}'
                }).background_gradient(subset=['ผลต่าง'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ ระบบขัดข้อง: {e}")
    else:
        st.info("💡 กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 ไฟล์ เพื่อเริ่มต้นการคำนวณ")

if __name__ == "__main__":
    main()
