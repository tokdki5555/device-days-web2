import streamlit as st
import pandas as pd
import plotly.express as px

# 1. การตั้งค่าหน้าจอและสไตล์ (Custom CSS)
st.set_page_config(page_title="Medical Unit Dashboard", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stMetric { 
        background-color: #ffffff; 
        padding: 20px; 
        border-radius: 15px; 
        box-shadow: 0 4px 12px rgba(0,0,0,0.05); 
        border-left: 5px solid #007bff;
    }
    div[data-testid="stExpander"] {
        background-color: #ffffff;
        border-radius: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

def get_sheet_summary(file):
    """ฟังก์ชันสรุปยอดจากทุก Sheet ในไฟล์"""
    xls = pd.ExcelFile(file)
    summary_list = []
    grand_total = 0
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        # กรองเฉพาะคอลัมน์ตัวเลข และจัดการ Error str+int ด้วย to_numeric
        num_cols = df.select_dtypes(include=['number', 'object']).columns
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # หายอดรวมจากคอลัมน์ตัวเลขแรกที่พบ
        numeric_only = df.select_dtypes(include=['number'])
        if not numeric_only.empty:
            val_col = numeric_only.columns[0]
            s_total = numeric_only[val_col].sum()
            grand_total += s_total
            summary_list.append({'Department': sheet, 'Total': s_total, 'Data': df})
            
    return grand_total, summary_list

def main():
    st.title("🏥 Medical Unit Performance Dashboard")
    st.markdown("---")

    # Sidebar สำหรับการจัดการไฟล์
    with st.sidebar:
        st.header("📂 Data Center")
        f1 = st.file_uploader("เดือนที่ 1 (Base)", type=['xlsx'])
        f2 = st.file_uploader("เดือนที่ 2 (Compare)", type=['xlsx'])
        st.divider()
        if f1 and f2:
            st.success("✅ ข้อมูลพร้อมวิเคราะห์")
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ Excel")

    if f1 and f2:
        try:
            # ประมวลผลข้อมูล
            total1, list1 = get_sheet_summary(f1)
            total2, list2 = get_sheet_summary(f2)
            
            diff = total2 - total1
            pct = (diff / total1 * 100) if total1 != 0 else 0

            # --- ส่วนที่ 1: ยอดรวมทั้งหมด (Grand Total) ---
            st.subheader("🌐 สรุปภาพรวมทุกแผนก (Grand Total)")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("ยอดรวมรวม M1", f"{total1:,.2f}")
            c2.metric("ยอดรวมรวม M2", f"{total2:,.2f}")
            c3.metric("ผลต่างรวม", f"{diff:+,.2f}")
            c4.metric("Growth (%)", f"{pct:+.2f}%", delta=f"{pct:+.2f}%")
            
            st.divider()

            # --- ส่วนที่ 2: เปรียบเทียบแยกรายแผนก (Department Breakdown) ---
            st.subheader("📊 เปรียบเทียบรายแผนก (By Department)")
            
            df_comp = pd.DataFrame(list1)[['Department', 'Total']].merge(
                pd.DataFrame(list2)[['Department', 'Total']], 
                on='Department', suffixes=('_M1', '_M2')
            )
            
            # กราฟแท่งเปรียบเทียบ
            plot_df = df_comp.melt(id_vars='Department', var_name='Period', value_name='Amount')
            plot_df['Period'] = plot_df['Period'].map({'Total_M1': 'เดือนที่ 1', 'Total_M2': 'เดือนที่ 2'})
            
            fig = px.bar(plot_df, x='Department', y='Amount', color='Period',
                         barmode='group', text_auto='.2s',
                         color_discrete_map={'เดือนที่ 1': '#A0AEC0', 'เดือนที่ 2': '#3182CE'},
                         template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

            # --- ส่วนที่ 3: เจาะลึกรายแผนก (Deep Dive) ---
            st.divider()
            st.subheader("🔍 เจาะลึกรายละเอียดรายแผนก")
            
            selected_dept = st.selectbox("เลือกแผนกที่ต้องการดู (CCU / ICU / Ward):", df_comp['Department'])
            
            # ดึง Data ของแผนกที่เลือกมาโชว์
            d1_df = next(item['Data'] for item in list1 if item['Department'] == selected_dept)
            d2_df = next(item['Data'] for item in list2 if item['Department'] == selected_dept)
            
            col_id = d1_df.columns[0] # คอลัมน์รายการ
            col_val = d1_df.select_dtypes(include=['number']).columns[0] # คอลัมน์ตัวเลข
            
            # สรุปตารางเปรียบเทียบในแผนก
            t1 = d1_df.groupby(col_id)[col_val].sum().reset_index()
            t2 = d2_df.groupby(col_id)[col_val].sum().reset_index()
            
            final_table = t1.merge(t2, on=col_id, how='outer', suffixes=(' (M1)', ' (M2)')).fillna(0)
            final_table['Diff'] = final_table.iloc[:, 2] - final_table.iloc[:, 1]
            
            st.write(f"วิเคราะห์ข้อมูลในแผนก: **{selected_dept}**")
            st.dataframe(
                final_table.style.format({
                    final_table.columns[1]: '{:,.2f}',
                    final_table.columns[2]: '{:,.2f}',
                    'Diff': '{:+,.2f}'
                }).background_gradient(subset=['Diff'], cmap='RdYlGn'),
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ ระบบเกิดข้อผิดพลาด: {e}")
    else:
        # หน้าต้อนรับ (Empty State)
        st.info("👋 ยินดีต้อนรับ! กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือนที่ Sidebar เพื่อเริ่มการวิเคราะห์")
        st.image("https://img.freepik.com/free-vector/flat-design-business-reporting-concept_23-2149156402.jpg", width=500)

if __name__ == "__main__":
    main()
