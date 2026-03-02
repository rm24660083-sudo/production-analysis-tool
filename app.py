import pandas as pd
import streamlit as st
import io
import re

def _parse_number(x):
    if pd.isna(x):
        return pd.NA
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "":
        return pd.NA
    s = s.replace(",", "")
    if s.endswith("%"):
        s = s[:-1]
    try:
        return float(s)
    except Exception:
        return pd.NA

def _normalize_numeric_columns(df):
    percent_cols = [c for c in df.columns if "% Waste" in c]
    value_cols = [c for c in df.columns if any(k in c for k in ["Consumption_", "Output_", "Waste_"])]
    for c in value_cols:
        df[c] = pd.to_numeric(df[c].map(_parse_number), errors="coerce")
    for c in percent_cols:
        series = pd.to_numeric(df[c].map(_parse_number), errors="coerce")
        sample = series.dropna()
        if not sample.empty and (sample.abs() <= 1).mean() > 0.5:
            series = series * 100
        df[c] = series
    return df

def _build_format_map(df):
    fmt = {}
    for c in df.columns:
        if "% Waste" in c:
            fmt[c] = "{:.2f}%"
        elif any(k in c for k in ["Consumption_", "Output_", "Waste_"]):
            fmt[c] = "{:,.2f}"
    if "Day of Posting Date" in df.columns:
        fmt["Day of Posting Date"] = lambda x: (
            "" if pd.isna(x)
            else (x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else str(x))
        )
    return fmt

def _get_series(df, name):
    return df[name] if name in df.columns else pd.Series(0, index=df.index)

def compute_special_conditions(df):
    cons_cols = [c for c in df.columns if c.startswith("Consumption_")]
    out_cols = [c for c in df.columns if c.startswith("Output_")]
    total_cons = df[cons_cols].fillna(0).sum(axis=1) if cons_cols else pd.Series(0, index=df.index)
    total_out = df[out_cols].fillna(0).sum(axis=1) if out_cols else pd.Series(0, index=df.index)
    c420 = _get_series(df, "Consumption_42000").fillna(0)
    o420 = _get_series(df, "Output_42000").fillna(0)
    w420 = _get_series(df, "Waste_42000").fillna(0)
    p420 = _get_series(df, "% Waste_42000").fillna(0)
    excl_420_all_zero = (c420 == 0) & (o420 == 0) & (w420 == 0) & (p420 == 0)
    excl_rows = excl_420_all_zero & (total_cons == 0) & (total_out == 0)
    cond_cons_zero_out_pos = (total_cons == 0) & (total_out > 0) & (~excl_rows)
    cond_out_zero_cons_pos = (total_out == 0) & (total_cons > 0) & (~excl_rows)
    a_rows = df[cond_cons_zero_out_pos].copy()
    b_rows = df[cond_out_zero_cons_pos].copy()
    def summarize(rows):
        if rows.empty:
            return rows
        if "Day of Posting Date" in rows.columns and pd.api.types.is_datetime64_any_dtype(rows["Day of Posting Date"]):
            dates = rows.groupby("Order No")["Day of Posting Date"].apply(lambda x: ", ".join(x.dt.strftime("%d/%m/%Y")))
        else:
            dates = rows.groupby("Order No")["Day of Posting Date"].apply(lambda x: ", ".join(x.astype(str)))
        summary = rows.groupby("Order No").size().reset_index(name="จำนวนครั้งที่พบ")
        summary["วันที่ที่พบ"] = summary["Order No"].map(dates.to_dict())
        return summary
    a_summary = summarize(a_rows)
    b_summary = summarize(b_rows)
    return a_rows, b_rows, a_summary, b_summary

def process_excel(file):
    df_raw = pd.read_excel(file, header=None)
    header_row = df_raw.iloc[2].tolist()
    cost_centers = []
    cc_mapping = {}
    current_cc = None
    for i, val in enumerate(df_raw.iloc[1]):
        if pd.notna(val) and str(val).isdigit():
            current_cc = str(val)
            if current_cc not in cost_centers:
                cost_centers.append(current_cc)
        if current_cc:
            cc_mapping[i] = current_cc
    df = pd.read_excel(file, skiprows=2)
    df = df.dropna(subset=['Order No'])
    new_columns = []
    for i, col in enumerate(df.columns):
        base = str(col)
        base = re.sub(r"\.\d+$", "", base)
        if i in cc_mapping:
            new_columns.append(f"{base}_{cc_mapping[i]}")
        else:
            new_columns.append(base)
    df.columns = new_columns
    if 'Day of Posting Date' in df.columns:
        df['Day of Posting Date'] = pd.to_datetime(df['Day of Posting Date'], errors='coerce')
    df = _normalize_numeric_columns(df)
    fmt_map = _build_format_map(df)
    duplicates = df[df.duplicated(subset=['Order No'], keep=False)].sort_values('Order No').copy()
    return df, duplicates, cost_centers, fmt_map

def main():
    st.set_page_config(page_title="Production Order Analysis Tool", layout="wide")
    st.title("🔍 วิเคราะห์รายงานการผลิต (Deep Scan)")
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            with st.spinner('กำลังวิเคราะห์ข้อมูลอย่างละเอียด...'):
                df, duplicates, cost_centers, fmt_map = process_excel(uploaded_file)
                a_rows, b_rows, a_summary, b_summary = compute_special_conditions(df)
            st.subheader("📋 ภาพรวมไฟล์")
            m1, m2, m3 = st.columns(3)
            m1.metric("จำนวนรายการทั้งหมด", len(df))
            m2.metric("จำนวน Order ที่ซ้ำ", len(duplicates['Order No'].unique()) if not duplicates.empty else 0)
            m3.metric("จำนวน Cost Center", len(cost_centers))
            st.subheader("🔎 เงื่อนไขพิเศษ (Consumption/Output)")
            c1, c2 = st.columns(2)
            c1.metric("Consumption = 0 แต่ Output > 0", len(a_summary))
            c2.metric("Output = 0 แต่ Consumption > 0", len(b_summary))
            with st.expander("ดูรายการ: Consumption = 0 แต่ Output > 0"):
                if len(a_summary) > 0:
                    st.dataframe(a_summary, use_container_width=True)
                    st.download_button(
                        label="📥 ดาวน์โหลดรายการ (Consumption=0, Output>0)",
                        data=a_rows.to_csv(index=False).encode('utf-8-sig'),
                        file_name='cons0_outpos_rows.csv',
                        mime='text/csv',
                    )
                else:
                    st.info("ไม่พบรายการตามเงื่อนไขนี้")
            with st.expander("ดูรายการ: Output = 0 แต่ Consumption > 0"):
                if len(b_summary) > 0:
                    st.dataframe(b_summary, use_container_width=True)
                    st.download_button(
                        label="📥 ดาวน์โหลดรายการ (Output=0, Consumption>0)",
                        data=b_rows.to_csv(index=False).encode('utf-8-sig'),
                        file_name='out0_conspos_rows.csv',
                        mime='text/csv',
                    )
                else:
                    st.info("ไม่พบรายการตามเงื่อนไขนี้")
            st.subheader("⚠️ รายละเอียดรายการ Order ที่ซ้ำกัน")
            if not duplicates.empty:
                st.warning("พบข้อมูลซ้ำ! ตารางด้านล่างแสดงข้อมูลเปรียบเทียบของ Order ที่มีหมายเลขเดียวกันในวันที่ต่างๆ")
                display_cols = ['Order No', 'Day of Posting Date']
                relevant_cols = [col for col in duplicates.columns if col not in display_cols and duplicates[col].notna().any()]
                disp = duplicates[display_cols + relevant_cols].copy()
                st.dataframe(
                    disp.style.format(fmt_map).highlight_null(color='#f8d7da'),
                    use_container_width=True
                )
                csv = duplicates.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 ดาวน์โหลดรายการซ้ำเป็น CSV",
                    data=csv,
                    file_name='duplicate_orders_report.csv',
                    mime='text/csv',
                )
            else:
                st.success("ยอดเยี่ยม! ไม่พบหมายเลข Order ที่ซ้ำกันในไฟล์นี้")
            st.subheader("🏢 ข้อมูลแยกตาม Cost Center")
            tabs = st.tabs([f"CC: {cc}" for cc in cost_centers])
            
            for i, cc in enumerate(cost_centers):
                with tabs[i]:
                    cc_cols = [col for col in df.columns if f"_{cc}" in col]
                    if cc_cols:
                        st.write(f"สรุปยอดรวมของ Cost Center: {cc}")
                        cc_data = df[['Order No', 'Day of Posting Date'] + cc_cols].dropna(subset=cc_cols, how='all')
                        st.dataframe(cc_data.style.format(fmt_map), use_container_width=True)
                    else:
                        st.info("ไม่มีข้อมูลการทำงานในศูนย์ต้นทุนนี้")
            with st.expander("📁 ดูข้อมูลดิบทั้งหมด (Raw Data)"):
                st.dataframe(df.style.format(fmt_map))

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.info("คำแนะนำ: ตรวจสอบให้แน่ใจว่าไฟล์ Excel มีโครงสร้างตามรูปแบบมาตรฐาน (Row 2 เป็น Header)")

    st.sidebar.markdown("---")
    st.sidebar.info("💡 **Tips:** คุณสามารถนำโค้ดนี้ไปรันบน **Streamlit Cloud** เพื่อให้ทีมงานใช้งานออนไลน์ได้ฟรี โดยไม่ต้องติดตั้งโปรแกรมในเครื่องครับ")

if __name__ == "__main__":
    main()
