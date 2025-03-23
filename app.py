import streamlit as st
import pandas as pd
import io

def preprocess_data(df):
    """Tiền xử lý dữ liệu input."""
    columns_to_drop = ["Member/Job Type", "Project"]
    last_column_name = df.columns[-1]
    if "Sum" in last_column_name:
        columns_to_drop.append(last_column_name)
    df = df.drop(columns=columns_to_drop, errors='ignore')
    if len(df) > 1:
        df = df.drop(0)
    return df

def convert_xlsx(uploaded_file):
    """Chuyển đổi file XLSX từ dạng ma trận sang dạng bảng dọc."""
    try:
        df = pd.read_excel(uploaded_file)
        df = preprocess_data(df)
        dates = df.columns[2:]
        result_df = pd.DataFrame(columns=["BP Ticket No.", "Summary", "Working (Hours)", "Actual Date"])
        for index, row in df.iterrows():
            ticket_no = row["ID"]
            summary = row["Title"]
            for date in dates:
                working_hours = row[date]
                if pd.notna(working_hours):
                    working_hours = str(working_hours).replace(" H", "")
                    result_df = pd.concat([result_df, pd.DataFrame({
                        "BP Ticket No.": [ticket_no],
                        "Summary": [summary],
                        "Working (Hours)": [working_hours],
                        "Actual Date": [date]
                    })], ignore_index=True)
        csv_buffer = io.StringIO()
        result_df.to_csv(csv_buffer, index=False)
        return csv_buffer.getvalue()
    except Exception as e:
        st.error(f"Lỗi: {e}")
        return None

st.title("Chuyển đổi Excel sang CSV")
uploaded_file = st.file_uploader("Chọn file Excel", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Chuyển đổi"):
        csv_data = convert_xlsx(uploaded_file)
        if csv_data:
            st.download_button(
                label="Tải file CSV",
                data=csv_data,
                file_name="output.csv",
                mime="text/csv"
            )