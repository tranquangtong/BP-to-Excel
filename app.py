import streamlit as st
import pandas as pd
import io
import zipfile
import os

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

st.title("Convert BP to Excel/Google Sheets")
uploaded_files = st.file_uploader("Chọn file Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Chuyển đổi và tải xuống tất cả"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for uploaded_file in uploaded_files:
                csv_data = convert_xlsx(uploaded_file)
                if csv_data:
                    zip_file.writestr(f"{uploaded_file.name.split('.')[0]}.csv", csv_data)
        st.download_button(
            label="Tải xuống tất cả file CSV",
            data=zip_buffer.getvalue(),
            file_name="output_csvs.zip",
            mime="application/zip"
        )
