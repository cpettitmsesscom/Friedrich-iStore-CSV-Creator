
import streamlit as st
import pandas as pd
from io import StringIO

def process_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = xls.parse(sheet)

    df_cleaned = df.iloc[:, 1:]
    tag_row = df_cleaned.columns[1:]  # Tags are column headers after removing col A and the label column
    model_numbers = df_cleaned.iloc[1, 1:]  # Row 2 (index 1) contains model numbers

    data = {}
    for tag, model in zip(tag_row, model_numbers):
        if pd.notna(model):
            data[model] = data.get(model, 0) + 1

    result_df = pd.DataFrame(list(data.items()), columns=["Model Number", "Quantity"])

    output = StringIO()
    result_df.to_csv(output, index=False)
    output.seek(0)
    return output

st.title("Friedrich iStore CSV Creator")
uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        processed_file = process_file(file)
        st.download_button(
            label=f"Download CSV: {file.name.replace('.xlsx', '.csv')}",
            data=processed_file,
            file_name=f"iStore_{file.name.replace('.xlsx', '.csv')}",
            mime="text/csv"
        )
