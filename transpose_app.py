
import streamlit as st
import pandas as pd
from io import StringIO

def process_file(uploaded_file):
    # Load Excel
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = xls.parse(sheet)

    # Remove first column (usually an index or empty)
    df_cleaned = df.iloc[:, 1:]

    # Tags are column headers, excluding the first column
    tag_row = df_cleaned.columns[1:]

    # Model numbers are in row index 2 (third row), skip first column
    model_numbers = df_cleaned.iloc[2, 1:]

    # Count occurrences of each model number
    model_counts = {}
    for tag, model in zip(tag_row, model_numbers):
        if pd.notna(model):
            model_counts[model] = model_counts.get(model, 0) + 1

    # Create result DataFrame
    result_df = pd.DataFrame(list(model_counts.items()), columns=["Model Number", "Quantity"])

    # Output CSV
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
