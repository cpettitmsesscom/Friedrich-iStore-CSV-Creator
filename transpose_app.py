
import streamlit as st
import pandas as pd
from io import StringIO
import re

def process_file(uploaded_file):
    # Load Excel
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = xls.parse(sheet)

    # Remove first column (assumed to be an index/label column)
    df_cleaned = df.iloc[:, 1:]

    # Extract tags from column headers (skip first column)
    tags = df_cleaned.columns[1:]

    # Extract model numbers from row index 1 (3rd visible row)
    model_numbers = df_cleaned.iloc[1, 1:]

    # Calculate quantities
    model_counts = {}
    for tag, model in zip(tags, model_numbers):
        if pd.isna(model):
            continue

        quantity = 1  # Default quantity

        # Handle "1thru5" format
        match_range = re.search(r'(\d+)\s*thru\s*(\d+)', tag, flags=re.IGNORECASE)
        if match_range:
            start, end = map(int, match_range.groups())
            quantity = end - start + 1
        else:
            # Handle "1,2,3" format
            match_list = re.findall(r'\d+', tag)
            if match_list:
                quantity = len(match_list)

        model_counts[model] = model_counts.get(model, 0) + quantity

    # Convert to CSV without headers
    output = StringIO()
    pd.DataFrame(model_counts.items()).to_csv(output, index=False, header=False)
    return output.getvalue().encode("utf-8")

st.title("Friedrich iStore CSV Creator")
uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        processed_csv = process_file(file)
        st.download_button(
            label=f"Download CSV: {file.name.replace('.xlsx', '.csv')}",
            data=processed_csv,
            file_name=f"iStore_{file.name.replace('.xlsx', '.csv')}",
            mime="text/csv"
        )
