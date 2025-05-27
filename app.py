import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers

st.title("TXT to Excel Combiner")

uploaded_files = st.file_uploader("Upload multiple .txt files (comma-separated)", type="txt", accept_multiple_files=True)

if uploaded_files:
    all_data = pd.DataFrame()

    for i, file in enumerate(uploaded_files):
        # Read each TXT file as a comma-separated table
        df = pd.read_csv(file, sep=",", dtype=str)

        # If it's not the first file, drop the header row
        if i > 0:
            df = df[1:]

        all_data = pd.concat([all_data, df], ignore_index=True)

    st.success(f"? {len(uploaded_files)} files combined successfully!")
    st.dataframe(all_data)

    # Create Excel with general format
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        all_data.to_excel(writer, index=False, sheet_name="Combined")

        # Optional: Apply "General" format (default) explicitly to all cells
        workbook = writer.book
        worksheet = writer.sheets["Combined"]
        for row in worksheet.iter_rows():
            for cell in row:
                cell.number_format = numbers.FORMAT_GENERAL

    st.download_button(
        "?? Download Combined Excel",
        data=output.getvalue(),
        file_name="Combined.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
