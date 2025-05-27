import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers

st.title("Pipe-Delimited TXT to Excel Combiner")

uploaded_files = st.file_uploader("Upload multiple .txt files (pipe-separated)", type="txt", accept_multiple_files=True)

if uploaded_files:
    all_data = pd.DataFrame()
    for i, file in enumerate(uploaded_files):
        try:
            # Read file with pipe separator
            df = pd.read_csv(file, sep="|", dtype=str, engine="python", on_bad_lines='skip')

            # Drop header row if it's not the first file
            if i > 0:
                df = df.iloc[1:]  # skip header row

            all_data = pd.concat([all_data, df], ignore_index=True)

        except Exception as e:
            st.error(f"❌ Error in {file.name}: {e}")

    if not all_data.empty:
        st.success(f"✅ {len(uploaded_files)} files combined successfully!")
        st.dataframe(all_data)

        # Save to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            all_data.to_excel(writer, index=False, sheet_name="Combined")
            worksheet = writer.sheets["Combined"]
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.number_format = numbers.FORMAT_GENERAL

        st.download_button(
            "📥 Download Combined Excel",
            data=output.getvalue(),
            file_name="Combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No data was combined — check your files for valid content.")
