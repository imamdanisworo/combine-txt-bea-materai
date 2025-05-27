import streamlit as st
import pandas as pd
import io
from openpyxl.styles import numbers

st.title("📄 Pipe-Delimited TXT to Excel Combiner")

uploaded_files = st.file_uploader(
    "Upload multiple .txt files (pipe-separated with headers)",
    type="txt",
    accept_multiple_files=True
)

if uploaded_files:
    all_data = pd.DataFrame()
    header_ref = None  # To store the reference header

    for i, file in enumerate(uploaded_files):
        try:
            # Read the file as pipe-separated
            df = pd.read_csv(file, sep="|", dtype=str, engine="python", on_bad_lines='skip')

            # First file: keep header and store it
            if i == 0:
                header_ref = list(df.columns)
                all_data = df.copy()
            else:
                # Compare first row to header — if it matches, it's a duplicate header row
                first_row = df.iloc[0].tolist()
                if first_row == header_ref:
                    df = df.iloc[1:]  # Skip only the header row
                all_data = pd.concat([all_data, df], ignore_index=True)

        except Exception as e:
            st.error(f"❌ Error processing {file.name}: {e}")

    if not all_data.empty:
        st.success(f"✅ Combined {len(uploaded_files)} files successfully.")
        st.dataframe(all_data)

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
        st.warning("⚠️ No valid data found in uploaded files.")
