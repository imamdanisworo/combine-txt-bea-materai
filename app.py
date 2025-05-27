import streamlit as st
import pandas as pd
import io
from openpyxl.styles import numbers
from openpyxl.utils import get_column_letter

st.title("📄 Pipe-Delimited TXT to Excel Combiner")

uploaded_files = st.file_uploader(
    "Upload multiple .txt files (pipe-separated with headers)",
    type="txt",
    accept_multiple_files=True
)

if uploaded_files:
    all_data = pd.DataFrame()
    header_ref = None  # Store header reference from the first file

    for i, file in enumerate(uploaded_files):
        try:
            df = pd.read_csv(file, sep="|", dtype=str, engine="python", on_bad_lines='skip')

            if i == 0:
                header_ref = list(df.columns)
                all_data = df.copy()
            else:
                first_row = df.iloc[0].tolist()
                if first_row == header_ref:
                    df = df.iloc[1:]  # Skip duplicate header
                all_data = pd.concat([all_data, df], ignore_index=True)

        except Exception as e:
            st.error(f"❌ Error processing {file.name}: {e}")

    if not all_data.empty:
        st.success(f"✅ Combined {len(uploaded_files)} files successfully.")
        st.dataframe(all_data)

        # Write to Excel and apply formatting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            all_data.to_excel(writer, index=False, sheet_name="Combined")

            # Apply number formatting to selected columns: A, B, O, P
            worksheet = writer.sheets["Combined"]
            columns_to_format = [1, 2, 15, 16]  # A=1, B=2, O=15, P=16

            for row in worksheet.iter_rows(min_row=2):  # Skip header
                for col_idx in columns_to_format:
                    if col_idx <= len(row):  # Make sure column exists
                        cell = row[col_idx - 1]
                        try:
                            float(cell.value)  # Check if value is numeric
                            cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                        except:
                            pass  # Ignore non-numeric

        st.download_button(
            "📥 Download Combined Excel",
            data=output.getvalue(),
            file_name="Combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid data found in uploaded files.")
