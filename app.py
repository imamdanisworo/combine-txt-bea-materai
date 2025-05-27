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
    header_ref = None

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
        st.success(f"✅ Combined {len(uploaded_files)} files successfully!")
        st.dataframe(all_data)

        # Convert only O and P columns (indexes 14, 15) to numeric
        for col_index in [14, 15]:  # 0-based index
            if col_index < len(all_data.columns):
                try:
                    all_data.iloc[:, col_index] = pd.to_numeric(
                        all_data.iloc[:, col_index].str.replace(",", "").str.strip(),
                        errors='coerce'
                    )
                except Exception as e:
                    st.warning(f"⚠️ Column {col_index + 1} couldn't be converted: {e}")

        # Export to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            all_data.to_excel(writer, index=False, sheet_name="Combined")

            worksheet = writer.sheets["Combined"]
            for row in worksheet.iter_rows(min_row=2):  # skip header
                for col_idx in [15, 16]:  # Excel columns O (15), P (16)
                    if col_idx <= len(row):
                        cell = row[col_idx - 1]
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

        st.download_button(
            "📥 Download Combined Excel",
            data=output.getvalue(),
            file_name="Combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid data found in uploaded files.")
