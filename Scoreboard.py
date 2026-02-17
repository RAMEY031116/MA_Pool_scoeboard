import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Sheet Combiner", layout="centered")

# --- Header ---
st.title("📊 Excel Sheet Combiner")
st.write("Upload multiple Excel files and combine selected sheets into one workbook.")

uploaded_files = st.file_uploader(
    "📁 Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

sheet_to_extract = st.text_input(
    "🔎 Sheet name to extract (leave blank to extract **all sheets**)",
    value=""
)

st.markdown("---")

# --- Main Logic ---
if st.button("🚀 Combine Excel Files"):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one Excel file.")
    else:
        progress = st.progress(0)
        status = st.empty()

        output_path = "combined.xlsx"

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            total_files = len(uploaded_files)

            for idx, file in enumerate(uploaded_files):
                file_name = file.name.replace(".xlsx", "")
                excel_file = pd.ExcelFile(file)

                status.text(f"Processing **{file.name}**...")

                if sheet_to_extract:
                    # Extract specific sheet
                    if sheet_to_extract in excel_file.sheet_names:
                        df = pd.read_excel(excel_file, sheet_name=sheet_to_extract)

                        # Clean sheet names to avoid Excel errors
                        sheet_name = f"{sheet_to_extract[:20]}_{file_name[:10]}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        st.warning(f"⚠️ Sheet **{sheet_to_extract}** not found in **{file.name}**")
                else:
                    # Extract all sheets
                    for sheet in excel_file.sheet_names:
                        df = pd.read_excel(excel_file, sheet_name=sheet)

                        sheet_name = f"{sheet[:20]}_{file_name[:10]}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                progress.progress((idx + 1) / total_files)

        # Download button
        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Download Combined Excel File",
                data=f,
                file_name="combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.success("🎉 Excel files combined successfully!")
