import streamlit as st
from openpyxl import load_workbook, Workbook
import re
import io

st.set_page_config(page_title="Excel Sheet Combiner", layout="centered")

st.title("📊 Excel Sheet Combiner (Keeps Formatting + Formulas)")
st.write("Upload Excel files and extract sheets with **all formatting preserved**.")

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

def extract_clean_name(filename):
    # Remove leading numbers and underscores
    clean = re.sub(r"^\d+[_-]*\s*", "", filename)
    clean = clean.replace(".xlsx", "")
    return clean.strip()

if st.button("🚀 Combine Excel Files"):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one Excel file.")
    else:
        progress = st.progress(0)
        status = st.empty()

        output_wb = Workbook()
        # Remove default empty sheet
        default_sheet = output_wb.active
        output_wb.remove(default_sheet)

        total_files = len(uploaded_files)

        for idx, uploaded in enumerate(uploaded_files):
            status.text(f"Processing **{uploaded.name}**...")

            wb = load_workbook(uploaded, data_only=False)  # keeps formulas

            base_name = extract_clean_name(uploaded.name)

            if sheet_to_extract:
                if sheet_to_extract in wb.sheetnames:
                    source_sheet = wb[sheet_to_extract]
                    new_sheet = output_wb.create_sheet(f"{base_name} - {sheet_to_extract}")

                    # Copy cells, formatting, merges, column widths
                    for row in source_sheet:
                        for cell in row:
                            new_cell = new_sheet[cell.coordinate]
                            new_cell.value = cell.value
                            if cell.has_style:
                                new_cell._style = cell._style
                            new_cell.number_format = cell.number_format

                    # Copy column widths
                    for col_letter, col_dim in source_sheet.column_dimensions.items():
                        new_sheet.column_dimensions[col_letter].width = col_dim.width

                    # Copy row heights
                    for row_idx, row_dim in source_sheet.row_dimensions.items():
                        new_sheet.row_dimensions[row_idx].height = row_dim.height

                else:
                    st.warning(f"⚠️ Sheet '{sheet_to_extract}' not found in {uploaded.name}")

            else:
                # Copy all sheets
                for sheet_name in wb.sheetnames:
                    source_sheet = wb[sheet_name]
                    new_sheet = output_wb.create_sheet(f"{base_name} - {sheet_name}")

                    for row in source_sheet:
                        for cell in row:
                            new_cell = new_sheet[cell.coordinate]
                            new_cell.value = cell.value
                            if cell.has_style:
                                new_cell._style = cell._style
                            new_cell.number_format = cell.number_format

                    for col_letter, col_dim in source_sheet.column_dimensions.items():
                        new_sheet.column_dimensions[col_letter].width = col_dim.width

                    for row_idx, row_dim in source_sheet.row_dimensions.items():
                        new_sheet.row_dimensions[row_idx].height = row_dim.height

            progress.progress((idx + 1) / total_files)

        # Save combined file
        output_path = "combined.xlsx"
        output_wb.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                "⬇️ Download Combined Excel File",
                f,
                file_name="combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("🎉 Sheets merged with full formatting + formulas preserved!")
