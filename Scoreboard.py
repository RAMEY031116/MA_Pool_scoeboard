import streamlit as st
from openpyxl import load_workbook, Workbook
import re
import io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="Excel Sheet Combiner", layout="centered")

st.title("📊 Excel Sheet Combiner (Preserves Formulas + Basic Formatting)")
st.write("Upload Excel files and extract sheets while preserving formulas and standard formatting.")

uploaded_files = st.file_uploader(
    "📁 Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

sheet_to_extract = st.text_input(
    "🔎 Sheet name to extract (leave blank to extract ALL sheets)",
    value=""
)

st.markdown("---")

# ----------- Helper: Clean filename to readable name -----------
def clean_filename(filename: str) -> str:
    """
    Convert a filename like '20260211_4467 South Acton ESG.xlsx'
    into 'South Acton ESG'.
    """
    name = filename.replace(".xlsx", "")
    name = re.sub(r"^\d+[ _-]*\d*\s*", "", name)  # remove numeric prefixes
    return name.strip()


# ----------- MAIN LOGIC ----------------
if st.button("🚀 Combine Excel Files"):
    if not uploaded_files:
        st.error("⚠ Please upload at least one file.")
        st.stop()

    progress = st.progress(0)
    status = st.empty()

    # Create final workbook
    output_wb = Workbook()
    # Remove default empty sheet
    output_wb.remove(output_wb.active)

    total_files = len(uploaded_files)

    for index, uploaded in enumerate(uploaded_files):
        status.text(f"Processing **{uploaded.name}** ...")

        try:
            wb = load_workbook(uploaded, data_only=False)
        except Exception as e:
            st.error(f"❌ Error loading {uploaded.name}: {e}")
            continue

        base_name = clean_filename(uploaded.name)

        # ----------- Option: Extract only one sheet -----------
        if sheet_to_extract:
            if sheet_to_extract not in wb.sheetnames:
                st.warning(f"⚠ Sheet '{sheet_to_extract}' not found in {uploaded.name}")
            else:
                source_sheet = wb[sheet_to_extract]
                new_sheet_name = f"{base_name} - {sheet_to_extract}"
                new_sheet = output_wb.create_sheet(new_sheet_name)

                # Copy cell values + styles safely
                for row in source_sheet.iter_rows():
                    for cell in row:
                        new_cell = new_sheet[cell.coordinate]
                        new_cell.value = cell.value

                        if cell.has_style:
                            new_cell.font = cell.font.copy()
                            new_cell.border = cell.border.copy()
                            new_cell.fill = cell.fill.copy()
                            new_cell.alignment = cell.alignment.copy()
                            new_cell.number_format = cell.number_format
                            new_cell.protection = cell.protection.copy()

                # Copy merged cells
                for merged_range in source_sheet.merged_cells.ranges:
                    new_sheet.merge_cells(str(merged_range))

                # Copy dimensions
                for col_letter, dim in source_sheet.column_dimensions.items():
                    new_sheet.column_dimensions[col_letter].width = dim.width

                for row_idx, dim in source_sheet.row_dimensions.items():
                    new_sheet.row_dimensions[row_idx].height = dim.height

        # ----------- Option: Extract ALL sheets -----------
        else:
            for sheet_name in wb.sheetnames:
                source_sheet = wb[sheet_name]
                new_sheet_name = f"{base_name} - {sheet_name}"
                new_sheet = output_wb.create_sheet(new_sheet_name)

                # Copy cell values + styles
                for row in source_sheet.iter_rows():
                    for cell in row:
                        new_cell = new_sheet[cell.coordinate]
                        new_cell.value = cell.value

                        if cell.has_style:
                            new_cell.font = cell.font.copy()
                            new_cell.border = cell.border.copy()
                            new_cell.fill = cell.fill.copy()
                            new_cell.alignment = cell.alignment.copy()
                            new_cell.number_format = cell.number_format
                            new_cell.protection = cell.protection.copy()

                # Copy merged cells
                for merged_range in source_sheet.merged_cells.ranges:
                    new_sheet.merge_cells(str(merged_range))

                # Copy dimensions
                for col_letter, dim in source_sheet.column_dimensions.items():
                    new_sheet.column_dimensions[col_letter].width = dim.width

                for row_idx, dim in source_sheet.row_dimensions.items():
                    new_sheet.row_dimensions[row_idx].height = dim.height

        progress.progress((index + 1) / total_files)

    # Save workbook
    output_path = "combined.xlsx"
    output_wb.save(output_path)

    with open(output_path, "rb") as f:
        st.download_button(
            "⬇ Download Combined Excel File",
            f,
            file_name="combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.success("🎉 Done! Your combined Excel file is ready.")
