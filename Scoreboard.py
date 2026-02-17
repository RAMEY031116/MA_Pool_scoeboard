import streamlit as st
import pandas as pd
import io

st.title("Excel Sheet Combiner")
st.write("Upload multiple Excel files and combine sheets into one workbook.")

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

sheet_to_extract = st.text_input(
    "Sheet name to extract (leave blank to extract ALL sheets)",
    value=""
)

if st.button("Combine Excel Files"):
    if not uploaded_files:
        st.error("Please upload at least one Excel file.")
    else:
        output = pd.ExcelWriter("combined.xlsx", engine="openpyxl")

        for file in uploaded_files:
            file_name = file.name.replace(".xlsx", "")
            excel_file = pd.ExcelFile(file)

            # If user wants a specific sheet
            if sheet_to_extract:
                if sheet_to_extract in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_to_extract)
                    sheet_name = f"{sheet_to_extract}_{file_name}"
                    df.to_excel(output, sheet_name=sheet_name, index=False)
                else:
                    st.warning(f"{sheet_to_extract} not found in {file.name}")
            else:
                # Extract all sheets
                for sheet in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet)
                    sheet_name = f"{sheet}_{file_name}"
                    df.to_excel(output, sheet_name=sheet_name, index=False)

        output.save()

        # Download link
        with open("combined.xlsx", "rb") as f:
            st.download_button(
                label="Download Combined Excel File",
                data=f,
                file_name="combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("Excel files combined successfully!")
``
