import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
import os

st.set_page_config(page_title="APL P-Card Processor", layout="centered")

st.title("🧾 APL P-Card Processor")
st.markdown("Upload the daily **APL_PCARD.xlsx** file to generate a formatted version for EBSS.")

uploaded_file = st.file_uploader("📤 Upload APL_PCARD Excel file", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded! Processing...")

    # Load workbook and worksheets
    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.active
    sheet1.title = "Sheet1"

    # Create Sheet2 if not already there
    if "Sheet2" in wb.sheetnames:
        del wb["Sheet2"]
    sheet2 = wb.create_sheet("Sheet2")

    max_row = sheet1.max_row
    max_col = 17  # Columns A to Q

    # Copy header row from A1:Q1 in Sheet1 to Sheet2
    for col in range(1, max_col + 1):
        sheet2.cell(row=1, column=col).value = sheet1.cell(row=1, column=col).value

    # Formula in A2 (F&G&H) for Sheet1 and Sheet2
    for row in range(2, max_row + 1):
        formula = f"=F{row}&G{row}&H{row}"
        sheet1[f"A{row}"] = formula
        sheet2[f"A{row}"] = formula

    # Add formulas to Sheet2
    for row in range(2, max_row + 1):
        sheet2[f"P{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:P),FALSE),"")'
        sheet2[f"Q{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:Q),FALSE),"")'
        sheet2[f"R{row}"] = f'=IF(P{row}=0,"",P{row})'
        sheet2[f"S{row}"] = f'=IF(Q{row}=0,"",Q{row})'

    # Save to buffer
    output = BytesIO()
    wb.save(output)
    st.success("✅ Processing complete!")

    # Download button
    st.download_button(
        label="📥 Download Processed File",
        data=output.getvalue(),
        file_name="APL_PCARD_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
