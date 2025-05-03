import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import openpyxl

st.set_page_config(page_title="🧾 APL PCARD + EBBS Processor", layout="centered")
st.title("📁 APL PCARD + EBBS Reconciler")

st.markdown("Upload your daily **APL_PCARD.xlsx** file and your current **EBBS Unreconciled Transactions_FYxx.xlsx** file.")

# Upload files
apl_file = st.file_uploader("Upload APL_PCARD.xlsx", type="xlsx")
ebbs_file = st.file_uploader("Upload EBBS Unreconciled Transactions file", type="xlsx")

if apl_file and ebbs_file:
    st.success("✅ Files uploaded successfully. Processing...")

    # Load APL_PCARD workbook
    apl_wb = openpyxl.load_workbook(apl_file)
    apl_ws1 = apl_wb.active
    apl_ws1.title = "Sheet1"

    # Create Sheet2 cleanly
    if "Sheet2" in apl_wb.sheetnames:
        del apl_wb["Sheet2"]
    apl_ws2 = apl_wb.create_sheet("Sheet2")

    max_row = apl_ws1.max_row
    max_col = 17  # Columns A to Q

    # Copy headers A1:Q1 to Sheet2
    for col in range(1, max_col + 1):
        apl_ws2.cell(row=1, column=col).value = apl_ws1.cell(row=1, column=col).value

    # Add formula to column A (F&G&H) in both sheets
    for row in range(2, max_row + 1):
        formula = f"=F{row}&G{row}&H{row}"
        apl_ws1[f"A{row}"] = formula
        apl_ws2[f"A{row}"] = formula

    # Add VLOOKUP and cleanup formulas in Sheet2
    for row in range(2, max_row + 1):
        apl_ws2[f"P{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:P),FALSE),"")'
        apl_ws2[f"Q{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:Q),FALSE),"")'
        apl_ws2[f"R{row}"] = f'=IF(P{row}=0,"",P{row})'
        apl_ws2[f"S{row}"] = f'=IF(Q{row}=0,"",Q{row})'

    # Copy only the final columns C to Q to merge into EBBS
    columns_to_copy = list("CDEFGHIJKLMNOPQ")
    copied_data = []
    for row in range(2, max_row + 1):
        row_data = [apl_ws2[f"{col}{row}"].value for col in columns_to_copy]
        copied_data.append(row_data)

    # Load EBBS workbook
    ebbs_wb = openpyxl.load_workbook(ebbs_file)
    ebbs_ws = ebbs_wb.active

    # Find the next empty row in EBBS
    next_row = ebbs_ws.max_row + 1

    # Add today's date in column A
    today_str = datetime.today().strftime('%m/%d/%Y')
    for i, row_data in enumerate(copied_data):
        ebbs_ws.cell(row=next_row + i, column=1).value = today_str  # Date in column A
        for j, value in enumerate(row_data):
            ebbs_ws.cell(row=next_row + i, column=2 + j).value = value  # Data in columns B onward

    # Copy L–N formulas from previous row (if needed)
    # Example: just reusing formulas from the last valid row
    for col in range(12, 15):  # Columns L (12) to N (14)
        last_formula = ebbs_ws.cell(row=next_row - 1, column=col).value
        if last_formula and isinstance(last_formula, str) and last_formula.startswith("="):
            for i in range(len(copied_data)):
                ebbs_ws.cell(row=next_row + i, column=col).value = last_formula

    # Save EBBS output
    output = BytesIO()
    ebbs_wb.save(output)

    st.success("🎉 Done! Download your updated EBBS file below:")
    st.download_button(
        label="📥 Download Updated EBBS File",
        data=output.getvalue(),
        file_name="EBBS_Unreconciled_Transactions_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Log summary (optional)
    st.markdown(f"**📌 {len(copied_data)} new transactions added on {today_str}**")

else:
    st.info("👆 Upload both Excel files to begin.")
