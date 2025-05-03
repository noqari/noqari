import streamlit as st
import openpyxl
from io import BytesIO

st.set_page_config(page_title="welcome to noqari!", layout="centered")
st.title("👋 welcome to noqari!")
st.markdown("Upload the **PCARD_OPEN.xlsx** file below and get back a fully processed version.")

uploaded_file = st.file_uploader("Upload PCARD_OPEN.xlsx", type="xlsx")

if uploaded_file:
    st.success("✅ File uploaded! Processing...")

    # Load workbook
    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.worksheets[0]
    sheet2 = wb.worksheets[1]

    max_row = sheet1.max_row

    # Step 1: =F2&G2&H2 in A2 down in both sheets
    for sheet in [sheet1, sheet2]:
        for row in range(2, max_row + 1):
            sheet[f"A{row}"] = f"=F{row}&G{row}&H{row}"

    # Step 2: VLOOKUP formulas in Sheet2 columns P and Q
    for row in range(2, max_row + 1):
        sheet2[f"P{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:P),FALSE),"")'
        sheet2[f"Q{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:Q),FALSE),"")'

    # Step 3: Cleanup formulas in R and S
    for row in range(2, max_row + 1):
        sheet2[f"R{row}"] = f'=IF(P{row}=0,"",P{row})'
        sheet2[f"S{row}"] = f'=IF(Q{row}=0,"",Q{row})'

    # Step 4: Copy VALUES from R and S over P and Q
    for row in range(2, max_row + 1):
        r_val = sheet2[f"R{row}"].value
        s_val = sheet2[f"S{row}"].value
        sheet2[f"P{row}"].value = r_val
        sheet2[f"Q{row}"].value = s_val

    # Save result
    output = BytesIO()
    wb.save(output)

    st.success("🎉 Done! Download your updated PCARD_OPEN file:")
    st.download_button(
        label="📥 Download Updated File",
        data=output.getvalue(),
        file_name="PCARD_OPEN_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Upload the PCARD_OPEN.xlsx file to get started.")
