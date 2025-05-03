import streamlit as st
import openpyxl
from io import BytesIO

# ---------------- Custom Styles ---------------- #
st.set_page_config(page_title="welcome to noqari 1.0!!!", layout="centered")

custom_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Lexend', sans-serif;
    background-color: #f8f9fb;
}

h1 {
    background: linear-gradient(90deg, #7F5AF0, #2CB67D);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.8rem;
    font-weight: 700;
    margin-bottom: 0.5rem;
    text-align: center;
}

.uploadbox {
    padding: 1rem;
    border-radius: 12px;
    background-color: #ffffff;
    border: 1px solid #e6e6e6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ---------------- Header & Tagline ---------------- #
st.markdown("""
<h1>🌸 welcome to noqari 1.0!!! 🌸</h1>
<div style="text-align:center; font-size:1.1rem; margin-bottom:20px;">
    <em>noqari: saving lives and reputations since its inception on May 3rd, 2025.</em>
</div>
""", unsafe_allow_html=True)

# ---------------- File Upload UI ---------------- #
with st.container():
    st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📤 Upload your PCARD_OPEN.xlsx file here:", type="xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- Excel Logic ---------------- #
if uploaded_file:
    st.success("✅ File uploaded! Processing...")

    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.worksheets[0]
    sheet2 = wb.worksheets[1]

    max_row = sheet1.max_row

    # Step 1: F&G&H in column A (Sheet1 and Sheet2)
    for sheet in [sheet1, sheet2]:
        for row in range(2, max_row + 1):
            sheet[f"A{row}"] = f"=F{row}&G{row}&H{row}"

    # Step 2: VLOOKUP formulas in P & Q
    for row in range(2, max_row + 1):
        sheet2[f"P{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:P),FALSE),"")'
        sheet2[f"Q{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:Q),FALSE),"")'

    # Step 3: Clean-up in R & S
    for row in range(2, max_row + 1):
        sheet2[f"R{row}"] = f'=IF(P{row}=0,"",P{row})'
        sheet2[f"S{row}"] = f'=IF(Q{row}=0,"",Q{row})'

    # Step 4: Paste values from R & S over P & Q
    for row in range(2, max_row + 1):
        r_val = sheet2[f"R{row}"].value
        s_val = sheet2[f"S{row}"].value
        sheet2[f"P{row}"].value = r_val
        sheet2[f"Q{row}"].value = s_val

    # Save result
    output = BytesIO()
    wb.save(output)

    st.success("🎉 All done! Your file is ready to download:")
    st.download_button(
        label="📥 Download Updated File",
        data=output.getvalue(),
        file_name="PCARD_OPEN_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Upload the PCARD_OPEN.xlsx file to get started.")


