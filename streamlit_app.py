import streamlit as st
import openpyxl
from io import BytesIO

# ---------------- Custom Styles ---------------- #
st.set_page_config(page_title="noqari 1.0", layout="centered")

custom_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;700&family=DM+Serif+Display&display=swap');

html, body, [class*="css"] {
    font-family: 'Lexend', sans-serif;
    background-color: #fefdf8;
    border: 24px solid transparent;
    padding: 24px;
    background-image: 
        linear-gradient(#fefdf8, #fefdf8),
        url("https://i.imgur.com/0R0o0Aw.png");
    background-origin: border-box;
    background-clip: content-box, border-box;
    background-repeat: no-repeat;
    background-size: cover;
}

.title-text {
    font-family: 'Georgia', serif;
    font-size: 3rem;
    font-weight: bold;
    color: #111111;
    text-align: center;
    margin-bottom: 0.2rem;
}

.tagline {
    font-family: 'DM Serif Display', serif;
    text-align: center;
    font-size: 1.4rem;
    margin-bottom: 20px;
    color: #FF69B4;
}

.uploadbox {
    padding: 1rem;
    border-radius: 12px;
    background-color: #ffffff;
    border: 1px solid #e6e6e6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.footer-note {
    font-size: 0.95rem;
    text-align: center;
    margin-top: 50px;
    color: #333;
}

.thank-you {
    font-family: 'Georgia', serif;
    text-align: center;
    font-size: 1.1rem;
    color: #FF69B4;
    margin-top: 16px;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ---------------- Header & Tagline ---------------- #
st.markdown("""
<div class="title-text">noqari 1.0</div>
<div class="tagline">sincerely, your tiny tab fairy</div>
""", unsafe_allow_html=True)

# ---------------- File Upload UI ---------------- #
with st.container():
    st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload your PCARD_OPEN.xlsx file here:", type="xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- Excel Logic (Untouched) ---------------- #
if uploaded_file:
    st.success("💌 File uploaded! Processing...")

    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.worksheets[0]
    sheet2 = wb.worksheets[1]

    max_row = sheet1.max_row

    for sheet in [sheet1, sheet2]:
        for row in range(2, max_row + 1):
            sheet[f"A{row}"] = f"=F{row}&G{row}&H{row}"

    for row in range(2, max_row + 1):
        sheet2[f"P{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:P),FALSE),"")'
        sheet2[f"Q{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:Q),FALSE),"")'
        sheet2[f"R{row}"] = f'=IF(P{row}=0,"",P{row})'
        sheet2[f"S{row}"] = f'=IF(Q{row}=0,"",Q{row})'

    for row in range(2, max_row + 1):
        r_val = sheet2[f"R{row}"].value
        s_val = sheet2[f"S{row}"].value
        sheet2[f"P{row}"].value = r_val
        sheet2[f"Q{row}"].value = s_val

    output = BytesIO()
    wb.save(output)

    st.success("✨ All done! Your file is ready to download:")
    st.download_button(
        label="📥 Download Updated File",
        data=output.getvalue(),
        file_name="PCARD_OPEN_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Please upload your PCARD_OPEN.xlsx file to get started!")

# ---------------- Footer Note ---------------- #
st.markdown("""
<div class="footer-note">
    <strong>Note:</strong> Please ensure your file is renamed to <code>PCARD_OPEN.xlsx</code> before uploading,<br>
    or the code will not be able to process it.
</div>
<div class="thank-you">Thanks so much!</div>
""", unsafe_allow_html=True)
