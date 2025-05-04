import streamlit as st
import openpyxl
from io import BytesIO
import base64

# ---------------- Custom Styles ---------------- #
st.set_page_config(page_title="noqari 1.0", layout="centered")

custom_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;700&family=DM+Serif+Display&display=swap');

html, body, [class*="css"] {
    font-family: 'Lexend', sans-serif;
    background-color: #ffffff;
    padding: 24px;
}

section.main {
    background-color: #ffffff !important;
}

.block-container {
    background-color: #ffffff;
    border-radius: 18px;
    padding: 3rem 2rem;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
    max-width: 800px;
    margin: auto;
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
    margin-top: 0.5rem;
    margin-bottom: 0.5rem;
    color: #FF69B4;
}

.uploadbox {
    padding: 1rem;
    border-radius: 12px;
    background-color: #ffffff;
    border: 1px solid #e6e6e6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    margin-top: 0.5rem;
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
    margin-top: 8px;
}

section[data-testid="stFileUploader"] label {
    display: none !important;
    margin: 0 !important;
    padding: 0 !important;
}

div[data-testid="stAlert"] {
    text-align: center;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ---------------- Header & Tagline ---------------- #
st.markdown("""
<div class="title-text">noqari 1.0</div>
<div style="text-align:center; font-size:1.6rem;">💌</div>
<div class="tagline">the happiest place on earth (for VLOOKUP formulas).</div>
""", unsafe_allow_html=True)

# ---------------- File Upload ---------------- #
with st.container():
    st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type="xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- Info Message ---------------- #
st.info("Please upload your PCARD_OPEN.xlsx file to get started!")

# ---------------- Excel Logic (100% Clean) ---------------- #
if uploaded_file:
    # File uploaded
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
        sheet2[f"P{row}"].value = sheet2[f"R{row}"].value
        sheet2[f"Q{row}"].value = sheet2[f"S{row}"].value

    output = BytesIO()
    wb.save(output)

    # ✨ Custom success message
    st.markdown("<div style='text-align:center; font-size: 1.2rem; margin-top: 1.2rem;'>✨ All yours! Your file is ready to go!! ✨</div>", unsafe_allow_html=True)

    # 🎀 Gradient download button
    b64 = base64.b64encode(output.getvalue()).decode()
    st.markdown(f"""
        <div style="text-align:center; margin-top: 2rem;">
            <a href="data:application/octet-stream;base64,{b64}" download="PCARD_OPEN_Processed.xlsx"
               style="
                   display: inline-block;
                   padding: 0.75rem 1.5rem;
                   font-size: 1rem;
                   font-weight: 600;
                   color: white;
                   background: linear-gradient(90deg, #FF69B4, #FFD700);
                   border: none;
                   border-radius: 10px;
                   text-decoration: none;
                   box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
                   transition: all 0.3s ease-in-out;
               "
               onmouseover="this.style.opacity=0.9"
               onmouseout="this.style.opacity=1"
            >Download Processed File</a>
        </div>
    """, unsafe_allow_html=True)

# ---------------- Footer ---------------- #
st.markdown("""
<div class="footer-note">
    <strong>NOTE:</strong> To ensure the code runs correctly, the file must be renamed to <code>PCARD_OPEN</code> and saved in <code>.xlsx</code> format.<br>
    Files with a different name or format will not be processed.
</div>
<div class="thank-you">sincerely, your tiny tab fairy</div>
""", unsafe_allow_html=True)
