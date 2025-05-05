import streamlit as st
import openpyxl
from openpyxl.styles import Font
from io import BytesIO
import base64

# ---------------- Page Config ---------------- #
st.set_page_config(page_title="noqari 1.0", layout="centered")

# ---------------- Custom CSS ---------------- #
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;700&family=DM+Serif+Display&display=swap');

/* Base font & white background */
html, body, [class*="css"] {
    font-family: 'Lexend', sans-serif;
    background-color: #ffffff;
    padding: 24px;
}

/* Content container styling */
section.main { background-color: #ffffff !important; }
.block-container {
    background-color: #ffffff;
    border-radius: 18px;
    padding: 3rem 2rem;
    box-shadow: 0 8px 24px rgba(0,0,0,0.08);
    max-width: 800px;
    margin: auto;
}

/* Title & tagline */
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
    margin: 0.5rem 0;
    color: #FF69B4;
}

/* Uploader box */
.uploadbox {
    padding: 1rem;
    border-radius: 12px;
    background-color: #ffffff;
    border: 1px solid #e6e6e6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    margin-top: 0.5rem;
}

/* Info alert centering */
div[data-testid="stAlert"] { text-align: center; }

/* Hide the default empty label */
section[data-testid="stFileUploader"] label { display: none !important; }

/* “Browse files” button styling */
div[data-testid="stFileUploader"] button {
    background-color: #FF69B4 !important;
    color: #ffffff !important;
    border: none !important;
    position: relative;
    overflow: hidden;
}
div[data-testid="stFileUploader"] button::after {
    content: "";
    position: absolute;
    top: 0; left: -100%;
    width: 100%; height: 100%;
    background: linear-gradient(120deg,
      rgba(255,255,255,0.2),
      rgba(255,255,255,0.5),
      rgba(255,255,255,0.2));
    transition: all 0.5s ease-in-out;
}
div[data-testid="stFileUploader"] button:hover::after {
    left: 100%;
}

/* === Clear-file “X” as a plain pink letter === */
/* 1) Remove any background/padding from the progress container */
div[data-testid="stFileUploadProgress"] {
    background: none !important;
    box-shadow: none !important;
}
/* 2) Remove padding on the wrapper */
div[data-testid="stFileUploadProgress"] > div {
    background: none !important;
    padding: 0 !important;
}
/* 3) Style the Clear button itself */
button[aria-label^="Clear"],
button[title^="Clear"] {
    background: none !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
    color: #FF69B4 !important;    /* pink X */
    font-size: inherit !important; /* keep default size */
}

/* Footer */
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
</style>
""", unsafe_allow_html=True)

# ---------------- Header & Tagline ---------------- #
st.markdown("""
<div class="title-text">noqari 1.0</div>
<div style="text-align:center; font-size:1.6rem;">💌</div>
<div class="tagline">the happiest place on earth (for VLOOKUP formulas).</div>
""", unsafe_allow_html=True)

# ---------------- File Upload (Hidden Label) ---------------- #
st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "PCARD file uploader for accessibility",
    type="xlsx",
    label_visibility="hidden"
)
st.markdown('</div>', unsafe_allow_html=True)

# ---------------- Info Message ---------------- #
st.markdown(
    '<div style="text-align:center; background-color:#eaf3fc; padding:1rem; '
    'border-radius:8px; margin-top:1rem;">'
    'Please upload your PCARD_OPEN.xlsx file to get started!'
    '</div>',
    unsafe_allow_html=True
)

# ---------------- Excel Logic (100% Untouched) ---------------- #
if uploaded_file:
    # suppress default Streamlit alert
    st.markdown("<div></div>", unsafe_allow_html=True)

    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.worksheets[0]
    sheet2 = wb.worksheets[1]
    max_row = sheet1.max_row

    # A-column concatenation + Calibri 11
    for sheet in (sheet1, sheet2):
        for row in range(2, max_row + 1):
            cell = sheet[f"A{row}"]
            cell.value = f"=F{row}&G{row}&H{row}"
            cell.font = Font(name="Calibri", size=11)

    # Sheet2 P/Q/R/S formulas + values-only paste
    for row in range(2, max_row + 1):
        sheet2[f"P{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:P),FALSE),"")'
        sheet2[f"Q{row}"] = f'=IFERROR(VLOOKUP($A{row},Sheet1!$A:$Q,COLUMNS(Sheet1!$A:Q),FALSE),"")'
        sheet2[f"R{row}"] = f'=IF(P{row}=0,"",P{row})'
        sheet2[f"S{row}"] = f'=IF(Q{row}=0,"",Q{row})'
        sheet2[f"P{row}"].value = sheet2[f"R{row}"].value
        sheet2[f"Q{row}"].value = sheet2[f"S{row}"].value

    output = BytesIO()
    wb.save(output)
    b64 = base64.b64encode(output.getvalue()).decode()

    # ✨ Success message
    st.markdown(
        "<div style='text-align:center; font-size:1.2rem; margin-top:1.2rem;'>"
        "✨ All yours! Your file is ready to go!! ✨</div>",
        unsafe_allow_html=True
    )

    # 🎀 Solid pink download button
    st.markdown(f"""
    <div style="text-align:center; margin-top:2rem;">
      <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
         download="PCARD_OPEN_Processed.xlsx"
         style="
           display:inline-block;
           padding:0.75rem 1.5rem;
           font-size:1rem;
           font-weight:600;
           color:white;
           background-color:#FF69B4;
           border:none;
           border-radius:10px;
           text-decoration:none;
           box-shadow:0 4px 12px rgba(0,0,0,0.15);
           transition:all 0.3s ease-in-out;
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
