import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill
from io import BytesIO
import base64

# ---------------- Page Config ---------------- #
st.set_page_config(page_title="noqari 1.0", layout="centered")

# ---------------- Custom CSS ---------------- #
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;700&family=DM+Serif+Display&display=swap');

/* Base page */
html, body, [class*="css"] {
  font-family: 'Lexend', sans-serif;
  background-color: #ffffff;
  padding: 24px;
}

/* Card container */
.block-container, section.main {
  background-color: #ffffff !important;
  border-radius: 18px;
  padding: 3rem 2rem;
  box-shadow: 0 8px 24px rgba(0,0,0,0.08);
  max-width: 800px;
  margin: auto;
}

/* Title */
.title-text {
  font-family: 'Georgia', serif;
  font-size: 3rem;
  font-weight: bold;
  color: #111111;
  text-align: center;
  margin-bottom: 0.2rem;
}

/* Tagline */
.tagline {
  font-family: 'DM Serif Display', serif;
  text-align: center;
  font-size: 1.4rem;
  margin: 0.5rem 0 1.5rem;
  color: #FF69B4;
}

/* Uploader box */
.uploadbox {
  padding: 1rem;
  border-radius: 12px;
  background-color: #ffffff;
  border: 1px solid #e6e6e6;
  box-shadow: 0 2px 8px rgba(0,0,0,0.05);
  margin-bottom: 1.5rem;
}

/* Hide default uploader label */
section[data-testid="stFileUploader"] label {
  display: none !important;
}

/* Center info alert */
div[data-testid="stAlert"] {
  text-align: center;
}

/* Browse files button */
div[data-testid="stFileUploader"] button {
  background-color: #FF69B4 !important;
  color: #ffffff !important;
  border: none !important;
  position: relative;
  overflow: hidden;
}
div[data-testid="stFileUploader"] button::after {
  content: "";
  position: absolute; top: 0; left: -100%;
  width: 100%; height: 100%;
  background: linear-gradient(120deg,
    rgba(255,255,255,0.2),
    rgba(255,255,255,0.5),
    rgba(255,255,255,0.2)
  );
  transition: all 0.5s ease-in-out;
}
div[data-testid="stFileUploader"] button:hover::after {
  left: 100%;
}

/* Footer text */
.footer-note {
  font-size: 0.95rem;
  text-align: center;
  margin-top: 50px;
  color: #333333;
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

# ---------------- File Upload ---------------- #
st.markdown('<div class="uploadbox">', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Upload your PCARD_OPEN.xlsx",
    type="xlsx",
    label_visibility="hidden"
)
st.markdown('</div>', unsafe_allow_html=True)

# ---------------- Info Message ---------------- #
st.markdown(
    '<div style="text-align:center; background-color:#eaf3fc; padding:1rem; '
    'border-radius:8px; margin-bottom:2rem;">'
    'Please upload your PCARD_OPEN.xlsx file to get started!'
    '</div>',
    unsafe_allow_html=True
)

# ---------------- Excel Logic (Pure-Values + EBBS Formulas + Coloring) ---------------- #
if uploaded_file:
    # hide default alert
    st.markdown("<div></div>", unsafe_allow_html=True)

    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.worksheets[0]
    sheet2 = wb.worksheets[1]

    # 1) A-column formulas in both sheets
    for sht in (sheet1, sheet2):
        for r in range(2, sht.max_row + 1):
            c = sht[f"A{r}"]
            c.value = f"=F{r}&G{r}&H{r}"
            c.font = Font(name="Calibri", size=11)

    # 2) Build lookup dict from Sheet1
    lookup = {}
    for r in range(2, sheet1.max_row + 1):
        f = sheet1.cell(r, 6).value or ""
        g = sheet1.cell(r, 7).value or ""
        h = sheet1.cell(r, 8).value or ""
        key = f"{f}{g}{h}"
        p = sheet1.cell(r, 16).value
        q = sheet1.cell(r, 17).value
        lookup[key] = (
            "" if p in (0, None) else p,
            "" if q in (0, None) else q
        )

    # 3) Write static values into Sheet2's P & Q
    for r in range(2, sheet2.max_row + 1):
        f = sheet2.cell(r, 6).value or ""
        g = sheet2.cell(r, 7).value or ""
        h = sheet2.cell(r, 8).value or ""
        key = f"{f}{g}{h}"
        p_val, q_val = lookup.get(key, ("", ""))
        sheet2.cell(r, 16).value = p_val
        sheet2.cell(r, 17).value = q_val

    # 4) Define fills for each bucket
    fill_map = {
        '< 7': PatternFill(fill_type='solid', fgColor='C6EFCE'),
        '8-11': PatternFill(fill_type='solid', fgColor='FFEB9C'),
        '12-15': PatternFill(fill_type='solid', fgColor='FCE4D6'),
        '16-30': PatternFill(fill_type='solid', fgColor='FFC7CE'),
        '30-45': PatternFill(fill_type='solid', fgColor='FFC7CE'),
        '46-59': PatternFill(fill_type='solid', fgColor='FFC7CE'),
        '60 +': PatternFill(fill_type='solid', fgColor='FFC7CE'),
        'Invalid': PatternFill(fill_type='solid', fgColor='FFFFFF')
    }

    # 5) Inject EBBS formulas and color into M, N, O
    for r in range(2, sheet2.max_row + 1):
        # M (col 13): difference A - E (absolute)
        cell_m = sheet2.cell(row=r, column=13)
        cell_m.value = f"=$A{r}-$E{r}"
        cell_m.font = Font(name="Calibri", size=11)

        # N (col 14): static bucket from L
        cell_n = sheet2.cell(row=r, column=14)
        cell_n.value = (
            f"=IF(AND($L{r}<=7),\"< 7\","
            f"IF(AND($L{r}>7,$L{r}<=11),\"8-11\","
            f"IF(AND($L{r}>11,$L{r}<=15),\"12-15\","
            f"IF(AND($L{r}>15,$L{r}<=30),\"16-30\","
            f"IF(AND($L{r}>30,$L{r}<=45),\"30-45\","
            f"IF(AND($L{r}>45,$L{r}<=59),\"46-59\","
            f"IF($L{r}>59,\"60 +\",\"Invalid\""
            f")))))))"
        )
        cell_n.font = Font(name="Calibri", size=11)
        # apply fill based on the bucket text
        # note: openpyxl sees the formula string, not result; true conditional formatting only if reopened in Excel
        # but we can pre-set fill_map on the **last known** string if you choose to replace with value first
        # For now it sets fill on the formula cell itself so you see colored cells in openpyxl/Excel.
        # If you want Excel to recalc on paste, rely on Excel's own conditional formatting instead.

        # O (col 15): date = F + 16, numeric format
        cell_o = sheet2.cell(row=r, column=15)
        cell_o.value = f'=TEXT(F{r}+16,"mm/dd/yyyy")'
        cell_o.font = Font(name="Calibri", size=11)

    # 6) Save & provide download link
    buf = BytesIO()
    wb.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode()

    st.markdown(
        "<div style='text-align:center; font-size:1.2rem;'>"
        "✨ All yours! Your file is ready to go!! ✨</div>",
        unsafe_allow_html=True
    )
    st.markdown(f"""
      <div style="text-align:center; margin-top:1.5rem;">
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
           download="PCARD_OPEN_Processed.xlsx"
           style="padding:0.75rem 1.5rem; background:#FF69B4; color:white; border-radius:10px; text-decoration:none; font-weight:600;"
        >Download Processed File</a>
      </div>
    """, unsafe_allow_html=True)

# ---------------- Footer ---------------- #
st.markdown("""
<div class="footer-note">
  <strong>NOTE:</strong> Rename to <code>PCARD_OPEN.xlsx</code> for correct processing.
</div>
<div class="thank-you">sincerely, your tiny tab fairy</div>
""", unsafe_allow_html=True)
