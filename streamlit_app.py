import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import CellIsRule
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

# ---------------- Excel Logic (Pure-Values + EBBS Formulas + Conditional Formatting) ---------------- #
if uploaded_file:
    st.markdown("<div></div>", unsafe_allow_html=True)

    wb = openpyxl.load_workbook(uploaded_file)
    sheet1 = wb.worksheets[0]
    sheet2 = wb.worksheets[1]

    # 1) A-column formulas
    for sht in (sheet1, sheet2):
        for r in range(2, sht.max_row + 1):
            cell = sht[f"A{r}"]
            cell.value = f"=F{r}&G{r}&H{r}"
            cell.font = Font(name="Calibri", size=11)

    # 2) Build lookup from Sheet1
    lookup = {}
    for r in range(2, sheet1.max_row + 1):
        f = sheet1.cell(r, 6).value or ""
        g = sheet1.cell(r, 7).value or ""
        h = sheet1.cell(r, 8).value or ""
        key = f"{f}{g}{h}"
        lookup[key] = (
            sheet1.cell(r, 16).value,
            sheet1.cell(r, 17).value
        )

    # 3) Write P & Q in Sheet2
    for r in range(2, sheet2.max_row + 1):
        f = sheet2.cell(r, 6).value or ""
        g = sheet2.cell(r, 7).value or ""
        h = sheet2.cell(r, 8).value or ""
        key = f"{f}{g}{h}"
        p_val, q_val = lookup.get(key, ("", ""))
        sheet2.cell(r, 16).value = p_val
        sheet2.cell(r, 17).value = q_val

    # 4) Inject EBBS formulas into M, N, O
    max_r = sheet2.max_row
    for r in range(2, max_r + 1):
        # M: difference A - E
        cell_m = sheet2.cell(row=r, column=13)
        cell_m.value = f"=$A{r}-$E{r}"
        cell_m.font = Font(name="Calibri", size=11)
        # N: bucket based on L
        cell_n = sheet2.cell(row=r, column=14)
        cell_n.value = (
            f"=IF(AND($L{r}<=7),(\"< 7\"),"
            f"IF(AND($L{r}>7,$L{r}<=11),(\"8-11\"),"
            f"IF(AND($L{r}>11,$L{r}<=15),(\"12-15\"),"
            f"IF(AND($L{r}>15,$L{r}<=30),(\"16-30\"),"
            f"IF(AND($L{r}>30,$L{r}<=45),(\"30-45\"),"
            f"IF(AND($L{r}>45,$L{r}<=59),(\"46-59\"),"
            f"IF($L{r}>59,(\"60 +\"),(\"Invalid\")))))))"
        )
        cell_n.font = Font(name="Calibri", size=11)
        # O: date = F + 16
        cell_o = sheet2.cell(row=r, column=15)
        cell_o.value = f'=TEXT(F{r}+16,"mm/dd/yyyy")'
        cell_o.font = Font(name="Calibri", size=11)

    # 5) Apply STATIC conditional-formatting to column L (days-difference)
    cf_range = f"L2:L{max_r}"
    rules = [
        ("lessThan",    ["7"],    "C6EFCE"),   # L<7
        ("between",     ["8","11"], "FFEB9C"),   # 8≤L≤11
        ("between",     ["12","15"],"FCE4D6"),   # 12≤L≤15
        ("between",     ["16","30"],"FFC7CE"),   # 16≤L≤30
        ("between",     ["30","45"],"FFC7CE"),   # 30≤L≤45
        ("between",     ["46","59"],"FFC7CE"),   # 46≤L≤59
        ("greaterThan", ["60"],   "FFC7CE"),   # L>60
    ]
    for op, vals, colour in rules:
        rule = CellIsRule(
            operator=op,
            formula=vals,
            stopIfTrue=True,
            fill=PatternFill(
                fill_type="solid",
                start_color=f"FF{colour}",
                end_color=f"FF{colour}"
            )
        )
        sheet2.conditional_formatting.add(cf_range, rule)

    # 6) Save & download
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
  <strong>NOTE:</strong> To ensure the code runs correctly, the file must be renamed to 
  <code>PCARD_OPEN</code> and saved in <code>.xlsx</code> format.<br>
  Files with a different name or format will not be processed.
</div>
<div class="thank-you">sincerely, your tiny tab fairy</div>
""", unsafe_allow_html=True)
