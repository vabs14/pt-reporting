"""
A&G Physiotherapy Report Generator — Streamlit web app
Upload PT_<Quarter>_<Year>_<HomeName>.xlsx → get .docx + .pdf
"""
import re, os, sys, tempfile, hashlib
import streamlit as st

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
from generate_pt_report import generate

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="A&G Physio Report Generator",
    page_icon="🏥",
    layout="centered",
)

LOGO = os.path.join(os.path.dirname(__file__), "src", "logo_clean.png")

# ── Auth ──────────────────────────────────────────────────────────────────────
def _hash(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def check_credentials(username: str, password: str) -> bool:
    users: dict = st.secrets.get("users", {})
    if username not in users:
        return False
    stored = users[username]
    # Accept plaintext or pre-hashed (sha256 hex, 64 chars)
    if len(stored) == 64 and all(c in "0123456789abcdef" for c in stored):
        return _hash(password) == stored
    return password == stored

def show_login():
    if os.path.exists(LOGO):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(LOGO, width=200)
    st.markdown("<h3 style='text-align:center'>Sign in to continue</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;color:grey'>A & G Physiotherapy Inc. — Internal Tool</p>",
                unsafe_allow_html=True)
    st.divider()

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in", use_container_width=True, type="primary")

    if submitted:
        if check_credentials(username.strip().lower(), password):
            st.session_state["authenticated"] = True
            st.session_state["username"] = username.strip().lower()
            st.rerun()
        else:
            st.error("Invalid username or password.")

# Gate the entire app
if not st.session_state.get("authenticated"):
    show_login()
    st.stop()

# ── App (authenticated users only) ────────────────────────────────────────────
if os.path.exists(LOGO):
    col1, col2 = st.columns([1, 3])
    with col1:
        st.image(LOGO, width=160)
    with col2:
        st.title("Quarterly Report Generator")
        st.caption("A & G Physiotherapy Inc.")
else:
    st.title("A&G Physio — Quarterly Report Generator")

# Logout button in top-right
with st.sidebar:
    st.markdown(f"Signed in as **{st.session_state.get('username', '')}**")
    if st.button("Sign out", use_container_width=True):
        st.session_state.clear()
        st.rerun()

st.divider()

# ── Filename parser ────────────────────────────────────────────────────────────
def parse_filename(name: str) -> tuple[str, str, str]:
    stem = re.sub(r"\.xlsx?$", "", name, flags=re.IGNORECASE)
    parts = stem.split("_")
    if len(parts) >= 4 and parts[0].upper() == "PT":
        raw_q   = parts[1].upper()
        raw_yr  = parts[2]
        raw_loc = "_".join(parts[3:])
        spaced  = re.sub(r"([a-z])([A-Z])", r"\1 \2", raw_loc)
        home    = re.sub(r"[-_]+", " ", spaced).strip().title()
        return f"{raw_q} {raw_yr}", raw_yr, home
    return "", "", stem.title()

# ── Upload widget ─────────────────────────────────────────────────────────────
st.markdown("### Upload Excel data file")
st.markdown(
    "File must be named **`PT_<Quarter>_<Year>_<HomeName>.xlsx`**  \n"
    "e.g. `PT_Q1_2026_BurtonManor.xlsx` or `PT_Q2_2026_Sunrise_LTC.xlsx`"
)

uploaded = st.file_uploader("Choose file", type=["xlsx", "xls"], label_visibility="collapsed")

if uploaded:
    quarter, year, home_name = parse_filename(uploaded.name)
    st.success(f"Detected — **{home_name}** · **{quarter}**")

    with st.expander("Edit details (optional)"):
        home_name = st.text_input("Care home name", value=home_name)
        quarter   = st.text_input("Quarter label (used in report)", value=quarter)

    if st.button("Generate Report", type="primary", use_container_width=True):
        with st.spinner("Generating charts and building document…"):
            with tempfile.TemporaryDirectory() as tmp:
                xl_path = os.path.join(tmp, uploaded.name)
                with open(xl_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                try:
                    docx_path = generate(xl_path, home_name=home_name)
                    pdf_path  = docx_path.replace(".docx", ".pdf")
                except Exception as e:
                    st.error(f"Report generation failed: {e}")
                    st.stop()

                with open(docx_path, "rb") as f:
                    docx_bytes = f.read()
                pdf_bytes = None
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()

        st.success("Done! Download your report below.")
        safe_home = re.sub(r"[^\w]+", "_", home_name).strip("_")
        base = f"PT_{quarter.replace(' ', '_')}_{safe_home}_Report"

        col_d, col_p = st.columns(2)
        with col_d:
            st.download_button(
                "⬇ Download .docx", data=docx_bytes, file_name=f"{base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with col_p:
            if pdf_bytes:
                st.download_button(
                    "⬇ Download .pdf", data=pdf_bytes, file_name=f"{base}.pdf",
                    mime="application/pdf", use_container_width=True,
                )
            else:
                st.info("PDF not generated (LibreOffice unavailable on this host).")

st.divider()
st.caption("A & G Physiotherapy Inc. · Internal tool")
