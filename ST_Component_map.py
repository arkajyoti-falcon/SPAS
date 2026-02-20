import io
import re
from pathlib import Path
import pandas as pd
import streamlit as st


# ===================== CONFIG =====================
DEFAULT_EXCEL_PATH = Path("CBS_Component_Mapping_WITH_DESCRIPTION.xlsx")
SHEET_NAME = "CBS Component Map"   # if not found, falls back to first sheet

IMAGE_DIR = Path("Component Image")
ALLOWED_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}  # case-insensitive

# UI tuning
THUMB_WIDTH = 150     # <-- bigger thumbnail
ROW_DIVIDER = True


# Backend column names
COL_MECH = "Mechanical Name"
COL_ACTUAL = "Actual Name"
COL_SECTION = "Section"
COL_DESC = "description"


# ===================== UTILS =====================
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())

def ensure_dirs():
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)

def load_excel_bytes_to_disk(file_bytes: bytes, target_path: Path):
    target_path.write_bytes(file_bytes)

def read_components(excel_path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(excel_path)
    sheet = SHEET_NAME if SHEET_NAME in xls.sheet_names else xls.sheet_names[0]
    raw = pd.read_excel(excel_path, sheet_name=sheet)

    cols = {str(c).strip().lower(): c for c in raw.columns}
    def pick(name: str) -> str | None:
        return cols.get(name.strip().lower())

    mech = pick(COL_MECH)
    actl = pick(COL_ACTUAL)
    sect = pick(COL_SECTION)
    desc = pick(COL_DESC) or pick("Description") or pick("description")

    missing = [n for n, v in [(COL_MECH, mech), (COL_ACTUAL, actl), (COL_SECTION, sect)] if v is None]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    df = pd.DataFrame({
        "Component_Code": raw[mech].fillna("").astype(str).map(_norm),
        "Component_name": raw[actl].fillna("").astype(str).map(_norm),
        "Section": raw[sect].fillna("").astype(str).map(_norm),
        "Description": (raw[desc].fillna("").astype(str).map(_norm) if desc else "")
    })

    df["Description"] = df["Description"].astype(str).str.slice(0, 250)
    return df

def write_components(excel_path: Path, df: pd.DataFrame):
    out = pd.DataFrame({
        COL_MECH: df["Component_Code"].fillna("").astype(str).map(_norm),
        COL_ACTUAL: df["Component_name"].fillna("").astype(str).map(_norm),
        COL_SECTION: df["Section"].fillna("").astype(str).map(_norm),
        COL_DESC: df["Description"].fillna("").astype(str).map(_norm).str.slice(0, 250),
    })

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name=SHEET_NAME, index=False)

def excel_bytes_for_download(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    out = pd.DataFrame({
        COL_MECH: df["Component_Code"].fillna("").astype(str).map(_norm),
        COL_ACTUAL: df["Component_name"].fillna("").astype(str).map(_norm),
        COL_SECTION: df["Section"].fillna("").astype(str).map(_norm),
        COL_DESC: df["Description"].fillna("").astype(str).map(_norm).str.slice(0, 250),
    })
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    buffer.seek(0)
    return buffer.getvalue()

def find_image_for_code(code: str) -> Path | None:
    code = _norm(code)
    if not code or not IMAGE_DIR.exists():
        return None

    # exact stem match (case-insensitive ext)
    for ext in ALLOWED_IMAGE_EXTS:
        for variant in (ext, ext.upper()):
            p = IMAGE_DIR / f"{code}{variant}"
            if p.exists() and p.is_file():
                return p

    # if user saved as any ext but same stem
    for p in IMAGE_DIR.iterdir():
        if p.is_file() and p.stem == code and p.suffix.lower() in ALLOWED_IMAGE_EXTS:
            return p

    return None

def save_uploaded_image(component_code: str, uploaded_file):
    if uploaded_file is None:
        return

    code = _norm(component_code)
    if not code:
        raise ValueError("Component_Code is required to save image.")

    ext = Path(uploaded_file.name).suffix.lower()
    if ext not in ALLOWED_IMAGE_EXTS:
        raise ValueError(f"Unsupported image type: {ext}")

    ensure_dirs()

    # remove existing images for same code
    for p in IMAGE_DIR.iterdir():
        if p.is_file() and p.stem == code and p.suffix.lower() in ALLOWED_IMAGE_EXTS:
            try:
                p.unlink()
            except:
                pass

    (IMAGE_DIR / f"{code}{ext}").write_bytes(uploaded_file.getvalue())


# ===================== UI =====================
st.set_page_config(page_title="CBS Component Mapping", layout="wide")
ensure_dirs()

# Small CSS polish (cleaner spacing + nicer buttons)
st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 1.2rem; }
      div[data-testid="stHorizontalBlock"] { align-items: center; }
      button[kind="secondary"] { border-radius: 12px; }
      button[kind="primary"] { border-radius: 12px; }
      .small-muted { color: #9aa0a6; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("CBS Component Mapping")

uploaded = st.file_uploader("Upload Excel (optional)", type=["xlsx", "xls"])

if "excel_path" not in st.session_state:
    st.session_state.excel_path = DEFAULT_EXCEL_PATH

if uploaded is not None:
    load_excel_bytes_to_disk(uploaded.getvalue(), DEFAULT_EXCEL_PATH)
    st.session_state.excel_path = DEFAULT_EXCEL_PATH
    st.session_state.reload_df = True

if "df" not in st.session_state or st.session_state.get("reload_df") is True:
    if not st.session_state.excel_path.exists():
        st.warning(f"Place `{DEFAULT_EXCEL_PATH.name}` next to `app.py` or upload it above.")
        st.stop()

    try:
        st.session_state.df = read_components(st.session_state.excel_path)
        st.session_state.reload_df = False
    except Exception as e:
        st.error(f"Failed to load Excel: {e}")
        st.stop()

df = st.session_state.df


@st.dialog("Image Preview")
def image_preview_dialog():
    p = st.session_state.get("preview_path", None)
    if not p:
        st.info("No image selected.")
        return
    st.image(p, use_column_width=True)


@st.dialog("Add Component")
def add_component_dialog():
    sections = sorted(set([s for s in df["Section"].dropna().tolist() if _norm(s)] + ["Infeed System", "Induct", "CBS", "Output Chutes"]))

    code = st.text_input("Component_Code (Mechanical Name) *")
    name = st.text_input("Component_name (Actual Name) *")
    section = st.selectbox("Section *", options=sections)
    desc = st.text_area("Description (max 250 chars)", height=100)
    img = st.file_uploader("Upload Image (saved as Component Image/<Component_Code>.<ext>)", type=["png", "jpg", "jpeg", "webp"])

    a, b = st.columns(2)
    with a:
        if st.button("Cancel", use_container_width=True):
            st.rerun()
    with b:
        if st.button("Add", type="primary", use_container_width=True):
            code_n, name_n, section_n = _norm(code), _norm(name), _norm(section)
            desc_n = _norm(desc)[:250]

            if not code_n or not name_n or not section_n:
                st.error("Component_Code, Component_name, and Section are required.")
                st.stop()

            if (df["Component_Code"].str.lower() == code_n.lower()).any():
                st.error("Component_Code already exists.")
                st.stop()

            if img is not None:
                save_uploaded_image(code_n, img)

            new_row = pd.DataFrame([{
                "Component_Code": code_n,
                "Component_name": name_n,
                "Section": section_n,
                "Description": desc_n
            }])

            st.session_state.df = pd.concat([df, new_row], ignore_index=True)
            write_components(st.session_state.excel_path, st.session_state.df)

            st.session_state.reload_df = True
            st.rerun()


@st.dialog("Edit Component")
def edit_component_dialog():
    idx = st.session_state.get("edit_idx", None)
    if idx is None or idx < 0 or idx >= len(st.session_state.df):
        st.error("Invalid selection.")
        st.stop()

    row = st.session_state.df.iloc[idx]
    old_code = row["Component_Code"]

    sections = sorted(set([s for s in st.session_state.df["Section"].dropna().tolist() if _norm(s)] + ["Infeed System", "Induct", "CBS", "Output Chutes"]))

    code = st.text_input("Component_Code (Mechanical Name) *", value=row["Component_Code"])
    name = st.text_input("Component_name (Actual Name) *", value=row["Component_name"])
    section = st.selectbox("Section *", options=sections, index=sections.index(row["Section"]) if row["Section"] in sections else 0)
    desc = st.text_area("Description (max 250 chars)", value=row["Description"], height=100)

    cur_img = find_image_for_code(old_code)
    if cur_img:
        st.image(str(cur_img), width=220, caption=f"Current: {cur_img.name}")
    else:
        st.info("No image found for this component.")

    img = st.file_uploader("Upload / Replace Image", type=["png", "jpg", "jpeg", "webp"])

    a, b = st.columns(2)
    with a:
        if st.button("Cancel", use_container_width=True):
            st.rerun()
    with b:
        if st.button("Save", type="primary", use_container_width=True):
            code_n, name_n, section_n = _norm(code), _norm(name), _norm(section)
            desc_n = _norm(desc)[:250]

            if not code_n or not name_n or not section_n:
                st.error("Component_Code, Component_name, and Section are required.")
                st.stop()

            # uniqueness if code changed
            if code_n.lower() != _norm(old_code).lower():
                if (st.session_state.df["Component_Code"].str.lower() == code_n.lower()).any():
                    st.error("New Component_Code already exists.")
                    st.stop()

                # rename existing image if any
                old_img = find_image_for_code(old_code)
                if old_img:
                    old_img.rename(IMAGE_DIR / f"{code_n}{old_img.suffix.lower()}")

            if img is not None:
                save_uploaded_image(code_n, img)

            st.session_state.df.at[idx, "Component_Code"] = code_n
            st.session_state.df.at[idx, "Component_name"] = name_n
            st.session_state.df.at[idx, "Section"] = section_n
            st.session_state.df.at[idx, "Description"] = desc_n

            write_components(st.session_state.excel_path, st.session_state.df)
            st.session_state.reload_df = True
            st.rerun()


# ---------- Top bar ----------
t1, t2, t3, t4 = st.columns([1.3, 2.6, 1.6, 2.0])
with t1:
    if st.button("➕ Add Component", type="primary", use_container_width=True):
        add_component_dialog()
with t2:
    search = st.text_input("Search", value="", placeholder="code / name / section")
with t3:
    st.download_button(
        "⬇️ Download Excel",
        data=excel_bytes_for_download(df),
        file_name="CBS_Component_Mapping_UPDATED.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with t4:
    st.markdown(f"<div class='small-muted'>Images: <code>{IMAGE_DIR}</code> | Thumb: {THUMB_WIDTH}px</div>", unsafe_allow_html=True)

st.divider()

# ---------- Filter ----------
if search.strip():
    s = search.strip().lower()
    view = df[
        df["Component_Code"].str.lower().str.contains(s, na=False) |
        df["Component_name"].str.lower().str.contains(s, na=False) |
        df["Section"].str.lower().str.contains(s, na=False)
    ].copy()
else:
    view = df.copy()

# ---------- Header ----------
hdr = st.columns([2.2, 2.4, 1.4, 3.2, 2.0, 0.7, 0.7])  # <-- wider Image col
hdr[0].markdown("**Component_Code**")
hdr[1].markdown("**Component_name**")
hdr[2].markdown("**Section**")
hdr[3].markdown("**Description**")
hdr[4].markdown("**Image**")
hdr[5].markdown("**Edit**")
hdr[6].markdown("**Delete**")

# ---------- Rows ----------
for _, r in view.reset_index().iterrows():
    idx = int(r["index"])
    code = r["Component_Code"]

    cols = st.columns([2.2, 2.4, 1.4, 3.2, 2.0, 0.7, 0.7])  # <-- match header

    cols[0].write(r["Component_Code"])
    cols[1].write(r["Component_name"])
    cols[2].write(r["Section"])
    cols[3].write(r["Description"])

    img_path = find_image_for_code(code)
    if img_path:
        cols[4].image(str(img_path), width=THUMB_WIDTH)
        # zoom icon
        if cols[4].button("🔍", key=f"zoom_{idx}", help="View larger"):
            st.session_state.preview_path = str(img_path)
            image_preview_dialog()
    else:
        cols[4].markdown("<span class='small-muted'>—</span>", unsafe_allow_html=True)

    # edit
    if cols[5].button("✏️", key=f"edit_{idx}", help="Edit"):
        st.session_state.edit_idx = idx
        edit_component_dialog()

    # delete
    if cols[6].button("🗑️", key=f"del_{idx}", help="Delete"):
        st.session_state.df = st.session_state.df.drop(index=idx).reset_index(drop=True)
        write_components(st.session_state.excel_path, st.session_state.df)
        st.session_state.reload_df = True
        st.rerun()

    if ROW_DIVIDER:
        st.divider()