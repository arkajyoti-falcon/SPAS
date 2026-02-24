# app.py
# Structured output only (no debug tables), with:
# 1) Infeed System -> VDS (has / not has)
# 2) Induct -> components + counts
# 3) Main Loop -> type based on visible sheets (Linear CBS preferred if visible, else Loop CBS)
# 4) Output Chutes -> chute types + counts (excluding sensors/IO)
# 5) Others (placeholder)

import re
import math
from typing import Dict, Any, Optional, Tuple, List

import pandas as pd
import streamlit as st
import openpyxl


# ----------------------------- Parsing helpers -----------------------------
def normalize_sheet_name(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def _is_excel_error(v) -> bool:
    return isinstance(v, str) and v.strip().startswith("#")


def _to_float_money(v) -> float:
    if v is None:
        return 0.0
    try:
        if pd.isna(v):
            return 0.0
    except Exception:
        pass
    if _is_excel_error(v):
        return 0.0
    if isinstance(v, bool):
        return float(int(v))
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return 0.0
        return float(v)

    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "null"):
        return 0.0
    s = re.sub(r"[^\d\-,.]", "", s).replace(",", "")
    if s in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _to_int(v) -> int:
    if v is None:
        return 0
    try:
        if pd.isna(v):
            return 0
    except Exception:
        pass
    if _is_excel_error(v):
        return 0
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return 0
        return int(round(float(v)))

    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "null"):
        return 0
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return 0
    return int(round(float(m.group(0))))


# ----------------------------- Workbook helpers -----------------------------
def visible_sheetnames(wb: openpyxl.Workbook) -> List[str]:
    out = []
    for name in wb.sheetnames:
        ws = wb[name]
        if getattr(ws, "sheet_state", "visible") == "visible":
            out.append(name)
    return out


def find_sheet_by_name_tolerant(wb: openpyxl.Workbook, desired: str) -> Optional[str]:
    want = normalize_sheet_name(desired)
    norm_map = {normalize_sheet_name(n): n for n in wb.sheetnames}
    if want in norm_map:
        return norm_map[want]
    for n in wb.sheetnames:
        if normalize_sheet_name(n).startswith(want):
            return n
    return None


def choose_main_loop_type_from_visible(wb: openpyxl.Workbook) -> str:
    """
    Rule per your ask:
      - If visible sheet 'Linear CBS' exists => Main Loop: Linear CBS
      - Else => Main Loop: Loop CBS
    """
    vis = visible_sheetnames(wb)
    vis_norm = {normalize_sheet_name(n) for n in vis}

    if normalize_sheet_name("Linear CBS") in vis_norm:
        return "Linear CBS"
    return "Loop CBS"


def pick_costing_sheet_visible_only(wb: openpyxl.Workbook) -> str:
    """For Induct extraction only: use visible Loop CBS else visible Linear CBS."""
    vis = visible_sheetnames(wb)
    if not vis:
        raise ValueError("No visible worksheets found in the workbook.")

    want_loop = normalize_sheet_name("Loop CBS")
    want_linear = normalize_sheet_name("Linear CBS")
    norm_map = {normalize_sheet_name(n): n for n in vis}

    if want_loop in norm_map:
        return norm_map[want_loop]
    if want_linear in norm_map:
        return norm_map[want_linear]

    for n in vis:
        if normalize_sheet_name(n).startswith(want_loop):
            return n
    for n in vis:
        if normalize_sheet_name(n).startswith(want_linear):
            return n

    raise ValueError(
        "Could not find a visible sheet matching 'Loop CBS' or 'Linear CBS'. "
        f"Visible sheets: {vis}"
    )


def find_cell(ws, needle: str) -> Optional[Tuple[int, int]]:
    tgt = needle.strip().lower()
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.strip().lower() == tgt:
                return (c.row, c.column)
    return None


def read_section_kv(ws, section_header: str) -> Dict[str, Any]:
    pos = find_cell(ws, section_header)
    if not pos:
        raise ValueError(f"Could not find '{section_header}' header in the selected sheet.")

    start_row, header_col = pos
    label_col = header_col
    value_col = header_col + 1

    kv: Dict[str, Any] = {}
    blanks = 0

    for r in range(start_row + 1, ws.max_row + 1):
        label = ws.cell(r, label_col).value
        value = ws.cell(r, value_col).value

        if label is None or (isinstance(label, str) and label.strip() == ""):
            blanks += 1
            if blanks >= 2:
                break
            continue
        blanks = 0

        if isinstance(label, str):
            s = label.strip().lower()
            if not (s.startswith("enter") or s.startswith("select")):
                break

        kv[str(label).strip()] = value

    return kv


# ----------------------------- Induct -----------------------------
def extract_induct_counts_no_calc(kv: Dict[str, Any]) -> Dict[str, int]:
    out = {
        "intelligent_angle_merge": 0,
        "buffer_conveyor": 0,
        "orientation_loading_conveyor": 0,
        "spacing_conveyor": 0,
        "receiving_conveyor": 0,
    }

    for k, v in kv.items():
        kl = k.lower()

        if "angle merge" in kl:  # Angle Merge == Intelligent Angle Merge
            out["intelligent_angle_merge"] = _to_int(v)
        elif "short buffer" in kl:
            out["buffer_conveyor"] = _to_int(v)
        elif "spacing conveyor" in kl:
            out["spacing_conveyor"] = _to_int(v)
        elif "orientation conveyor" in kl:  # Orientation == Loading
            out["orientation_loading_conveyor"] = _to_int(v)
        elif "single operator loading conveyor" in kl:
            out["orientation_loading_conveyor"] = _to_int(v)
        elif "loading conveyor" in kl and "operator" not in kl:
            out["orientation_loading_conveyor"] = _to_int(v)
        elif "receiving conveyor" in kl:
            out["receiving_conveyor"] = _to_int(v)

    return out


# ----------------------------- Output Chutes -----------------------------
EXCLUDE_CHUTE_KEYWORDS = [
    "sensor", "sensors", "i/o", " io ", "input", "output",
    "wiring", "cable", "connector", "terminal",
    "plc", "relay", "module", "panel",
    "lamp", "light", "indicator", "beacon",
    "switch", "photoeye", "photocell", "scanner",
    "power", "supply", "psu",
]

def _looks_like_accessory(desc: str) -> bool:
    d = " " + desc.lower().strip() + " "
    return any(k in d for k in EXCLUDE_CHUTE_KEYWORDS)

def _looks_like_chute(desc: str) -> bool:
    d = desc.lower()
    if "chute" not in d:
        return False
    if _looks_like_accessory(desc):
        return False
    return True


def extract_output_chutes(uploaded_file, wb: openpyxl.Workbook) -> List[Dict[str, Any]]:
    sheet = find_sheet_by_name_tolerant(wb, "Destinations")
    if not sheet:
        return []

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    desc_col = None
    qty_col = None
    for c in df.columns:
        cl = c.lower()
        if cl == "description":
            desc_col = c
        if cl in ("qty", "quantity"):
            qty_col = c
    if desc_col is None or qty_col is None:
        return []

    base = df[[desc_col, qty_col]].copy()
    base[desc_col] = base[desc_col].astype(str).fillna("").str.strip()
    base[qty_col] = base[qty_col].apply(_to_int)

    base = base[(base[qty_col] != 0) & (base[desc_col].apply(_looks_like_chute))]
    if base.empty:
        return []

    out = (
        base.groupby(desc_col, as_index=False)[qty_col]
        .sum()
        .rename(columns={desc_col: "Chute", qty_col: "Count"})
        .sort_values("Chute")
        .reset_index(drop=True)
    )
    out["Count"] = out["Count"].astype(int)

    # Return list of "4.1 Output Chutes 1", "4.2 Output Chutes 2" style items
    items = []
    for i, row in out.iterrows():
        items.append({"title": f"4.{i+1}. Output Chutes {i+1}", "type": row["Chute"], "count": int(row["Count"])})
    return items


# ----------------------------- VDS detection (Arm Sorter) -----------------------------
def detect_vds_presence_from_arm_sorter(wb: openpyxl.Workbook) -> str:
    """
    Your rule:
      - Arm Sorter sheet may be hidden/unhidden
      - Table: S.No., Description, Quantity, UOM, Unit Price, Total
      - If ANY Total value > 0 => HAS VDS
      - Else NO VDS
    """
    sheet = find_sheet_by_name_tolerant(wb, "Arm Sorter")
    if not sheet:
        return "unknown"

    ws = wb[sheet]

    header_row = None
    col_total = None

    # Find a row containing header "Total"
    for r in range(1, min(ws.max_row, 200) + 1):
        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() == "total":
                # likely header cell
                header_row = r
                col_total = c
                break
        if header_row:
            break

    if not header_row or not col_total:
        return "unknown"

    # Scan totals below header (stop on 3 blank rows or on next 'Total' summary line)
    blank_streak = 0
    found_any_row = False
    has_nonzero = False

    for r in range(header_row + 1, min(ws.max_row, header_row + 80) + 1):
        tot_raw = ws.cell(r, col_total).value

        # blank stopping
        if tot_raw is None or str(tot_raw).strip() == "":
            blank_streak += 1
            if blank_streak >= 3 and found_any_row:
                break
            continue
        blank_streak = 0

        found_any_row = True
        tot_val = _to_float_money(tot_raw)
        if tot_val > 0:
            has_nonzero = True
            break

    if not found_any_row:
        return "unknown"
    return "has" if has_nonzero else "not_has"


# ----------------------------- UI -----------------------------
st.set_page_config(page_title="CBS Structured Extractor", layout="wide")
st.title("CBS Structured Output")

uploaded = st.file_uploader("Upload costing workbook (.xlsx)", type=["xlsx"])
hide_zero_induct = st.checkbox("Hide induct components with 0 count", value=True)

if uploaded:
    try:
        # data_only=True is important for totals (cached formula results)
        wb = openpyxl.load_workbook(uploaded, data_only=True)

        # MAIN LOOP TYPE from visible sheets
        main_loop_type = choose_main_loop_type_from_visible(wb)

        # VDS
        vds_status = detect_vds_presence_from_arm_sorter(wb)

        # INDUCT from visible Loop/Linear costing sheet
        costing_sheet = pick_costing_sheet_visible_only(wb)
        kv = read_section_kv(wb[costing_sheet], "Main Inducts")
        induct_counts = extract_induct_counts_no_calc(kv)

        induct_components = [
            {"component": "Intelligent Angle Merge Conveyor", "count": induct_counts["intelligent_angle_merge"]},
            {"component": "Buffer Conveyor", "count": induct_counts["buffer_conveyor"]},
            {"component": "Orientation / Loading Conveyor", "count": induct_counts["orientation_loading_conveyor"]},
            {"component": "Spacing Conveyor", "count": induct_counts["spacing_conveyor"]},
            {"component": "Receiving Conveyor", "count": induct_counts["receiving_conveyor"]},
        ]
        if hide_zero_induct:
            induct_components = [x for x in induct_components if int(x["count"]) != 0]

        # OUTPUT CHUTES
        # pandas needs file pointer reset
        uploaded.seek(0)
        output_chute_items = extract_output_chutes(uploaded, wb)

        # -------- Render in your requested outline style --------
        st.subheader("Structure")

        st.markdown("**1. Infeed System**")
        st.markdown("   **1.1. Infeed Conveyors**")
        st.markdown("   **1.2. VDS Loop**")
        st.markdown(f"   - VDS: **{'HAS' if vds_status == 'has' else ('NOT HAS' if vds_status == 'not_has' else 'UNKNOWN')}**")

        st.markdown("**2. Induct**")
        if induct_components:
            for i, item in enumerate(induct_components, start=1):
                st.markdown(f"   **2.{i}. Induct Components {i}**")
                st.markdown(f"   - {item['component']}: **{item['count']}**")
        else:
            st.markdown("   - (No induct components found)")

        st.markdown(f"**3. Main Loop : {main_loop_type}**")

        st.markdown("**4. Output Chutes**")
        if output_chute_items:
            for item in output_chute_items:
                st.markdown(f"   **{item['title']}**")
                st.markdown(f"   - {item['type']}: **{item['count']}**")
        else:
            st.markdown("   - (No output chutes found)")

        st.markdown("**5. Others**")

        # Optional machine-readable JSON (kept compact)
        st.subheader("Structured JSON")
        st.json({
            "infeed_system": {"vds": vds_status},
            "induct": induct_components,
            "main_loop": {"type": main_loop_type},
            "output_chutes": output_chute_items,
            "others": []
        })

    except Exception as e:
        st.error(str(e))
else:
    st.caption("Upload the costing workbook to generate the outline.")