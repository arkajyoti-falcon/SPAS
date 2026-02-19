# app.py
# Stage 3 (fixed):
# Upload DXF -> extract INSERT blocks -> remove noise -> map Mechanical->Actual using backend XLSX
# -> detect unmapped (by Mechanical/Component Name) -> call Groq API to predict Actual Name for unmapped
# -> final table (PER COMPONENT NAME): Component Name, Actual Name, Count, X, Y, Z, AI_Generated, AI_Confidence
# -> sorted by (X,Y,Z)
#
# Install:
#   pip install streamlit ezdxf pandas openpyxl requests python-dotenv
#
# Files needed next to app.py:
#   CBS_Component_Mapping.xlsx
#   DrINsaNE - JUST A BOY (Lyrics) Japanese Rap Song.mp3   (optional)
#
# Run:
#   streamlit run app.py
#
# API Key:
#   Windows PowerShell:  $env:GROQ_API_KEY="YOUR_KEY"
#   Linux/macOS:        export GROQ_API_KEY="YOUR_KEY"

import base64
import json
import os
import re
import tempfile
from collections import defaultdict

import ezdxf
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

load_dotenv()

# -------------------- CONFIG --------------------
MAPPING_FILE = "CBS_Component_Mapping.xlsx"  # backend mapping file
GROQ_BASE_URL = "https://api.groq.com/openai/v1"
GROQ_MODEL = "openai/gpt-oss-120b"

# Noise rules (aggressive)
NOISE_PREFIXES = ("*",)        # drops *U### and any *... anonymous/dynamic blocks
NOISE_CONTAINS = ("$0$",)      # drops xref-style names
NOISE_EXACT = {"GENAXEH"}      # drafting helpers
NOISE_REGEX = re.compile(r"^(TITLE BLOCK.*|FAL_TTL_.*|_ACMFILLED.*)$", re.IGNORECASE)

FAL_CODE_RE = re.compile(r"(FAL_[A-Z0-9_]+)(?:\.\d+)?", re.IGNORECASE)

# --- Background music (optional) ---
BACKGROUND_MUSIC_FILE = "DrINsaNE - JUST A BOY (Lyrics) Japanese Rap Song.mp3"


# -------------------- Background music --------------------
def play_background_music():
    """Play background music hidden, using components.html with base64-encoded audio."""
    if not os.path.exists(BACKGROUND_MUSIC_FILE):
        return

    with open(BACKGROUND_MUSIC_FILE, "rb") as f:
        audio_b64 = base64.b64encode(f.read()).decode()

    # hide the iframe container
    st.markdown(
        """
        <style>
            iframe[title="streamlit_components_v1.html"] { display:none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    components.html(
        f"""
        <audio id="bgm" autoplay loop>
            <source src="data:audio/mp3;base64,{audio_b64}" type="audio/mpeg">
        </audio>
        <script>
            const a = document.getElementById('bgm');
            if (a) {{
                a.volume = 1.0;
                const p = a.play();
                if (p) {{
                    p.catch(function() {{
                        // Autoplay blocked — start on first click anywhere
                        try {{
                            window.parent.document.addEventListener('click', function go() {{
                                a.play();
                                window.parent.document.removeEventListener('click', go);
                            }}, {{ once: true }});
                        }} catch(e) {{}}
                    }});
                }}
            }}
        </script>
        """,
        height=10,
    )


# -------------------- HELPERS --------------------
def is_noise_block_name(name: str) -> bool:
    if not name:
        return True
    n = str(name).strip()
    if not n:
        return True
    up = n.upper()

    if up in NOISE_EXACT:
        return True

    for p in NOISE_PREFIXES:
        if up.startswith(p):
            return True

    for c in NOISE_CONTAINS:
        if c in n:
            return True

    if NOISE_REGEX.match(n):
        return True

    if up.startswith("FAL_BLK_ATT_"):
        return True

    return False


def extract_base_fal_code(s: str) -> str:
    if not s:
        return ""
    m = FAL_CODE_RE.search(str(s).upper())
    return m.group(1).upper() if m else ""


def normalize_mech_key(s: str) -> str:
    """
    Normalize a mechanical name to a stable key.
    Examples:
      'FAL_FS002V02(Without weighing)' -> 'FAL_FS002V02'
      'FAL_FS002V02.1' -> 'FAL_FS002V02'
      'FAL_FS002V02' -> 'FAL_FS002V02'
    """
    if not s:
        return ""
    s = str(s).strip().upper()
    fal = extract_base_fal_code(s)
    if fal:
        return fal
    m = re.match(r"^[A-Z0-9_]+", s)
    return m.group(0) if m else s


@st.cache_data(show_spinner=False)
def load_mapping():
    if not os.path.exists(MAPPING_FILE):
        raise FileNotFoundError(
            f"Backend mapping file not found: {MAPPING_FILE}. "
            f"Keep it in the same folder as app.py (or update MAPPING_FILE)."
        )

    df = pd.read_excel(MAPPING_FILE)
    if "Mechanical Name" not in df.columns or "Actual Name" not in df.columns:
        raise ValueError("Mapping XLSX must contain columns: 'Mechanical Name', 'Actual Name'")

    exact_lookup = {}
    norm_lookup = {}

    for _, r in df.iterrows():
        mech = str(r.get("Mechanical Name") or "").strip()
        actual = str(r.get("Actual Name") or "").strip()
        if not mech or not actual:
            continue

        mech_up = mech.upper()
        exact_lookup.setdefault(mech_up, actual)

        norm = normalize_mech_key(mech_up)
        if norm:
            norm_lookup.setdefault(norm, actual)

    actual_names = sorted(set(exact_lookup.values()))

    mapping_pairs = df[["Mechanical Name", "Actual Name"]].dropna().copy()
    mapping_pairs["Mechanical Name"] = mapping_pairs["Mechanical Name"].astype(str)
    mapping_pairs["Actual Name"] = mapping_pairs["Actual Name"].astype(str)

    return exact_lookup, norm_lookup, actual_names, mapping_pairs


def read_dxf_inserts(dxf_path: str):
    try:
        doc, _aud = ezdxf.recover.readfile(dxf_path)
    except Exception:
        doc = ezdxf.readfile(dxf_path)

    msp = doc.modelspace()

    inserts = []
    total_inserts = 0
    removed_noise = 0

    for e in msp.query("INSERT"):
        total_inserts += 1
        name = getattr(e.dxf, "name", None)

        if is_noise_block_name(name):
            removed_noise += 1
            continue

        x = y = z = None
        try:
            p = e.dxf.insert
            x = float(getattr(p, "x", 0.0))
            y = float(getattr(p, "y", 0.0))
            z = float(getattr(p, "z", 0.0))
        except Exception:
            pass

        inserts.append({"component_name": str(name).strip(), "x": x, "y": y, "z": z})

    return inserts, {"total_inserts": total_inserts, "removed_noise": removed_noise, "kept": len(inserts)}


def groq_api_key():
    try:
        k = st.secrets.get("GROQ_API_KEY", None)
        if k:
            return k
    except Exception:
        pass
    return os.getenv("GROQ_API_KEY")


def build_mapping_context_text(mapping_pairs_df: pd.DataFrame, max_chars: int = 60000) -> str:
    lines = []
    used = 0
    for mech, actual in zip(mapping_pairs_df["Mechanical Name"], mapping_pairs_df["Actual Name"]):
        line = f"{str(mech).strip()} => {str(actual).strip()}"
        if used + len(line) + 1 > max_chars:
            break
        lines.append(line)
        used += len(line) + 1
    return "\n".join(lines)


def build_system_prompt() -> str:
    return (
        "Role: Senior Warehouse Automation Solution Architect (CBS domain) focused on CAD-to-BOM normalization.\n"
        "Objective: Given unmapped DXF block names (mechanical identifiers), predict the correct standardized "
        "'Actual Name' used in CBS proposals.\n\n"
        "You will be provided:\n"
        "1) Reference mapping: Mechanical Name => Actual Name (ground truth examples).\n"
        "2) Allowed Actual Names list (you MUST pick from this list or output UNKNOWN).\n"
        "3) Unmapped mechanical names to classify.\n\n"
        "Decision rules:\n"
        "- Use reference mapping patterns first (same code family, suffix patterns, bracket hints).\n"
        "- Use mechanical-code similarity (e.g., FAL_FSxxx family) only as secondary evidence.\n"
        "- Never invent a new Actual Name.\n"
        "- If uncertain, output UNKNOWN.\n\n"
        "Output:\n"
        "- Return JSON only with one prediction per mechanical_name.\n"
        "- Provide confidence in [0.0, 1.0]."
    )


def groq_predict_actual_names(unmapped_components, actual_names, mapping_context):
    """
    unmapped_components: list[str]
    returns dict[MECH_UPPER -> {"actual_name": str, "confidence": float}]
    """
    key = groq_api_key()
    if not key:
        raise RuntimeError("GROQ_API_KEY not found in environment or st.secrets.")

    url = f"{GROQ_BASE_URL}/chat/completions"
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}

    system = build_system_prompt()
    CHUNK = 40
    results = {}

    # send original strings, but store results in UPPER for stable lookup
    for i in range(0, len(unmapped_components), CHUNK):
        chunk = unmapped_components[i : i + CHUNK]

        user = (
            "REFERENCE MAPPING (Mechanical Name => Actual Name):\n"
            f"{mapping_context}\n\n"
            "ALLOWED ACTUAL NAMES (choose one of these or UNKNOWN):\n"
            f"{json.dumps(actual_names, ensure_ascii=False)}\n\n"
            "UNMAPPED MECHANICAL COMPONENT NAMES:\n"
            f"{json.dumps(chunk, ensure_ascii=False)}\n\n"
            "Return JSON ONLY in this exact format:\n"
            "{\n"
            "  \"predictions\": [\n"
            "    {\"mechanical_name\": \"...\", \"actual_name\": \"...\", \"confidence\": 0.0}\n"
            "  ]\n"
            "}\n"
        )

        payload = {
            "model": GROQ_MODEL,
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
            "temperature": 0,
            "response_format": {"type": "json_object"},
        }

        resp = requests.post(url, headers=headers, json=payload, timeout=90)
        resp.raise_for_status()

        content = resp.json()["choices"][0]["message"]["content"] or ""
        try:
            obj = json.loads(content)
        except Exception:
            for name in chunk:
                results[str(name).strip().upper()] = {"actual_name": "UNKNOWN", "confidence": 0.0}
            continue

        preds = obj.get("predictions", []) if isinstance(obj, dict) else []
        seen = set()

        for p in preds:
            mn = str(p.get("mechanical_name", "")).strip()
            an = str(p.get("actual_name", "UNKNOWN")).strip()
            cf = p.get("confidence", 0.0)

            try:
                cf = float(cf)
            except Exception:
                cf = 0.0

            mn_up = mn.upper()
            if not mn_up:
                continue

            # enforce allow-list
            if an != "UNKNOWN" and an not in actual_names:
                an = "UNKNOWN"
                cf = 0.0

            results[mn_up] = {"actual_name": an, "confidence": max(0.0, min(1.0, cf))}
            seen.add(mn_up)

        for name in chunk:
            nm_up = str(name).strip().upper()
            if nm_up not in results and nm_up not in seen:
                results[nm_up] = {"actual_name": "UNKNOWN", "confidence": 0.0}

    return results


def map_component_to_actual(component_name, exact_lookup, norm_lookup):
    """
    Returns (actual_name or None)
    """
    if not component_name:
        return None

    comp_up = str(component_name).strip().upper()
    if comp_up in exact_lookup:
        return exact_lookup[comp_up]

    norm = normalize_mech_key(comp_up)
    if norm and norm in norm_lookup:
        return norm_lookup[norm]

    return None


def build_component_summary(inserts):
    """
    Aggregate raw inserts per COMPONENT NAME (mechanical block name):
      component_name -> count + avg xyz
    """
    agg = defaultdict(lambda: {"count": 0, "sx": 0.0, "sy": 0.0, "sz": 0.0, "pos_n": 0})

    for it in inserts:
        mech = str(it["component_name"]).strip()
        a = agg[mech]
        a["count"] += 1

        x, y, z = it.get("x"), it.get("y"), it.get("z")
        if x is not None and y is not None and z is not None:
            a["sx"] += x
            a["sy"] += y
            a["sz"] += z
            a["pos_n"] += 1

    rows = []
    for mech, a in agg.items():
        if a["pos_n"] > 0:
            ax = a["sx"] / a["pos_n"]
            ay = a["sy"] / a["pos_n"]
            az = a["sz"] / a["pos_n"]
        else:
            ax = ay = az = None

        rows.append({"Component Name": mech, "Count": a["count"], "X": ax, "Y": ay, "Z": az})

    df = pd.DataFrame(rows)
    if not df.empty:
        df["_X"] = df["X"].fillna(1e18)
        df["_Y"] = df["Y"].fillna(1e18)
        df["_Z"] = df["Z"].fillna(1e18)
        df = df.sort_values(by=["_X", "_Y", "_Z", "Component Name"], ascending=[True, True, True, True])
        df = df.drop(columns=["_X", "_Y", "_Z"])
    return df


def build_final_table(component_df, exact_lookup, norm_lookup, ai_predictions):
    """
    Final table is PER COMPONENT NAME.
    Columns:
      Component Name (raw mechanical)
      Actual Name (from Excel, else AI, else UNKNOWN)
      Count, X,Y,Z
      AI_Generated, AI_Confidence
    """
    rows = []
    for _, r in component_df.iterrows():
        mech = str(r["Component Name"]).strip()
        mech_up = mech.upper()

        actual = map_component_to_actual(mech, exact_lookup, norm_lookup)
        ai_generated = False
        ai_conf = None

        if actual is None:
            pred = ai_predictions.get(mech_up, {"actual_name": "UNKNOWN", "confidence": 0.0})
            if pred.get("actual_name") and pred["actual_name"] != "UNKNOWN":
                actual = pred["actual_name"]
                ai_generated = True
                try:
                    ai_conf = float(pred.get("confidence", 0.0) or 0.0)
                except Exception:
                    ai_conf = 0.0
            else:
                actual = "UNKNOWN"

        rows.append({
            "Component Name": mech,
            "Actual Name": actual,
            "Count": int(r["Count"]),
            "X": r["X"], "Y": r["Y"], "Z": r["Z"],
            "AI_Generated": bool(ai_generated),
            "AI_Confidence": ai_conf,
        })

    df = pd.DataFrame(rows)

    # Sort by position
    if not df.empty:
        df["_X"] = df["X"].fillna(1e18)
        df["_Y"] = df["Y"].fillna(1e18)
        df["_Z"] = df["Z"].fillna(1e18)
        df = df.sort_values(by=["_X", "_Y", "_Z", "Actual Name", "Component Name"], ascending=[True, True, True, True, True])
        df = df.drop(columns=["_X", "_Y", "_Z"])

    return df


# -------------------- Streamlit UI --------------------
st.set_page_config(page_title="CBS DXF → Final Summary (Stage 3)", layout="wide")

play_background_music()

st.title("CBS DXF → Final Summary (Stage 3)")
st.caption("Upload DXF → extract components → remove noise → map via Excel → AI for unmapped → final table sorted by (X,Y,Z).")

uploaded = st.file_uploader("Upload DXF", type=["dxf"])

if uploaded:
    try:
        exact_lookup, norm_lookup, actual_names, mapping_pairs = load_mapping()
    except Exception as ex:
        st.error(str(ex))
        st.stop()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
        tmp.write(uploaded.getbuffer())
        dxf_path = tmp.name

    try:
        inserts, stats = read_dxf_inserts(dxf_path)
    except Exception as ex:
        st.error(f"Failed to process DXF: {ex}")
        st.stop()

    # Stage: aggregate per mechanical component name
    comp_df = build_component_summary(inserts)

    # Detect unmapped component types (unique mechanical names not in excel)
    unmapped = []
    for mech in comp_df["Component Name"].tolist():
        if map_component_to_actual(mech, exact_lookup, norm_lookup) is None:
            unmapped.append(mech)

    st.info(
        f"Modelspace INSERTs: {stats['total_inserts']} | "
        f"Noise removed: {stats['removed_noise']} | "
        f"Kept INSERTs: {stats['kept']} | "
        f"Unmapped component types: {len(unmapped)}"
    )

    # AI predictions only for unmapped component types
    ai_predictions = {}
    if len(unmapped) > 0:
        try:
            mapping_context = build_mapping_context_text(mapping_pairs, max_chars=60000)
            with st.spinner("Predicting Actual Names for unmapped components using Groq AI..."):
                ai_predictions = groq_predict_actual_names(unmapped, actual_names, mapping_context)
        except Exception as ex:
            st.error(f"Groq AI mapping failed: {ex}")
            ai_predictions = {}

    df_final = build_final_table(comp_df, exact_lookup, norm_lookup, ai_predictions)

    st.subheader("Final Table (Component Name, Actual Name, Count, X, Y, Z) — Sorted by Position")
    st.dataframe(df_final, use_container_width=True, hide_index=True)
