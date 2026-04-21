# Address Cleaning Agent (Enhanced)
# ------------------------------------------------------------
# Purpose:
#   Streamlit app to clean global addresses, with strong support for
#   "CITY./VILL./SETTLEMENT" prefix patterns as seen in add1.xlsx.
#   Produces:
#     - Cleaned Data (same columns as input + POSTAL_FLAG, DATA_QUALITY)
#     - Changes Log (row/field/before/after)
#
# Notes:
#   - Offline-first: no external API calls.
#   - Deterministic rules: explicit country mentions override ambiguous city mappings.
#   - Designed to reduce false country inference (e.g., Odessa -> US) by enforcing
#     explicit-country-wins.
# ------------------------------------------------------------

import streamlit as st
import pandas as pd
import re
import io
from typing import Dict, Tuple, List

st.set_page_config(page_title="Address Cleaning Agent (Enhanced)", page_icon="🗺️", layout="wide")

# -------------------------------
# Helpers
# -------------------------------

def norm_ws(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    s = str(x)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[\t\r\n]+", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s


def safe_upper(s: str) -> str:
    return norm_ws(s).upper()


# -------------------------------
# Core rule sets (enhancements)
# -------------------------------

# Prefixes seen in datasets like add1.xlsx
PREFIX_LOCALITY_RE = re.compile(
    r"^\s*(CITY|CIT|VILL|VILLAGE|SETTLEMENT|N\.P|N\.P\.|NP|TOWN|CITY,)\.?\s*[:\-]?\s*",
    re.I,
)

# Administrative noise that should never become CITY
ADMIN_NOISE_RE = re.compile(r"\b(ASSR|SSR|USSR)\b;?", re.I)

# Protect Saint cities so they don't get split into ST. + PETERSBURG
SAINT_FIX_RE = re.compile(r"\bST\.\s*(PETERSBURG|LOUIS|PAUL|HELENS)\b", re.I)

# General junk city tokens
JUNK_CITY = {"ASSR", "SSR", "USSR", "REGION", "DISTRICT", "OBLAST", "KRAI", "KRAY", "REPUBLIC"}

# Explicit country detection: explicit mention always wins
EXPLICIT_COUNTRY_HINTS: List[Tuple[re.Pattern, str, str]] = [
    (re.compile(r"\bUKRAINE\b|\bUKRAINIAN\b", re.I), "UA", "Ukraine"),
    (re.compile(r"\bREPUBLIC OF AZERBAIJAN\b|\bAZERBAIJAN\b", re.I), "AZ", "Azerbaijan"),
    (re.compile(r"\bBELARUS\b", re.I), "BY", "Belarus"),
    (re.compile(r"\bTAJIKISTAN\b", re.I), "TJ", "Tajikistan"),
    (re.compile(r"\bUZBEKISTAN\b", re.I), "UZ", "Uzbekistan"),
    (re.compile(r"\bKAZAKHSTAN\b", re.I), "KZ", "Kazakhstan"),
    (re.compile(r"\bRUSSIAN FEDERATION\b|\bRUSSIA\b|\bRSFSR\b", re.I), "RU", "Russia"),
    (re.compile(r"\bCOTE\s*[- ]?IVOIRE\b|\bCÔTE\s*D['’]IVOIRE\b", re.I), "CI", "Côte d'Ivoire"),
]

# Ambiguity guards: Odessa exists in multiple countries; if text indicates Ukraine, force UA.
AMBIG_CITY_OVERRIDES = {
    "odessa": (re.compile(r"\bODESSA\b.*\bUKRAINE\b|\bODESSA REGION\b.*\bUKRAINE\b", re.I), "UA", "Ukraine"),
}

# Country normalization (keep minimal; extend as needed)
COUNTRY_NAME_TO_ISO2 = {
    "ukraine": "UA",
    "azerbaijan": "AZ",
    "belarus": "BY",
    "tajikistan": "TJ",
    "uzbekistan": "UZ",
    "kazakhstan": "KZ",
    "russia": "RU",
    "côte d'ivoire": "CI",
    "cote d'ivoire": "CI",
}
ISO2_TO_COUNTRY = {
    "UA": "Ukraine",
    "AZ": "Azerbaijan",
    "BY": "Belarus",
    "TJ": "Tajikistan",
    "UZ": "Uzbekistan",
    "KZ": "Kazakhstan",
    "RU": "Russia",
    "CI": "Côte d'Ivoire",
}

# Postal validation (basic; country-aware)
POSTAL_RE = {
    "US": re.compile(r"^\d{5}(-\d{4})?$", re.I),
    "GB": re.compile(r"^[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}$", re.I),
    "CA": re.compile(r"^[A-Z]\d[A-Z]\s*\d[A-Z]\d$", re.I),
    "IN": re.compile(r"^\d{6}$", re.I),
    "RU": re.compile(r"^\d{6}$", re.I),
    "UA": re.compile(r"^\d{5}$", re.I),
    "AZ": re.compile(r"^[A-Z]{2}\d{4}$|^\d{4}$", re.I),
    "BY": re.compile(r"^\d{6}$", re.I),
    "TJ": re.compile(r"^\d{6}$", re.I),
    "UZ": re.compile(r"^\d{6}$", re.I),
    "KZ": re.compile(r"^\d{6}$", re.I),
}


def protect_saint(s: str) -> str:
    s = norm_ws(s)
    if not s:
        return s
    return SAINT_FIX_RE.sub(lambda m: "Saint " + m.group(1).title(), s)


def strip_admin_noise(s: str) -> str:
    s = norm_ws(s)
    if not s:
        return s
    return ADMIN_NOISE_RE.sub("", s).strip(" ,;")


def extract_locality_from_address1(addr1: str) -> Tuple[str, str]:
    """Extract locality and remainder from ADDRESS1 for patterns like:
       'CITY. IMISHLI, REPUBLIC OF AZERBAIJAN;'
       'VILL. CHORKUKH, ISFARA DISTRICT, ...;'
    """
    s = protect_saint(addr1)
    if not s:
        return "", ""

    s2 = PREFIX_LOCALITY_RE.sub("", s.strip())
    parts = re.split(r"[;,]\s*", s2, maxsplit=1)
    loc = parts[0].strip(" ,;")
    rem = parts[1].strip() if len(parts) > 1 else ""

    # avoid junk tokens becoming city
    if safe_upper(loc) in JUNK_CITY:
        loc = ""

    loc = strip_admin_noise(loc)
    rem = strip_admin_noise(rem)
    return loc, rem


def detect_country_from_text(text: str) -> Tuple[str, str]:
    """Return (iso2, country_name) if explicitly mentioned in text."""
    if not text:
        return "", ""
    for pat, iso2, name in EXPLICIT_COUNTRY_HINTS:
        if pat.search(text):
            return iso2, name
    return "", ""


def normalize_country(country_id: str, country: str) -> Tuple[str, str]:
    cid = norm_ws(country_id).upper()
    cname = strip_admin_noise(country)

    if cname and not cid:
        cid = COUNTRY_NAME_TO_ISO2.get(cname.lower(), "")

    if cid and not cname:
        cname = ISO2_TO_COUNTRY.get(cid, "")

    # last cleanup
    cname = strip_admin_noise(cname)
    return cid, cname


def validate_postal(postal: str, country_id: str) -> str:
    p = norm_ws(postal)
    if not p:
        return "MISSING"
    cid = norm_ws(country_id).upper()
    if cid in POSTAL_RE:
        return "VALID" if POSTAL_RE[cid].match(p) else "INVALID"
    return "UNKNOWN"


def data_quality(addr1: str, addr2: str, city: str, country: str) -> str:
    addr_present = bool(norm_ws(addr1) or norm_ws(addr2))
    city_present = bool(norm_ws(city))
    ctry_present = bool(norm_ws(country))
    score = sum([addr_present, city_present, ctry_present])
    if score == 3:
        return "COMPLETE"
    if score == 0:
        return "MISSING"
    return "PARTIAL"


def clean_row(row: Dict, colmap: Dict[str, str]) -> Tuple[Dict, List[Tuple[str, str, str]]]:
    """Clean a single row (dict). Returns updated row and list of field-level changes."""

    def get(col_key: str) -> str:
        col = colmap.get(col_key, col_key)
        return row.get(col, "")

    changes: List[Tuple[str, str, str]] = []

    # pull
    addr1 = protect_saint(get("ADDRESS1"))
    addr2 = protect_saint(get("ADDRESS2"))
    city = protect_saint(get("CITY"))
    state = norm_ws(get("STATE"))
    country = strip_admin_noise(get("COUNTRY"))
    country_id = norm_ws(get("COUNTRY_ID")).upper()
    postal = norm_ws(get("POSTAL"))

    # 1) extract city from address1 if city missing
    loc, rem = extract_locality_from_address1(addr1)
    if not city and loc:
        city = loc
        if rem:
            addr1 = rem

    # 2) admin noise cleanup
    city = strip_admin_noise(city)
    addr1 = strip_admin_noise(addr1)

    # 3) country resolution
    combined = " ".join([addr1, addr2, city, state, country, country_id]).strip()

    # Ambiguity override (Odessa)
    if city.lower() in AMBIG_CITY_OVERRIDES:
        pat, iso2o, nameo = AMBIG_CITY_OVERRIDES[city.lower()]
        if pat.search(combined):
            country_id, country = iso2o, nameo

    # Explicit country wins
    if not country_id or not country:
        iso2, name = detect_country_from_text(combined)
        if iso2:
            country_id, country = iso2, name

    # Normalize
    country_id, country = normalize_country(country_id, country)

    # 4) postal flag & data quality
    postal_flag = validate_postal(postal, country_id)
    dq = data_quality(addr1, addr2, city, country)

    # write & track changes
    def commit(key, value):
        col = colmap.get(key, key)
        before = "" if row.get(col, "") is None else str(row.get(col, ""))
        after = "" if value is None else str(value)
        if before != after:
            changes.append((col, before, after))
        row[col] = value

    commit("ADDRESS1", addr1)
    commit("ADDRESS2", addr2)
    commit("CITY", city)
    commit("STATE", state)
    commit("COUNTRY_ID", country_id)
    commit("COUNTRY", country)
    commit("POSTAL", postal)

    # Add computed fields if missing
    if "POSTAL_FLAG" not in row:
        row["POSTAL_FLAG"] = ""
    if "DATA_QUALITY" not in row:
        row["DATA_QUALITY"] = ""

    commit("POSTAL_FLAG", postal_flag)
    commit("DATA_QUALITY", dq)

    return row, changes


def build_default_colmap(df: pd.DataFrame) -> Dict[str, str]:
    """Try to map expected keys to actual columns."""
    cols = list(df.columns)
    u = {c.upper(): c for c in cols}

    def pick(*names):
        for n in names:
            if n in u:
                return u[n]
        return ""

    mapping = {
        "ADDRESS1": pick("ADDRESS1", "ADDRESS_1", "ADDR1"),
        "ADDRESS2": pick("ADDRESS2", "ADDRESS_2", "ADDR2"),
        "CITY": pick("CITY", "TOWN"),
        "STATE": pick("STATE", "PROVINCE", "REGION"),
        "POSTAL": pick("POSTAL", "ZIP", "POSTCODE"),
        "COUNTRY": pick("COUNTRY"),
        "COUNTRY_ID": pick("COUNTRY_ID", "COUNTRYCODE", "COUNTRY_CODE"),
        "POSTAL_FLAG": "POSTAL_FLAG",
        "DATA_QUALITY": "DATA_QUALITY",
    }

    # Fill missing mappings with key itself (so code works if columns exist)
    for k, v in list(mapping.items()):
        if not v and k not in {"POSTAL_FLAG", "DATA_QUALITY"}:
            mapping[k] = k

    return mapping


def ensure_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    if "POSTAL_FLAG" not in df.columns:
        df["POSTAL_FLAG"] = ""
    if "DATA_QUALITY" not in df.columns:
        df["DATA_QUALITY"] = ""
    return df


def clean_dataframe(df: pd.DataFrame, colmap: Dict[str, str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df2 = df.copy()
    df2 = ensure_output_columns(df2)

    change_rows = []

    for idx, r in df2.iterrows():
        row = r.to_dict()
        new_row, changes = clean_row(row, colmap)
        df2.loc[idx] = pd.Series(new_row)
        for field, before, after in changes:
            if before != after:
                change_rows.append({"Row": idx + 1, "Field": field, "Before": before, "After": after})

    changes_df = pd.DataFrame(change_rows)
    return df2, changes_df


def to_excel_bytes(cleaned: pd.DataFrame, changes: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Data", index=False)
        changes.to_excel(writer, sheet_name="Changes Log", index=False)
    return buf.getvalue()


# -------------------------------
# UI
# -------------------------------

st.title("🗺️ Address Cleaning Agent (Enhanced)")
st.caption("Offline, deterministic address cleanup with improved handling for CITY./VILL. style inputs.")

with st.sidebar:
    st.header("1) Upload")
    up = st.file_uploader("Drop Excel or CSV", type=["xlsx", "xls", "csv"])

    st.header("2) Options")
    st.checkbox("Extract locality from ADDRESS1 (CITY./VILL./SETTLEMENT)", value=True)
    st.checkbox("Explicit country mention overrides ambiguous inference", value=True)
    st.caption("(These options are always enabled in the enhanced rules; toggles are for future extensions.)")

    run = st.button("▶ Clean Addresses", type="primary", use_container_width=True)

if not up:
    st.info("Upload add1.xlsx (or any Excel/CSV) to begin.")
    st.stop()

# Load file
if up.name.lower().endswith(".csv"):
    df_in = pd.read_csv(up)
else:
    df_in = pd.read_excel(up, engine="openpyxl")

st.subheader("Raw preview")
st.dataframe(df_in.head(30), use_container_width=True)

colmap = build_default_colmap(df_in)

with st.expander("Column mapping (auto-detected)", expanded=False):
    st.write("If your headers differ, adjust them here.")
    cols = [""] + list(df_in.columns)
    for k in ["ADDRESS1", "ADDRESS2", "CITY", "STATE", "POSTAL", "COUNTRY", "COUNTRY_ID"]:
        current = colmap.get(k, "")
        idx = cols.index(current) if current in cols else 0
        colmap[k] = st.selectbox(k, options=cols, index=idx)

if not run:
    st.stop()

st.subheader("Cleaning output")
with st.spinner("Cleaning..."):
    cleaned_df, changes_df = clean_dataframe(df_in, colmap)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    st.metric("Rows", len(cleaned_df))
with c2:
    st.metric("Changed fields", len(changes_df))
with c3:
    st.metric("Complete", int((cleaned_df.get("DATA_QUALITY", "") == "COMPLETE").sum()) if "DATA_QUALITY" in cleaned_df.columns else 0)

st.markdown("#### Cleaned preview")
st.dataframe(cleaned_df.head(50), use_container_width=True)

st.markdown("#### Changes log (preview)")
st.dataframe(changes_df.head(200), use_container_width=True)

excel_bytes = to_excel_bytes(cleaned_df, changes_df)

st.download_button(
    "⬇ Download cleaned Excel",
    data=excel_bytes,
    file_name="cleaned_addresses.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
