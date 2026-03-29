import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Address Cleaning Agent", page_icon="🗺️", layout="wide")

# ── Cleaning Engine ────────────────────────────────────────────────
ABBREV = [
    (r"\bul\.\s*", "Ulitsa "), (r"\bulitsa\b", "Ulitsa"),
    (r"\bd\.(?=\s*[\d/])", "Dom "), (r"\bdom\b(?=\s)", "Dom"),
    (r"\bpom\.\s*", "Pomeshchenie "), (r"\bpomeshchenie\b", "Pomeshchenie"),
    (r"\bkorp\.\s*", "Korpus "), (r"\bstr\.\s*(?=\d)", "Stroenie "),
    (r"\bet\.\s*", "Etazh "), (r"\bof\.\s*", "Office "),
    (r"\bkab\.\s*", "Kabinet "), (r"\blit\.\s*", "Liter "),
    (r"\bkv\.\s*", "Kvartira "), (r"\bkom\.\s*", "Komnata "),
    (r"\bk\.\s*(?=\d)", "Korpus "),
]

STATE_KW  = re.compile(r"\b(Oblast|Penzenskaya|Novosibirskaya|Volgogradskaya|Saratovskaya)\b", re.I)
ADDR3_KW  = re.compile(r"\b(Krai|Kray|Okrug|Republic|Autonomous|Rayon|Raion|Province|Region|Territory|County|Prefecture|Governorate)\b", re.I)
ANY_RGN   = re.compile(r"\b(Oblast|Krai|Kray|Okrug|Republic|Autonomous|Rayon|Raion|Province|Region|Territory|County|Prefecture|Governorate|Penzenskaya|Novosibirskaya|Volgogradskaya|Saratovskaya)\b", re.I)
KNOWN_STATES = {
    "shandong","guangdong","zhejiang","jiangsu","sichuan","hubei","hunan","henan",
    "hebei","fujian","liaoning","yunnan","guangxi","colorado","california","texas",
    "florida","illinois","ohio","georgia","michigan","virginia","washington","arizona",
    "massachusetts","tennessee","indiana","missouri","maryland","wisconsin","minnesota",
    "alabama","louisiana","kentucky","oregon","oklahoma","utah","iowa","nevada",
    "arkansas","mississippi","kansas","nebraska","idaho","connecticut","delaware",
    "south carolina","north carolina","pennsylvania","west virginia",
    "hessen","bayern","hamburg","sachsen","maharashtra","karnataka","gujarat",
    "uttar pradesh","west bengal","andhra pradesh","telangana","kerala","punjab",
    "haryana","rajasthan","khanh hoa","tortola","crimea","corsica","sardinia","sicily",
}
CAPITAL_CITIES = {
    "moscow","athens","belgrade","beijing","shanghai","london","paris","dubai",
    "singapore","hong kong","mumbai","delhi","istanbul","tokyo","seoul","bangkok",
    "shenzhen","minsk","kyiv","kiev","pskov","yaroslavl","tver","vladimir","ryazan",
    "ivanovo","kazan","ufa","perm","samara","voronezh","saratov","omsk","novosibirsk",
    "krasnoyarsk","irkutsk","vladivostok","khabarovsk","yekaterinburg",
}
FEDERAL_CITIES  = {"moscow"}
DISTRICT_RE     = re.compile(r"\bDistrict\b", re.I)
SETTLEMENT_PFX  = re.compile(r"^(Rabochiy\s+Poselok|Poselok\s+Gorodskogo|Poselok\s+Kresty|Selo\s+|Derevnya\s+|Stanitsya\s+)", re.I)
POSTAL_FORMATS  = {
    "russia": r"^\d{6}$", "ukraine": r"^\d{5,6}$", "belarus": r"^\d{6}$",
    "kazakhstan": r"^\d{6}$", "germany": r"^\d{5}$", "france": r"^\d{5}$",
    "united kingdom": r"^[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}$",
    "china": r"^\d{6}$", "india": r"^\d{6}$",
    "united states": r"^\d{5}(-\d{4})?$",
    "bermuda": r"^[A-Z]{2}\s?\d{2}$", "luxembourg": r"^(L-?)?\d{4}$",
    "switzerland": r"^(CH-?)?\d{4}$", "moldova": r"^(MD-?)?\d{4}$",
    "azerbaijan": r"^(AZ-?)?\d{4}$", "malta": r"^[A-Z]{3}\s?\d{4}$",
    "isle of man": r"^IM\d\s?\d[A-Z]{2}$", "serbia": r"^\d{5}$",
    "cyprus": r"^\d{4}$", "kyrgyzstan": r"^\d{6}$", "armenia": r"^\d{4}$",
    "austria": r"^\d{4}$", "greece": r"^\d{5}$", "romania": r"^\d{6}$",
    "vietnam": r"^\d{6}$",
}

def clean_addr(raw):
    if not raw or str(raw).strip() in ("", "nan"): return ""
    s = str(raw).strip()
    for pat, rep in ABBREV: s = re.sub(pat, rep, s, flags=re.IGNORECASE)
    s = re.sub(r" +", " ", s).strip()
    return s[0].upper() + s[1:] if s else s

def is_region(p, city_cand=""):
    if p.lower() == "new york":
        return bool(re.match(r"^New York City$", city_cand.strip(), re.I))
    return bool(STATE_KW.search(p) or ADDR3_KW.search(p) or p.lower() in KNOWN_STATES)

def classify_region(p):
    if STATE_KW.search(p) or p.lower() in KNOWN_STATES or p.lower() == "new york": return "state"
    return "addr3"

def parse_city(raw):
    if not raw or str(raw).strip() in ("", "nan"): return "", "", ""
    s = re.sub(r"^[Gg]orod\s+", "", str(raw).strip(), flags=re.I)
    s = re.sub(r"^[Gg]\.\s*", "", s)
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) == 1:
        p = parts[0]
        if STATE_KW.search(p): return "", p, ""
        if ADDR3_KW.search(p): return "", "", p
        if p.lower() in KNOWN_STATES: return "", p, ""
        return p, "", ""
    city_cand = parts[0]
    for p in reversed(parts):
        if p.lower() in CAPITAL_CITIES: city_cand = p; break
    region_idx = {i for i, p in enumerate(parts) if is_region(p, city_cand)}
    other_idx  = [i for i in range(len(parts)) if i not in region_idx]
    other      = [parts[i] for i in other_idx]
    if not other: return parts[0], "", ""
    city = ""; city_oi = -1
    for i in range(len(other)-1, -1, -1):
        if other[i].lower() in CAPITAL_CITIES:
            city = other[i]; city_oi = other_idx[i]; break
    if not city:
        if SETTLEMENT_PFX.match(other[0]) and len(other) > 1:
            city = other[-1]; city_oi = other_idx[-1]
        else:
            city = other[0]; city_oi = other_idx[0]
    extra_state, extra_addr3 = [], []
    for i, p in enumerate(parts):
        if i in region_idx or i == city_oi: continue
        if p.lower() == city.lower():
            (extra_state if city.lower() in FEDERAL_CITIES else extra_addr3).append(p)
        elif DISTRICT_RE.search(p) and i > city_oi: extra_state.append(p)
        else: extra_addr3.append(p)
    st_parts = [parts[i] for i in sorted(region_idx) if classify_region(parts[i]) == "state"] + extra_state
    a3_parts = [parts[i] for i in sorted(region_idx) if classify_region(parts[i]) == "addr3"] + extra_addr3
    def ordered(lst):
        seen_v, seen_i, out = set(), set(), []
        for i, p in enumerate(parts):
            if p in lst and i not in seen_i and p not in seen_v:
                out.append(p); seen_i.add(i); seen_v.add(p)
        return out
    sv = ", ".join(ordered(st_parts))
    av = ", ".join(ordered(a3_parts))
    if sv.strip().lower() == city.lower(): sv = ""
    return city, sv, av

def validate_postal(postal, country):
    p = str(postal or "").strip()
    if not p or p == "nan": return "MISSING"
    fmt = POSTAL_FORMATS.get(str(country or "").lower().strip())
    if not fmt: return "VALID"
    return "VALID" if re.match(fmt, p, re.I) else "INVALID"

def quality(row, cols):
    has = sum([
        bool(str(row.get(cols.get("address",""), "")).strip()),
        bool(str(row.get(cols.get("city",""), "")).strip()),
        bool(str(row.get(cols.get("postal",""), "")).strip() not in ("","nan")),
        bool(str(row.get(cols.get("country",""), "")).strip()),
    ])
    return "COMPLETE" if has == 4 else "PARTIAL" if has >= 2 else "MISSING"

HINTS = {
    "address": ["address1","address","addr1","street","line1"],
    "city":    ["city","town","locality"],
    "state":   ["state","province","region","oblast"],
    "postal":  ["postal","zip","postcode","postalcode"],
    "country": ["country","countryname","nation"],
}
def guess_col(headers, field):
    for h in headers:
        if any(h.lower().replace(" ","").replace("_","") == hint for hint in HINTS[field]):
            return h
    return ""

# ── UI ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
.stApp { max-width: 1400px; }
.metric-card { background: #f8f8f6; border: 1px solid #e0dfd8; border-radius: 8px; padding: 12px 16px; text-align: center; }
.metric-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: .05em; margin-bottom: 4px; }
.metric-value { font-size: 28px; font-weight: 300; color: #1a1a18; }
.tag-complete { background: #e8f5ee; color: #1a7a4a; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.tag-partial  { background: #fef9e8; color: #8a5f00; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.tag-missing  { background: #fdecea; color: #c0392b; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.tag-valid    { background: #e8f5ee; color: #1a7a4a; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.tag-invalid  { background: #fdecea; color: #c0392b; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.tag-missing2 { background: #fef9e8; color: #8a5f00; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.title("🗺️ Address Cleaning Agent")
st.caption("Upload any Excel or CSV with address data — cleans city, state, postal and address fields automatically.")

# ── Sidebar ─────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx","xls","csv"])

    col_map = {}
    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded)
            else:
                df_raw = pd.read_excel(uploaded)
            df_raw = df_raw.fillna("").astype(str).replace("nan","")
            headers = df_raw.columns.tolist()
            st.success(f"✓ {len(df_raw)} rows · {len(headers)} columns")
            st.divider()
            st.subheader("Column Mapping")
            for field in ["address","city","state","postal","country"]:
                default = guess_col(headers, field)
                idx = headers.index(default) + 1 if default in headers else 0
                col_map[field] = st.selectbox(
                    field.title(), ["— skip —"] + headers,
                    index=idx, key=f"col_{field}"
                )
                col_map[field] = "" if col_map[field] == "— skip —" else col_map[field]
        except Exception as e:
            st.error(f"Error reading file: {e}")
            uploaded = None

    st.divider()
    st.subheader("Cleaning Tasks")
    o_addr   = st.checkbox("Expand address abbreviations", value=True, help="ul.→Ulitsa, d.→Dom, pom.→Pomeshchenie…")
    o_city   = st.checkbox("Parse city field", value=True, help="Splits City + STATE (Oblast) + ADDRESS3 (Krai/Okrug)")
    o_postal = st.checkbox("Validate postal code", value=True, help="Country-aware validation")
    o_quality= st.checkbox("Data quality score", value=True, help="COMPLETE / PARTIAL / MISSING per row")

    run = st.button("▶  Run Cleaning Agent", type="primary", use_container_width=True, disabled=uploaded is None)

# ── Main ────────────────────────────────────────────────────────────
if not uploaded:
    st.info("👈 Upload a file in the sidebar to get started")
    st.stop()

if not run and "cleaned_df" not in st.session_state:
    st.info("Configure column mapping in the sidebar, then click **Run Cleaning Agent**")

    st.subheader("Preview")
    st.dataframe(df_raw.head(20), use_container_width=True)
    st.stop()

if run:
    with st.spinner("Cleaning addresses..."):
        df = df_raw.copy()
        diffs = []

        # Clean ADDRESS1
        if o_addr and col_map.get("address"):
            col = col_map["address"]
            orig = df[col].copy()
            df[col] = df[col].apply(clean_addr)
            changed = df[col] != orig
            for i in df[changed].index:
                diffs.append({"Row": i+1, "Field": col, "Before": orig[i], "After": df.at[i,col]})

        # Parse CITY
        if o_city and col_map.get("city"):
            col = col_map["city"]
            orig = df[col].copy()
            for idx, row in df.iterrows():
                city, state, addr3 = parse_city(row[col])
                df.at[idx, col] = city
                if state:
                    sc = col_map.get("state") or "STATE"
                    if sc not in df.columns: df[sc] = ""
                    df.at[idx, sc] = state
                if addr3:
                    if "ADDRESS3" not in df.columns: df["ADDRESS3"] = ""
                    existing = str(df.at[idx,"ADDRESS3"]).strip()
                    df.at[idx,"ADDRESS3"] = (existing + ", " + addr3).strip(", ") if existing else addr3
                if city != orig[idx] and orig[idx]:
                    after_lbl = city
                    if state: after_lbl += f" | state: {state}"
                    if addr3: after_lbl += f" | addr3: {addr3}"
                    diffs.append({"Row": idx+1, "Field": col, "Before": orig[idx], "After": after_lbl})

        # Postal flag
        if o_postal and col_map.get("postal"):
            df["POSTAL_FLAG"] = df.apply(
                lambda r: validate_postal(r[col_map["postal"]], r.get(col_map.get("country",""),"") if col_map.get("country") else ""),
                axis=1
            )

        # Quality
        if o_quality:
            df["DATA_QUALITY"] = df.apply(lambda r: quality(r.to_dict(), col_map), axis=1)

        st.session_state["cleaned_df"] = df
        st.session_state["diffs"] = diffs
        st.session_state["col_map"] = col_map

# Show results
if "cleaned_df" in st.session_state:
    df      = st.session_state["cleaned_df"]
    diffs   = st.session_state["diffs"]
    col_map = st.session_state["col_map"]

    n = len(df)
    complete = (df.get("DATA_QUALITY","") == "COMPLETE").sum() if "DATA_QUALITY" in df else 0
    partial  = (df.get("DATA_QUALITY","") == "PARTIAL").sum()  if "DATA_QUALITY" in df else 0
    missing  = (df.get("DATA_QUALITY","") == "MISSING").sum()  if "DATA_QUALITY" in df else 0
    invalid  = (df.get("POSTAL_FLAG","") == "INVALID").sum()   if "POSTAL_FLAG"  in df else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Rows cleaned", n)
    c2.metric("Changes made", len(diffs))
    c3.metric("Complete", complete, f"{round(complete/n*100)}%" if n else "")
    c4.metric("Partial",  partial,  f"{round(partial/n*100)}%"  if n else "")
    c5.metric("Invalid postal", invalid)

    tab1, tab2, tab3 = st.tabs(["📋 Cleaned Data", "🔍 Changes", "⬇️ Download"])

    with tab1:
        st.dataframe(df, use_container_width=True, height=500)

    with tab2:
        if diffs:
            st.dataframe(pd.DataFrame(diffs), use_container_width=True, height=400)
        else:
            st.info("No changes were made")

    with tab3:
        st.subheader("Download cleaned file")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Cleaned Data", index=False)
            if diffs:
                pd.DataFrame(diffs).to_excel(writer, sheet_name="Changes Log", index=False)
        st.download_button(
            "⬇️  Download Cleaned Excel",
            data=buf.getvalue(),
            file_name=f"cleaned_addresses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
