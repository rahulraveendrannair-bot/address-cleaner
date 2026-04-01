import streamlit as st
import pandas as pd
import re
import io
import time
import requests
from functools import lru_cache

st.set_page_config(page_title="Address Cleaning Agent", page_icon="🗺️", layout="wide")

# ── Geocoding via Nominatim ──────────────────────────────────────────────────
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
HEADERS = {"User-Agent": "address-cleaner-app/2.0 (contact@wisetech.com)"}

@st.cache_data(show_spinner=False, ttl=3600)
def geocode_place(name: str) -> dict:
    """
    Query Nominatim for a place name.
    Returns dict with: city, state, country, postcode, place_type, country_code
    """
    if not name or len(name.strip()) < 2:
        return {}
    try:
        resp = requests.get(
            NOMINATIM_URL,
            params={"q": name, "format": "json", "addressdetails": 1, "limit": 1},
            headers=HEADERS,
            timeout=5
        )
        time.sleep(0.2)  # respect Nominatim rate limit
        data = resp.json()
        if not data:
            return {}
        result = data[0]
        addr   = result.get("address", {})
        return {
            "city":         addr.get("city") or addr.get("town") or addr.get("village") or addr.get("municipality") or "",
            "suburb":       addr.get("suburb") or addr.get("neighbourhood") or addr.get("quarter") or "",
            "district":     addr.get("county") or addr.get("district") or "",
            "state":        addr.get("state") or addr.get("region") or "",
            "country":      addr.get("country") or "",
            "country_code": addr.get("country_code", "").upper(),
            "postcode":     addr.get("postcode") or "",
            "place_type":   result.get("type", ""),
            "place_class":  result.get("class", ""),
            "display":      result.get("display_name", ""),
        }
    except Exception:
        return {}

def classify_place(name: str, geo: dict) -> str:
    """
    Given a place name and its geocode result, return which column it belongs in:
    CITY, STATE, COUNTRY, POSTAL, ADDRESS3, or KEEP (leave where it is)
    """
    if not geo:
        return "UNKNOWN"

    ptype  = geo.get("place_type", "")
    pclass = geo.get("place_class", "")
    city   = geo.get("city", "").lower()
    state  = geo.get("state", "").lower()
    country= geo.get("country", "").lower()
    name_l = name.strip().lower()

    # Exact match to known types
    if pclass == "boundary" and ptype == "administrative":
        # Could be city, state, or country depending on admin level
        display = geo.get("display", "").lower()
        if country and name_l == country: return "COUNTRY"
        if state   and name_l == state:   return "STATE"
        if city    and name_l == city:    return "CITY"

    if ptype in ("city", "town", "village", "hamlet", "municipality"):
        return "CITY"
    if ptype in ("suburb", "neighbourhood", "quarter", "borough"):
        return "ADDRESS3"
    if ptype in ("state", "province", "region", "county"):
        return "STATE"
    if ptype == "country":
        return "COUNTRY"
    if ptype == "postcode":
        return "POSTAL"
    if ptype in ("district", "governorate"):
        return "ADDRESS3"

    # Fallback: check if name matches known geo fields
    if city   and name_l == city:    return "CITY"
    if state  and name_l == state:   return "STATE"
    if country and name_l == country: return "COUNTRY"

    return "UNKNOWN"

# ── Cleaning rules (pattern-based fallback) ──────────────────────────────────
ABBREV = [
    (r"\bul\.\s*","Ulitsa "),(r"\bulitsa\b","Ulitsa"),
    (r"\bd\.(?=\s*[\d/])","Dom "),(r"\bdom\b(?=\s)","Dom"),
    (r"\bpom\.\s*","Pomeshchenie "),(r"\bpomeshchenie\b","Pomeshchenie"),
    (r"\bkorp\.\s*","Korpus "),(r"\bstr\.\s*(?=\d)","Stroenie "),
    (r"\bet\.\s*","Etazh "),(r"\bof\.\s*","Office "),
    (r"\bkab\.\s*","Kabinet "),(r"\blit\.\s*","Liter "),
    (r"\bkv\.\s*","Kvartira "),(r"\bkom\.\s*","Komnata "),
    (r"\bk\.\s*(?=\d)","Korpus "),
]
COUNTRY_RE = re.compile(r"\b(Russia|China|United Kingdom|England|Scotland|Northern Ireland|UK|United States|USA|Germany|France|Ukraine|India|Iran|Iraq|Syria|Turkey|Türkiye|Japan|South Korea|North Korea|Myanmar|Malaysia|Indonesia|Pakistan|Egypt|Switzerland|Netherlands|Belgium|Italy|Spain|Portugal|Greece|Poland|Romania|Bulgaria|Hungary|Czech Republic|Austria|Sweden|Norway|Denmark|Finland|Ireland|Luxembourg|Cyprus|Serbia|Croatia|Belarus|Kazakhstan|Azerbaijan|Armenia|Georgia|Moldova|Singapore|Hong Kong|Taiwan|Philippines|Bangladesh|Afghanistan|Uzbekistan|Vietnam|Thailand|Cambodia|Nigeria|Kenya|South Africa|Ethiopia|Sudan|Yemen|Qatar|Kuwait|Oman|Bahrain|Jordan|Lebanon|Israel|Saudi Arabia|United Arab Emirates|Canada|Mexico|Brazil|Argentina|Venezuela|Colombia|Peru|Chile|Bolivia|Australia|New Zealand|Libya|Tunisia|Morocco|Algeria|Palestine|Palestinian|Seychelles|Latvia|Monaco|Marshall Islands|Laos|Kyrgyzstan|Tajikistan|Turkmenistan|Myanmar|Macau)\b", re.I)
STATE_KW   = re.compile(r"\b(Oblast|Penzenskaya|Novosibirskaya|Volgogradskaya|Saratovskaya)\b", re.I)
ADDR3_KW   = re.compile(r"\b(Krai|Kray|Okrug|Republic|Autonomous|Rayon|Raion|Province|Region|Territory|County|Prefecture|Governorate)\b", re.I)
ADDR_RE    = re.compile(r"\b(Street|St\.|Avenue|Ave|Road|Rd|Lane|Drive|Blvd|Floor|Suite|Block|No\.|Ulitsa|Dom|Korpus|Etazh|Strasse|Prospekt|Chemin|Rue|Boulevard|Weg|Gasse|Platz|Jalan|Building|Bldg|Unit|Room|Apt|Plaza|Tower|Mahallesi|Cad\.|Sok\.)\b", re.I)
JUNK_RE    = re.compile(r"^(XX|00000|0{4,}|N/A|NA|None|null|TBD)$", re.I)
NOTE_RE    = re.compile(r"^(Linked To:|Address redacted|Unknown|Located in|Resident in|Trust Company|Letter )", re.I)
ISO2_RE    = re.compile(r"^[A-Z]{2}$")
TERRITORY_RE = re.compile(r",?\s*\b(West Bank|Northern Gaza|Gaza Strip|Gaza|Crimea)\b\s*,?", re.I)

US_STATES  = {"AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"}
CA_PROVS   = {"AB","BC","MB","NB","NL","NS","NT","NU","ON","PE","QC","SK","YT"}
US_STATE_EXPAND = {"AL":"Alabama","AK":"Alaska","AZ":"Arizona","AR":"Arkansas","CA":"California","CO":"Colorado","CT":"Connecticut","DE":"Delaware","FL":"Florida","GA":"Georgia","HI":"Hawaii","ID":"Idaho","IL":"Illinois","IN":"Indiana","IA":"Iowa","KS":"Kansas","KY":"Kentucky","LA":"Louisiana","ME":"Maine","MD":"Maryland","MA":"Massachusetts","MI":"Michigan","MN":"Minnesota","MS":"Mississippi","MO":"Missouri","MT":"Montana","NE":"Nebraska","NV":"Nevada","NH":"New Hampshire","NJ":"New Jersey","NM":"New Mexico","NY":"New York","NC":"North Carolina","ND":"North Dakota","OH":"Ohio","OK":"Oklahoma","OR":"Oregon","PA":"Pennsylvania","RI":"Rhode Island","SC":"South Carolina","SD":"South Dakota","TN":"Tennessee","TX":"Texas","UT":"Utah","VT":"Vermont","VA":"Virginia","WA":"Washington","WV":"West Virginia","WI":"Wisconsin","WY":"Wyoming","DC":"District of Columbia"}
CA_PROV_EXPAND = {"AB":"Alberta","BC":"British Columbia","MB":"Manitoba","NB":"New Brunswick","NL":"Newfoundland and Labrador","NS":"Nova Scotia","NT":"Northwest Territories","NU":"Nunavut","ON":"Ontario","PE":"Prince Edward Island","QC":"Quebec","SK":"Saskatchewan","YT":"Yukon"}
COUNTRY_NORMALIZE = {"myanmar [burma]":"Myanmar","burma":"Myanmar","republic of türkiye":"Turkey","republic of turkey":"Turkey","türkiye":"Turkey","england":"United Kingdom","scotland":"United Kingdom","northern ireland":"United Kingdom","uk":"United Kingdom","great britain":"United Kingdom","viet nam":"Vietnam","south korea":"South Korea","republic of korea":"South Korea","north korea":"North Korea","democratic people's republic of korea":"North Korea","dprk":"North Korea","usa":"United States","u.s.a.":"United States","uae":"United Arab Emirates","irn":"Iran","iran, islamic republic of":"Iran","west bank":"Palestinian Territory","northern gaza":"Palestinian Territory","gaza":"Palestinian Territory","palestine":"Palestinian Territory","notk":""}
COUNTRY_TO_CODE = {"russia":"RU","china":"CN","united states":"US","united kingdom":"GB","germany":"DE","france":"FR","ukraine":"UA","india":"IN","iran":"IR","syria":"SY","iraq":"IQ","north korea":"KP","south korea":"KR","japan":"JP","turkey":"TR","saudi arabia":"SA","united arab emirates":"AE","pakistan":"PK","myanmar":"MM","thailand":"TH","vietnam":"VN","malaysia":"MY","indonesia":"ID","singapore":"SG","hong kong":"HK","taiwan":"TW","belarus":"BY","kazakhstan":"KZ","switzerland":"CH","austria":"AT","netherlands":"NL","sweden":"SE","norway":"NO","denmark":"DK","finland":"FI","ireland":"IE","portugal":"PT","spain":"ES","italy":"IT","greece":"GR","israel":"IL","egypt":"EG","libya":"LY","tunisia":"TN","morocco":"MA","algeria":"DZ","canada":"CA","australia":"AU","new zealand":"NZ","brazil":"BR","argentina":"AR","venezuela":"VE","colombia":"CO","peru":"PE","seychelles":"SC","latvia":"LV","monaco":"MC","laos":"LA","marshall islands":"MH","kyrgyzstan":"KG","tajikistan":"TJ","turkmenistan":"TM","uzbekistan":"UZ","cambodia":"KH","nigeria":"NG","kenya":"KE","south africa":"ZA","ethiopia":"ET","afghanistan":"AF","bangladesh":"BD","philippines":"PH","azerbaijan":"AZ","armenia":"AM","georgia":"GE","moldova":"MD","serbia":"RS","croatia":"HR","czech republic":"CZ","hungary":"HU","romania":"RO","bulgaria":"BG","poland":"PL","luxembourg":"LU","cyprus":"CY","belgium":"BE","yemen":"YE","qatar":"QA","kuwait":"KW","oman":"OM","bahrain":"BH","jordan":"JO","lebanon":"LB","palestinian territory":"PS","mexico":"MX","chile":"CL","bolivia":"BO","macau":"MO","estonia":"EE","lithuania":"LT","albania":"AL","montenegro":"ME","slovenia":"SI","slovakia":"SK","malta":"MT","iceland":"IS"}
ISO_TO_COUNTRY = {v.upper():k.title() for k,v in COUNTRY_TO_CODE.items()}
TERRITORY_TO_COUNTRY = {"west bank":"Palestinian Territory","northern gaza":"Palestinian Territory","gaza strip":"Palestinian Territory","gaza":"Palestinian Territory","crimea":"Russia"}

def clean_addr(raw):
    if not raw or raw.strip() == "": return ""
    s = raw.strip()
    for pat, rep in ABBREV: s = re.sub(pat, rep, s, flags=re.IGNORECASE)
    s = re.sub(r" +", " ", s).strip()
    return s[0].upper() + s[1:] if s else s

def normalize_country(c):
    c = c.strip()
    if c.isupper() and len(c) > 2: c = c.title()
    return COUNTRY_NORMALIZE.get(c.lower(), c.strip())

def parse_city_field(raw):
    """Split CITY field into CITY + STATE + ADDRESS3."""
    if not raw or raw.strip() == "": return "", "", ""
    s = re.sub(r"\s+Province\b", "", raw.strip(), flags=re.I)
    s = re.sub(r"\s+City\b(?=,|$)", "", s, flags=re.I)
    s = re.sub(r"^[Gg]orod\s+", "", s, flags=re.I)
    s = re.sub(r"^[Gg]\.\s*", "", s)
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) == 1:
        p = parts[0]
        if STATE_KW.search(p): return "", p, ""
        if ADDR3_KW.search(p): return "", "", p
        return p, "", ""
    # Find region parts
    KNOWN_STATES = {"shandong","guangdong","zhejiang","jiangsu","sichuan","hubei","hunan","henan","hebei","fujian","liaoning","yunnan","guangxi","anhui","jiangxi","shaanxi","xinjiang","beijing","shanghai","tianjin","chongqing","maharashtra","karnataka","gujarat","uttar pradesh","west bengal","andhra pradesh","telangana","kerala","punjab","haryana","rajasthan","colorado","california","texas","florida","illinois","ohio","georgia","michigan","virginia","washington","arizona","massachusetts","tennessee","indiana","missouri","maryland","wisconsin","minnesota","alabama","louisiana","kentucky","oregon","oklahoma","utah","iowa","nevada","arkansas","mississippi","kansas","nebraska","connecticut","west virginia","south carolina","north carolina","pennsylvania","new york","new jersey"}
    CAPITAL_CITIES = {"moscow","athens","belgrade","beijing","shanghai","london","paris","dubai","singapore","hong kong","mumbai","delhi","istanbul","tokyo","seoul","bangkok","shenzhen","minsk","kyiv","kiev","novosibirsk","yekaterinburg","kazan","ufa","perm","samara","voronezh","saratov","omsk","krasnoyarsk","irkutsk","vladivostok","khabarovsk","yangon","cairo","tehran","baghdad","wanchai","taipei","hanoi","jakarta","kuala lumpur","ramallah"}
    region_idx = {i for i,p in enumerate(parts) if STATE_KW.search(p) or ADDR3_KW.search(p) or p.lower() in KNOWN_STATES}
    other_idx  = [i for i in range(len(parts)) if i not in region_idx]
    other      = [parts[i] for i in other_idx]
    if not other: return parts[0], "", ""
    city = ""; city_oi = -1
    for i in range(len(other)-1,-1,-1):
        if other[i].lower() in CAPITAL_CITIES: city = other[i]; city_oi = other_idx[i]; break
    if not city: city = other[0]; city_oi = other_idx[0]
    extra_state, extra_addr3 = [], []
    for i,p in enumerate(parts):
        if i in region_idx or i == city_oi: continue
        extra_addr3.append(p)
    def classify(p):
        if STATE_KW.search(p) or p.lower() in KNOWN_STATES: return "state"
        return "addr3"
    st = [parts[i] for i in sorted(region_idx) if classify(parts[i])=="state"] + extra_state
    a3 = [parts[i] for i in sorted(region_idx) if classify(parts[i])=="addr3"] + extra_addr3
    def ordered(lst):
        sv=set();si=set();out=[]
        for i,p in enumerate(parts):
            if p in lst and i not in si and p not in sv: out.append(p);si.add(i);sv.add(p)
        return out
    sv = ", ".join(ordered(st)); av = ", ".join(ordered(a3))
    if sv.strip().lower() == city.lower(): sv = ""
    return city, sv, av

def parse_a1_with_geo(a1, ex_city="", ex_postal="", ex_country="", use_geo=True):
    """
    Parse ADDRESS1 using geocoding to correctly identify each comma-separated part.
    Falls back to pattern matching if geocoding unavailable.
    """
    if not a1 or NOTE_RE.match(a1):
        return a1, ex_city, ex_postal, ex_country

    city    = ex_city.strip()
    postal  = ex_postal.strip()
    country = ex_country.strip()

    parts = [p.strip() for p in a1.split(",") if p.strip()]

    # For each part that isn't clearly a street address, geocode it
    city_candidates    = []
    country_candidates = []
    postal_candidates  = []
    street_parts       = []

    for part in parts:
        part_clean = part.strip()

        # Pure postal code
        if re.match(r"^\d{4,6}$", part_clean):
            if not postal: postal = part_clean
            continue

        # Postal+city like "8600 Dübendorf"
        m = re.match(r"^(\d{4,5})\s+([A-Z\u00C0-\u024F]\S.+)$", part_clean)
        if m:
            if not postal: postal = m.group(1)
            if not city: city_candidates.append(m.group(2))
            continue

        # Postal at end: "Istanbul 34758"
        m = re.search(r"(?<!\d)(\d{5,6})\s*$", part_clean)
        if m:
            if not postal: postal = m.group(1)
            part_clean = part_clean[:m.start()].strip()
            if part_clean and not ADDR_RE.search(part_clean):
                city_candidates.append(part_clean)
            elif part_clean:
                street_parts.append(part_clean)
            continue

        # Known country
        if COUNTRY_RE.match(part_clean) and not ADDR_RE.search(part_clean):
            country_candidates.append(part_clean)
            continue

        # Clear street part
        if ADDR_RE.search(part_clean):
            street_parts.append(part_clean)
            continue

        # Ambiguous — geocode it
        if use_geo and len(part_clean) > 2 and not re.match(r"^\d+$", part_clean):
            geo = geocode_place(part_clean)
            classification = classify_place(part_clean, geo)
            if classification == "CITY":
                city_candidates.append(part_clean)
            elif classification == "STATE":
                # Will be handled by city field parser
                city_candidates.append(part_clean)
            elif classification == "COUNTRY":
                country_candidates.append(part_clean)
                if not country and geo.get("country"):
                    country = geo["country"]
                if not postal and geo.get("postcode"):
                    postal = geo["postcode"]
            elif classification in ("ADDRESS3",):
                city_candidates.append(part_clean)  # keep as part of address
            else:
                # Unknown — use pattern heuristic
                if not ADDR_RE.search(part_clean) and len(part_clean.split()) <= 3:
                    city_candidates.append(part_clean)
                else:
                    street_parts.append(part_clean)
        else:
            # No geocoding — pattern fallback
            if not ADDR_RE.search(part_clean) and not re.match(r"^\d+$", part_clean):
                if len(part_clean.split()) <= 4:
                    city_candidates.append(part_clean)
                else:
                    street_parts.append(part_clean)
            else:
                street_parts.append(part_clean)

    # Resolve country
    if not country and country_candidates:
        country = normalize_country(country_candidates[-1])
    elif not country:
        for cand in city_candidates:
            m = COUNTRY_RE.search(cand)
            if m: country = normalize_country(m.group(0)); break

    # Resolve city — last candidate that isn't a country
    if not city:
        for cand in reversed(city_candidates):
            if not COUNTRY_RE.fullmatch(cand):
                city = cand; break

    # Rebuild clean ADDRESS1 (only keep street parts)
    clean = ", ".join(street_parts)

    return clean, city, postal, country

def quality(row, col_map):
    has = sum([bool(str(row.get(col_map.get("address",""),"")).strip()),
               bool(str(row.get(col_map.get("city",""),"")).strip()),
               bool(str(row.get(col_map.get("postal",""),"") or "").strip() not in ("","nan")),
               bool(str(row.get(col_map.get("country",""),"")).strip())])
    return "COMPLETE" if has==4 else "PARTIAL" if has>=2 else "MISSING"

def validate_postal(postal, country):
    POSTAL_FORMATS = {"russia":r"^\d{6}$","ukraine":r"^\d{5,6}$","china":r"^\d{6}$","india":r"^\d{6}$","united states":r"^\d{5}(-\d{4})?$","united kingdom":r"^[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}$","germany":r"^\d{5}$","france":r"^\d{5}$","switzerland":r"^(CH-)?\d{4}$","austria":r"^\d{4}$","netherlands":r"^\d{4}\s?[A-Z]{2}$"}
    p = str(postal or "").strip()
    if not p or p == "nan": return "MISSING"
    fmt = POSTAL_FORMATS.get(str(country or "").lower().strip())
    if not fmt: return "VALID"
    return "VALID" if re.match(fmt, p, re.I) else "INVALID"

HINTS = {"address":["address1","address","addr1","street","line1"],"city":["city","town","locality"],"state":["state","province","region","oblast"],"postal":["postal","zip","postcode","postalcode"],"country":["country","countryname","nation"]}
def guess_col(headers, field):
    for h in headers:
        if any(h.lower().replace(" ","").replace("_","") == hint for hint in HINTS[field]): return h
    return ""

# ── UI ─────────────────────────────────────────────────────────────────────
st.markdown("""<style>
.stApp{max-width:1400px}
.tag-complete{background:#e8f5ee;color:#1a7a4a;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
.tag-partial{background:#fef9e8;color:#8a5f00;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
.tag-missing{background:#fdecea;color:#c0392b;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
</style>""", unsafe_allow_html=True)

st.title("🗺️ Address Cleaning Agent")
st.caption("Geocoding-enhanced — each place name is verified against real geographic data before routing to the correct column.")

with st.sidebar:
    st.header("⚙️ Settings")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx","xls","csv"])
    use_geocoding = st.toggle("🌍 Enable geocoding (slower, more accurate)", value=True,
        help="Uses OpenStreetMap Nominatim to verify place names. ~0.2s per unique value. Disable for large files first pass.")

    col_map = {}
    if uploaded:
        try:
            df_raw = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
            df_raw = df_raw.fillna("").astype(str).replace("nan","")
            headers = df_raw.columns.tolist()
            st.success(f"✓ {len(df_raw):,} rows · {len(headers)} columns")
            st.divider()
            st.subheader("Column Mapping")
            for field in ["address","city","state","postal","country"]:
                default = guess_col(headers, field)
                idx = headers.index(default)+1 if default in headers else 0
                col_map[field] = st.selectbox(field.title(), ["— skip —"]+headers, index=idx, key=f"col_{field}")
                col_map[field] = "" if col_map[field] == "— skip —" else col_map[field]
            # Also detect COUNTRY_ID column
            col_map["country_id"] = next((h for h in headers if h.upper() in ("COUNTRY_ID","COUNTRYID")), "")
        except Exception as e:
            st.error(f"Error: {e}"); uploaded = None

    st.divider()
    st.subheader("Cleaning Tasks")
    o_abbrev = st.checkbox("Expand address abbreviations", True)
    o_geo    = st.checkbox("Geocode & route place names", True)
    o_postal = st.checkbox("Validate postal codes", True)
    o_quality= st.checkbox("Data quality score", True)

    run = st.button("▶  Clean Addresses", type="primary", use_container_width=True, disabled=uploaded is None)

if not uploaded:
    st.info("👈 Upload a file in the sidebar to get started"); st.stop()

if not run and "cleaned_df" not in st.session_state:
    st.subheader("Preview")
    st.dataframe(df_raw.head(30), use_container_width=True); st.stop()

if run:
    df = df_raw.copy()
    diffs = []
    progress = st.progress(0, text="Starting...")
    total = len(df)

    # ── Pre-geocode all unique place candidates ────────────────────
    geo_cache = {}
    if o_geo and use_geocoding:
        candidates = set()
        addr_col = col_map.get("address","")
        city_col = col_map.get("city","")

        # Collect unique short comma parts from ADDRESS1 and CITY
        for _, row in df.iterrows():
            if addr_col:
                for part in str(row[addr_col]).split(","):
                    p = part.strip()
                    if p and not ADDR_RE.search(p) and not re.match(r"^\d+$", p) and 2 < len(p) <= 40:
                        candidates.add(p)
            if city_col:
                for part in str(row[city_col]).split(","):
                    p = part.strip()
                    if p and 2 < len(p) <= 40: candidates.add(p)

        st.info(f"🌍 Geocoding {len(candidates):,} unique place names... (this takes ~{len(candidates)//5} seconds)")
        geo_bar = st.progress(0)
        for i, cand in enumerate(candidates):
            if cand not in geo_cache:
                geo_cache[cand] = geocode_place(cand)
            geo_bar.progress((i+1)/len(candidates))
        geo_bar.empty()
        st.success(f"✓ Geocoded {len(geo_cache):,} place names")

    # ── Process each row ──────────────────────────────────────────
    for idx, row in df.iterrows():
        if idx % 500 == 0:
            progress.progress(idx/total, text=f"Processing row {idx:,} of {total:,}...")

        a1  = str(row.get(col_map.get("address",""),"") if col_map.get("address") else "").strip()
        a2  = str(row.get("ADDRESS2","") if "ADDRESS2" in df.columns else "").strip()
        a3  = str(row.get("ADDRESS3","") if "ADDRESS3" in df.columns else "").strip()
        cty = str(row.get(col_map.get("city",""),"") if col_map.get("city") else "").strip()
        st  = str(row.get(col_map.get("state",""),"") if col_map.get("state") else "").strip()
        pst = str(row.get(col_map.get("postal",""),"") if col_map.get("postal") else "").strip()
        cou = str(row.get(col_map.get("country",""),"") if col_map.get("country") else "").strip()
        cid = str(row.get(col_map.get("country_id",""),"") if col_map.get("country_id") else "").strip()

        def record(field, before, after):
            if before != after and after:
                diffs.append({"Row":idx+1,"Field":field,"Before":before,"After":after})

        # Normalize COUNTRY
        if cou:
            norm = normalize_country(cou)
            if norm != cou:
                record(col_map.get("country","COUNTRY"), cou, norm)
                if col_map.get("country"): df.at[idx,col_map["country"]] = norm; cou = norm

        # Fix STATE abbreviations
        sc = col_map.get("state","")
        if sc and ISO2_RE.match(st):
            if st in US_STATES:
                new_st = US_STATE_EXPAND[st]
                df.at[idx,sc] = new_st; st = new_st
                if not cou and sc: df.at[idx,col_map.get("country","")] = "United States"; cou = "United States"
            elif st in CA_PROVS:
                new_st = CA_PROV_EXPAND[st]
                df.at[idx,sc] = new_st; st = new_st
                if not cou: df.at[idx,col_map.get("country","")] = "Canada"; cou = "Canada"

        # Fix POSTAL junk
        pc = col_map.get("postal","")
        if pc and pst and JUNK_RE.match(pst):
            df.at[idx,pc] = ""; pst = ""

        # Fix ADDRESS2
        if a2:
            if not cty and not ADDR_RE.search(a2) and not re.match(r"^\d", a2):
                if col_map.get("city"): df.at[idx,col_map["city"]] = a2; cty = a2
                if "ADDRESS2" in df.columns: df.at[idx,"ADDRESS2"] = ""
            elif re.match(r"^\d{4,6}$", a2) and not pst:
                if pc: df.at[idx,pc] = a2; pst = a2
                if "ADDRESS2" in df.columns: df.at[idx,"ADDRESS2"] = ""

        # Fix ADDRESS3
        if a3 and "STATE_KW" and STATE_KW.search(a3) and not st:
            if sc: df.at[idx,sc] = a3; st = a3
            if "ADDRESS3" in df.columns: df.at[idx,"ADDRESS3"] = ""

        # Parse CITY field
        cc = col_map.get("city","")
        if cc and cty:
            city_val, state_val, addr3_val = parse_city_field(cty)
            if city_val != cty or state_val or addr3_val:
                df.at[idx,cc] = city_val; record(cc, cty, city_val); cty = city_val
                if state_val and sc and not st: df.at[idx,sc] = state_val; st = state_val
                if addr3_val and "ADDRESS3" in df.columns:
                    ex = str(df.at[idx,"ADDRESS3"]).strip()
                    df.at[idx,"ADDRESS3"] = (ex+", "+addr3_val).strip(", ") if ex else addr3_val

        # Parse ADDRESS1 with geocoding
        addr_col = col_map.get("address","")
        if addr_col and a1 and not NOTE_RE.match(a1):
            # Territory extraction
            t_match = TERRITORY_RE.search(a1)
            if t_match:
                territory = t_match.group(1).strip()
                if sc and not st: df.at[idx,sc] = territory; st = territory
                if not cou:
                    tc = TERRITORY_TO_COUNTRY.get(territory.lower(),"")
                    if tc: cou = tc; df.at[idx,col_map.get("country","")] = tc
                clean_a1 = TERRITORY_RE.sub("", a1).strip().strip(",").strip()
                df.at[idx,addr_col] = clean_a1; a1 = clean_a1

            # Use geocache for each part
            if use_geocoding and geo_cache:
                parts = [p.strip() for p in a1.split(",") if p.strip()]
                new_street = []; found_city = cty; found_postal = pst; found_country = cou

                for part in parts:
                    # Pure postal
                    if re.match(r"^\d{4,6}$", part):
                        if not found_postal: found_postal = part
                        continue
                    # Postal+city pattern
                    m = re.match(r"^(\d{4,5})\s+([A-Z\u00C0-\u024F]\S.+)$", part)
                    if m:
                        if not found_postal: found_postal = m.group(1)
                        if not found_city: found_city = m.group(2)
                        continue
                    # Postal at end
                    m2 = re.search(r"(?<!\d)(\d{5,6})\s*$", part)
                    if m2:
                        if not found_postal: found_postal = m2.group(1)
                        remainder = part[:m2.start()].strip()
                        if remainder: new_street.append(remainder)
                        continue
                    # Street part
                    if ADDR_RE.search(part):
                        new_street.append(part); continue
                    # Geocode this part
                    geo = geo_cache.get(part, {})
                    cl = classify_place(part, geo)
                    if cl == "CITY" and not found_city:
                        found_city = part
                    elif cl == "COUNTRY" and not found_country:
                        found_country = geo.get("country", part)
                        if not found_postal and geo.get("postcode"): found_postal = geo["postcode"]
                    elif cl in ("STATE",) and sc and not st:
                        df.at[idx,sc] = part; st = part
                    elif COUNTRY_RE.match(part) and not found_country:
                        found_country = normalize_country(part)
                    elif not found_city and not ADDR_RE.search(part):
                        found_city = part  # best guess
                    else:
                        new_street.append(part)

                # Write back
                clean_a1 = ", ".join(new_street)
                if clean_a1 != a1: df.at[idx,addr_col] = clean_a1; record(addr_col, a1, clean_a1)
                if found_city and not cty and cc:
                    df.at[idx,cc] = found_city; record(cc, cty, found_city); cty = found_city
                if found_postal and not pst and pc:
                    df.at[idx,pc] = found_postal; record(pc, pst, found_postal); pst = found_postal
                if found_country and not cou and col_map.get("country"):
                    norm = normalize_country(found_country)
                    df.at[idx,col_map["country"]] = norm; record(col_map["country"], cou, norm); cou = norm

        # Expand abbreviations
        if o_abbrev and addr_col:
            a1_now = str(df.at[idx,addr_col]).strip()
            cleaned = clean_addr(a1_now)
            if cleaned and cleaned != a1_now:
                df.at[idx,addr_col] = cleaned; record(addr_col, a1_now, cleaned)

        # Fill COUNTRY_ID
        cou_final = str(df.at[idx,col_map.get("country","")] if col_map.get("country") else "").strip()
        cid_col   = col_map.get("country_id","")
        cid_final = str(df.at[idx,cid_col] if cid_col else "").strip()
        if cou_final and not cid_final and cid_col:
            code = COUNTRY_TO_CODE.get(cou_final.lower(),"")
            if code: df.at[idx,cid_col] = code

        # Quality + postal flag
        if o_postal and pc:
            df.at[idx,"POSTAL_FLAG"] = validate_postal(str(df.at[idx,pc]).strip(), cou_final)
        if o_quality:
            df.at[idx,"DATA_QUALITY"] = quality(row, col_map)

    progress.progress(1.0, text="Complete!")
    st.session_state["cleaned_df"] = df
    st.session_state["diffs"]      = diffs
    st.session_state["col_map"]    = col_map

# Show results
if "cleaned_df" in st.session_state:
    df      = st.session_state["cleaned_df"]
    diffs   = st.session_state["diffs"]
    col_map = st.session_state["col_map"]
    n = len(df)
    complete = (df.get("DATA_QUALITY","") == "COMPLETE").sum() if "DATA_QUALITY" in df.columns else 0
    partial  = (df.get("DATA_QUALITY","") == "PARTIAL").sum()  if "DATA_QUALITY" in df.columns else 0
    missing  = (df.get("DATA_QUALITY","") == "MISSING").sum()  if "DATA_QUALITY" in df.columns else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Rows", n)
    c2.metric("Changes", len(diffs))
    c3.metric("Complete", complete)
    c4.metric("Partial",  partial)
    c5.metric("Missing",  missing)

    tab1, tab2, tab3 = st.tabs(["📋 Cleaned Data","🔍 Changes","⬇️ Download"])
    with tab1: st.dataframe(df, use_container_width=True, height=500)
    with tab2:
        if diffs: st.dataframe(pd.DataFrame(diffs), use_container_width=True, height=400)
        else: st.info("No changes recorded")
    with tab3:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Cleaned Data", index=False)
            if diffs: pd.DataFrame(diffs).to_excel(w, sheet_name="Changes Log", index=False)
        st.download_button("⬇️ Download Cleaned Excel", buf.getvalue(),
            file_name="cleaned_addresses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
