import streamlit as st
import pandas as pd
import re
import io
from collections import defaultdict

st.set_page_config(page_title="Address Cleaning Agent", page_icon="🗺️", layout="wide")

# ══════════════════════════════════════════════════════════════════
# CLEANING ENGINE
# Rules learned from Raw/Cleaned example data
# ══════════════════════════════════════════════════════════════════

COUNTRY_RE = re.compile(
    r'\b(Russia|Russian Federation|China|United Kingdom|England|Scotland|'
    r'Northern Ireland|UK|United States|United States of America|USA|Germany|'
    r'France|Ukraine|India|Iran|Iraq|Syria|Turkey|Türkiye|Japan|South Korea|'
    r'North Korea|Myanmar|Malaysia|Indonesia|Pakistan|Egypt|Switzerland|'
    r'Netherlands|Belgium|Italy|Spain|Portugal|Greece|Poland|Romania|Bulgaria|'
    r'Hungary|Czech Republic|Austria|Sweden|Norway|Denmark|Finland|Ireland|'
    r'Luxembourg|Cyprus|Serbia|Croatia|Belarus|Kazakhstan|Azerbaijan|Armenia|'
    r'Georgia|Moldova|Singapore|Hong Kong|Taiwan|Philippines|Bangladesh|'
    r'Afghanistan|Uzbekistan|Vietnam|Thailand|Cambodia|Nigeria|Kenya|'
    r'South Africa|Ethiopia|Sudan|Yemen|Qatar|Kuwait|Oman|Bahrain|Jordan|'
    r'Lebanon|Israel|Saudi Arabia|United Arab Emirates|Canada|Mexico|Brazil|'
    r'Argentina|Venezuela|Colombia|Peru|Chile|Bolivia|Australia|New Zealand|'
    r'Libya|Tunisia|Morocco|Algeria|Palestine|Palestinian Territory|Seychelles|'
    r'Latvia|Monaco|Marshall Islands|Laos|Kyrgyzstan|Tajikistan|Turkmenistan|'
    r'Macau|Albania|Montenegro|Bosnia|Slovenia|Slovakia|Estonia|Lithuania|'
    r'Malta|Iceland|Kosovo|Honduras|Guyana|Uganda|Tanzania|Ghana|Senegal|'
    r'Ecuador|Paraguay|Uruguay|Nepal|Sri Lanka|Cuba|Dominican Republic|'
    r"Democratic People's Republic of Korea|DPRK|SYRIE|PHILIPPINES)\b",
    re.I
)

ADDR_RE = re.compile(
    r'\b(Street|St\b|Avenue|Ave\b|Road|Rd\b|Lane|Drive|Blvd|Floor|Suite|'
    r'Block|No\b|Ulitsa|Dom\b|Korpus|Etazh|Prospekt|Pereulok|Nabereznaya|'
    r'Shosse|Mahallesi|Cad\b|Sok\b|Jalan|Room|Unit|Bldg|Building|Apt\b|'
    r'Plaza|Tower|Center|Centre|Park|Square|Close|Court|Crescent|Grove|'
    r'Way\b|Walk\b|Hill\b|Gate\b|Gardens|Terrace|Heights|Manor|House\b|'
    r'Chemin|Rue\b|Boulevard|Weg\b|Gasse|Platz|Allee|Zona|Promyshlennaya|'
    r'ul\b|d\b(?=\s*\d)|Str\b|Mah\b|Blok\b|Kat\b|Daire\b|Pasa\b)'
    r'|strasse\b|straße\b|gasse\b|\bplatz\b',
    re.I
)

STATE_KW = re.compile(
    r'\b(Oblast|Krai|Kray|Okrug|Republic|Autonomous|Rayon|Raion|Province|'
    r'Region|Territory|County|Prefecture|Governorate|Penzenskaya|'
    r'Novosibirskaya|Volgogradskaya|Saratovskaya|Township|Borough|Parish)\b',
    re.I
)

DISTRICT_RE = re.compile(
    r'\b(\w[\w\s\-]+(?:District|Qu\b|gu\b|dong\b|guyok\b|Ward\b|'
    r'Quarter\b|Arrondissement\b))\b', re.I
)

TERRITORY_RE = re.compile(
    r',?\s*\b(West Bank|Northern Gaza|Gaza Strip|Gaza|Crimea)\b\s*,?', re.I
)

JUNK_POSTAL = re.compile(r'^(XX|00000|0{4,}|N/A|NA|None|null|TBD)$', re.I)
NOTE_RE     = re.compile(
    r'^(Linked To:|Address redacted|Unknown|Located in|Resident in|'
    r'Trust Company Complex|Letter )', re.I
)
ISO2_RE = re.compile(r'^[A-Z]{2}$')

US_STATES = {
    'AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN',
    'IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV',
    'NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN',
    'TX','UT','VT','VA','WA','WV','WI','WY','DC'
}
CA_PROVS = {'AB','BC','MB','NB','NL','NS','NT','NU','ON','PE','QC','SK','YT'}

US_STATE_EXPAND = {
    'AL':'Alabama','AK':'Alaska','AZ':'Arizona','AR':'Arkansas','CA':'California',
    'CO':'Colorado','CT':'Connecticut','DE':'Delaware','FL':'Florida','GA':'Georgia',
    'HI':'Hawaii','ID':'Idaho','IL':'Illinois','IN':'Indiana','IA':'Iowa',
    'KS':'Kansas','KY':'Kentucky','LA':'Louisiana','ME':'Maine','MD':'Maryland',
    'MA':'Massachusetts','MI':'Michigan','MN':'Minnesota','MS':'Mississippi',
    'MO':'Missouri','MT':'Montana','NE':'Nebraska','NV':'Nevada','NH':'New Hampshire',
    'NJ':'New Jersey','NM':'New Mexico','NY':'New York','NC':'North Carolina',
    'ND':'North Dakota','OH':'Ohio','OK':'Oklahoma','OR':'Oregon','PA':'Pennsylvania',
    'RI':'Rhode Island','SC':'South Carolina','SD':'South Dakota','TN':'Tennessee',
    'TX':'Texas','UT':'Utah','VT':'Vermont','VA':'Virginia','WA':'Washington',
    'WV':'West Virginia','WI':'Wisconsin','WY':'Wyoming','DC':'District of Columbia'
}
CA_PROV_EXPAND = {
    'AB':'Alberta','BC':'British Columbia','MB':'Manitoba','NB':'New Brunswick',
    'NL':'Newfoundland and Labrador','NS':'Nova Scotia','NT':'Northwest Territories',
    'NU':'Nunavut','ON':'Ontario','PE':'Prince Edward Island','QC':'Quebec',
    'SK':'Saskatchewan','YT':'Yukon'
}

COUNTRY_NORMALIZE = {
    'myanmar [burma]':'Myanmar','burma':'Myanmar',
    'republic of türkiye':'Turkey','republic of turkey':'Turkey','türkiye':'Turkey',
    'england':'United Kingdom','scotland':'United Kingdom',
    'northern ireland':'United Kingdom','uk':'United Kingdom',
    'great britain':'United Kingdom',
    'viet nam':'Vietnam','south korea':'South Korea',
    'republic of korea':'South Korea','north korea':'North Korea',
    "democratic people's republic of korea":'North Korea','dprk':'North Korea',
    'usa':'United States','u.s.a.':'United States','u.s.':'United States',
    'united states of america':'United States',
    'uae':'United Arab Emirates','irn':'Iran',
    'iran, islamic republic of':'Iran','islamic republic of iran':'Iran',
    'syrian arab republic':'Syria','syrie':'Syria',
    'west bank':'Palestinian Territory','northern gaza':'Palestinian Territory',
    'gaza':'Palestinian Territory','palestine':'Palestinian Territory',
    'russian federation':'Russia','notk':'',
    'republic of china':'Taiwan',
}

COUNTRY_TO_CODE = {
    'russia':'RU','china':'CN','united states':'US','united kingdom':'GB',
    'germany':'DE','france':'FR','ukraine':'UA','india':'IN','iran':'IR',
    'syria':'SY','iraq':'IQ','north korea':'KP','south korea':'KR',
    'japan':'JP','turkey':'TR','saudi arabia':'SA','united arab emirates':'AE',
    'pakistan':'PK','myanmar':'MM','thailand':'TH','vietnam':'VN',
    'malaysia':'MY','indonesia':'ID','singapore':'SG','hong kong':'HK',
    'taiwan':'TW','belarus':'BY','kazakhstan':'KZ','switzerland':'CH',
    'austria':'AT','netherlands':'NL','sweden':'SE','norway':'NO',
    'denmark':'DK','finland':'FI','ireland':'IE','portugal':'PT',
    'spain':'ES','italy':'IT','greece':'GR','israel':'IL','egypt':'EG',
    'libya':'LY','tunisia':'TN','morocco':'MA','algeria':'DZ',
    'canada':'CA','australia':'AU','new zealand':'NZ','brazil':'BR',
    'argentina':'AR','venezuela':'VE','colombia':'CO','peru':'PE',
    'seychelles':'SC','latvia':'LV','monaco':'MC','laos':'LA',
    'marshall islands':'MH','kyrgyzstan':'KG','tajikistan':'TJ',
    'turkmenistan':'TM','uzbekistan':'UZ','cambodia':'KH',
    'nigeria':'NG','kenya':'KE','south africa':'ZA','ethiopia':'ET',
    'afghanistan':'AF','bangladesh':'BD','philippines':'PH',
    'azerbaijan':'AZ','armenia':'AM','georgia':'GE','moldova':'MD',
    'serbia':'RS','croatia':'HR','czech republic':'CZ','hungary':'HU',
    'romania':'RO','bulgaria':'BG','poland':'PL','luxembourg':'LU',
    'cyprus':'CY','belgium':'BE','yemen':'YE','qatar':'QA',
    'kuwait':'KW','oman':'OM','bahrain':'BH','jordan':'JO','lebanon':'LB',
    'palestinian territory':'PS','mexico':'MX','chile':'CL','bolivia':'BO',
    'macau':'MO','estonia':'EE','lithuania':'LT','albania':'AL',
    'montenegro':'ME','slovenia':'SI','slovakia':'SK','malta':'MT',
    'iceland':'IS','kosovo':'XK','honduras':'HN','guyana':'GY',
    'uganda':'UG','tanzania':'TZ','ghana':'GH','senegal':'SN',
    'ecuador':'EC','nepal':'NP','sri lanka':'LK','cuba':'CU',
    'dominican republic':'DO','paraguay':'PY','uruguay':'UY',
    'british virgin islands':'VG','isle of man':'IM','bermuda':'BM',
}

ISO_TO_COUNTRY = {
    'RU':'Russia','CN':'China','US':'United States','GB':'United Kingdom',
    'DE':'Germany','FR':'France','UA':'Ukraine','IN':'India','IR':'Iran',
    'SY':'Syria','IQ':'Iraq','KP':'North Korea','KR':'South Korea',
    'JP':'Japan','TR':'Turkey','SA':'Saudi Arabia','AE':'United Arab Emirates',
    'PK':'Pakistan','MM':'Myanmar','TH':'Thailand','VN':'Vietnam',
    'MY':'Malaysia','ID':'Indonesia','SG':'Singapore','HK':'Hong Kong',
    'TW':'Taiwan','BY':'Belarus','KZ':'Kazakhstan','CH':'Switzerland',
    'AT':'Austria','NL':'Netherlands','SE':'Sweden','NO':'Norway',
    'DK':'Denmark','FI':'Finland','IE':'Ireland','PT':'Portugal',
    'ES':'Spain','IT':'Italy','GR':'Greece','IL':'Israel','EG':'Egypt',
    'CA':'Canada','AU':'Australia','NZ':'New Zealand','BR':'Brazil',
    'AR':'Argentina','VG':'British Virgin Islands','IM':'Isle of Man',
    'BM':'Bermuda','PS':'Palestinian Territory','RO':'Romania',
    'BG':'Bulgaria','PL':'Poland','LU':'Luxembourg','CY':'Cyprus',
    'BE':'Belgium','YE':'Yemen','QA':'Qatar','KW':'Kuwait','OM':'Oman',
    'BH':'Bahrain','JO':'Jordan','LB':'Lebanon','MX':'Mexico',
    'CL':'Chile','BO':'Bolivia','HN':'Honduras','GY':'Guyana',
    'AF':'Afghanistan','BD':'Bangladesh','PH':'Philippines','LK':'Sri Lanka',
    'NP':'Nepal','GE':'Georgia','AZ':'Azerbaijan','AM':'Armenia',
    'MD':'Moldova','RS':'Serbia','HR':'Croatia','CZ':'Czech Republic',
    'HU':'Hungary','AL':'Albania','ME':'Montenegro','SI':'Slovenia',
    'SK':'Slovakia','MT':'Malta','IS':'Iceland','XK':'Kosovo',
    'NG':'Nigeria','KE':'Kenya','ZA':'South Africa','ET':'Ethiopia',
    'SD':'Sudan','TZ':'Tanzania','UG':'Uganda','GH':'Ghana',
    'SN':'Senegal','EC':'Ecuador','CU':'Cuba','DO':'Dominican Republic',
    'PY':'Paraguay','UY':'Uruguay',
}

TERRITORY_TO_COUNTRY = {
    'west bank':'Palestinian Territory',
    'northern gaza':'Palestinian Territory',
    'gaza strip':'Palestinian Territory',
    'gaza':'Palestinian Territory',
    'crimea':'Russia',
}

KNOWN_STATES = {
    # China
    'shandong','guangdong','zhejiang','jiangsu','sichuan','hubei','hunan',
    'henan','hebei','fujian','liaoning','yunnan','guangxi','anhui','jiangxi',
    'shaanxi','gansu','guizhou','hainan','qinghai','xinjiang','heilongjiang',
    'jilin','beijing','shanghai','tianjin','chongqing',
    # India
    'maharashtra','karnataka','gujarat','uttar pradesh','west bengal',
    'andhra pradesh','telangana','kerala','punjab','haryana','rajasthan',
    'tamil nadu','madhya pradesh','odisha','jharkhand','chhattisgarh',
    'assam','bihar','himachal pradesh','uttarakhand','goa',
    # US
    'colorado','california','texas','florida','illinois','ohio','georgia',
    'michigan','virginia','washington','arizona','massachusetts','tennessee',
    'indiana','missouri','maryland','wisconsin','minnesota','alabama',
    'louisiana','kentucky','oregon','oklahoma','utah','iowa','nevada',
    'arkansas','mississippi','kansas','nebraska','idaho','connecticut',
    'delaware','south carolina','north carolina','pennsylvania','west virginia',
    'new york','new jersey','new mexico','north dakota','south dakota',
    'rhode island','vermont','maine','hawaii','alaska','montana','wyoming',
    'district of columbia',
    # Canada
    'ontario','quebec','alberta','british columbia','manitoba',
    'new brunswick','nova scotia','saskatchewan','newfoundland and labrador',
    # UK
    'essex','cambridgeshire','cheshire','west yorkshire','south yorkshire',
    'lancashire','kent','surrey','hampshire','hertfordshire','norfolk',
    'suffolk','derbyshire','nottinghamshire','leicestershire','warwickshire',
    'staffordshire','oxfordshire','berkshire','buckinghamshire','wiltshire',
    'somerset','devon','cornwall','dorset','gloucestershire','lincolnshire',
    'cumbria','county down','county antrim','middlesex',
    # Myanmar
    'yangon','mandalay',
    # Brazil
    'parana','sao paulo','minas gerais','rio de janeiro','bahia',
    'rio grande do sul','santa catarina','goias','pernambuco','ceara',
    # Other
    'khanh hoa','tortola',
}

CITY_TO_COUNTRY = {
    'moscow':'Russia','saint petersburg':'Russia','novosibirsk':'Russia',
    'yekaterinburg':'Russia','krasnodar':'Russia','kazan':'Russia',
    'samara':'Russia','omsk':'Russia','chelyabinsk':'Russia','ufa':'Russia',
    'perm':'Russia','volgograd':'Russia','rostov-on-don':'Russia',
    'voronezh':'Russia','saratov':'Russia','tyumen':'Russia',
    'tolyatti':'Russia','izhevsk':'Russia','barnaul':'Russia',
    'ulyanovsk':'Russia','irkutsk':'Russia','khabarovsk':'Russia',
    'yaroslavl':'Russia','vladivostok':'Russia','tomsk':'Russia',
    'orenburg':'Russia','kemerovo':'Russia','tula':'Russia','kirov':'Russia',
    'ryazan':'Russia','lipetsk':'Russia','penza':'Russia','astrakhan':'Russia',
    'naberezhnye chelny':'Russia','bryansk':'Russia','tver':'Russia',
    'ivanovo':'Russia','surgut':'Russia','belgorod':'Russia',
    'kursk':'Russia','murmansk':'Russia','nizhny novgorod':'Russia',
    'novorossiysk':'Russia','makhachkala':'Russia','novy urengoy':'Russia',
    'luhansk':'Ukraine','donetsk':'Ukraine',
    'zürich':'Switzerland','zurich':'Switzerland','geneva':'Switzerland',
    'bern':'Switzerland','luzern':'Switzerland','lugano':'Switzerland',
    'zug':'Switzerland','baar':'Switzerland','dübendorf':'Switzerland',
    'dubendorf':'Switzerland','basel':'Switzerland',
    'london':'United Kingdom','manchester':'United Kingdom',
    'birmingham':'United Kingdom','leeds':'United Kingdom',
    'glasgow':'United Kingdom','liverpool':'United Kingdom',
    'hartlepool':'United Kingdom','chesterfield':'United Kingdom',
    'high wycombe':'United Kingdom','peterborough':'United Kingdom',
    'wakefield':'United Kingdom','norwich':'United Kingdom',
    'newark':'United Kingdom','chorley':'United Kingdom',
    'paris':'France','berlin':'Germany','munich':'Germany',
    'hamburg':'Germany','frankfurt':'Germany','cologne':'Germany',
    'beijing':'China','shanghai':'China','shenzhen':'China',
    'guangzhou':'China','chengdu':'China','wuhan':'China',
    'nanjing':'China','xi\'an':'China','hangzhou':'China',
    'tokyo':'Japan','osaka':'Japan','seoul':'South Korea',
    'pyongyang':'North Korea','tehran':'Iran','baghdad':'Iraq',
    'damascus':'Syria','istanbul':'Turkey','ankara':'Turkey',
    'dubai':'United Arab Emirates','abu dhabi':'United Arab Emirates',
    'riyadh':'Saudi Arabia','doha':'Qatar','muscat':'Oman',
    'amman':'Jordan','beirut':'Lebanon','tel aviv':'Israel',
    'cairo':'Egypt','tripoli':'Libya','tunis':'Tunisia',
    'algiers':'Algeria','rabat':'Morocco','khartoum':'Sudan',
    'nairobi':'Kenya','lagos':'Nigeria','accra':'Ghana',
    'addis ababa':'Ethiopia','kampala':'Uganda',
    'dar es salaam':'Tanzania','dakar':'Senegal',
    'toronto':'Canada','montreal':'Canada','vancouver':'Canada',
    'sydney':'Australia','melbourne':'Australia',
    'minsk':'Belarus','kyiv':'Ukraine','kiev':'Ukraine',
    'tashkent':'Uzbekistan','almaty':'Kazakhstan','baku':'Azerbaijan',
    'yangon':'Myanmar','bangkok':'Thailand','hanoi':'Vietnam',
    'jakarta':'Indonesia','kuala lumpur':'Malaysia',
    'ramallah':'Palestinian Territory','kabul':'Afghanistan',
    'mumbai':'India','delhi':'India','bangalore':'India',
    'rawalpindi':'Pakistan','lahore':'Pakistan','karachi':'Pakistan',
    'caracas':'Venezuela','bogota':'Colombia','lima':'Peru',
    'buenos aires':'Argentina','mexico city':'Mexico',
    'tegucigalpa':'Honduras','curitiba':'Brazil','sao paulo':'Brazil',
    'rio de janeiro':'Brazil','brasilia':'Brazil',
    'indianapolis':'United States','houston':'United States',
    'virginia beach':'United States','richmond':'United States',
    'alexandria':'United States',
}

def normalize_country(c):
    c = c.strip()
    if not c: return c
    if c.upper() == c and len(c) > 2:
        c = c.title()
    return COUNTRY_NORMALIZE.get(c.lower(), c)

def get_country_code(country):
    return COUNTRY_TO_CODE.get(country.lower().strip(), '')

def parse_city_field(cty):
    """Split CITY into city + state + addr3."""
    if not cty: return '', '', ''
    s = re.sub(r'\s+Province\b', '', cty.strip(), flags=re.I)
    s = re.sub(r'\s+City\b(?=,|$)', '', s, flags=re.I)
    s = re.sub(r'^[Gg]orod\s+', '', s, flags=re.I)
    s = re.sub(r'^[Gg]\.\s*', '', s)
    parts = [p.strip() for p in s.split(',') if p.strip()]
    if len(parts) == 1:
        p = parts[0]
        if STATE_KW.search(p): return '', p, ''
        if p.lower() in KNOWN_STATES: return '', p, ''
        return p, '', ''
    region_idx = {
        i for i, p in enumerate(parts)
        if STATE_KW.search(p) or p.lower() in KNOWN_STATES
    }
    other = [parts[i] for i in range(len(parts)) if i not in region_idx]
    if not other: return parts[0], '', ''
    # City is first non-region part (or a known capital)
    city = other[0]
    for p in other:
        if p.lower() in CITY_TO_COUNTRY: city = p; break
    state_parts = [parts[i] for i in sorted(region_idx)]
    addr3_parts = [p for p in other if p != city]
    return city, ', '.join(state_parts), ', '.join(addr3_parts)

def extract_from_address1(a1, ex_city='', ex_postal='', ex_country=''):
    """
    THE RULE (learned from 200+ Raw/Cleaned examples):
    For each comma-separated part, strip from the RIGHT:
      Country → Postal → City (always just the last 1 word)
    Everything before stays in ADDRESS1.
    District/neighbourhood before city STAYS in ADDRESS1 — never moves to CITY.
    Does NOT expand abbreviations.
    """
    if not a1: return a1, ex_city, ex_postal, ex_country, '', ''
    city=ex_city.strip(); postal=ex_postal.strip(); country=ex_country.strip()
    parts=[p.strip() for p in a1.split(',') if p.strip()]
    street_parts=[]

    def is_city_part(part):
        """Short standalone part that is just a city name (not a phrase or country)."""
        p=part.strip()
        if not p or ADDR_RE.search(p) or re.match(r'^\d',p): return False
        if COUNTRY_RE.search(p): return False
        if ISO2_RE.match(p) and p in ISO_TO_COUNTRY: return False
        words=p.split()
        if len(words)>2: return False
        return bool(re.match(r'^[A-ZÀ-ɏ]',p) and len(p)<40)

    def has_geo(part):
        return (bool(COUNTRY_RE.search(part)) or
                bool(re.search(r'\d{4,7}\s*$',part)) or
                bool(ISO2_RE.match(part.strip()) and part.strip() in ISO_TO_COUNTRY))

    def strip_tail(s, c='', p='', co=''):
        """Strip Country/Postal/City from right end. Returns (street, city, postal, country)."""
        # Country+Postal together: "Turkey 34758" or "Russia 119071"
        mc_p=re.search(r'\s+('+COUNTRY_RE.pattern+r')\s+(\d{4,7})\s*$',s,re.I)
        if mc_p:
            if not co: co=mc_p.group(1)
            if not p:  p=mc_p.group(mc_p.lastindex)
            s=s[:mc_p.start()].strip()
        else:
            # ISO at end
            iso_m=re.search(r'\s+([A-Z]{2})\s*$',s)
            if iso_m and iso_m.group(1) in ISO_TO_COUNTRY and iso_m.group(1) not in ('No','D'):
                if not co: co=ISO_TO_COUNTRY[iso_m.group(1)]
                s=s[:iso_m.start()].strip()
            else:
                # Country alone at end
                mc=re.search(r'\s+('+COUNTRY_RE.pattern+r')\s*$',s,re.I)
                if mc:
                    if not co: co=mc.group(1)
                    s=s[:mc.start()].strip()
            # Postal alone at end
            mp=re.search(r'\s+(\d{4,7})\.?\d*\s*$',s)
            if mp and not re.match(r'^\d',s):
                if not p: p=re.sub(r'\.0+$','',mp.group(1))
                s=s[:mp.start()].strip()
        # City: last 1 word (or 2 for known multi-word cities)
        if not c:
            words=s.split()
            if len(words)>=2:
                two=' '.join(words[-2:]).lower()
                if two in MULTIWORD_CITIES:
                    c=' '.join(words[-2:]); s=' '.join(words[:-2]).strip()
                else:
                    last=words[-1]
                    if re.match(r'^[A-ZÀ-ɏ]',last) and not re.match(r'^\d',last) and not COUNTRY_RE.match(last) and len(last)>1:
                        c=last; s=' '.join(words[:-1]).strip()
        return s,c,p,co

    MULTIWORD_CITIES={'saint petersburg','new york','hong kong','kuala lumpur','buenos aires',
        'mexico city','rio de janeiro','sao paulo','los angeles','las vegas','tel aviv',
        'abu dhabi','dar es salaam','addis ababa','new delhi','cape town','ho chi minh',
        'sakarya adapazari',
        # Russian 2-word cities
        'novyy urengoy','novy urengoy','novyy urengoi','nizhny novgorod',
        'naberezhnye chelny','veliky novgorod','komsomolsk-on-amur',
        'yuzhno-sakhalinsk','petropavlovsk-kamchatsky','khanty-mansiysk',
        'yoshkar-ola','ulan-ude','rostov-on-don'}

    for i,part in enumerate(parts):
        # Pure postal
        if re.match(r'^\d{4,7}\.?\d*$',part):
            if not postal: postal=re.sub(r'\.0+$','',part)
            continue
        # "8600 Dübendorf"
        m1=re.match(r'^(\d{4,5})\s+([A-ZÀ-ɏ]\S.+)$',part)
        if m1:
            if not postal: postal=m1.group(1)
            if not city:   city=m1.group(2)
            continue
        # ISO alone
        if ISO2_RE.match(part) and part in ISO_TO_COUNTRY and part not in ('No','D'):
            if not country: country=ISO_TO_COUNTRY[part]
            continue
        # Country+Postal BEFORE Country-alone
        mc_p=re.match(r'^('+COUNTRY_RE.pattern+r')\s+(\d{4,7})\s*$',part,re.I)
        if mc_p:
            if not country: country=mc_p.group(1)
            if not postal:  postal=mc_p.group(mc_p.lastindex)
            continue
        # Country alone (fullmatch)
        clean_p=re.sub(r'\d+\.?\d*\s*$','',part).strip()
        if COUNTRY_RE.search(clean_p) and re.fullmatch(COUNTRY_RE.pattern,clean_p,re.I):
            if not country: country=clean_p
            continue
        # Standalone city adjacent to geo part
        if is_city_part(part) and not city:
            nxt=i+1<len(parts) and has_geo(parts[i+1])
            prv=i>0 and has_geo(parts[i-1])
            if nxt or prv: city=part; continue
        # Embedded country/postal/ISO
        has_c=bool(COUNTRY_RE.search(part)) and not ISO2_RE.match(part.strip())
        has_p=bool(re.search(r'\s\d{4,7}\.?\d*\s*$',part))
        ie=part.strip().split()[-1]
        has_i=bool(ISO2_RE.match(ie) and ie in ISO_TO_COUNTRY and ie not in ('No','D'))
        if has_c or has_p or has_i:
            rem,nc,np,nco=strip_tail(part,city,postal,country)
            if rem: street_parts.append(rem)
            if not city and nc:    city=nc
            if not postal and np:  postal=np
            if not country and nco: country=nco
            continue
        # Street part with embedded postal+city: "Str 30. 8600 Dübendorf"
        if re.search(r'\d{4,5}\s+[A-ZÀ-ɏ]',part):
            m2=re.search(r'(\d{4,5})\s+([A-ZÀ-ɏ][a-zÀ-ɏ]+(?:\s+[A-Z][a-z]+)?)\s*$',part)
            if m2:
                if not postal: postal=m2.group(1)
                if not city:   city=m2.group(2)
                street_parts.append(part[:m2.start()].strip()); continue
        # City appended to street: "...Cad. Izmir"
        words=part.split()
        if len(words)>1 and not city:
            last=words[-1]
            if (re.match(r'^[A-ZÀ-ɏ]',last) and not re.match(r'^\d',last)
                    and not ADDR_RE.search(last) and not COUNTRY_RE.match(last) and len(last)>2):
                rest=' '.join(words[:-1])
                if ADDR_RE.search(rest): city=last; street_parts.append(rest); continue
        street_parts.append(part)

    # No-comma fallback
    if not street_parts and len(parts)==1:
        rem,nc,np,nco=strip_tail(a1,city,postal,country)
        if not city and nc:    city=nc
        if not postal and np:  postal=np
        if not country and nco: country=nco
        if rem: street_parts.append(rem)

    if country:
        country=COUNTRY_NORMALIZE.get(country.lower().strip(),country) or country
    if country and ISO2_RE.match(country) and country in ISO_TO_COUNTRY:
        country=ISO_TO_COUNTRY[country]

    return ', '.join(p for p in street_parts if p.strip()) or a1, city, postal, country, '', ''


def infer_country(city, state, postal):
    c = str(city  or '').strip().lower()
    s = str(state or '').strip().lower()
    p = str(postal or '').strip()
    if c in CITY_TO_COUNTRY: return CITY_TO_COUNTRY[c]
    if s in TERRITORY_TO_COUNTRY: return TERRITORY_TO_COUNTRY[s]
    if s in {'maharashtra','karnataka','gujarat','uttar pradesh','west bengal',
             'andhra pradesh','telangana','kerala','punjab','haryana','rajasthan',
             'tamil nadu','madhya pradesh','odisha','goa','assam','bihar'}:
        return 'India'
    if s in {'shandong','guangdong','zhejiang','jiangsu','sichuan','hubei',
             'hunan','henan','hebei','fujian','liaoning','yunnan','guangxi',
             'anhui','jiangxi','shaanxi','xinjiang','beijing','shanghai',
             'tianjin','chongqing'}:
        return 'China'
    if re.search(r'\b(Oblast|Krai|Okrug|Bashkortostan|Tatarstan|Dagestan|'
                 r'Komi|Udmurt|Leningrad|Sverdlovsk)\b', s, re.I):
        return 'Russia'
    if s in US_STATE_EXPAND.values() or s in {v.lower() for v in US_STATE_EXPAND.values()}:
        return 'United States'
    if s in {v.lower() for v in CA_PROV_EXPAND.values()}:
        return 'Canada'
    if s in {'yangon region','mandalay region','yangon','mandalay'}:
        return 'Myanmar'
    if s in {'ontario','quebec','alberta','british columbia','manitoba',
             'new brunswick','nova scotia','saskatchewan',
             'newfoundland and labrador'}:
        return 'Canada'
    if re.match(r'^\d{6}$', p): return 'Russia'
    return ''

def validate_postal(p, country):
    POSTAL_FMTS = {
        'russia':r'^\d{6}$','ukraine':r'^\d{5,6}$','china':r'^\d{6}$',
        'india':r'^\d{6}$','united states':r'^\d{5}(-\d{4})?$',
        'united kingdom':r'^[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}$',
        'germany':r'^\d{5}$','france':r'^\d{5}$',
        'switzerland':r'^(CH-?)?\d{4}$','austria':r'^\d{4}$',
        'netherlands':r'^\d{4}\s?[A-Z]{2}$',
    }
    p = str(p or '').strip()
    if not p: return 'MISSING'
    if JUNK_POSTAL.match(p): return 'INVALID'
    fmt = POSTAL_FMTS.get(str(country or '').lower().strip())
    if not fmt: return 'VALID'
    return 'VALID' if re.match(fmt, p, re.I) else 'INVALID'

def data_quality(row, col_map):
    has = sum([
        bool(str(row.get(col_map.get('address',''),'') or '').strip()),
        bool(str(row.get(col_map.get('city',''),'')   or '').strip()),
        bool(str(row.get(col_map.get('postal',''),'') or '').strip()),
        bool(str(row.get(col_map.get('country',''),'')or '').strip()),
    ])
    return 'COMPLETE' if has==4 else 'PARTIAL' if has>=2 else 'MISSING'

# ── Column hint detection ─────────────────────────────────────────
HINTS = {
    'address':    ['address1','address','addr1','street','line1'],
    'address2':   ['address2','addr2','line2'],
    'address3':   ['address3','addr3','line3'],
    'city':       ['city','town','locality'],
    'state':      ['state','province','region','oblast'],
    'postal':     ['postal','zip','postcode','postalcode'],
    'country':    ['country','countryname','nation'],
    'country_id': ['countryid','country_id','countrycode','country_code','iso'],
}
def guess_col(headers, field):
    for h in headers:
        clean = h.lower().replace(' ','').replace('_','')
        if any(clean == hint for hint in HINTS[field]):
            return h
    return ''

# ══════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
.stApp { max-width: 1400px }
.metric-row { display: flex; gap: 12px; margin-bottom: 16px }
</style>
""", unsafe_allow_html=True)

st.title("🗺️ Address Cleaning Agent")
st.caption("Upload any Excel or CSV — automatically extracts city, postal, country and routes each value to the correct column.")

# ── Sidebar ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    uploaded = st.file_uploader(
        "Upload file to clean",
        type=["xlsx","xls","csv"],
        help="Excel or CSV with address data"
    )

    col_map = {}
    df_raw  = None

    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded)
            else:
                df_raw = pd.read_excel(uploaded)
            df_raw = df_raw.fillna('').astype(str).replace('nan','')
            headers = df_raw.columns.tolist()
            st.success(f"✓ {len(df_raw):,} rows · {len(headers)} columns")

            st.divider()
            st.subheader("Column Mapping")
            for field in ['address','address2','address3','city','state','postal','country','country_id']:
                default = guess_col(headers, field)
                idx = headers.index(default)+1 if default in headers else 0
                col_map[field] = st.selectbox(
                    field.replace('_',' ').title(),
                    ['— skip —']+headers,
                    index=idx, key=f'col_{field}'
                )
                col_map[field] = '' if col_map[field]=='— skip —' else col_map[field]

        except Exception as e:
            st.error(f"Error reading file: {e}")
            uploaded = None

    st.divider()
    st.subheader("Options")
    o_postal  = st.checkbox("Validate postal codes",  True)
    o_quality = st.checkbox("Add data quality score", True)
    o_infer   = st.checkbox("Infer missing countries", True)

    run = st.button(
        "▶  Clean Addresses", type="primary",
        use_container_width=True, disabled=uploaded is None
    )

# ── Main ──────────────────────────────────────────────────────────
if not uploaded:
    st.info("👈 Upload a file in the sidebar to get started")
    st.stop()

if not run and 'cleaned_df' not in st.session_state:
    st.subheader("Preview")
    st.dataframe(df_raw.head(30), use_container_width=True)
    st.caption(f"Showing 30 of {len(df_raw):,} rows")
    st.stop()

if run:
    df     = df_raw.copy()
    diffs  = []
    total  = len(df)
    prog   = st.progress(0, text="Starting...")

    def C(col): return col_map.get(col,'')
    def get(row, col): return str(row.get(C(col),'') or '').strip()
    def record(idx, field_key, before, after):
        if before != after and after and C(field_key):
            diffs.append({'Row':idx+1,'Field':C(field_key),'Before':before,'After':after})
    def setcol(idx, field_key, val):
        if C(field_key): df.at[idx, C(field_key)] = val

    for idx, row in df.iterrows():
        if idx % 1000 == 0:
            prog.progress(idx/total, text=f"Processing row {idx:,} of {total:,}…")

        a1  = get(row,'address')
        a2  = get(row,'address2')
        a3  = get(row,'address3')
        cty = get(row,'city')
        st_ = get(row,'state')
        pst = get(row,'postal')
        cou = get(row,'country')
        cid = get(row,'country_id')

        # ── 1. Normalize COUNTRY ──────────────────────────────────
        if cou:
            norm = normalize_country(cou)
            if norm != cou:
                record(idx,'country',cou,norm)
                setcol(idx,'country',norm); cou=norm

        # ── 2. Fix STATE abbreviations ────────────────────────────
        if ISO2_RE.match(st_):
            if st_ in US_STATES:
                new_st = US_STATE_EXPAND[st_]
                record(idx,'state',st_,new_st)
                setcol(idx,'state',new_st); st_=new_st
                if not cou:
                    setcol(idx,'country','United States'); cou='United States'
                    setcol(idx,'country_id','US');         cid='US'
            elif st_ in CA_PROVS:
                new_st = CA_PROV_EXPAND[st_]
                record(idx,'state',st_,new_st)
                setcol(idx,'state',new_st); st_=new_st
                if not cou:
                    setcol(idx,'country','Canada'); cou='Canada'
                    setcol(idx,'country_id','CA'); cid='CA'
            elif st_ in ISO_TO_COUNTRY:
                if not cou:
                    setcol(idx,'country',ISO_TO_COUNTRY[st_]); cou=ISO_TO_COUNTRY[st_]
                if not cid:
                    setcol(idx,'country_id',st_); cid=st_
                setcol(idx,'state',''); st_=''

        # ── 3. Fix POSTAL junk ────────────────────────────────────
        if pst and JUNK_POSTAL.match(pst):
            setcol(idx,'postal',''); pst=''

        # ── 4. Postal in STATE → move ─────────────────────────────
        if re.match(r'^\d{4,6}$', st_) and not pst:
            record(idx,'postal','',st_)
            setcol(idx,'postal',st_); pst=st_
            setcol(idx,'state','');   st_=''

        # ── 5. Country name in STATE → move ──────────────────────
        if st_ and COUNTRY_RE.match(st_) and not ADDR_RE.search(st_):
            norm = normalize_country(st_)
            if not cou:
                record(idx,'country','',norm)
                setcol(idx,'country',norm); cou=norm
            setcol(idx,'state',''); st_=''

        # ── 6. Territory in STATE → infer COUNTRY ─────────────────
        if st_.lower() in TERRITORY_TO_COUNTRY and not cou:
            tc = TERRITORY_TO_COUNTRY[st_.lower()]
            setcol(idx,'country',tc); cou=tc

        # ── 7. Fix ADDRESS2 ───────────────────────────────────────
        if a2:
            if STATE_KW.search(a2) and not ADDR_RE.search(a2) and not st_:
                record(idx,'state','',a2)
                setcol(idx,'state',a2); st_=a2
                setcol(idx,'address2',''); a2=''
            elif re.match(r'^\d{4,6}$',a2) and not pst:
                setcol(idx,'postal',a2); pst=a2
                setcol(idx,'address2',''); a2=''
            elif (not cty and not ADDR_RE.search(a2) and not COUNTRY_RE.search(a2)
                    and not re.match(r'^\d',a2) and len(a2.split(','))<=2):
                record(idx,'city','',a2)
                setcol(idx,'city',a2); cty=a2
                setcol(idx,'address2',''); a2=''

        # ── 8. Fix ADDRESS3 ───────────────────────────────────────
        if a3:
            if re.compile(r'\b(Oblast|Penzenskaya|Novosibirskaya)\b',re.I).search(a3) and not st_:
                setcol(idx,'state',a3); st_=a3
                setcol(idx,'address3',''); a3=''
            elif re.match(r'^\d{4,6}$',a3) and not pst:
                setcol(idx,'postal',a3); pst=a3
                setcol(idx,'address3',''); a3=''

        # ── 9. Parse CITY field ───────────────────────────────────
        if cty:
            # Remove embedded postal
            city_parts = [p.strip() for p in cty.split(',')]
            clean_city = []
            for part in city_parts:
                if re.match(r'^\d{5,6}$',part) and not pst:
                    setcol(idx,'postal',part); pst=part
                else:
                    clean_city.append(part)
            if len(clean_city) != len(city_parts):
                cty = ', '.join(clean_city); setcol(idx,'city',cty)

            city_v, state_v, addr3_v = parse_city_field(cty)
            if city_v != cty or state_v or addr3_v:
                record(idx,'city',cty,city_v)
                setcol(idx,'city',city_v); cty=city_v
                if state_v:
                    setcol(idx,'state',state_v); st_=state_v
                if addr3_v and C('address3'):
                    ex = str(df.at[idx,C('address3')]).strip()
                    setcol(idx,'address3',(ex+', '+addr3_v).strip(', ') if ex else addr3_v)

        # ── 10. Extract territory from ADDRESS1 ───────────────────
        if a1 and not NOTE_RE.match(a1):
            tm = TERRITORY_RE.search(a1)
            if tm:
                territory = tm.group(1).strip()
                if not st_: setcol(idx,'state',territory); st_=territory
                if not cou:
                    tc = TERRITORY_TO_COUNTRY.get(territory.lower(),'')
                    if tc: setcol(idx,'country',tc); cou=tc
                clean_a1 = TERRITORY_RE.sub('',a1).strip().strip(',').strip()
                setcol(idx,'address',clean_a1); a1=clean_a1

        # ── 11. Parse ADDRESS1 ────────────────────────────────────
        if a1 and not NOTE_RE.match(a1):
            # Clear CITY if it equals raw ADDRESS1 (bug from earlier)
            if cty and (cty == a1 or ADDR_RE.search(cty)):
                setcol(idx,'city',''); cty=''

            clean_a1, new_city, new_postal, new_country, new_a2, new_a3 = \
                extract_from_address1(a1, cty, pst, cou)

            if clean_a1 and clean_a1 != a1:
                record(idx,'address',a1,clean_a1)
                setcol(idx,'address',clean_a1)

            if new_city and not cty:
                record(idx,'city','',new_city)
                setcol(idx,'city',new_city); cty=new_city

            if new_postal and not pst:
                record(idx,'postal','',new_postal)
                setcol(idx,'postal',new_postal); pst=new_postal

            if new_country and not cou:
                norm = normalize_country(new_country)
                record(idx,'country','',norm)
                setcol(idx,'country',norm); cou=norm

            if new_a2 and not a2 and C('address2'):
                setcol(idx,'address2',new_a2)

            if new_a3 and C('address3'):
                ex = str(df.at[idx,C('address3')] if C('address3') else '').strip()
                if not ex: setcol(idx,'address3',new_a3)

        # ── 12. Fix CITY = country name ───────────────────────────
        cty_now = str(df.at[idx,C('city')] if C('city') else '').strip()
        if cty_now and COUNTRY_RE.fullmatch(cty_now):
            if not cou:
                norm = normalize_country(cty_now)
                setcol(idx,'country',norm); cou=norm
            setcol(idx,'city','')

        # ── 13. Infer COUNTRY if missing ──────────────────────────
        if o_infer:
            cou_now = str(df.at[idx,C('country')] if C('country') else '').strip()
            if not cou_now:
                inferred = infer_country(
                    str(df.at[idx,C('city')]  if C('city')  else ''),
                    str(df.at[idx,C('state')] if C('state') else ''),
                    str(df.at[idx,C('postal')]if C('postal')else ''),
                )
                if inferred:
                    setcol(idx,'country',inferred); cou=inferred

        # ── 14. Fill COUNTRY_ID ───────────────────────────────────
        fc = str(df.at[idx,C('country')] if C('country') else '').strip()
        fi = str(df.at[idx,C('country_id')] if C('country_id') else '').strip()
        if fc and not fi and C('country_id'):
            code = get_country_code(fc)
            if code: setcol(idx,'country_id',code)

        # ── 15. Postal validation ─────────────────────────────────
        if o_postal and C('postal'):
            df.at[idx,'POSTAL_FLAG'] = validate_postal(
                str(df.at[idx,C('postal')]).strip(),
                str(df.at[idx,C('country')] if C('country') else '').strip()
            )

        # ── 16. Data quality ──────────────────────────────────────
        if o_quality:
            df.at[idx,'DATA_QUALITY'] = data_quality(row, col_map)

    prog.progress(1.0, text="Complete!")
    st.session_state['cleaned_df'] = df
    st.session_state['diffs']      = diffs
    st.session_state['col_map']    = col_map

# ── Show results ──────────────────────────────────────────────────
if 'cleaned_df' in st.session_state:
    df      = st.session_state['cleaned_df']
    diffs   = st.session_state['diffs']
    col_map = st.session_state['col_map']
    n = len(df)

    complete = (df.get('DATA_QUALITY','')=='COMPLETE').sum() if 'DATA_QUALITY' in df.columns else 0
    partial  = (df.get('DATA_QUALITY','')=='PARTIAL').sum()  if 'DATA_QUALITY' in df.columns else 0
    missing  = (df.get('DATA_QUALITY','')=='MISSING').sum()  if 'DATA_QUALITY' in df.columns else 0
    invalid  = (df.get('POSTAL_FLAG','')=='INVALID').sum()   if 'POSTAL_FLAG'  in df.columns else 0

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Rows",         f"{n:,}")
    c2.metric("Changes",      f"{len(diffs):,}")
    c3.metric("Complete",     f"{complete:,}", f"{round(complete/n*100) if n else 0}%")
    c4.metric("Partial",      f"{partial:,}",  f"{round(partial/n*100)  if n else 0}%")
    c5.metric("Missing",      f"{missing:,}",  f"{round(missing/n*100)  if n else 0}%")
    c6.metric("Invalid postal",f"{invalid:,}")

    tab1,tab2,tab3 = st.tabs(["📋 Cleaned Data","🔍 Changes","⬇️ Download"])

    with tab1:
        st.dataframe(df, use_container_width=True, height=500)

    with tab2:
        if diffs:
            st.dataframe(pd.DataFrame(diffs), use_container_width=True, height=400)
        else:
            st.info("No changes recorded")

    with tab3:
        st.subheader("Download cleaned file")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='Cleaned Data', index=False)
            if diffs:
                pd.DataFrame(diffs).to_excel(w, sheet_name='Changes Log', index=False)
        st.download_button(
            "⬇️  Download Cleaned Excel",
            data=buf.getvalue(),
            file_name="cleaned_addresses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
