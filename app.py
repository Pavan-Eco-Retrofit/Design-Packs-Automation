import streamlit as st
import os
import glob
import re
import math
import requests
import fitz  # PyMuPDF
import pandas as pd
from openpyxl import load_workbook
from copy import copy
from openpyxl.utils import range_boundaries
import duckdb
import time
import tempfile
import io

# ================================================================
# PAGE CONFIG
# ================================================================
st.set_page_config(
    page_title="Retrofit Design Exporter",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ================================================================
# CUSTOM CSS — dark industrial aesthetic
# ================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@700;800&family=Inter:wght@300;400;500&display=swap');

:root {
    --bg: #0d0f12;
    --surface: #141720;
    --surface2: #1c2030;
    --border: #2a2f3e;
    --accent: #00e5a0;
    --accent2: #3b82f6;
    --warn: #f59e0b;
    --danger: #ef4444;
    --text: #e2e8f0;
    --muted: #64748b;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: var(--bg);
    color: var(--text);
}

.stApp { background: var(--bg); }

section[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border);
}

h1, h2, h3 {
    font-family: 'Syne', sans-serif;
    color: var(--text);
    letter-spacing: -0.02em;
}

.hero-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.8rem;
    font-weight: 800;
    background: linear-gradient(135deg, #00e5a0 0%, #3b82f6 60%, #a855f7 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height: 1.1;
    margin-bottom: 0.2rem;
}

.hero-sub {
    font-family: 'DM Mono', monospace;
    font-size: 0.85rem;
    color: var(--muted);
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 2rem;
}

.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
    position: relative;
    overflow: hidden;
}

.card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
}

.card-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.6rem;
}

.badge {
    display: inline-block;
    background: rgba(0,229,160,0.1);
    border: 1px solid rgba(0,229,160,0.3);
    color: var(--accent);
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 0.72rem;
    font-family: 'DM Mono', monospace;
    letter-spacing: 0.05em;
}

.badge-blue {
    background: rgba(59,130,246,0.1);
    border-color: rgba(59,130,246,0.3);
    color: var(--accent2);
}

.badge-warn {
    background: rgba(245,158,11,0.1);
    border-color: rgba(245,158,11,0.3);
    color: var(--warn);
}

.log-box {
    background: #0a0c10;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 1rem 1.2rem;
    font-family: 'DM Mono', monospace;
    font-size: 0.78rem;
    color: #94a3b8;
    min-height: 120px;
    max-height: 340px;
    overflow-y: auto;
    line-height: 1.7;
}

.log-ok   { color: var(--accent); }
.log-warn { color: var(--warn); }
.log-err  { color: var(--danger); }
.log-info { color: var(--accent2); }

.stButton > button {
    background: linear-gradient(135deg, #00e5a0, #3b82f6) !important;
    color: #0d0f12 !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.65rem 2rem !important;
    letter-spacing: 0.02em !important;
    transition: all 0.2s !important;
    width: 100%;
}

.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 24px rgba(0,229,160,0.3) !important;
}

.dl-btn > button {
    background: var(--surface2) !important;
    color: var(--text) !important;
    border: 1px solid var(--border) !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
}

div[data-testid="stFileUploader"] {
    background: var(--surface2);
    border: 1px dashed var(--border);
    border-radius: 10px;
    padding: 0.5rem;
}

.stTextInput > div > div > input,
.stSelectbox > div > div,
.stNumberInput > div > div > input {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    color: var(--text) !important;
    border-radius: 8px !important;
    font-family: 'Inter', sans-serif !important;
}

label, .stTextInput label, .stSelectbox label, .stFileUploader label {
    color: var(--muted) !important;
    font-size: 0.82rem !important;
    font-family: 'DM Mono', monospace !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
}

.stProgress > div > div > div {
    background: linear-gradient(90deg, var(--accent), var(--accent2)) !important;
}

hr { border-color: var(--border) !important; }

.stExpander {
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    background: var(--surface) !important;
}
</style>
""", unsafe_allow_html=True)

# ================================================================
# CUSTOM ROUND
# ================================================================
def custom_round(x) -> float:
    try:
        x = float(x)
    except (TypeError, ValueError):
        return x
    floor_val = math.floor(x)
    return float(math.ceil(x) if (x - floor_val) >= 0.5 else floor_val)

# ================================================================
# ROOM CLASSIFICATION & SORT
# ================================================================
WET_ROOM_KEYWORDS = ["kitchen", "bathroom", "wetroom", "wet room", "wc", "w/c", "toilet", "utility"]
WET_ROOM_PRIORITY = {
    "kitchen": 0, "utility": 1, "bathroom": 2,
    "wetroom": 3, "wet room": 3, "wc": 4, "w/c": 4, "toilet": 5,
}

def is_wet_room(room_name): return any(kw in room_name.lower() for kw in WET_ROOM_KEYWORDS)

def is_wc_or_utility(room_name):
    return any(kw in room_name.lower().strip() for kw in ["wc", "w/c", "toilet", "utility"])

def wet_room_sort_key(room_name):
    name = room_name.lower()
    for kw, priority in WET_ROOM_PRIORITY.items():
        if kw in name:
            return (0, priority)
    if any(kw in name for kw in ["living", "lounge", "dining", "reception"]): return (1, 0)
    if "bedroom" in name: return (2, 0)
    return (3, 0)

# ================================================================
# PDF HELPERS
# ================================================================
def read_pdf_text(path):
    with fitz.open(path) as doc:
        return "\n".join(page.get_text() for page in doc)

def read_pdf_page_lines(path, page_num):
    with fitz.open(path) as doc:
        if page_num < 0 or page_num >= len(doc):
            return []
        return (doc[page_num].get_text() or "").splitlines()

# ================================================================
# ADDRESS EXTRACTION
# ================================================================
def extract_property_address_from_cr(path):
    lines = [l.strip() for l in read_pdf_page_lines(path, 0) if l.strip()]
    addr_lines = []
    for i, line in enumerate(lines):
        if "Property Address" in line:
            for l in lines[i + 1: i + 7]:
                if any(x in l.lower() for x in ["assessment", "reference", "date"]):
                    break
                addr_lines.append(l)
            break
    addr = " ".join(addr_lines)
    addr = re.sub(r"\b([A-Z]{1,2}\d{1,2}[A-Z]?)\s+(\d[A-Z]{2})\b", r"\1 \2", addr, flags=re.I)
    return re.sub(r"\s{2,}", " ", addr).strip()

def extract_postcode(text):
    m = re.search(r"[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}", str(text), re.I)
    return m.group(0).upper() if m else ""

def extract_house_number(address):
    if pd.isna(address): return None
    address = str(address).strip()
    m = re.search(r"(?i)(?:flat|apartment)\s+([A-Za-z0-9]{1,5})", address)
    if m: return m.group(1)
    m = re.match(r"^([A-Za-z]?\d+[A-Za-z]?)\b", address)
    if m: return m.group(1)
    m = re.search(r"\b(\d+[A-Za-z]?)\b", address)
    return m.group(1) if m else None

def extract_street_number(address, house_number):
    if not address or not house_number: return None
    remaining = re.sub(rf"\b{re.escape(str(house_number))}\b", "", address, count=1, flags=re.I)
    m = re.search(r"\b(\d+[A-Za-z]?)\b", remaining)
    return m.group(1) if m else None

def normalize_text(x):
    if pd.isna(x): return ""
    return re.sub(r"[^a-z0-9]", "", str(x).lower())

def normalize_postcode(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+", "", str(x).lower())

# ================================================================
# TENURE NORMALISATION
# ================================================================
def normalise_tenure(raw_tenure):
    if not raw_tenure: return raw_tenure
    cleaned = raw_tenure.strip()
    if re.search(r"rent(?:al|ed)\s*\(?\s*social\s*\)?", cleaned, re.I):
        return "Social Housing"
    return cleaned

# ================================================================
# EXTRACT PROPERTY DETACHMENT FROM CR PDF PAGE 3
# ================================================================
def extract_property_detachment_from_cr(path):
    lines = read_pdf_page_lines(path, 2)
    for line in lines:
        stripped = line.strip()
        if stripped.lower().startswith("property detachment"):
            value = re.sub(r"(?i)property\s+detachment\s*", "", stripped).strip()
            if value: return value
    return ""

# ================================================================
# SMARTSHEET
# ================================================================
def fetch_smartsheet_df(ss_token, sheet_id):
    url = f"https://api.smartsheet.eu/2.0/sheets/{sheet_id}"
    headers = {"Authorization": f"Bearer {ss_token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    sheet = r.json()
    col_map = {c["id"]: c["title"] for c in sheet["columns"]}
    rows = []
    for row in sheet["rows"]:
        d = {col_map[cell["columnId"]]: cell.get("displayValue", "") for cell in row["cells"]}
        rows.append(d)
    return pd.DataFrame(rows)

def fetch_client_spec_df(ss_token, second_sheet_id):
    url = f"https://api.smartsheet.eu/2.0/sheets/{second_sheet_id}"
    headers = {"Authorization": f"Bearer {ss_token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    sheet2 = resp.json()
    col_map2 = {c["id"]: c["title"] for c in sheet2["columns"]}
    rows2 = []
    for r2 in sheet2["rows"]:
        rows2.append({col_map2.get(c["columnId"], ""): c.get("displayValue", c.get("value", "")) for c in r2["cells"]})
    return pd.DataFrame(rows2)

# ================================================================
# COORDINATES
# ================================================================
def dms_to_dd(deg, minutes, sec, direction):
    dd = float(deg) + float(minutes) / 60 + float(sec) / 3600
    return round(-dd if direction in ["S", "W"] else dd, 5)

def extract_lat_lon_from_cr(path):
    text = read_pdf_text(path)
    lat_match = re.search(r"Latitude[: ]\s*(\d+)[° ]\s*(\d+)'?\s*(\d+)''?\s*([NSEW])", text)
    lon_match = re.search(r"Longitude[: ]\s*(\d+)[° ]\s*(\d+)'?\s*(\d+)''?\s*([NSEW])", text)
    lat = dms_to_dd(*lat_match.groups()) if lat_match else ""
    lon = dms_to_dd(*lon_match.groups()) if lon_match else ""
    return lat, lon

# ================================================================
# EXTRACTION HELPERS
# ================================================================
def safe_search(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""

def extract_date_built_from_sn(path):
    lines = read_pdf_page_lines(path, 0)
    for i, line in enumerate(lines):
        if "Date Built" in line:
            for l in lines[i: i + 6]:
                m = re.search(r"(\d{4})\s*[-–]\s*(\d{4})", l)
                if m: return f"{m.group(1)}-{m.group(2)}"
            m2 = re.search(r"(\d{4})\s*[-–]\s*(\d{4})", read_pdf_text(path))
            if m2: return f"{m2.group(1)} - {m2.group(2)}"
            break
    return ""

def extract_num_storeys_from_sn(path):
    lines = read_pdf_page_lines(path, 0)
    for i, line in enumerate(lines):
        if "Number of" in line and "Storeys" in " ".join(lines[i: i + 3]):
            for j in range(i + 1, min(i + 6, len(lines))):
                m = re.search(r"\b(\d+)\b", lines[j])
                if m: return m.group(1)
            break
    return ""

def extract_exposure_zone_from_cr(path):
    for line in read_pdf_page_lines(path, 2):
        if "Exposure" in line and "Zone" in line:
            m = re.search(r"Zone\s*([A-Za-z]+)", line)
            if m: return m.group(1).strip()
    return ""

def extract_vulnerable_flag_from_cr(path):
    with fitz.open(path) as doc:
        last_page = len(doc) - 1
    lines = read_pdf_page_lines(path, last_page)
    questions = [
        "How many are children under the age of 18?",
        "How many of a pensionable age?",
        "How many with disabilities?"
    ]
    values = []
    for line in lines:
        s = line.strip()
        for q in questions:
            if s.startswith(q):
                part = s.split("?", 1)[1].strip()
                values.append(int(part) if part.isdigit() else 0)
    return "Yes" if any(v > 0 for v in values) else "No"

def extract_constraints_and_construction_from_cr(path):
    text = "\n".join(read_pdf_page_lines(path, 2))
    m_con  = re.search(r"Property\s+Constraints\s*([^\n\r]+)", text)
    m_cons = re.search(r"Property\s+Construction\s*([^\n\r]+)", text)
    constraints_val   = m_con.group(1).strip() if m_con else ""
    construction_flag = "Yes" if m_cons and m_cons.group(1).strip().lower().startswith("traditional") else "No"
    return constraints_val, construction_flag

def get_building_orientation_direction(path):
    text = "\n".join(read_pdf_page_lines(path, 2))
    m = re.search(r"Building\s+Orientation\s*\(degrees\)\s*([0-9]+(?:\.[0-9]+)?)", text)
    if not m: return ""
    angle = float(m.group(1))
    table = [
        (15, "North"), (45, "North East"), (75, "East"), (105, "East"),
        (135, "South East"), (165, "South"), (195, "South"), (225, "South West"),
        (255, "West"), (285, "West"), (315, "North West"), (345, "North"), (360, "North")
    ]
    for heading, direction in table:
        if angle <= heading: return direction
    return ""

def extract_highest_bedroom_from_contents(cr_page2_lines):
    highest_bedroom = 0
    found_plain_bedroom = False
    for line in cr_page2_lines:
        m = re.search(r"\bBedroom\s+(\d+)\b", line.strip(), flags=re.IGNORECASE)
        if m:
            highest_bedroom = max(highest_bedroom, int(m.group(1)))
        elif re.search(r"\bBedroom\b", line.strip(), flags=re.IGNORECASE):
            found_plain_bedroom = True
    if highest_bedroom == 0 and found_plain_bedroom:
        highest_bedroom = 1
    return highest_bedroom

# ================================================================
# ROOM EXTRACTION FROM CONTENTS PAGE
# ================================================================
_NON_ROOM_EXACT = {
    "contents", "property details", "property summary", "occupancy",
    "elevations", "elevation", "significance survey", "additional information",
    "annex", "appendix", "notes", "floor plan", "floorplan",
}
_NON_ROOM_SUBSTRINGS = [
    "element no", "significance", "elevation", "floorplan", "floor plan",
    "property detail", "property summary", "additional info", "annex",
    "appendix", "occupancy",
]
_END_SECTION_KEYWORDS = [
    "occupancy", "significance", "elevation", "additional", "annex",
    "appendix", "summary", "notes",
]
_ROOM_INDICATOR_KEYWORDS = [
    "kitchen", "bathroom", "bedroom", "living", "lounge", "dining",
    "hall", "hallway", "wc", "toilet", "utility", "wetroom", "wet room",
    "study", "office", "conservatory", "landing", "store", "cupboard",
    "reception", "playroom", "cellar", "loft", "attic", "en suite",
    "en-suite", "ensuite", "cloakroom",
]
_PAGE_NOISE_PATTERN = re.compile(r"^page\s+\d+(\s+of\s+\d+)?$", re.IGNORECASE)

def _is_non_room(text):
    t = text.strip().lower()
    if _PAGE_NOISE_PATTERN.match(t): return True
    if t in _NON_ROOM_EXACT: return True
    return any(sub in t for sub in _NON_ROOM_SUBSTRINGS)

def _is_end_section(text):
    t = text.strip().lower()
    return any(kw in t for kw in _END_SECTION_KEYWORDS)

def _looks_like_room(text):
    t = text.strip().lower()
    return any(kw in t for kw in _ROOM_INDICATOR_KEYWORDS)

def extract_rooms_from_contents(cr_pdf_path):
    raw_lines = read_pdf_page_lines(cr_pdf_path, 1)
    cleaned = []
    for line in raw_lines:
        s = line.strip()
        if not s or _PAGE_NOISE_PATTERN.match(s): continue
        cleaned.append(s)
    if not cleaned: return []

    pairs = []
    i = 0
    while i < len(cleaned):
        token = cleaned[i]
        if token.isdigit():
            i += 1
            continue
        if i + 1 < len(cleaned) and cleaned[i + 1].isdigit():
            pairs.append((token, int(cleaned[i + 1])))
            i += 2
        else:
            pairs.append((token, -1))
            i += 1
    if not pairs: return []

    start_idx = None
    for idx, (name, _) in enumerate(pairs):
        if _is_end_section(name): break
        if _looks_like_room(name) and not _is_non_room(name):
            start_idx = idx
            break

    if start_idx is None:
        rooms_raw = [(name, pg) for name, pg in pairs[2:] if not _is_non_room(name) and not _is_end_section(name)]
    else:
        rooms_raw = []
        for name, pg in pairs[start_idx:]:
            if _is_end_section(name): break
            if _is_non_room(name): continue
            rooms_raw.append((name, pg))

    seen, dedup = set(), []
    for name, pg in rooms_raw:
        key = name.strip().lower()
        if key not in seen:
            seen.add(key)
            dedup.append((name, pg))
    return dedup

# ================================================================
# ROOM EXTRACTION QUESTIONS
# ================================================================
Q_CONDENSATION = "Is there any evidence of condensation, damp or mould growth in this room?"
Q_BG_AREA      = "Background Ventilation Area (mm2)"
Q_TRICKLE      = "Do they have trickle vents?"
Q_UNDERCUTS    = "Are the undercuts on all the internal doors"
Q_WINDOWS      = "Does the room have any windows?"
Q_FANS         = "Does the room have fans fitted?"

def extract_yes_no_strict(lines, question):
    for line in lines:
        if line.startswith(question):
            parts = line.split("?", 1)
            if len(parts) == 2 and parts[1].strip() in ("Yes", "No"):
                return parts[1].strip()
            return None
    return None

def extract_mm2_strict(lines):
    for line in lines:
        if line.startswith(Q_BG_AREA):
            tail = line.replace(Q_BG_AREA, "").strip()
            return tail if tail.isdigit() else None
    return None

def clean_lines(lines):
    return [l.strip() for l in lines if l.strip() and not _PAGE_NOISE_PATTERN.match(l.strip().lower())]

def extract_room_data(cr_pdf_path, rooms_with_pages, room_name):
    empty = {"condensation": None, "background": None, "trickle": None,
             "undercuts": None, "windows": None, "fans": None}
    target_idx = None
    for i, (name, pg) in enumerate(rooms_with_pages):
        if name == room_name:
            target_idx = i
            break
    if target_idx is None: return empty
    _, start_page_no = rooms_with_pages[target_idx]
    if start_page_no == -1: return empty

    next_page_no = None
    for i in range(target_idx + 1, len(rooms_with_pages)):
        _, pg = rooms_with_pages[i]
        if pg != -1:
            next_page_no = pg
            break

    start_idx = max(0, start_page_no - 1)
    end_idx   = (next_page_no - 2) if next_page_no is not None else (start_idx + 3)
    end_idx   = max(start_idx, end_idx)

    all_lines = []
    try:
        with fitz.open(cr_pdf_path) as doc:
            total_pages = len(doc)
        for p in range(start_idx, min(end_idx + 1, total_pages)):
            all_lines.extend(read_pdf_page_lines(cr_pdf_path, p))
    except Exception:
        return empty

    lines = clean_lines(all_lines)
    return {
        "condensation": extract_yes_no_strict(lines, Q_CONDENSATION),
        "background":   extract_mm2_strict(lines),
        "trickle":      extract_yes_no_strict(lines, Q_TRICKLE),
        "undercuts":    extract_yes_no_strict(lines, Q_UNDERCUTS),
        "windows":      extract_yes_no_strict(lines, Q_WINDOWS),
        "fans":         extract_yes_no_strict(lines, Q_FANS),
    }

# ================================================================
# FLOW RATE HELPERS
# ================================================================
def get_min_lps_for_room(room_name):
    name = room_name.lower()
    if "kitchen" in name:                                      return "60 l/s minimum"
    if "bathroom" in name or "wetroom" in name:                return "15 l/s minimum"
    if "wc" in name or "toilet" in name or "w/c" in name:     return "6 l/s minimum"
    if "utility" in name:                                      return "30 l/s minimum"
    return "N/A"

def get_min_lps_for_room_q(room_name):
    name = room_name.lower()
    if "kitchen" in name:                                      return "13 l/s minimum"
    if "bathroom" in name or "wetroom" in name:                return "8 l/s minimum"
    if "wc" in name or "toilet" in name or "w/c" in name:     return "6 l/s minimum"
    if "utility" in name:                                      return "6 l/s minimum"
    return "N/A"

def extract_lps_number(value):
    if not value: return 0
    if isinstance(value, str) and "l/s" in value:
        try:    return int(value.split()[0])
        except: return 0
    return 0

# ================================================================
# EXCEL HELPERS
# ================================================================
def find_row_by_label(ws, label, col="A"):
    for r in range(1, ws.max_row + 1):
        val = ws[f"{col}{r}"].value
        if isinstance(val, str) and val.strip() == label:
            return r
    raise ValueError(f"Row with label '{label}' not found")

def parse_area_m2(value):
    if value is None: return 0.0
    if isinstance(value, (int, float)): return float(value)
    if isinstance(value, str):
        num = "".join(ch for ch in value if ch.isdigit() or ch == ".")
        return float(num) if num else 0.0
    return 0.0

def get_master_cell(ws, cell_ref):
    for merged in ws.merged_cells.ranges:
        if cell_ref in merged:
            return merged.coord.split(":")[0]
    return cell_ref

def unmerge_in_range(ws, start_row):
    affected = []
    for m in list(ws.merged_cells.ranges):
        c1, r1, c2, r2 = range_boundaries(str(m))
        if r1 >= start_row:
            affected.append((c1, r1, c2, r2))
            ws.unmerge_cells(str(m))
    return affected

def remerge(ws, ranges):
    for c1, r1, c2, r2 in ranges:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

# ================================================================
# WRITE ROOM ROW
# ================================================================
def write_room_row(ws, row, d, room):
    def v(x): return x if x is not None else "Not Extracted"
    room_lower = room.lower()
    bg_value   = d["background"]
    fans_value = v(d["fans"])
    wet_room_keys = ["kitchen", "bathroom", "wetroom", "wet room", "wc", "w/c", "toilet", "utility"]
    is_wet = any(w in room_lower for w in wet_room_keys)

    if is_wc_or_utility(room):
        ws[f"B{row}"].value = "No minimum"
    else:
        ws[f"B{row}"].value = v(d["condensation"])

    if bg_value is None:
        ws[f"C{row}"].value = "No"
    elif "bathroom" in room_lower or "wetroom" in room_lower:
        ws[f"C{row}"].value = "4000mm2"
    else:
        ws[f"C{row}"].value = "8000mm2"

    ws[f"D{row}"].value = "No" if bg_value is None else f"{bg_value}mm2"
    ws[f"E{row}"].value = v(d["undercuts"])
    ws[f"F{row}"].value = v(d["windows"])

    if is_wet:
        ws[f"G{row}"].value = "IEV Present" if fans_value == "Yes" else "None"
    else:
        ws[f"G{row}"].value = "N/A"

    ws[f"H{row}"].value = get_min_lps_for_room(room)

    cond_val = d["condensation"]
    ws[f"K{row}"].value = (
        "No" if cond_val == "No" else
        "Clean" if cond_val == "Yes" else
        "Not Extracted"
    )
    ws[f"L{row}"].value = "N/A" if is_wet else "4000mm2"

    col_c_val     = ws[f"C{row}"].value
    undercuts_val = ws[f"E{row}"].value
    if is_wet:
        ws[f"M{row}"].value = "Sealed" if fans_value == "Yes" else "No Works"
    else:
        ws[f"M{row}"].value = "Required" if col_c_val == "No" and undercuts_val == "Yes" else "No Works"

    ws[f"N{row}"].value = (
        "Required" if d["undercuts"] == "No" else
        "No Works" if d["undercuts"] == "Yes" else
        "Not Extracted"
    )
    ws[f"O{row}"].value = (
        "Required" if d["windows"] == "No" else
        "No Works" if d["windows"] == "Yes" else
        "Not Extracted"
    )

    if is_wet:
        ws[f"P{row}"].value = "Upgrade dMEV" if fans_value == "Yes" else "Install dMEV"
    else:
        ws[f"P{row}"].value = "N/A"

    ws[f"Q{row}"].value = get_min_lps_for_room_q(room)
    ws[f"R{row}"].value = "Information Unavailable"

# ================================================================
# MAIN ROOM INSERTION (Sheet 11.0)
# ================================================================
def insert_rooms_with_template(cr_pdf_path, workbook, sheet_name="11.0",
                                insert_at_row=5, template_row=4,
                                room_col_left="A", room_col_right="J",
                                tfa=None, highest_bedroom=0, log=None):
    def _log(msg):
        if log is not None: log(msg)

    ws = workbook[sheet_name]
    rooms_with_pages = extract_rooms_from_contents(cr_pdf_path)
    _log(f"📋 Rooms extracted ({len(rooms_with_pages)}): {[r for r, _ in rooms_with_pages]}")

    if not rooms_with_pages:
        raise ValueError("No rooms found in Contents page.")

    rooms_with_pages = sorted(rooms_with_pages, key=lambda rp: wet_room_sort_key(rp[0]))
    _log(f"📋 Rooms after sort: {[r for r, _ in rooms_with_pages]}")
    room_names = [r for r, _ in rooms_with_pages]

    original_heights = {r: d.height for r, d in ws.row_dimensions.items() if d.height is not None}
    affected_merges  = unmerge_in_range(ws, insert_at_row)
    ws.insert_rows(insert_at_row, amount=len(room_names))

    for r, h in original_heights.items():
        ws.row_dimensions[r + len(room_names) if r >= insert_at_row else r].height = h

    tpl_height = ws.row_dimensions[template_row].height
    for i in range(len(room_names)):
        ws.row_dimensions[insert_at_row + i].height = tpl_height
        for c in range(1, ws.max_column + 1):
            ws.cell(insert_at_row + i, c)._style = copy(ws.cell(template_row, c)._style)

    for i, room in enumerate(room_names):
        row  = insert_at_row + i
        ws[f"{room_col_left}{row}"].value = room
        ws[f"{room_col_right}{row}"].value = room
        data = extract_room_data(cr_pdf_path, rooms_with_pages, room)
        write_room_row(ws, row, data, room)
        found = " ".join(f"{k}={v}" for k, v in data.items() if v is not None)
        _log(f"  ✓ {room}: {found or 'No questions found'}")

    ws4       = workbook["4.0"]
    room_start = insert_at_row
    room_end   = insert_at_row + len(room_names) - 1

    ws4["H5"].value = (
        "Yes" if any(ws[f"P{r}"].value == "Upgrade dMEV" for r in range(room_start, room_end + 1))
        else "No"
    )
    ws4["B8"].value = (
        "Yes" if any(ws[f"B{r}"].value == "Yes" for r in range(room_start, room_end + 1))
        else "No"
    )

    total_lps_h = sum(extract_lps_number(ws[f"H{insert_at_row + i}"].value) for i in range(len(room_names)))
    total_lps_q = sum(extract_lps_number(ws[f"Q{insert_at_row + i}"].value) for i in range(len(room_names)))
    total_row   = insert_at_row + len(room_names) + 1
    ws[f"H{total_row}"].value = f"{total_lps_h} l/s"
    ws[f"Q{total_row}"].value = f"{total_lps_q} l/s"

    floor_area_row = find_row_by_label(ws, "Whole building ventilation rate (l/s) - floor area (m2)", col="A")
    ws[f"G{floor_area_row}"].value = tfa
    ws[f"P{floor_area_row}"].value = tfa
    floor_rate = custom_round(parse_area_m2(tfa) * 0.3)
    ws[f"H{floor_area_row}"].value = f"{floor_rate} l/s"
    ws[f"Q{floor_area_row}"].value = f"{floor_rate} l/s"

    bedroom_row  = find_row_by_label(ws, "Whole building ventilation rate (l/s) - bedrooms", col="A")
    ws[f"G{bedroom_row}"].value = f"{highest_bedroom} bed"
    ws[f"P{bedroom_row}"].value = f"{highest_bedroom} bed"
    bed_rate_raw = {1: 19, 2: 25, 3: 31, 4: 37, 5: 43}.get(highest_bedroom, 0)
    bed_rate     = custom_round(bed_rate_raw)
    ws[f"H{bedroom_row}"].value = f"{bed_rate} l/s"
    ws[f"Q{bedroom_row}"].value = f"{bed_rate} l/s"

    final_row    = find_row_by_label(ws, "Whole building ventilation rate - final", col="A")
    final_rate_h = custom_round(max(floor_rate, bed_rate, total_lps_h))
    final_rate_q = custom_round(max(floor_rate, bed_rate, total_lps_q))
    ws[f"H{final_row}"].value = f"{final_rate_h} l/s"
    ws[f"Q{final_row}"].value = f"{final_rate_q} l/s"

    remerge(ws, [(c1, r1 + len(room_names), c2, r2 + len(room_names)) for c1, r1, c2, r2 in affected_merges])
    return room_names

# ================================================================
# WRITE SHEET 4.0
# ================================================================
def write_sheet_4_base(ws, wb_data):
    d = wb_data
    ws[get_master_cell(ws, "B3")]  = d["dwelling_address"]
    ws[get_master_cell(ws, "B4")]  = d["property_id_value"]
    ws[get_master_cell(ws, "B5")]  = d["local_authority_name"]
    ws[get_master_cell(ws, "B6")]  = d["sap_current"]
    ws[get_master_cell(ws, "E3")]  = d["sap_current"]
    ws[get_master_cell(ws, "E9")]  = d["sap_potential"]
    ws[get_master_cell(ws, "B7")]  = d["tenure"]
    ws[get_master_cell(ws, "B14")] = d["date_built"]
    ws[get_master_cell(ws, "B15")] = d["wall_type"]
    ws[get_master_cell(ws, "B16")] = d["property_type"]
    ws[get_master_cell(ws, "B17")] = d["num_storeys"]
    ws[get_master_cell(ws, "B18")] = d["highest_bedroom"] if d["highest_bedroom"] > 0 else ""
    ws[get_master_cell(ws, "B19")] = d["exposure_zone"]
    ws[get_master_cell(ws, "B20")] = d["property_construction_flag"]
    ws[get_master_cell(ws, "B21")] = d["conservation_flag"]
    ws[get_master_cell(ws, "B22")] = "No"
    ws[get_master_cell(ws, "B24")] = d["property_constraints_val"]
    ws[get_master_cell(ws, "B25")] = d["vulnerable_flag"]
    ws[get_master_cell(ws, "E4")]  = d["tfa_sqm"]
    ws[get_master_cell(ws, "E10")] = d["tfa_sqm"]
    ws[get_master_cell(ws, "E5")]  = d["extra1_last"]
    ws[get_master_cell(ws, "E6")]  = d["extra1"]
    ws[get_master_cell(ws, "E11")] = d["extra2_last"]
    ws[get_master_cell(ws, "E12")] = d["extra2"]
    conservation = d.get("conservation_flag", "No")
    if str(conservation).strip().lower() == "yes":
        for cell in ["B10", "B12"]: ws[get_master_cell(ws, cell)] = "TBC"
        for cell in ["B11", "B13"]: ws[get_master_cell(ws, cell)] = "Pending"
    else:
        for cell in ["B10", "B11", "B12", "B13"]: ws[get_master_cell(ws, cell)] = "No"

def write_sheet_4_climate_risk(ws, wb_data):
    d = wb_data
    ws[get_master_cell(ws, "E16")] = d["radon_risk"]
    ws[get_master_cell(ws, "E17")] = d["flood_risk"]
    ws[get_master_cell(ws, "E18")] = d["subs2030"]
    ws[get_master_cell(ws, "E19")] = d["subs2070"]
    ws[get_master_cell(ws, "E20")] = d["wdr_risk"]
    ws[get_master_cell(ws, "E21")] = d["rainfall_risk"]
    ws[get_master_cell(ws, "E22")] = d["wind_speed_risk"]
    ws[get_master_cell(ws, "E23")] = d["wildfire_risk"]

# ================================================================
# WRITE SHEET 10.0
# ================================================================
def write_sheet_10(workbook, b3_value, f3_value, log=None):
    if "10.0" not in workbook.sheetnames:
        if log: log("⚠️  Sheet '10.0' not found — skipping.")
        return
    ws = workbook["10.0"]
    ws[get_master_cell(ws, "B3")] = b3_value
    ws[get_master_cell(ws, "F3")] = f3_value
    if log: log(f"  ✓ Sheet 10.0: B3='{b3_value}', F3='{f3_value}'")

# ================================================================
# WRITE SHEET 7.0
# ================================================================
def write_sheet_7(ws, wb_data, df_specs, selected_measures, orientation_direction, ws4, spec_col_order=None, log=None):
    def _log(msg):
        if log: log(msg)

    measure_to_cell = {
        "Trickle Vents": "A2", "Door undercuts": "A4", "dMEV": "A3",
        "LI (B9)": "A5", "CWI": "A6", "EWI": "A7", "IWI": "A8",
        "UFI": "A9", "FRI": "A10", "RiR and/or Skeiling": "A11",
        "Windows": "A12", "Doors": "A13", "Boiler/FTCH": "A17",
        "ESH": "A18", "ASHP": "A19", "GSHP": "A20",
        "Solar Thermal": "A21", "SPV": "A22",
    }
    abbrev_map = {
        "Trickle Vents": "TV", "dMEV": "dMEV", "Door undercuts": "DRU",
        "LI (B9)": "LI", "CWI": "CWI", "EWI": "EWI", "IWI": "IWI",
        "UFI": "UFI", "FRI": "FRI", "RiR and/or Skeiling": "RiR/SKI",
        "Windows": "Windows", "Doors": "Doors", "Boiler/FTCH": "BLR",
        "ESH": "ESH", "ASHP": "ASHP", "GSHP": "GSHP",
        "Solar Thermal": "ST", "SPV": "SPV", "LEL": "LEL", "HWC": "HWC",
        "Draught-Proofing": "DPRF",
        "Upgrade boiler, same fuel": "UBLR",
        "Heat Recovery System for Mixer Showers": "HRSMS",
        "Communal PV": "SPV",
    }
    SEQUENCE_ORDER = [
        ["dMEV"], ["CWI"], ["EWI"], ["IWI"], ["LI (B9)"], ["FRI"],
        ["RiR and/or Skeiling"], ["Windows", "Doors"], ["Boiler/FTCH"],
        ["ESH"], ["ASHP"], ["GSHP"], ["Solar Thermal"], ["HWC"], ["LEL"], ["SPV"],
    ]

    measure_seq_lookup = {}
    _seq = 1
    for group in SEQUENCE_ORDER:
        for name in group:
            measure_seq_lookup[name] = _seq
        _seq += 1
    next_unknown_seq = _seq

    # Default column priority if not provided from UI
    _default_col_order = [
        "Construction Design Specification Notes",
        "Client Specification",
        "Design Specification Notes",
    ]
    _col_order = spec_col_order if spec_col_order else _default_col_order

    def get_spec(abbr):
        if df_specs.empty or "Abbreviation" not in df_specs.columns:
            return ""
        spec_row = df_specs[df_specs["Abbreviation"] == abbr]
        if spec_row.empty:
            return ""
        r = spec_row.iloc[0]
        for col in _col_order:
            val = r.get(col, "")
            if val and str(val).strip():
                return str(val).strip()
        return ""

    def find_empty_d_row():
        for r in range(17, 23):
            if ws[get_master_cell(ws, f"D{r}")].value in (None, ""):
                return r
        for r in range(2, 15):
            if ws[get_master_cell(ws, f"D{r}")].value in (None, ""):
                return r
        return None

    measures_list = [m.strip() for m in selected_measures.split(",") if m.strip()]
    spv_included = False

    for measure in measures_list:
        is_communal_pv = "communal pv" in measure.lower()
        if is_communal_pv:
            spv_included = True
            row_num = int(measure_to_cell["SPV"][1:])
            ws[get_master_cell(ws, f"D{row_num}")] = get_spec("SPV")
            c22_master = get_master_cell(ws, "C22")
            for merged in list(ws.merged_cells.ranges):
                c1, r1, c2, r2 = range_boundaries(str(merged))
                if r1 <= 22 <= r2 and c1 <= 3 <= c2:
                    ws.unmerge_cells(str(merged))
                    break
            ws[c22_master] = "Communal PV"
            seq = measure_seq_lookup.get("SPV", next_unknown_seq)
            ws[get_master_cell(ws, f"G{row_num}")] = seq
            _log(f"  ✓ Communal PV → row {row_num} (SPV spec), C22='Communal PV', seq {seq}")
            continue

        matched_key = None
        for key in measure_to_cell:
            if key.lower() in measure.lower():
                matched_key = key
                break

        if matched_key == "SPV":
            spv_included = True

        if matched_key is not None:
            row_num = int(measure_to_cell[matched_key][1:])
            abbr    = abbrev_map.get(matched_key, "")
            ws[get_master_cell(ws, f"D{row_num}")] = get_spec(abbr)
            seq = measure_seq_lookup.get(matched_key, next_unknown_seq)
            ws[get_master_cell(ws, f"G{row_num}")] = seq
            _log(f"  ✓ Known measure '{matched_key}' → row {row_num}, seq {seq}")
        else:
            row_num = find_empty_d_row()
            if row_num is None:
                _log(f"  ⚠️  No empty row found for '{measure}' — skipped")
                continue
            b_master = get_master_cell(ws, f"B{row_num}")
            for merged in list(ws.merged_cells.ranges):
                c1, r1, c2, r2 = range_boundaries(str(merged))
                if r1 <= row_num <= r2 and c1 <= 2 <= c2:
                    ws.unmerge_cells(str(merged))
                    break
            ws[b_master] = measure
            abbr     = abbrev_map.get(measure, measure[:4].upper())
            c_master = get_master_cell(ws, f"C{row_num}")
            for merged in list(ws.merged_cells.ranges):
                c1, r1, c2, r2 = range_boundaries(str(merged))
                if r1 <= row_num <= r2 and c1 <= 3 <= c2:
                    ws.unmerge_cells(str(merged))
                    break
            ws[c_master] = abbr
            ws[get_master_cell(ws, f"D{row_num}")] = get_spec(abbr)
            ws[get_master_cell(ws, f"G{row_num}")] = next_unknown_seq
            ws[get_master_cell(ws, f"H{row_num}")] = "Minor"
            _log(f"  ℹ️  Unknown measure '{measure}' → row {row_num}, seq {next_unknown_seq}")
            next_unknown_seq += 1

    if spv_included:
        if orientation_direction:
            ws[get_master_cell(ws, "F22")] = orientation_direction
        ws[get_master_cell(ws, "F23")] = "None or little"
        _log(f"  ✓ PV included → F22='{orientation_direction}', F23='None or little'")
    else:
        ws[get_master_cell(ws, "F22")] = None
        ws[get_master_cell(ws, "F23")] = None

    condition_met = any(
        ws[get_master_cell(ws, cell)].value not in (None, False, "", "FALSE")
        for cell in ["A5", "A6", "A7", "A8", "A10", "A12", "A13"]
    )
    ws4["H3"].value = "Yes" if condition_met else "No"
    ws4["H4"].value = "Yes" if condition_met else "No"

    seq_number = 1
    for group in SEQUENCE_ORDER:
        group_has_content = False
        for measure_name in group:
            if measure_name not in measure_to_cell: continue
            row_num = int(measure_to_cell[measure_name][1:])
            if ws[get_master_cell(ws, f"D{row_num}")].value not in (None, ""):
                ws[get_master_cell(ws, f"G{row_num}")] = seq_number
                group_has_content = True
        if group_has_content:
            seq_number += 1

# ================================================================
# CORE PROCESSING FUNCTION (replaces the main loop)
# ================================================================
def process_property(
    cr_path, sn_path, epr_path,
    template_bytes,
    ss_token, sheet_id, second_sheet_id,
    sheet10_b3, sheet10_f3,
    prop_name,
    duckdb_path,
    spec_col_order,
    log
):
    start_time = time.time()

    # -- Read PDFs ---
    sn_text  = read_pdf_text(sn_path)
    epr_text = read_pdf_text(epr_path)

    # -- Address matching --
    dwelling_address = extract_property_address_from_cr(cr_path)
    log(f"📍 Extracted address: {dwelling_address}")

    df1 = pd.DataFrame({"Address": [dwelling_address]})
    df1["Postcode"]      = df1["Address"].apply(extract_postcode)
    df1["house_number"]  = df1["Address"].apply(extract_house_number)
    df1["street_number"] = df1.apply(lambda r: extract_street_number(r["Address"], r["house_number"]), axis=1)
    df1["n_postcode"]    = df1["Postcode"].apply(normalize_postcode)
    df1["n_house"]       = df1["house_number"].apply(normalize_text)
    df1["n_street"]      = df1["street_number"].apply(normalize_text)

    log("🔗 Fetching Smartsheet IOE data...")
    df2 = fetch_smartsheet_df(ss_token, sheet_id)
    df2.rename(columns={
        "Full Address":    "Address",
        "Postal Code":     "PostalCode",
        "Select Measures": "Selected_Measures"
    }, inplace=True)

    df2["Address"] = df2.apply(
        lambda r: r["Address"] if extract_postcode(r["Address"]) else f'{r["Address"]}, {r["PostalCode"]}',
        axis=1
    )
    df2["Postcode"]      = df2["Address"].apply(extract_postcode)
    df2["house_number"]  = df2["Address"].apply(extract_house_number)
    df2["street_number"] = df2.apply(lambda r: extract_street_number(r["Address"], r["house_number"]), axis=1)
    df2["n_postcode"]    = df2["Postcode"].apply(normalize_postcode)
    df2["n_house"]       = df2["house_number"].apply(normalize_text)
    df2["n_street"]      = df2["street_number"].apply(normalize_text)

    matched_row = None
    for _, row in df2[df2["n_postcode"] == df1.loc[0, "n_postcode"]].iterrows():
        if row["n_house"] == df1.loc[0, "n_house"] and row["n_street"] == df1.loc[0, "n_street"]:
            matched_row = row
            break

    if matched_row is None:
        log(f"⚠️  No address match found in Smartsheet — continuing with blank Smartsheet fields")
        matched_address = selected_measures = property_id_value = sap_potential = sap_current = ""
    else:
        log(f"✅ ADDRESS MATCH found in Smartsheet")
        matched_address   = matched_row["Address"]
        selected_measures = matched_row["Selected_Measures"]
        property_id_value = matched_row.get("Property ID", "")
        sap_potential     = matched_row.get("Proposed Post SAP Score", "")
        sap_current       = matched_row.get("Pre SAP Score (Full SAP)", "")

    # -- Coordinates --
    latitude, longitude = extract_lat_lon_from_cr(cr_path)
    if latitude == "" or longitude == "":
        raise ValueError("Latitude/Longitude not found in CR PDF.")
    log(f"🌍 Coordinates: {latitude}, {longitude}")

    # -- EPR additional ratings --
    epr_lines = epr_text.splitlines()
    nums, capture = [], False
    for line in epr_lines:
        if "Additional ratings for your home" in line:
            capture = True
            continue
        if capture:
            lc = line.strip()
            if not lc or "office use only" in lc.lower(): break
            m = re.search(r"\d+(\.\d+)?", lc)
            if m: nums.append(m.group(0))
            if len(nums) >= 10: break

    if len(nums) >= 10:
        extra1      = custom_round(float(nums[0]))
        extra2      = custom_round(float(nums[1]))
        extra1_last = custom_round(float(nums[-2]))
        extra2_last = custom_round(float(nums[-1]))
    else:
        extra1 = extra2 = extra1_last = extra2_last = ""

    # -- SN page 1 data --
    sn_page1_text = "\n".join(read_pdf_page_lines(sn_path, 0))
    m_tenure      = re.search(r"Property\s+Tenure:\s*([^\n\r]+)", sn_page1_text)
    raw_tenure    = m_tenure.group(1).strip() if m_tenure else ""
    tenure        = normalise_tenure(raw_tenure)
    date_built    = extract_date_built_from_sn(sn_path)
    num_storeys   = extract_num_storeys_from_sn(sn_path)
    wall_type     = safe_search(r"Type\s+[A-Z]+\s+(.*)", sn_text)
    property_type = extract_property_detachment_from_cr(cr_path)
    exposure_zone = extract_exposure_zone_from_cr(cr_path)
    tfa           = safe_search(r"Total Floor Area\s*([\d\.]+\s*m²?)", epr_text)
    if tfa:        tfa = tfa.replace("m²", "m2")

    vulnerable_flag                                       = extract_vulnerable_flag_from_cr(cr_path)
    property_constraints_val, property_construction_flag  = extract_constraints_and_construction_from_cr(cr_path)
    orientation_direction                                 = get_building_orientation_direction(cr_path)
    highest_bedroom = extract_highest_bedroom_from_contents(read_pdf_page_lines(cr_path, 1))
    log(f"🏠 Bedrooms: {highest_bedroom} | TFA: {tfa} | Tenure: {tenure}")

    # -- Geo risks via DuckDB --
    log("🗺️  Querying geo risk database...")
    lon_f = float(longitude)
    lat_f = float(latitude)
    con = duckdb.connect(duckdb_path)
    con.execute("INSTALL spatial;")
    con.execute("LOAD spatial;")

    query = f"""
    WITH point AS (SELECT ST_Point({lon_f}, {lat_f}) AS geom),
    la AS (SELECT * FROM la_districts WHERE ST_Contains(geom, (SELECT geom FROM point)))
    SELECT
        CASE WHEN EXISTS (
            SELECT 1 FROM conservation WHERE ST_Contains(geom, (SELECT geom FROM point))
        ) THEN 'Yes' ELSE 'No' END AS conservation,
        (SELECT CASE CLASS_MAX
            WHEN 1 THEN 'Very Low' WHEN 2 THEN 'Low' WHEN 3 THEN 'Moderate'
            WHEN 4 THEN 'High' WHEN 5 THEN 'Very High' WHEN 6 THEN 'Extremely High'
            ELSE 'Unknown' END
         FROM radon WHERE ST_Contains(geom, (SELECT geom FROM point)) LIMIT 1) AS radon,
        (SELECT "flood-risk-level"
         FROM flood WHERE ST_Contains(geom, (SELECT geom FROM point))
         ORDER BY "flood-risk-level" DESC LIMIT 1) AS flood,
        (SELECT CLASS FROM subs2030 WHERE ST_Contains(geom, (SELECT geom FROM point)) LIMIT 1) AS subs2030,
        (SELECT CLASS FROM subs2070 WHERE ST_Contains(geom, (SELECT geom FROM point)) LIMIT 1) AS subs2070,
        (SELECT CASE
            WHEN AVG_WDR_baseline_Median <= 100 THEN 'Very Low'
            WHEN AVG_WDR_baseline_Median <= 200 THEN 'Low'
            WHEN AVG_WDR_baseline_Median <= 300 THEN 'Moderate'
            WHEN AVG_WDR_baseline_Median <= 400 THEN 'High'
            WHEN AVG_WDR_baseline_Median <= 500 THEN 'Very High'
            ELSE 'Extremely High' END
         FROM wdr WHERE ST_Contains(geom, (SELECT geom FROM point)) LIMIT 1) AS wdr,
        (SELECT CASE
            WHEN avg_pr < 60  THEN 'Very Low'    WHEN avg_pr < 80  THEN 'Low'
            WHEN avg_pr < 100 THEN 'Below Moderate' WHEN avg_pr < 120 THEN 'Moderate'
            WHEN avg_pr < 140 THEN 'Above Moderate' WHEN avg_pr < 160 THEN 'High'
            WHEN avg_pr < 180 THEN 'Very High'   WHEN avg_pr < 200 THEN 'Extremely High'
            WHEN avg_pr < 240 THEN 'Severe'      ELSE 'Exceptional' END
         FROM (
            SELECT (prJan+prFeb+prMar+prApr+prMay+prJun+prJul+prAug+prSep+prOct+prNov+prDec)/12.0 AS avg_pr, geom
            FROM rainfall
         ) r WHERE ST_Contains(geom, (SELECT geom FROM point)) LIMIT 1) AS rainfall,
        (SELECT CASE
            WHEN avg_ws < 4  THEN 'Very Low' WHEN avg_ws < 6  THEN 'Low'
            WHEN avg_ws < 8  THEN 'Moderate' WHEN avg_ws < 10 THEN 'High'
            ELSE 'Very High' END
         FROM (
            SELECT (ws_winter_baseline_median+ws_spring_baseline_median+ws_summer_baseline_median+ws_autumn_baseline_median)/4.0 AS avg_ws, geom
            FROM windspeed
         ) w WHERE ST_Contains(geom, (SELECT geom FROM point)) LIMIT 1) AS wind,
        (SELECT CASE
            WHEN w."Wildfire incidents (number)" / w."Area (square kilometres)" < 1.18  THEN 'Low'
            WHEN w."Wildfire incidents (number)" / w."Area (square kilometres)" < 4.88  THEN 'Moderate'
            WHEN w."Wildfire incidents (number)" / w."Area (square kilometres)" < 13.5  THEN 'High'
            ELSE 'Very High' END
         FROM wildfire w JOIN la ON la."name" = w."Local authority district" LIMIT 1) AS wildfire,
        (SELECT la."name" FROM la LIMIT 1) AS local_authority_name
    """
    result = con.execute(query).fetchdf().iloc[0].to_dict()
    con.close()

    conservation_flag    = result.get("conservation", "")
    radon_risk           = result.get("radon", "")
    flood_risk_raw       = result.get("flood")
    flood_risk           = "No" if (flood_risk_raw is None or pd.isna(flood_risk_raw) or str(flood_risk_raw).strip() == "") else flood_risk_raw
    subs2030             = result.get("subs2030", "")
    subs2070             = result.get("subs2070", "")
    wdr_risk             = result.get("wdr", "")
    rainfall_risk        = result.get("rainfall", "")
    wind_speed_risk      = result.get("wind", "")
    wildfire_risk        = result.get("wildfire", "")
    local_authority_name = result.get("local_authority_name", "")
    log(f"🏛️  Local Authority (DB): {local_authority_name} | Conservation: {conservation_flag}")

    # -- Fetch client spec sheet --
    log("🔗 Fetching Smartsheet Client Specification...")
    df_specs = fetch_client_spec_df(ss_token, second_sheet_id)

    # -- Bundle all extracted data --
    wb_data = {
        "dwelling_address":           dwelling_address,
        "property_id_value":          property_id_value,
        "local_authority_name":       sheet10_b3,   # from Sheet 10 B3 input
        "sap_current":                sap_current,
        "sap_potential":              sap_potential,
        "tenure":                     tenure,
        "date_built":                 date_built,
        "wall_type":                  wall_type,
        "property_type":              property_type,
        "num_storeys":                num_storeys,
        "highest_bedroom":            highest_bedroom,
        "exposure_zone":              exposure_zone,
        "property_construction_flag": property_construction_flag,
        "conservation_flag":          conservation_flag,
        "property_constraints_val":   property_constraints_val,
        "vulnerable_flag":            vulnerable_flag,
        "tfa_sqm":                    tfa.replace("m", "sqm") if tfa else "",
        "extra1":                     extra1,
        "extra2":                     extra2,
        "extra1_last":                extra1_last,
        "extra2_last":                extra2_last,
        "radon_risk":                 radon_risk,
        "flood_risk":                 flood_risk,
        "subs2030":                   subs2030,
        "subs2070":                   subs2070,
        "wdr_risk":                   wdr_risk,
        "rainfall_risk":              rainfall_risk,
        "wind_speed_risk":            wind_speed_risk,
        "wildfire_risk":              wildfire_risk,
    }

    # ── PRELIM OUTPUT — no climate risk ──────────────────────────
    log("📄 Building Prelim output...")
    wb_prelim  = load_workbook(io.BytesIO(template_bytes))
    ws4_prelim = wb_prelim["4.0"]
    write_sheet_4_base(ws4_prelim, wb_data)
    ws7_prelim = wb_prelim["7.0"]
    write_sheet_7(ws7_prelim, wb_data, df_specs, selected_measures, orientation_direction, ws4_prelim, spec_col_order=spec_col_order, log=log)
    write_sheet_10(wb_prelim, sheet10_b3, sheet10_f3, log=log)
    insert_rooms_with_template(
        cr_pdf_path=cr_path, workbook=wb_prelim, sheet_name="11.0",
        insert_at_row=5, template_row=4, room_col_left="A", room_col_right="J",
        tfa=tfa, highest_bedroom=highest_bedroom, log=log
    )
    prelim_buffer = io.BytesIO()
    wb_prelim.save(prelim_buffer)
    prelim_buffer.seek(0)
    log("✅ Prelim output ready")

    # ── CONSTRUCTION OUTPUT — with climate risk ───────────────────
    log("🏗️  Building Construction output...")
    wb_construction  = load_workbook(io.BytesIO(template_bytes))
    ws4_construction = wb_construction["4.0"]
    write_sheet_4_base(ws4_construction, wb_data)
    write_sheet_4_climate_risk(ws4_construction, wb_data)
    ws7_construction = wb_construction["7.0"]
    write_sheet_7(ws7_construction, wb_data, df_specs, selected_measures, orientation_direction, ws4_construction, spec_col_order=spec_col_order, log=log)
    write_sheet_10(wb_construction, sheet10_b3, sheet10_f3, log=log)
    insert_rooms_with_template(
        cr_pdf_path=cr_path, workbook=wb_construction, sheet_name="11.0",
        insert_at_row=5, template_row=4, room_col_left="A", room_col_right="J",
        tfa=tfa, highest_bedroom=highest_bedroom, log=log
    )
    construction_buffer = io.BytesIO()
    wb_construction.save(construction_buffer)
    construction_buffer.seek(0)
    log("✅ Construction output ready")

    elapsed = time.time() - start_time
    log(f"⏱️  Completed in {elapsed:.1f}s")

    return prelim_buffer, construction_buffer, prop_name, dwelling_address

# ================================================================
# STREAMLIT UI
# ================================================================

# Hero header
st.markdown('<div class="hero-title">Retrofit Design Exporter</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-sub">Elmhurst CR · SN · EPR → Excel Design Pack</div>', unsafe_allow_html=True)

# ── SIDEBAR — Credentials & Config ──────────────────────────────
with st.sidebar:
    st.markdown('<div class="card-label">⚙ Configuration</div>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-label">Smartsheet</div>', unsafe_allow_html=True)
    ss_token = st.text_input("API Token", type="password", placeholder="rWR95nWpa9TMOOs...", key="ss_token")
    sheet_id = st.text_input("IOE Sheet ID", placeholder="74429503768452", key="sheet_id")
    second_sheet_id = st.text_input("Client Spec Sheet ID", placeholder="1408360446560132", key="second_sheet_id")

    st.markdown("---")
    st.markdown('<div class="card-label">Sheet 10 — Local Authority</div>', unsafe_allow_html=True)
    sheet10_b3 = st.text_input("Local Authority Name (B3)", placeholder="e.g. Croydon Council", key="b3")
    sheet10_f3 = st.text_input("Year (F3)", placeholder="e.g. 2026", key="f3")

    st.markdown("---")
    st.markdown('<div class="card-label">Database</div>', unsafe_allow_html=True)
    duckdb_path = st.text_input("DuckDB File Path", placeholder=r"data/uk_risk.duckdb", key="duckdb")

    st.markdown("---")
    st.markdown('<div class="card-label">📑 Client Spec — Column Priority</div>', unsafe_allow_html=True)
    st.caption("Drag to reorder — top column is checked first. First non-empty value wins.")

    ALL_SPEC_COLS = [
        "Construction Design Specification Notes",
        "Client Specification",
        "Design Specification Notes",
    ]

    # Let user pick which columns to include and in what order via multiselect
    # (Streamlit doesn't have drag-to-reorder natively, so we use ordered multiselect)
    spec_col_order = st.multiselect(
        "Column lookup order (top = highest priority)",
        options=ALL_SPEC_COLS,
        default=ALL_SPEC_COLS,
        help="Select columns in the priority order you want. Deselect any column to skip it entirely.",
        key="spec_col_order",
    )
    if not spec_col_order:
        st.warning("⚠️ No spec columns selected — spec cells will be blank.")
        spec_col_order = ALL_SPEC_COLS  # fallback to default

    # Show current active priority as numbered badges
    for i, col in enumerate(spec_col_order, 1):
        short = col.replace(" Specification", " Spec").replace("Construction Design ", "CD ")
        st.markdown(
            f'<div style="font-family:\'DM Mono\',monospace;font-size:0.72rem;color:#94a3b8;padding:2px 0">'
            f'<span style="color:#00e5a0;font-weight:700">{i}.</span> {short}</div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown('<div class="card-label">Excel Template</div>', unsafe_allow_html=True)
    template_file = st.file_uploader("Upload DESIGN EXPORT TEMPLATE.xlsx", type=["xlsx"], key="template")
    if template_file:
        st.markdown('<span class="badge">✓ Template loaded</span>', unsafe_allow_html=True)

# ── MAIN AREA ────────────────────────────────────────────────────
col_left, col_right = st.columns([1.1, 0.9], gap="large")

with col_left:
    st.markdown('<div class="card"><div class="card-label">📂 PDF Upload — Three Files Required</div>', unsafe_allow_html=True)

    prop_name_input = st.text_input(
        "Property Name / Reference",
        placeholder="e.g. 12 Oak Street NW10 4AB",
        key="prop_name"
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        cr_file  = st.file_uploader("CR PDF", type=["pdf"], help="Condition Report — filename must contain 'cr' or 'condition'")
    with c2:
        sn_file  = st.file_uploader("SN PDF", type=["pdf"], help="Survey Notes — filename must contain 'sn'")
    with c3:
        epr_file = st.file_uploader("EPR PDF", type=["pdf"], help="Energy Performance Report — filename must contain 'epr'")

    ready_flags = {
        "Template": template_file is not None,
        "CR PDF": cr_file is not None,
        "SN PDF": sn_file is not None,
        "EPR PDF": epr_file is not None,
        "Smartsheet Token": bool(ss_token),
        "IOE Sheet ID": bool(sheet_id),
        "Client Spec Sheet ID": bool(second_sheet_id),
        "DuckDB Path": bool(duckdb_path),
        "Property Name": bool(prop_name_input),
    }

    st.markdown("</div>", unsafe_allow_html=True)

    # Checklist
    st.markdown('<div class="card"><div class="card-label">✔ Pre-flight Checklist</div>', unsafe_allow_html=True)
    chk_cols = st.columns(3)
    for i, (label, ok) in enumerate(ready_flags.items()):
        icon = "🟢" if ok else "🔴"
        chk_cols[i % 3].markdown(f"{icon} **{label}**")
    st.markdown("</div>", unsafe_allow_html=True)

    all_ready = all(ready_flags.values())

    run_btn = st.button("🚀 Generate Excel Reports", disabled=not all_ready)

with col_right:
    st.markdown('<div class="card"><div class="card-label">📋 Processing Log</div>', unsafe_allow_html=True)
    log_placeholder = st.empty()
    st.markdown("</div>", unsafe_allow_html=True)

    result_placeholder = st.empty()

# ── PROCESS ──────────────────────────────────────────────────────
if run_btn:
    log_lines = []

    def log(msg):
        log_lines.append(msg)
        html_lines = []
        for line in log_lines[-60:]:
            if line.startswith("✅") or line.startswith("✓"):
                cls = "log-ok"
            elif line.startswith("⚠️") or line.startswith("ℹ️"):
                cls = "log-warn"
            elif line.startswith("❌"):
                cls = "log-err"
            else:
                cls = "log-info"
            escaped = line.replace("<", "&lt;").replace(">", "&gt;")
            html_lines.append(f'<div class="{cls}">{escaped}</div>')
        log_placeholder.markdown(
            f'<div class="log-box">{"".join(html_lines)}</div>',
            unsafe_allow_html=True
        )

    try:
        # Save uploaded PDFs to temp files
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
            f.write(cr_file.read())
            cr_path = f.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
            f.write(sn_file.read())
            sn_path = f.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
            f.write(epr_file.read())
            epr_path = f.name

        template_bytes = template_file.read()

        log(f"🚀 Starting processing: {prop_name_input}")

        prelim_buf, construction_buf, prop_name, dwelling_address = process_property(
            cr_path=cr_path,
            sn_path=sn_path,
            epr_path=epr_path,
            template_bytes=template_bytes,
            ss_token=ss_token,
            sheet_id=int(sheet_id),
            second_sheet_id=int(second_sheet_id),
            sheet10_b3=sheet10_b3,
            sheet10_f3=sheet10_f3,
            prop_name=prop_name_input,
            duckdb_path=duckdb_path,
            spec_col_order=spec_col_order,
            log=log,
        )

        log("🎉 Both outputs generated successfully!")

        # Store results in session state for download buttons
        st.session_state["prelim_buf"]        = prelim_buf.getvalue()
        st.session_state["construction_buf"]  = construction_buf.getvalue()
        st.session_state["prop_name_result"] = prop_name_input
        st.session_state["dwelling_address"]  = dwelling_address
        st.session_state["results_ready"]     = True

        # Cleanup temp files
        for p in [cr_path, sn_path, epr_path]:
            try: os.unlink(p)
            except: pass

    except Exception as e:
        import traceback
        log(f"❌ ERROR: {e}")
        log(traceback.format_exc())
        st.session_state["results_ready"] = False

# ── DOWNLOAD SECTION ──────────────────────────────────────────────
if st.session_state.get("results_ready"):
    addr   = st.session_state.get("dwelling_address", "property")
    pname = st.session_state.get("prop_name_result", "property")
    safe   = re.sub(r"[^\w\s-]", "", pname).strip().replace(" ", "_")

    with result_placeholder.container():
        st.markdown("""
        <div class="card">
        <div class="card-label">⬇ Download Outputs</div>
        """, unsafe_allow_html=True)

        st.markdown(f'<span class="badge badge-blue">📍 {addr}</span>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        dcol1, dcol2 = st.columns(2)

        with dcol1:
            st.markdown("""
            <div style="background:#1c2030;border:1px solid #2a2f3e;border-radius:10px;padding:1rem;text-align:center;margin-bottom:0.5rem">
                <div style="font-size:2rem">📄</div>
                <div style="font-family:'Syne',sans-serif;font-size:1rem;font-weight:700;color:#e2e8f0">Prelim Report</div>
                <div style="font-size:0.72rem;color:#64748b;font-family:'DM Mono',monospace;margin-top:0.3rem">No climate risk data</div>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="⬇ Download Prelim",
                data=st.session_state["prelim_buf"],
                file_name=f"{safe}_Prelim.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_prelim",
                use_container_width=True,
            )

        with dcol2:
            st.markdown("""
            <div style="background:#1c2030;border:1px solid #2a2f3e;border-radius:10px;padding:1rem;text-align:center;margin-bottom:0.5rem">
                <div style="font-size:2rem">🏗️</div>
                <div style="font-family:'Syne',sans-serif;font-size:1rem;font-weight:700;color:#e2e8f0">Construction Report</div>
                <div style="font-size:0.72rem;color:#64748b;font-family:'DM Mono',monospace;margin-top:0.3rem">Includes climate risk data</div>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="⬇ Download Construction",
                data=st.session_state["construction_buf"],
                file_name=f"{safe}_Construction.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_construction",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<div style="text-align:center;color:#334155;font-family:\'DM Mono\',monospace;font-size:0.72rem;letter-spacing:0.08em">'
    'RETROFIT DESIGN EXPORTER · ELMHURST CR / SN / EPR → EXCEL'
    '</div>',
    unsafe_allow_html=True
)