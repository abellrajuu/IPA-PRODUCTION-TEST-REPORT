# -*- coding: utf-8 -*-
"""
ST-ERP Production Data Integration Engine
=========================================
Extracts live technical specifications, customer details, and system configurations
directly from the ST-ERP SQL Server database (C:\\Users\\Admin\\Desktop\\ST-ERP.exe).

Authenticates application user (Abel / Abel@123) and queries active financial year
databases dynamically from Company.dbo.YearMaster.
"""

import os
import re
from datetime import datetime

# Load pythonnet CLR for native ADO.NET SQL Server communication
try:
    import clr
    clr.AddReference("System.Data")
    from System.Data.SqlClient import SqlConnection, SqlCommand
    HAS_CLR_SQL = True
except Exception as _e:
    HAS_CLR_SQL = False
    SqlConnection = None
    SqlCommand = None

# ST-ERP Database & Server Credentials
STERP_SERVER   = os.environ.get("STERP_SERVER", "4.213.6.174")
STERP_DB_USER  = os.environ.get("STERP_DB_USER", "sterp")
STERP_DB_PASS  = os.environ.get("STERP_DB_PASS", "TechWeb@12345")
STERP_APP_USER = os.environ.get("STERP_APP_USER", "Abel")
STERP_APP_PASS = os.environ.get("STERP_APP_PASS", "Abel@123")


def get_connection(db_name="Company", timeout=3):
    """Create and open a SqlConnection to ST-ERP SQL Server with fast fail-over timeout."""
    if not HAS_CLR_SQL:
        raise RuntimeError("pythonnet / System.Data.SqlClient is not available in this environment.")
    conn_str = (
        f"Data Source={STERP_SERVER};"
        f"Initial Catalog={db_name};"
        f"User ID={STERP_DB_USER};"
        f"Password={STERP_DB_PASS};"
        f"Connect Timeout={timeout};"
    )
    conn = SqlConnection(conn_str)
    conn.Open()
    return conn


def verify_sterp_user(user_id=None, password=None):
    """
    Verifies that the application user (e.g. Abel / Abel@123) exists and is active
    in ST-ERP UserMaster.
    """
    uid = user_id or STERP_APP_USER
    pwd = password or STERP_APP_PASS
    try:
        # Check current year DB
        dbs = get_available_year_databases()
        conn = get_connection(dbs[0] if dbs else "I0012627")
        try:
            cmd = SqlCommand("SELECT UserName, Active FROM UserMaster WHERE UserID=@uid AND UPassword=@pwd", conn)
            cmd.Parameters.AddWithValue("@uid", uid)
            cmd.Parameters.AddWithValue("@pwd", pwd)
            r = cmd.ExecuteReader()
            valid = False
            uname = ""
            if r.Read():
                uname = str(r.GetValue(0))
                active = str(r.GetValue(1))
                valid = (active.strip().lower() == "yes")
            r.Close()
            return {"authenticated": valid, "user_id": uid, "user_name": uname}
        finally:
            conn.Close()
    except Exception as e:
        return {"authenticated": False, "error": str(e)}


def get_available_year_databases():
    """
    Retrieves active financial year databases from Company.dbo.YearMaster,
    ordered with the most recent financial year first.
    """
    dbs = []
    try:
        conn = get_connection("Company")
        try:
            cmd = SqlCommand(
                "SELECT CompYear FROM YearMaster WHERE Active='Yes' ORDER BY YearID DESC",
                conn
            )
            r = cmd.ExecuteReader()
            while r.Read():
                val = str(r.GetValue(0)).strip()
                if val and val not in dbs:
                    dbs.append(val)
            r.Close()
        finally:
            conn.Close()
    except Exception:
        pass

    # Ensure fallback order if YearMaster query fails
    for default_db in ["I0012627", "I0012526", "I0012425", "I0012324"]:
        if default_db not in dbs:
            dbs.append(default_db)
    return dbs


def _clean_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s == "None" else s


def _normalize_jbox_model(desc, detail="", model="", code=""):
    """
    Normalizes junction box descriptions, BOM items, and JobCard details into
    proper manufacturer model numbers (e.g., JT041M9, JT031M9, JT011M0, JT022M0, JT081M9,
    JC041M9, JC081M9, JB-04, JB-08, etc.)
    Strictly avoids bogus values like 'JB 0', 'JB-0', 'JB0'.
    """
    full = f"{desc} {detail} {model}".strip()

    # 1. Exact catalog code match (priority on detail and model from Work Order / Job Card)
    for target in [detail, model, full]:
        if not target:
            continue
        m_mod = re.search(r'(?:MODEL|M/N|CAT(?:ALOG)?(?:\s*NO)?)\s*[:=-]?\s*([A-Za-z0-9_\-\/]+)', target, re.I)
        if m_mod:
            cand = m_mod.group(1).strip()
            if cand.upper() not in ('NO', 'TYPE', 'MAKE', 'FOR', 'MOC', 'STD', 'AS', 'PER', 'NA', 'N/A', 'REPUTED', 'JB-0', 'JB0', 'JB', '0'):
                if len(cand) >= 3 and cand[0].upper() == 'J' and cand[2].upper() == 'O':
                    cand = cand[0:2] + '0' + cand[3:]
                return cand

        # Specific IPA Junction Box catalog code patterns: JT041M9, JC041M9, JT011M0, JB-04, etc.
        m_cat = re.search(r'\b(J[TC][0-9O]{2,4}[A-Z0-9]*|JB[-_][0-9]{2}[A-Z0-9]*)\b', target, re.I)
        if m_cat:
            cand = m_cat.group(1).strip()
            if len(cand) >= 3 and cand[0].upper() == 'J' and cand[2].upper() == 'O':
                cand = cand[0:2] + '0' + cand[3:]
            if cand.upper() not in ('JB-0', 'JB0', 'JB-00', 'JB00', '0'):
                return cand

    # 2. Item code mapping from ERP Item Master
    code_clean = (code or '').strip().upper()
    if code_clean == 'IJBX00000419': return 'JT041M9'
    if code_clean == 'IJBX00000610': return 'JT031M9'
    if code_clean == 'IJBX00000219': return 'JT022M0'
    if code_clean == 'IJBX00033E30': return 'JT081M9'
    if code_clean == 'IJBX00033F30': return 'JT061M9'
    if code_clean == 'IJBX00000000': return 'JT011M0'

    # 3. Port configuration mapping to standard IPA JB models (fallback if no catalog model)
    port_m = re.search(r'(\d+)\s*IN\s*(\d+)\s*OUT', full, re.I)
    if port_m:
        in_p = port_m.group(1)
        if in_p == '4': return 'JT041M9'
        elif in_p == '3': return 'JT031M9'
        elif in_p == '8': return 'JT081M9'
        elif in_p == '6': return 'JT061M9'
        elif in_p == '2': return 'JT022M0'
        elif in_p == '1': return 'JT011M0'

    return ""


def _extract_specs_from_text(combined_text, default_job="", customer_name=""):
    """
    Parses engineering specifications from text descriptions in SalesOrderDetail,
    ItemMaster, and WorkOrderDetail.
    """
    spec = {
        "job_no": f"JOB-{default_job}" if default_job and not default_job.upper().startswith("JOB-") else default_job,
        "customer": customer_name,
        "customer_name": customer_name,
        "capacity": "",
        "conveyor_no": "",
        "conveyor_tag": "",
        "tag_no": "",
        "crane_id": "",
        "tracking_no": "",
        "equipment_id": "",
        "belt_speed": "",
        "belt_width": "",
        "material": "",
        "bs_model": "",
        "di_model": "",
        "indicator_model": "",
        "sc_model": "",
        "remote_model": "",
        "remote_display_model": "",
        "sensor_model": "",
        "load_cell_model": "",
        "lc_model": "",
        "lc_capacity": "",
        "lc_make": "IPA",
        "lc_qty": "",
        "tacho_model": "",
        "jbox_model": "",
        "jbox_model1": "",
        "num_outputs": "",
        "voltage": "",
        "sw_version": "",
        # Standard Routine Test & Inspection Defaults
        "spec_1": "230V AC",
        "act_1": "230V AC",
        "spec_2": "24V DC",
        "act_2": "24V DC",
        "spec_3": "OK",
        "act_3": "OK",
        "spec_4": "",
        "act_4": "",
        "spec_5": "NA",
        "act_5": "NA",
        "spec_6": "OK",
        "check_6": "OK",
        "act_6": "OK",
        "spec_7": "OK",
        "check_7": "OK",
        "act_7": "OK",
        "spec_8": "RS485",
        "check_8": "OK",
        "act_8": "OK",
        "spec_9": "NA",
        "check_9": "NA",
        "act_9": "NA",
        "spec_10": "1A",
        "check_10": "YES",
        "act_10": "YES"
    }

    txt = combined_text

    # 1. Capacity
    m_cap = re.search(r'Capacity\s*(?:\([^)]*\))?\s*[:=-]?\s*([0-9\.\/]+[A-Za-z0-9\.\/]*\s*(?:TPH|T|KG|T/hr|Ton|Tons|Tonnes)?)\b', txt, re.I)
    if not m_cap:
        m_cap = re.search(r'\b(?:CAP|CAPACITY|SWL|GROSS WT|GROSS WEIGHT)\s*[:=-]?\s*([0-9\.\/]+[A-Za-z0-9\.\/]*\s*(?:TPH|T|KG|Ton|Tons|Tonnes)?)\b', txt, re.I)
    if not m_cap:
        m_cap = re.search(r'\b([0-9\.]+\s*(?:TPH|Ton))\b', txt, re.I)
    if m_cap:
        spec["capacity"] = m_cap.group(1).strip()

    # 2. Belt Speed
    m_speed = re.search(r'Belt\s*Speed\s*[:=-]?\s*([0-9\.]+\s*(?:m/sec|m/s|mps)?)\b', txt, re.I)
    if not m_speed:
        m_speed = re.search(r'\bSpeed\s*[:=-]?\s*([0-9\.]+\s*(?:m/sec|m/s|mps))\b', txt, re.I)
    if m_speed:
        spec["belt_speed"] = m_speed.group(1).strip()

    # 3. Belt Width
    m_bw = re.search(r'\bBW\s*[-:=]?\s*([0-9]+(?:\s*MM)?)\b', txt, re.I)
    if not m_bw:
        m_bw = re.search(r'Belt\s*Width\s*[:=-]?\s*([0-9]+(?:\s*MM)?)\b', txt, re.I)
    if m_bw:
        spec["belt_width"] = m_bw.group(1).strip()

    # 4. Material
    m_mat = re.search(r'material\s*[-:=]?\s*([A-Za-z0-9_\-\s]{3,30})(?:,|\n|$)', txt, re.I)
    if m_mat:
        spec["material"] = m_mat.group(1).strip()

    # 5. Tracking No / Tag No / Conveyor No / Crane ID
    m_crane_no = re.search(r'Crane\s*(?:No\.?|ID)\s*[:=-]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if m_crane_no:
        c_id = m_crane_no.group(1).strip()
        spec["crane_id"] = c_id
        spec["tag_no"] = c_id
        spec["tracking_no"] = c_id
        spec["equipment_id"] = c_id

    m_track = re.search(r'Tracking\s*No\.?\s*[-:=]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if m_track:
        spec["tracking_no"] = m_track.group(1).strip()
        spec["tag_no"] = spec["tracking_no"]
        spec["crane_id"] = spec["tracking_no"]
        spec["conveyor_no"] = spec["tracking_no"]
        spec["equipment_id"] = spec["tracking_no"]

    m_tag = re.search(r'Tag\s*No\.?\s*[-:=]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if m_tag:
        spec["tag_no"] = m_tag.group(1).strip()
        if not spec["conveyor_no"]:
            spec["conveyor_no"] = spec["tag_no"]
            spec["equipment_id"] = spec["tag_no"]

    m_conv = re.search(r'Conveyor\s*(?:No\.?)?\s*[-:=]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if m_conv:
        spec["conveyor_no"] = m_conv.group(1).strip()

    # 6. Indicator / Controller Model
    m_ind = re.search(r'(?:INDICATOR|CONTROLLER|CWS INDICATOR).*?MODEL\s*(?:NO\.?)?\s*[:=-]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if not m_ind:
        m_ind = re.search(r'MODEL\s*[:=-]?\s*([A-Za-z0-9_\-\/]+)\s*(?:Series|INDICATOR|CONTROLLER)', txt, re.I)
    if not m_ind:
        m_ind = re.search(r'\b(BCW\s*(?:Series)?|LH23200|LH2320E|DI[-_]?[0-9A-Z]+|SC200[A-Z0-9]*|LC200[A-Z0-9]*)\b', txt, re.I)
    if m_ind:
        model = m_ind.group(1).strip()
        if model.upper() not in ('DIA', 'STD', 'DISPLAY', 'SERIES', 'MOC', 'NO', 'MODEL', 'TYPE', 'MAKE'):
            spec["bs_model"] = model
            spec["di_model"] = model
            spec["indicator_model"] = model
            if "SC" in model.upper():
                spec["sc_model"] = model

    # 7. Signal Conditioner
    m_sc = re.search(r'SIGNAL\s*CONDITIONER.*?MODEL\s*(?:NO\.?)?\s*[:=-]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if m_sc:
        spec["sc_model"] = m_sc.group(1).strip()

    # 8. Load Cell Model & Capacity
    m_lc_model = re.search(r'(?:LOADCELL|LOAD\s*CELL).*?MODEL\s*[-:=]?\s*([A-Za-z0-9_\-\/]+)', txt, re.I)
    if not m_lc_model:
        m_lc_model = re.search(r'\b(SP\s*[0-9]+[A-Z0-9]*|SB\s*[0-9]+[A-Z0-9]*|TH\s*[0-9]+[A-Z0-9]*|DS\s*[0-9]+[A-Z0-9]*|DLC\s*[0-9]+[A-Z0-9]*)\b', txt, re.I)
    if m_lc_model:
        spec["sensor_model"] = m_lc_model.group(1).strip()
        spec["load_cell_model"] = spec["sensor_model"]
        spec["lc_model"] = spec["sensor_model"]

    m_lc_cap = re.search(r'(?:LOADCELL|LOAD\s*CELL).*?(?:CAPACITY|CAP)\s*[-:=]?\s*([0-9\.]+\s*(?:T|KG|Ton)?)\b', txt, re.I)
    if not m_lc_cap:
        m_lc_cap = re.search(r'(?:SHEAR\s*BEAM|DOUBLE\s*S|TENSION\s*HOLLOW|SHEAR\s*PIN)\s*LOADCELL\s*([0-9\.]+\s*(?:T|KG|Ton)?)\b', txt, re.I)
    if m_lc_cap:
        spec["lc_capacity"] = m_lc_cap.group(1).strip()

    m_lc_qty = re.search(r'(?:LOADCELL|LOAD\s*CELL).*?\[(\d+)(?:\s*no\.?)?\]', txt, re.I)
    if m_lc_qty:
        spec["lc_qty"] = m_lc_qty.group(1).strip()

    # 9. Remote Display Model
    m_rd = re.search(r'(?:REMOTE\s*DISPLAY|7\s*SEGMENT\s*DISPLAY).*?(RD[A-Za-z0-9_\-\/]+|\d+mm)', txt, re.I)
    if m_rd:
        spec["remote_model"] = m_rd.group(1).strip()
        spec["remote_display_model"] = spec["remote_model"]

    # 10. Junction Box Model (Normalized)
    m_jb = re.search(r'junction\s*box\s*[-:=]?\s*([A-Za-z0-9_\-\/\[\]\s]+)', txt, re.I)
    if m_jb:
        raw_jb = m_jb.group(1).strip().replace("\n", " ")
        spec["jbox_model"] = _normalize_jbox_model(raw_jb)
        spec["jbox_model1"] = spec["jbox_model"]
        spec["jbox_model_1"] = spec["jbox_model"]

    # 11. Outputs & Voltage
    m_out = re.search(r'(?:NO:?\s*OF\s*)?(?:RS\s*485|ANALOG(?:UE)?)\s*O/?P\s*[:=-]?\s*(\d+)', txt, re.I)
    if m_out:
        spec["num_outputs"] = m_out.group(1).strip()

    m_volt = re.search(r'(?:SUPPLY\s*(?:VOLTAGE|V)\s*[:=-]?\s*([0-9]+\s*(?:VDC|VAC|V\s*AC|V\s*DC)))', txt, re.I)
    if m_volt:
        spec["voltage"] = m_volt.group(1).strip()
    if not spec.get("spec_1"): spec["spec_1"] = "230V AC"
    if not spec.get("act_1"): spec["act_1"] = spec["spec_1"]
    if not spec.get("spec_2"): spec["spec_2"] = "24V DC"
    if not spec.get("act_2"): spec["act_2"] = spec["spec_2"]

    return spec


def _enrich_specs_from_components(spec, components, default_wo_qty=1):
    """
    Enriches specifications with structured BOM, Work Order, and Job Card components
    (Load Cells, Number of Loadcells, Indicators, Junction Box, Speed Sensor, Remote Display, Mechanical Frames).
    """
    if not components:
        return spec

    lc_model_totals = {}
    jboxes = []
    lc_models = []
    lc_caps = []

    for c in components:
        desc = (c.get("ItemDesc") or "").strip().upper()
        detail = (c.get("DetailDesc") or "").strip()
        model = (c.get("ModelNo") or "").strip()
        code = (c.get("BOMItem") or c.get("code") or c.get("ItemCode") or "").strip()
        
        qty_val = 1
        try:
            qty_raw = str(c.get("ReqQty") or c.get("Quantity") or c.get("qty") or "1")
            qty_val = float(qty_raw)
        except Exception:
            pass

        full_desc = f"{desc} {detail} {model}".strip()

        # Check if item is a mounting accessory, hardware, plate, or consumable to exclude from loadcell
        is_lc_accessory = any(w in desc for w in [
            "PLATE", "ACCESSOR", "BRACKET", "STAND", "COVER", "MOUNT", 
            "ARRESTOR", "CLAMP", "PIN", "CASING", "BOLT", "SCREW", "WASHER", 
            "BUSH", "SLEAVE", "TIE WIRE", "GEL", "RTV", "FLUX", "SOLDER", 
            "SHEET", "STRAIN GAUGE", "TISSUE", "COTTON", "TOLUENE", "SOLUTION",
            "M-BOND", "CYNOACRYLITE", "FLEXIBLE PCB", "TRIMMING SMD", "DIAPHRAGM"
        ]) or code.startswith(("ILMP", "ILMA", "IMTB", "IMTP", "IBOL", "LCAC", "ILCAC"))

        # 1. Loadcell & No of Loadcells
        if (code.startswith("ILDC") or ("LOADCELL" in desc or "LOAD CELL" in desc)) and not is_lc_accessory:
            m_model = re.search(r'\b(SC\s*[0-9]+[A-Z0-9]*|SP\s*[0-9]+[A-Z0-9]*|SB\s*[0-9]+[A-Z0-9]*|TH\s*[0-9]+[A-Z0-9]*|BR\s*[0-9]*|DS\s*[0-9]+[A-Z0-9]*|DLC\s*[0-9]+[A-Z0-9]*)\b', full_desc)
            lc_m = m_model.group(0).strip() if m_model else (model or desc)
            if lc_m:
                lc_model_totals[lc_m] = lc_model_totals.get(lc_m, 0) + qty_val
                if lc_m not in lc_models:
                    lc_models.append(lc_m)
            m_cap = re.search(r'\b([0-9\.]+\s*(?:T|KG|Ton))\b', full_desc, re.I)
            if m_cap:
                c_val = m_cap.group(1).strip()
                if c_val not in lc_caps:
                    lc_caps.append(c_val)

        # 2. Indicator / Controller
        elif ("INDICATOR" in desc or "CONTROLLER" in desc or code.startswith("IINR") or ("DS" in desc and "CONDUIT" not in desc)) and not any(w in desc for w in ["CLAMP", "BRACKET", "PLATE", "PEDESTAL", "STAND", "CANOPY", "CHANNEL", "BOX"]):
            m_ind = re.search(r'\b(DS[-_ ]?[0-9]+[A-Z0-9]*|LH[-_ ]?[0-9]+[A-Z0-9]*|BCW[-_ ]?[0-9]+[A-Z0-9]*|BCW\s*Series|MT4W[A-Z0-9\-]*|DI[-_ ]?[0-9]+[A-Z0-9]*)\b', full_desc, re.I)
            if m_ind:
                ind_m = m_ind.group(0).strip()
                if ind_m.upper() not in ('DIA', 'STD', 'DISPLAY', 'SERIES', 'MOC', 'NO', 'MODEL', 'TYPE', 'MAKE'):
                    spec["bs_model"] = ind_m
                    spec["di_model"] = ind_m
                    spec["indicator_model"] = ind_m

        # 3. Junction Box (Strict filter: exclude mounting brackets, plates, cards, canopies, stickers)
        elif (code.startswith("IJBX") or ("JUNCTION BOX" in desc and not any(w in desc for w in ["BRACKET", "PLATE", "CARD", "CANOPY", "STAND", "STICKER", "GLAND"]))) and not code.startswith(("IMTB", "IMTP", "ICRD", "ISKR", "IPAE", "ICLG", "IBOL", "ICHN")):
            jb_norm = _normalize_jbox_model(desc, detail, model, code)
            if jb_norm and jb_norm not in jboxes:
                jboxes.append(jb_norm)

        # 4. Tacho / Speed Sensor
        elif "TACHO" in desc or "SPEED SENSOR" in desc or code.startswith("ITAC") or code.startswith("ISBW"):
            if not spec.get("tacho_model"):
                m_tacho = re.search(r'\b(BW-PROXIMITY|PROXIMITY|TACHO[-_]?[0-9A-Z]+)\b', full_desc, re.I)
                spec["tacho_model"] = m_tacho.group(0).strip() if m_tacho else (model or desc)

        # 5. Remote Display (Exclude enclosures, brackets, cards)
        elif (code.startswith("ILRD") or "REMOTE DISPLAY" in desc or (("REMOTE" in desc or "DISPLAY" in desc) and not any(w in desc for w in ["ENCLOSURE", "HOUSING", "CARD", "BRACKET", "STAND", "PLATE", "INDICATOR", "CONTROLLER"]))) or "RD2" in full_desc:
            if not spec.get("remote_model"):
                m_rd = re.search(r'\b(RD[A-Za-z0-9_\-\/]+|\d+\s*mm)\b', full_desc, re.I)
                raw_rd = m_rd.group(1).strip() if m_rd else (model or desc)
                if re.match(r'^\d+\s*mm$', raw_rd, re.I):
                    spec["remote_model"] = f"RD {raw_rd}"
                else:
                    spec["remote_model"] = raw_rd
                spec["remote_display_model"] = spec["remote_model"]

        # 6. Mechanical Weighing Frame / Belt Width / Material
        elif "WEIGHING FRAME" in desc or "BELT WEIGHER" in desc or code.startswith("IWBW"):
            m_bw = re.search(r'BW[-:=]?\s*([0-9]+(?:\s*MM)?)', full_desc, re.I)
            if m_bw and not spec.get("belt_width"):
                spec["belt_width"] = m_bw.group(1).strip()
            m_moc = re.search(r'MOC[-:=]?\s*([A-Za-z0-9_\-]+)', full_desc, re.I)
            if m_moc and not spec.get("material"):
                spec["material"] = m_moc.group(1).strip()

        # 7. Crane Scale / Crane Weighing System / Load Limiter
        elif "CRANE WEIGHING" in desc or "CRANE SCALE" in desc or "LOAD LIMITER" in desc or code.startswith("ISYS"):
            if not spec.get("crane_model"):
                m_crane = re.search(r'\b(CRANE\s*WEIGHING\s*SYSTEM[^\,\n]*|LOAD\s*LIMITER\s*SYSTEM[^\,\n]*|CWS[A-Za-z0-9\-]*)\b', full_desc, re.I)
                spec["crane_model"] = m_crane.group(0).strip() if m_crane else desc

    def _jbox_priority(jb):
        jb_u = jb.upper()
        if any(k in jb_u for k in ['08', 'JB-08', '8IN', 'JT08', 'JC08']): return 8
        if any(k in jb_u for k in ['06', 'JB-06', '6IN', 'JT06', 'JC06']): return 6
        if any(k in jb_u for k in ['04', 'JB-04', '4IN', 'JT04', 'JC04']): return 4
        if any(k in jb_u for k in ['03', 'JB-03', '3IN', 'JT03', 'JC03']): return 3
        if any(k in jb_u for k in ['02', 'JB-02', '2IN', 'JT02', 'JC02']): return 2
        if any(k in jb_u for k in ['01', 'JB-01', '1IN', 'JT01', 'JC01']): return 1
        return 0

    if jboxes:
        jboxes.sort(key=_jbox_priority, reverse=True)
        spec["jbox_model"] = jboxes[0]
        spec["jbox_model1"] = jboxes[0]
        spec["jbox_model_1"] = jboxes[0]
        if len(jboxes) > 1:
            spec["jbox_model2"] = jboxes[1]
            spec["jbox_model_2"] = jboxes[1]
        if len(jboxes) > 2:
            spec["jbox_model3"] = jboxes[2]
            spec["jbox_model_3"] = jboxes[2]

    if lc_models:
        # Sort load cell models by total quantity descending so real weighing load cells are primary
        lc_models.sort(key=lambda m: lc_model_totals.get(m, 0), reverse=True)
        primary_model = lc_models[0]
        spec["sensor_model"] = primary_model
        spec["load_cell_model"] = primary_model
        spec["lc_model"] = primary_model
        spec["load_cell_model1"] = primary_model
        spec["lc_model1"] = primary_model
        if len(lc_models) > 1:
            spec["load_cell_model2"] = lc_models[1]
            spec["lc_model2"] = lc_models[1]

        # Calculate per-unit loadcell quantity
        tot = lc_model_totals.get(primary_model, 0)
        per_unit = int(round(tot / max(default_wo_qty, 1)))
        if per_unit > 0:
            spec["lc_qty"] = str(per_unit)
            if per_unit in (1, 2, 3, 4, 6, 8):
                spec["lc_system"] = f"{per_unit}LC"

    if lc_caps:
        spec["lc_capacity"] = lc_caps[0]
        spec["lc_capacity1"] = lc_caps[0]
        if len(lc_caps) > 1:
            spec["lc_capacity2"] = lc_caps[1]
        if not spec.get("capacity"):
            spec["capacity"] = lc_caps[0]

    return spec



def fetch_sterp_job_data(job_query):
    """
    Main entry point: fetches job information from ST-ERP database.
    
    Accepts:
        job_query (str): Job Number, Sales Order Number, Work Order Number, or Customer PO.
                         Examples: "SY1605", "JOB-SY1605", "1605", "SP1630", "WY00002512"

    Returns:
        dict: Standardized payload containing 'status', 'source', 'extracted',
              'multi_systems' (if applicable), and raw records.
    """
    if not job_query or not str(job_query).strip():
        return {"status": "error", "message": "Job Number is required."}

    raw = str(job_query).strip()
    clean = re.sub(r'^(JOB[-_]?)', '', raw, flags=re.I).strip()
    clean = re.sub(r'[^a-zA-Z0-9_\-]', '', clean)
    if not clean:
        return {"status": "error", "message": "Invalid Job Number format."}

    # ── OFFLINE RESILIENCE: Check local SQLite cache first ──
    try:
        import database_python
        cached_data = database_python.get_sterp_cache(clean)
    except Exception:
        cached_data = None

    year_dbs = get_available_year_databases()
    found_data = None
    server_connection_error = None
    connected_any = False
    matched_db = None

    for db_name in year_dbs:
        try:
            conn = get_connection(db_name, timeout=3)
            connected_any = True
        except Exception as _net_err:
            server_connection_error = str(_net_err)
            if not connected_any:
                # If cannot connect to the server on the first attempt, server is unreachable
                break
            continue

        try:
            # 1. Search SalesOrderMaster by SONo or CustPONo (Exact or Prefix)
            cmd_so = SqlCommand(
                "SELECT TOP 1 SONo, SODate, CustPONo, PODate, PartyCode, SPartyName, SAddress, SCity, SState "
                "FROM SalesOrderMaster "
                "WHERE SONo=@q1 OR CustPONo=@q2 OR SONo=@clean_pref "
                "ORDER BY SODate DESC",
                conn
            )
            cmd_so.Parameters.AddWithValue("@q1", clean)
            cmd_so.Parameters.AddWithValue("@q2", raw)
            cmd_so.Parameters.AddWithValue("@clean_pref", f"SY{clean}" if clean.isdigit() else clean)
            r_so = cmd_so.ExecuteReader()
            so_row = None
            if r_so.Read():
                so_row = {
                    "SONo": _clean_str(r_so.GetValue(0)),
                    "SODate": _clean_str(r_so.GetValue(1)),
                    "CustPONo": _clean_str(r_so.GetValue(2)),
                    "PODate": _clean_str(r_so.GetValue(3)),
                    "PartyCode": _clean_str(r_so.GetValue(4)),
                    "SPartyName": _clean_str(r_so.GetValue(5)),
                    "SAddress": _clean_str(r_so.GetValue(6)),
                    "SCity": _clean_str(r_so.GetValue(7)),
                    "SState": _clean_str(r_so.GetValue(8)),
                }
            r_so.Close()

            # 1b. If not exact, check fuzzy search in SalesOrderMaster
            if not so_row:
                cmd_fuzzy = SqlCommand(
                    "SELECT TOP 1 SONo, SODate, CustPONo, PODate, PartyCode, SPartyName, SAddress, SCity, SState "
                    "FROM SalesOrderMaster "
                    "WHERE SONo LIKE @fuzzy OR CustPONo LIKE @fuzzy "
                    "ORDER BY SODate DESC",
                    conn
                )
                cmd_fuzzy.Parameters.AddWithValue("@fuzzy", f"%{clean}%")
                r_f = cmd_fuzzy.ExecuteReader()
                if r_f.Read():
                    so_row = {
                        "SONo": _clean_str(r_f.GetValue(0)),
                        "SODate": _clean_str(r_f.GetValue(1)),
                        "CustPONo": _clean_str(r_f.GetValue(2)),
                        "PODate": _clean_str(r_f.GetValue(3)),
                        "PartyCode": _clean_str(r_f.GetValue(4)),
                        "SPartyName": _clean_str(r_f.GetValue(5)),
                        "SAddress": _clean_str(r_f.GetValue(6)),
                        "SCity": _clean_str(r_f.GetValue(7)),
                        "SState": _clean_str(r_f.GetValue(8)),
                    }
                r_f.Close()

            # 2. Check WorkOrderMaster if not found in SalesOrderMaster
            wo_sono = None
            if not so_row:
                cmd_wo = SqlCommand(
                    "SELECT TOP 1 WONo, SONo, PartyCode FROM WorkOrderMaster WHERE WONo=@w OR WONo LIKE @wf",
                    conn
                )
                cmd_wo.Parameters.AddWithValue("@w", clean)
                cmd_wo.Parameters.AddWithValue("@wf", f"%{clean}%")
                r_wo = cmd_wo.ExecuteReader()
                if r_wo.Read():
                    wo_sono = _clean_str(r_wo.GetValue(1))
                r_wo.Close()

                if wo_sono:
                    cmd_so2 = SqlCommand(
                        "SELECT TOP 1 SONo, SODate, CustPONo, PODate, PartyCode, SPartyName, SAddress, SCity, SState "
                        "FROM SalesOrderMaster WHERE SONo=@s",
                        conn
                    )
                    cmd_so2.Parameters.AddWithValue("@s", wo_sono)
                    r_so2 = cmd_so2.ExecuteReader()
                    if r_so2.Read():
                        so_row = {
                            "SONo": _clean_str(r_so2.GetValue(0)),
                            "SODate": _clean_str(r_so2.GetValue(1)),
                            "CustPONo": _clean_str(r_so2.GetValue(2)),
                            "PODate": _clean_str(r_so2.GetValue(3)),
                            "PartyCode": _clean_str(r_so2.GetValue(4)),
                            "SPartyName": _clean_str(r_so2.GetValue(5)),
                            "SAddress": _clean_str(r_so2.GetValue(6)),
                            "SCity": _clean_str(r_so2.GetValue(7)),
                            "SState": _clean_str(r_so2.GetValue(8)),
                        }
                    r_so2.Close()

            if so_row:
                target_so = so_row["SONo"]

                # 3. Retrieve Line Items from SalesOrderDetail & ItemMaster
                cmd_items = SqlCommand(
                    "SELECT sod.SrNo, sod.ItemCode, im.ItemDesc, im.AddDesc, im.ModelNo, im.DrawingNo, sod.DetailDesc, sod.Quantity, sod.UOM, sod.AutoNo "
                    "FROM SalesOrderDetail sod "
                    "LEFT JOIN ItemMaster im ON sod.ItemCode = im.ItemCode "
                    "WHERE sod.SONo=@so "
                    "ORDER BY sod.SrNo",
                    conn
                )
                cmd_items.Parameters.AddWithValue("@so", target_so)
                r_items = cmd_items.ExecuteReader()
                items = []
                while r_items.Read():
                    items.append({
                        "SrNo": _clean_str(r_items.GetValue(0)),
                        "ItemCode": _clean_str(r_items.GetValue(1)),
                        "ItemDesc": _clean_str(r_items.GetValue(2)),
                        "AddDesc": _clean_str(r_items.GetValue(3)),
                        "ModelNo": _clean_str(r_items.GetValue(4)),
                        "DrawingNo": _clean_str(r_items.GetValue(5)),
                        "DetailDesc": _clean_str(r_items.GetValue(6)),
                        "Quantity": _clean_str(r_items.GetValue(7)),
                        "UOM": _clean_str(r_items.GetValue(8)),
                        "AutoNo": _clean_str(r_items.GetValue(9)),
                    })
                r_items.Close()

                # 4. Retrieve Associated Work Orders & Subitems
                cmd_wos = SqlCommand(
                    "SELECT wo.WONo, wo.ItemCode, im.ItemDesc, wo.WOQty, wo.BOMNo, wo.AutoNo "
                    "FROM WorkOrderMaster wo "
                    "LEFT JOIN ItemMaster im ON wo.ItemCode = im.ItemCode "
                    "WHERE wo.SONo=@so",
                    conn
                )
                cmd_wos.Parameters.AddWithValue("@so", target_so)
                r_w = cmd_wos.ExecuteReader()
                wos = []
                while r_w.Read():
                    wos.append({
                        "WONo": _clean_str(r_w.GetValue(0)),
                        "ItemCode": _clean_str(r_w.GetValue(1)),
                        "ItemDesc": _clean_str(r_w.GetValue(2)),
                        "WOQty": _clean_str(r_w.GetValue(3)),
                        "BOMNo": _clean_str(r_w.GetValue(4)),
                        "AutoNo": _clean_str(r_w.GetValue(5)),
                        "subitems": []
                    })
                r_w.Close()

                for w in wos:
                    cmd_subs = SqlCommand(
                        "SELECT wod.BOMItem, im.ItemDesc, im.ModelNo, wod.DetailDesc, wod.ReqQty, im.DrawingNo "
                        "FROM WorkOrderDetail wod "
                        "LEFT JOIN ItemMaster im ON wod.BOMItem = im.ItemCode "
                        "WHERE wod.WONo=@wono",
                        conn
                    )
                    cmd_subs.Parameters.AddWithValue("@wono", w["WONo"])
                    r_sub = cmd_subs.ExecuteReader()
                    while r_sub.Read():
                        w["subitems"].append({
                            "BOMItem": _clean_str(r_sub.GetValue(0)),
                            "ItemDesc": _clean_str(r_sub.GetValue(1)),
                            "ModelNo": _clean_str(r_sub.GetValue(2)),
                            "DetailDesc": _clean_str(r_sub.GetValue(3)),
                            "ReqQty": _clean_str(r_sub.GetValue(4)),
                            "DrawingNo": _clean_str(r_sub.GetValue(5)),
                        })
                    r_sub.Close()

                # 5. Retrieve BOM Components for Systems (Belt Scale, Weigh Feeder, Indicators, etc.)
                for itm in items:
                    itm["bom_components"] = []
                    bom_no = ""
                    try:
                        cmd_blookup = SqlCommand("SELECT TOP 1 BOMNo FROM BOMMaster WHERE ItemCode=@icode ORDER BY BOMRevision DESC", conn)
                        cmd_blookup.Parameters.AddWithValue("@icode", itm["ItemCode"])
                        res_b = cmd_blookup.ExecuteScalar()
                        if res_b:
                            bom_no = _clean_str(res_b)
                    except Exception:
                        pass

                    if not bom_no:
                        for w in wos:
                            if (w.get("ItemCode") == itm.get("ItemCode") or len(items) == 1) and w.get("BOMNo"):
                                bom_no = w.get("BOMNo")
                                break

                    if bom_no:
                        try:
                            cmd_bdetails = SqlCommand(
                                "SELECT bd.ACode, im.ItemDesc, im.ModelNo, bd.DetailDesc, bd.Quantity, im.DrawingNo "
                                "FROM BOMDetails bd "
                                "LEFT JOIN ItemMaster im ON bd.ACode = im.ItemCode "
                                "WHERE bd.BOMNo=@bno",
                                conn
                            )
                            cmd_bdetails.Parameters.AddWithValue("@bno", bom_no)
                            r_bd = cmd_bdetails.ExecuteReader()
                            while r_bd.Read():
                                itm["bom_components"].append({
                                    "BOMItem": _clean_str(r_bd.GetValue(0)),
                                    "ItemDesc": _clean_str(r_bd.GetValue(1)),
                                    "ModelNo": _clean_str(r_bd.GetValue(2)),
                                    "DetailDesc": _clean_str(r_bd.GetValue(3)),
                                    "ReqQty": _clean_str(r_bd.GetValue(4)),
                                    "DrawingNo": _clean_str(r_bd.GetValue(5)),
                                })
                            r_bd.Close()
                        except Exception:
                            pass

                # Fallback: retrieve BOM details only for Work Orders that had no WorkOrderDetail rows
                for w in wos:
                    w_bom = w.get("BOMNo")
                    if w_bom and not w["subitems"]:
                        try:
                            wo_qty_scale = float(w.get("WOQty") or 1) or 1.0
                        except Exception:
                            wo_qty_scale = 1.0
                        try:
                            cmd_wbom = SqlCommand(
                                "SELECT bd.ACode, im.ItemDesc, im.ModelNo, bd.DetailDesc, bd.Quantity, im.DrawingNo "
                                "FROM BOMDetails bd "
                                "LEFT JOIN ItemMaster im ON bd.ACode = im.ItemCode "
                                "WHERE bd.BOMNo=@bno",
                                conn
                            )
                            cmd_wbom.Parameters.AddWithValue("@bno", w_bom)
                            r_wb = cmd_wbom.ExecuteReader()
                            while r_wb.Read():
                                raw_qty = _clean_str(r_wb.GetValue(4))
                                try:
                                    scaled_qty = str(float(raw_qty) * wo_qty_scale) if raw_qty else ""
                                except Exception:
                                    scaled_qty = raw_qty
                                w["subitems"].append({
                                    "BOMItem": _clean_str(r_wb.GetValue(0)),
                                    "ItemDesc": _clean_str(r_wb.GetValue(1)),
                                    "ModelNo": _clean_str(r_wb.GetValue(2)),
                                    "DetailDesc": _clean_str(r_wb.GetValue(3)),
                                    "ReqQty": scaled_qty,
                                    "DrawingNo": _clean_str(r_wb.GetValue(5)),
                                })
                            r_wb.Close()
                        except Exception:
                            pass

                # 6. Retrieve Shop-Floor Job Cards (JobCardMaster) for electrical and mechanical specs
                job_cards = []
                try:
                    cmd_jcs = SqlCommand(
                        "SELECT jc.JobCardNo, jc.WONo, jc.ItemCode, im.ItemDesc, im.ModelNo, jc.DetailDesc, jc.JCType, jc.Remark, jc.JobQty "
                        "FROM JobCardMaster jc "
                        "LEFT JOIN ItemMaster im ON jc.ItemCode = im.ItemCode "
                        "WHERE jc.WONo IN (SELECT WONo FROM WorkOrderMaster WHERE SONo=@so) "
                        "   OR jc.WONo=@so "
                        "   OR jc.RefNo=@so "
                        "ORDER BY jc.JobCardDate DESC",
                        conn
                    )
                    cmd_jcs.Parameters.AddWithValue("@so", target_so)
                    r_jc = cmd_jcs.ExecuteReader()
                    while r_jc.Read():
                        job_cards.append({
                            "JobCardNo": _clean_str(r_jc.GetValue(0)),
                            "WONo": _clean_str(r_jc.GetValue(1)),
                            "BOMItem": _clean_str(r_jc.GetValue(2)),
                            "ItemDesc": _clean_str(r_jc.GetValue(3)),
                            "ModelNo": _clean_str(r_jc.GetValue(4)),
                            "DetailDesc": _clean_str(r_jc.GetValue(5)),
                            "JCType": _clean_str(r_jc.GetValue(6)),
                            "Remark": _clean_str(r_jc.GetValue(7)),
                            "ReqQty": _clean_str(r_jc.GetValue(8)),
                        })
                    r_jc.Close()
                except Exception:
                    pass

                found_data = {
                    "so": so_row,
                    "items": items,
                    "work_orders": wos,
                    "job_cards": job_cards
                }
                matched_db = db_name
                # Only stop searching if this database actually contains line items or work orders!
                # If an order is an empty shell in a newer database (e.g. carry-over),
                # continue searching earlier databases where the real data was entered.
                if items or wos or job_cards:
                    break
        finally:
            conn.Close()

    if not found_data:
        if cached_data:
            cached_data["message"] = f"Loaded '{job_query}' from local cache (offline mode)"
            return cached_data

        if not connected_any:
            return {
                "status": "error",
                "message": f"Could not connect to ST-ERP SQL Server at {STERP_SERVER}. Please verify that your computer is connected to the company network or VPN."
            }

        return {
            "status": "error",
            "message": f"Job / Order '{job_query}' not found in ST-ERP databases."
        }

    so_info = found_data["so"]
    items = found_data["items"]
    wos = found_data["work_orders"]
    job_cards = found_data.get("job_cards", [])
    sono = so_info["SONo"]
    customer = so_info["SPartyName"]

    # ── ADVANCED MULTI-SYSTEM DETECTION ENGINE ──
    # Check if this order contains multiple distinct systems:
    # 1. Distinct System Line Items (ISYS... or complete weighing systems)
    # 2. Component Sets grouped by TRACKING NO. (e.g. Electromech Crane Sets)
    # 3. Otherwise: Single Unified System (avoiding false alarms from spare parts/accessories)
    
    tracking_groups = {}
    for itm in items:
        full = f"{itm.get('ItemDesc', '')} {itm.get('DetailDesc', '')}"
        m_trk = re.search(r'TRACKING\s*NO\.?\s*([A-Za-z0-9_\-]+)', full, re.I)
        if m_trk:
            trk = m_trk.group(1).strip().upper()
            tracking_groups.setdefault(trk, []).append(itm)

    system_items = []
    for itm in items:
        code_i = (itm.get("ItemCode") or "").upper()
        desc_i = (itm.get("ItemDesc") or "").upper()
        detail_i = (itm.get("DetailDesc") or "").upper()
        full_i = f"{desc_i} {detail_i}"
        
        # Exclude non-systems (cables, freight, services, standalone accessories)
        if any(ex in desc_i for ex in ["CABLE", "FREIGHT", "SERVICE", "COMMISSIONING", "CALIBRATION CHARGES", "STANDALONE", "JUNCTION BOX ONLY", "ACCESSORIES ONLY"]):
            continue

        if code_i.startswith("ISYS") or any(kw in full_i for kw in [
            "WEIGHING SYSTEM", "BELT WEIGHER", "WEIGH FEEDER", "CRANE SCALE",
            "CRANE WEIGHING", "LOSS IN WEIGHT", "BATCHING SYSTEM", "BIN WEIGHING",
            "SILO WEIGHING", "TANK WEIGHING", "HOPPER WEIGHING", "HOPPER/SILO"
        ]):
            system_items.append(itm)

    multi_systems = {}
    systems_meta = []

    if len(system_items) > 1:
        # STRATEGY 1: Multiple System Line Items (e.g. SY0192)
        for idx, itm in enumerate(system_items, 1):
            detail = itm.get("DetailDesc") or ""
            desc = itm.get("ItemDesc") or ""
            
            # Clean application name
            lines = [l.strip() for l in detail.splitlines() if l.strip()]
            app_name = ""
            for l in lines:
                if not any(skip in l.upper() for skip in ["TAG", "SCOPE", "REFER", "DRAWING", "ANNEXURE", "AS PER"]):
                    app_name = l
                    break
            if not app_name:
                app_name = desc

            clean_title = f"System {idx}: {app_name}"
            sys_key = f"SYSTEM {idx} - {app_name}"

            # Text extraction strictly for this line
            line_text = f"{desc} \n{detail}"
            sys_spec = _extract_specs_from_text(line_text, default_job=sono, customer_name=customer)

            # Match work orders strictly by AutoNo (foreign key in ST-ERP!)
            itm_auto = itm.get("AutoNo")
            matching_wos = [w for w in wos if w.get("AutoNo") and w.get("AutoNo") == itm_auto]
            if not matching_wos and len(system_items) == 1:
                matching_wos = wos

            sys_components = []
            total_wo_qty = sum(float(w.get("WOQty") or 1) for w in matching_wos) or float(itm.get("Quantity") or 1)
            for mw in matching_wos:
                sys_components.extend(mw.get("subitems", []))
            matching_jcs = [jc for jc in job_cards if jc.get("BOMItem") == itm.get("ItemCode")]
            sys_components.extend(matching_jcs)
            qty_divisor = total_wo_qty
            if not sys_components:
                sys_components.extend(itm.get("bom_components", []))
                qty_divisor = 1
            sys_components.append(itm)

            sys_spec = _enrich_specs_from_components(sys_spec, sys_components, default_wo_qty=qty_divisor)
            multi_systems[sys_key] = sys_spec

            cap = sys_spec.get("capacity") or ""
            tag = sys_spec.get("tag_no") or sys_spec.get("conveyor_no") or ""
            lc_m = sys_spec.get("lc_model") or sys_spec.get("sensor_model") or ""
            lc_q = sys_spec.get("lc_qty") or ""
            di_m = sys_spec.get("di_model") or sys_spec.get("indicator_model") or ""
            jb_m = sys_spec.get("jbox_model") or sys_spec.get("jbox_model1") or ""

            summary_parts = []
            if cap: summary_parts.append(cap)
            if lc_m: summary_parts.append(f"{lc_q}× {lc_m}" if lc_q else lc_m)
            if di_m: summary_parts.append(di_m)
            if jb_m: summary_parts.append(jb_m)

            systems_meta.append({
                "id": f"sys_{idx}",
                "key": sys_key,
                "title": clean_title,
                "tag_no": tag,
                "capacity": cap,
                "lc_model": lc_m,
                "lc_qty": lc_q,
                "di_model": di_m,
                "jbox_model": jb_m,
                "summary": " • ".join(summary_parts) or "Complete System"
            })

    elif len(tracking_groups) > 1:
        # STRATEGY 2: Grouped by Tracking Numbers (e.g. Electromech Crane Scale Sets like SY1246)
        for idx, (trk_no, g_items) in enumerate(tracking_groups.items(), 1):
            clean_title = f"Set {idx}: Tracking No. {trk_no}"
            sys_key = f"SET {idx} - TRACKING {trk_no}"

            group_text = "\n".join(f"{i.get('ItemDesc', '')} {i.get('DetailDesc', '')}" for i in g_items)
            sys_spec = _extract_specs_from_text(group_text, default_job=sono, customer_name=customer)
            sys_spec["tracking_no"] = trk_no
            sys_spec["tag_no"] = trk_no
            sys_spec["equipment_id"] = trk_no
            if not sys_spec.get("conveyor_no"):
                sys_spec["conveyor_no"] = trk_no

            sys_spec = _enrich_specs_from_components(sys_spec, g_items, default_wo_qty=1)
            multi_systems[sys_key] = sys_spec

            cap = sys_spec.get("capacity") or ""
            lc_m = sys_spec.get("lc_model") or sys_spec.get("sensor_model") or ""
            lc_q = sys_spec.get("lc_qty") or str(sum(1 for i in g_items if "LOADCELL" in (i.get("ItemDesc") or "").upper()))
            di_m = sys_spec.get("di_model") or sys_spec.get("indicator_model") or ""
            jb_m = sys_spec.get("jbox_model") or ""

            summary_parts = []
            if cap: summary_parts.append(cap)
            if lc_m: summary_parts.append(f"{lc_q}× {lc_m}" if lc_q else lc_m)
            if di_m: summary_parts.append(di_m)

            systems_meta.append({
                "id": f"sys_{idx}",
                "key": sys_key,
                "title": clean_title,
                "tag_no": trk_no,
                "capacity": cap,
                "lc_model": lc_m,
                "lc_qty": lc_q,
                "di_model": di_m,
                "jbox_model": jb_m,
                "summary": " • ".join(summary_parts) or f"Tracking {trk_no}"
            })

    # Overall combined extraction across all items, work orders, and shop-floor job cards
    all_blocks = []
    all_components = []
    for itm in items:
        all_blocks.append(f"{itm['ItemDesc']} {itm['AddDesc']} {itm['DetailDesc']}")
        all_components.append(itm)
        if not wos:
            all_components.extend(itm.get("bom_components", []))
    for w in wos:
        for sub in w["subitems"]:
            all_blocks.append(f"{sub['ItemDesc']} {sub['ModelNo']} {sub['DetailDesc']}")
            all_components.append(sub)
    for jc in job_cards:
        all_blocks.append(f"{jc['ItemDesc']} {jc['ModelNo']} {jc['DetailDesc']} {jc['Remark']}")
        all_components.append(jc)
    combined_master_text = "\n".join(all_blocks)

    overall_extracted = _extract_specs_from_text(combined_master_text, default_job=sono, customer_name=customer)
    overall_extracted = _enrich_specs_from_components(overall_extracted, all_components)

    # Ensure all systems in multi_systems inherit shared order-level parameters ONLY
    # (customer, job_no, voltages, checklist defaults), NOT system-specific hardware!
    SHARED_ORDER_KEYS = {
        "job_no", "customer", "customer_name", "voltage", "sw_version",
        "spec_1", "act_1", "spec_2", "act_2", "spec_3", "act_3",
        "spec_5", "act_5", "spec_6", "check_6", "act_6",
        "spec_7", "check_7", "act_7", "spec_8", "check_8", "act_8",
        "spec_9", "check_9", "act_9", "spec_10", "check_10", "act_10",
        "tested_by", "approved_by", "date"
    }
    for s_name, s_spec in multi_systems.items():
        for k in SHARED_ORDER_KEYS:
            if not s_spec.get(k) and overall_extracted.get(k):
                s_spec[k] = overall_extracted[k]
        # Also inherit indicator / junction box IF the system line didn't specify its own
        if not s_spec.get("di_model") and overall_extracted.get("di_model"):
            s_spec["di_model"] = overall_extracted["di_model"]
        if not s_spec.get("indicator_model") and overall_extracted.get("indicator_model"):
            s_spec["indicator_model"] = overall_extracted["indicator_model"]
        if not s_spec.get("jbox_model") and overall_extracted.get("jbox_model"):
            s_spec["jbox_model"] = overall_extracted["jbox_model"]

    has_multi = len(multi_systems) > 1
    if has_multi and multi_systems:
        first_sys_key = next(iter(multi_systems.keys()))
        extracted = dict(overall_extracted)
        for k, v in multi_systems[first_sys_key].items():
            if v and v not in ("None", "NA", "OK", ""):
                extracted[k] = v
    else:
        extracted = overall_extracted

    final_result = {
        "status": "success",
        "source": "sterp",
        "filename": f"ST-ERP [{sono}]",
        "so_no": sono,
        "database": matched_db,
        "customer": customer,
        "has_multiple_systems": has_multi,
        "system_names": list(multi_systems.keys()) if has_multi else [],
        "multi_systems": multi_systems if has_multi else {},
        "systems_meta": systems_meta if has_multi else [],
        "extracted": extracted,
        "items": items,
        "work_orders": wos,
        "raw_so": so_info
    }
    try:
        import database_python
        database_python.save_sterp_cache(clean, final_result)
        if sono and sono.upper() != clean:
            database_python.save_sterp_cache(sono, final_result)
    except Exception as _c_err:
        pass
    return final_result
