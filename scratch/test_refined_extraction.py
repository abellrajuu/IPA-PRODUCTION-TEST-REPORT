import sys, os
sys.path.insert(0, os.path.abspath('.'))
sys.stdout.reconfigure(encoding='utf-8')
import re
from sterp_data_fetcher import get_connection, _clean_str, _normalize_jbox_model, _extract_specs_from_text
import clr
clr.AddReference('System.Data')
from System.Data.SqlClient import SqlCommand

conn = get_connection('I0012627')

# Query SalesOrderDetail with AutoNo
cmd_items = SqlCommand(
    "SELECT sod.SrNo, sod.ItemCode, im.ItemDesc, im.AddDesc, im.ModelNo, im.DrawingNo, sod.DetailDesc, sod.Quantity, sod.UOM, sod.AutoNo "
    "FROM SalesOrderDetail sod "
    "LEFT JOIN ItemMaster im ON sod.ItemCode = im.ItemCode "
    "WHERE sod.SONo='SY0192' "
    "ORDER BY sod.SrNo",
    conn
)
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

# Query WorkOrderMaster with AutoNo
cmd_wos = SqlCommand(
    "SELECT wo.WONo, wo.ItemCode, im.ItemDesc, wo.WOQty, wo.BOMNo, wo.AutoNo "
    "FROM WorkOrderMaster wo "
    "LEFT JOIN ItemMaster im ON wo.ItemCode = im.ItemCode "
    "WHERE wo.SONo='SY0192'",
    conn
)
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

conn.Close()

print(f"Loaded {len(items)} items and {len(wos)} work orders.")

def refined_enrich(spec, components, default_wo_qty=1):
    if not components:
        return spec
    
    lc_count = 0
    jboxes = []
    lc_models = []
    lc_caps = []
    
    # Priority indicators
    for c in components:
        desc = (c.get("ItemDesc") or "").strip().upper()
        detail = (c.get("DetailDesc") or "").strip()
        model = (c.get("ModelNo") or "").strip()
        code = (c.get("BOMItem") or c.get("code") or c.get("ItemCode") or "").strip()
        full_desc = f"{desc} {detail} {model}".strip()

        qty_val = 1
        try:
            qty_raw = str(c.get("ReqQty") or c.get("Quantity") or c.get("qty") or "1")
            qty_val = float(qty_raw)
        except Exception:
            pass

        # Check if mounting accessory/hardware to exclude from loadcell
        is_accessory = any(w in desc for w in [
            "PLATE", "ACCESSOR", "BRACKET", "STAND", "COVER", "MOUNT", 
            "ARRESTOR", "CLAMP", "PIN", "CASING", "BOLT", "SCREW", "WASHER", 
            "BUSH", "SLEAVE", "TIE WIRE", "GEL", "RTV", "FLUX", "SOLDER", 
            "SHEET", "STRAIN GAUGE", "TISSUE", "COTTON", "TOLUENE", "SOLUTION",
            "M-BOND", "CYNOACRYLITE", "FLEXIBLE PCB", "TRIMMING SMD", "DIAPHRAGM"
        ]) or code.startswith(("ILMP", "ILMA", "IMTB", "IMTP", "IBOL", "LCAC", "ILCAC"))

        # 1. Loadcell & No of Loadcells
        if (code.startswith("ILDC") or ("LOADCELL" in desc or "LOAD CELL" in desc)) and not is_accessory:
            # Per-unit load cell count
            per_unit_lc = int(round(qty_val / max(default_wo_qty, 1)))
            lc_count = max(lc_count, per_unit_lc)
            m_model = re.search(r'\b(SC\s*[0-9]+[A-Z0-9]*|SP\s*[0-9]+[A-Z0-9]*|SB\s*[0-9]+[A-Z0-9]*|TH\s*[0-9]+[A-Z0-9]*|BR\s*[0-9]*|DS\s*[0-9]+[A-Z0-9]*|DLC\s*[0-9]+[A-Z0-9]*)\b', full_desc)
            lc_m = m_model.group(0).strip() if m_model else (model or desc)
            if lc_m and lc_m not in lc_models:
                lc_models.append(lc_m)
            m_cap = re.search(r'\b([0-9\.]+\s*(?:T|KG|Ton))\b', full_desc, re.I)
            if m_cap:
                c_val = m_cap.group(1).strip()
                if c_val not in lc_caps:
                    lc_caps.append(c_val)

        # 2. Indicator / Controller
        elif ("INDICATOR" in desc or "CONTROLLER" in desc or code.startswith("IINR")) and not any(w in desc for w in ["CLAMP", "BRACKET", "PLATE", "PEDESTAL", "STAND"]):
            m_ind = re.search(r'\b(DS\s*[0-9A-Z]+|LH\s*[0-9A-Z]+|BCW\s*(?:Series)?|MT4W[A-Z0-9\-]*|DI\s*[0-9A-Z]+)\b', full_desc, re.I)
            ind_m = m_ind.group(0).strip() if m_ind else (model or desc)
            spec["bs_model"] = ind_m
            spec["di_model"] = ind_m
            spec["indicator_model"] = ind_m

        # 3. Junction Box
        elif (code.startswith("IJBX") or ("JUNCTION BOX" in desc and not any(w in desc for w in ["BRACKET", "PLATE", "CARD", "CANOPY", "STAND", "STICKER", "GLAND"]))) and not code.startswith(("IMTB", "IMTP", "ICRD", "ISKR", "IPAE", "ICLG", "IBOL", "ICHN")):
            jb_norm = _normalize_jbox_model(desc, detail, model, code)
            if jb_norm and jb_norm not in jboxes:
                jboxes.append(jb_norm)

    if lc_count > 0:
        spec["lc_qty"] = str(lc_count)
    if lc_models:
        spec["load_cell_model"] = lc_models[0]
        spec["sensor_model"] = lc_models[0]
        spec["lc_model"] = lc_models[0]
    if lc_caps:
        spec["capacity"] = lc_caps[0]
        spec["lc_capacity"] = lc_caps[0]
    if jboxes:
        spec["jbox_model"] = jboxes[0]

    return spec

for idx, itm in enumerate(items):
    matched_wos = [w for w in wos if w.get("AutoNo") == itm.get("AutoNo")]
    total_wo_qty = sum(float(w.get("WOQty") or 1) for w in matched_wos)
    subcomponents = []
    for mw in matched_wos:
        subcomponents.extend(mw.get("subitems", []))
    
    spec = _extract_specs_from_text(f"{itm['ItemDesc']} {itm['AddDesc']} {itm['DetailDesc']}", default_job="SY0192", customer_name="Megha")
    spec = refined_enrich(spec, subcomponents, default_wo_qty=total_wo_qty)
    
    print(f"=== SYSTEM {idx+1}: {itm['DetailDesc'].splitlines()[0]} (Tag: {spec.get('tag_no')}) ===")
    print(f"   Indicator: {spec.get('indicator_model')}")
    print(f"   Load Cell: {spec.get('load_cell_model')} (Cap: {spec.get('capacity')}, Qty per bin: {spec.get('lc_qty')})")
    print(f"   Junction Box: {spec.get('jbox_model')}")
    print(f"   Matched WOs: {[w['WONo'] for w in matched_wos]}")
