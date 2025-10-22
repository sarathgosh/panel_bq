# panel_bq_app.py
# Single-file Streamlit app that auto-launches the server & browser when run directly.
# Features:
# - Option A/B: Manual rows or Excel upload
# - Option C: Totals -> auto-generate panels (with tail-merge to avoid 1-point panels)
# - 25% spare, DIN capacity check, PSU estimate, glands, enclosure suggestion
# - High-level devices: DPM, VAV, BTU, IAQ
# - Excel export (Panel_BQ, etc.) with reliable downloads via session_state

import io
import re
import os
import sys
from math import ceil, floor
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

import base64
import hashlib

def bytes_size_and_sha1(b: bytes) -> str:
    return f"{len(b)} bytes | sha1={hashlib.sha1(b).hexdigest()[:10]}"

def make_base64_xlsx_link(b: bytes, filename: str) -> str:
    b64 = base64.b64encode(b).decode()
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return f'<a download="{filename}" href="data:{mime};base64,{b64}">⬇ Download via fallback link</a>'


# ---------------- Assumptions / Config ----------------
ASSUMPTIONS = {
    # Spares & capacities
    "io_spare_pct": 0.25,
    "tb_spare_pct": 0.25,
    "di_module_capacity": 16,
    "do_module_capacity": 16,
    "ai_module_capacity": 8,
    "ao_module_capacity": 8,

    # DIN rail / layout heuristics
    "din_mm_per_terminal": 18,
    "din_mm_per_io_module_unit": 25,  # per 16 logical IO points (post-spare)
    "rail_pitch_mm": 120,
    "margin_mm": 30,
    "panel_fill_factor": 0.75,        # only use 75% of theoretical rail length

    # Electrical (PSU estimate only; BQ is quantity-only)
    "psu_voltage_V": 24,
    "psu_headroom_pct": 0.30,
    "typ_di_mA": 2,
    "typ_do_mA": 10,
    "typ_ai_mA": 5,
    "typ_ao_mA": 10,
    "modbus_device_typ_W": 2,
    "controller_base_W": 5,
    "misc_W": 10,
    "psu_std_sizes_A": [2.5, 5, 10, 20, 30, 40, 60, 80],

    # Device catalog (affects PSU + comms)
    "device_types": {
        "DPM": {"protocol": "modbus", "typ_power_W": 3, "item_code": "DEV-DPM", "desc": "Digital Power Meter"},
        "VAV": {"protocol": "modbus", "typ_power_W": 2, "item_code": "DEV-VAV", "desc": "VAV Controller"},
        "BTU": {"protocol": "modbus", "typ_power_W": 3, "item_code": "DEV-BTU", "desc": "BTU Meter"},
        "IAQ": {"protocol": "modbus", "typ_power_W": 2, "item_code": "DEV-IAQ", "desc": "Indoor Air Quality Sensor"},
    },

    # Gland heuristics
    "signals_per_gland": 8,
    "modbus_devices_per_gland": 4,
    "min_power_glands": 2,

    # Enclosure bins (Wmax, Hmax, Dmax, description)
    "enclosure_bins": [
        (600, 800, 250,  "Enclosure 600x800x250, IP54"),
        (800, 1200, 300, "Enclosure 800x1200x300, IP54"),
        (1000, 1400, 300,"Enclosure 1000x1400x300, IP54"),
        (1200, 1600, 400,"Enclosure 1200x1600x400, IP54"),
        (1600, 2000, 400,"Enclosure 1600x2000x400, IP54"),
    ],

    # Item codes
    "item_codes": {
        "controller": "CTRL-BASE",
        "di_module": "IOM-DI16",
        "do_module": "IOM-DO16",
        "ai_module": "IOM-AI8",
        "ao_module": "IOM-AO8",
        "psu": "PSU-24V",
        "terminal_block": "TB-2.5",
        "din_rail": "DIN-RAIL-TH35",
        "gland": "GLAND-M20",
        "enclosure": "ENC-STD",
        "modbus_gateway": "GW-MODBUS",
    },
}

INPUT_COLS = [
    "Project_Name","Panel_ID",
    "DI_Count","DO_Count","AI_Count","AO_Count",
    "Modbus_Device_Count",
    "DPM_Count","VAV_Count","BTU_Count","IAQ_Count",
    "Panel_Width_mm","Panel_Height_mm","Panel_Depth_mm",
]

# ---------------- Helpers ----------------
def number(x, default=0.0):
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)): return default
        s = str(x).strip()
        if s == "": return default
        return float(s)
    except Exception:
        return default

def choose_psu_size(required_A):
    for s in ASSUMPTIONS["psu_std_sizes_A"]:
        if s >= required_A:
            return s
    return ASSUMPTIONS["psu_std_sizes_A"][-1]

def pick_enclosure_desc(w, h, d):
    for Wmax, Hmax, Dmax, desc in ASSUMPTIONS["enclosure_bins"]:
        if w <= Wmax and h <= Hmax and d <= Dmax:
            return desc
    return f"Custom Enclosure {int(w)}x{int(h)}x{int(d)}"

def usable_din_length_mm(w, h):
    rail_pitch = ASSUMPTIONS["rail_pitch_mm"]
    margin = ASSUMPTIONS["margin_mm"]
    rails = max(1, floor((h - 2*margin) / rail_pitch))
    usable_width = max(0, (w - 2*margin))
    total = rails * usable_width * ASSUMPTIONS["panel_fill_factor"]
    return max(0, total), rails

def capacity_check_and_suggest(w, h, d, required_mm):
    usable_mm, _ = usable_din_length_mm(w, h)
    if required_mm <= usable_mm:
        return f"✅ OK (Used {int(required_mm)} / {int(usable_mm)} mm DIN)", pick_enclosure_desc(w, h, d)
    # Try larger bins
    for Wmax, Hmax, Dmax, desc in ASSUMPTIONS["enclosure_bins"]:
        usable_bin_mm, _ = usable_din_length_mm(Wmax, Hmax)
        if required_mm <= usable_bin_mm:
            return (f"⚠️ Too small. Need ~{int(required_mm - usable_mm)} mm more. "
                    f"Suggested: {desc} (usable ≈ {int(usable_bin_mm)} mm)"), desc
    # Custom fallback
    return "⚠️ Too small. Please engineer a custom enclosure.", "Custom Enclosure"

def sanitize_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", str(name))
    return name[:31] if len(name) > 31 else name

# DIN estimate for a raw allocation
def estimate_din_for_allocation(di_raw, do_raw, ai_raw, ao_raw):
    spare = ASSUMPTIONS["io_spare_pct"]
    di_eff, do_eff, ai_eff, ao_eff = [ceil(v*(1+spare)) for v in (di_raw,do_raw,ai_raw,ao_raw)]
    io_units = ceil((di_eff + do_eff + ai_eff + ao_eff) / 16)  # 16 logical points per "unit"
    din_for_io = io_units * ASSUMPTIONS["din_mm_per_io_module_unit"]
    tb_total = ceil((di_raw + do_raw + ai_raw + ao_raw) * (1 + ASSUMPTIONS["tb_spare_pct"]))
    din_for_tb = tb_total * ASSUMPTIONS["din_mm_per_terminal"]
    return din_for_io + din_for_tb

def panel_raw_totals(p):
    return int(p["DI_Count"]) + int(p["DO_Count"]) + int(p["AI_Count"]) + int(p["AO_Count"])

def panel_required_din(p):
    return estimate_din_for_allocation(int(p["DI_Count"]), int(p["DO_Count"]),
                                       int(p["AI_Count"]), int(p["AO_Count"]))

# Auto-split totals into multiple panels of given size, then squeeze tail panels
def autosplit_panels(project, base_prefix, totals, size):
    W, H, D = size
    usable_mm, _ = usable_din_length_mm(W, H)
    if usable_mm <= 0:
        raise ValueError("Panel usable DIN length is zero/negative. Check size.")

    rem = totals.copy()
    panels = []
    idx = 1

    # A) Allocate panels until IO placed (binary-search proportion per panel)
    while any(rem[k] > 0 for k in ("DI","DO","AI","AO")):
        lo, hi = 0.0, 1.0
        best = {"DI":0,"DO":0,"AI":0,"AO":0}
        for _ in range(24):
            s = (lo + hi) / 2.0
            guess = {k: int(max(0, rem[k] * s)) for k in ("DI","DO","AI","AO")}
            for k in ("DI","DO","AI","AO"):
                if rem[k] > 0 and guess[k] == 0:
                    guess[k] = 1
            din = estimate_din_for_allocation(guess["DI"], guess["DO"], guess["AI"], guess["AO"])
            if din <= usable_mm and all(guess[k] <= rem[k] for k in guess):
                best = guess; lo = s
            else:
                hi = s

        # If nothing fits, try single-point placement
        if sum(best.values()) == 0:
            placed = False
            for k in ("DI","DO","AI","AO"):
                if rem[k] > 0:
                    din = estimate_din_for_allocation(1 if k=="DI" else 0,
                                                      1 if k=="DO" else 0,
                                                      1 if k=="AI" else 0,
                                                      1 if k=="AO" else 0)
                    if din <= usable_mm:
                        best = {"DI":0,"DO":0,"AI":0,"AO":0}; best[k] = 1
                        placed = True; break
            if not placed:
                raise ValueError("Panel size too small to place any remaining I/O with current heuristics.")

        for k in ("DI","DO","AI","AO"):
            rem[k] -= best[k]

        panels.append({
            "Project_Name": project,
            "Panel_ID": f"{base_prefix}-{idx:02d}",
            "DI_Count": best["DI"],
            "DO_Count": best["DO"],
            "AI_Count": best["AI"],
            "AO_Count": best["AO"],
            "Panel_Width_mm": W,
            "Panel_Height_mm": H,
            "Panel_Depth_mm": D
        })
        idx += 1

    # B) Distribute devices evenly across panels
    n = max(1, len(panels))
    for key in ("Modbus","DPM","VAV","BTU","IAQ"):
        total = int(totals.get(key, 0))
        base = total // n
        extra = total % n
        for i, p in enumerate(panels):
            p[f"{key}_Count"] = base + (1 if i < extra else 0)

    # C) Squeeze small tail panel into earlier panels if possible
    MIN_RAW_IO_TO_KEEP = 8
    changed = True
    while changed and len(panels) >= 2:
        changed = False
        last = panels[-1]
        if panel_raw_totals(last) >= MIN_RAW_IO_TO_KEEP:
            break
        need = {k: int(last[f"{k}_Count"]) for k in ("DI","DO","AI","AO")}
        for j in range(len(panels)-2, -1, -1):
            if sum(need.values()) == 0: break
            pj = panels[j]
            usable_j, _ = usable_din_length_mm(pj["Panel_Width_mm"], pj["Panel_Height_mm"])

            def fits_with_add(di_add, do_add, ai_add, ao_add):
                tmp = dict(pj)
                tmp["DI_Count"] = int(tmp["DI_Count"]) + di_add
                tmp["DO_Count"] = int(tmp["DO_Count"]) + do_add
                tmp["AI_Count"] = int(tmp["AI_Count"]) + ai_add
                tmp["AO_Count"] = int(tmp["AO_Count"]) + ao_add
                return panel_required_din(tmp) <= usable_j

            add = need.copy()
            while sum(add.values()) > 0 and not fits_with_add(add["DI"], add["DO"], add["AI"], add["AO"]):
                if add["DI"] > 0: add["DI"] -= 1
                elif add["DO"] > 0: add["DO"] -= 1
                elif add["AI"] > 0: add["AI"] -= 1
                elif add["AO"] > 0: add["AO"] -= 1

            if sum(add.values()) > 0 and fits_with_add(add["DI"], add["DO"], add["AI"], add["AO"]):
                pj["DI_Count"] = int(pj["DI_Count"]) + add["DI"]
                pj["DO_Count"] = int(pj["DO_Count"]) + add["DO"]
                pj["AI_Count"] = int(pj["AI_Count"]) + add["AI"]
                pj["AO_Count"] = int(pj["AO_Count"]) + add["AO"]
                need["DI"] -= add["DI"]; need["DO"] -= add["DO"]; need["AI"] -= add["AI"]; need["AO"] -= add["AO"]

        if sum(need.values()) == 0:
            panels.pop(); changed = True
        else:
            last["DI_Count"] = need["DI"]; last["DO_Count"] = need["DO"]
            last["AI_Count"] = need["AI"]; last["AO_Count"] = need["AO"]
            break

    return panels

def build_panel_rows(rec: dict):
    proj = str(rec.get("Project_Name", ""))
    pid  = str(rec.get("Panel_ID", "PANEL"))
    di,do,ai,ao = map(lambda k:number(rec.get(k),0),["DI_Count","DO_Count","AI_Count","AO_Count"])
    mb_base = number(rec.get("Modbus_Device_Count"),0)
    dpm = number(rec.get("DPM_Count"),0)
    vav = number(rec.get("VAV_Count"),0)
    btu = number(rec.get("BTU_Count"),0)
    iaq = number(rec.get("IAQ_Count"),0)
    W,H,D = map(lambda k:number(rec.get(k),0),["Panel_Width_mm","Panel_Height_mm","Panel_Depth_mm"])

    spare = ASSUMPTIONS["io_spare_pct"]
    di_eff, do_eff, ai_eff, ao_eff = [ceil(v*(1+spare)) for v in (di,do,ai,ao)]
    n_di = ceil(di_eff / ASSUMPTIONS["di_module_capacity"]) if di_eff>0 else 0
    n_do = ceil(do_eff / ASSUMPTIONS["do_module_capacity"]) if do_eff>0 else 0
    n_ai = ceil(ai_eff / ASSUMPTIONS["ai_module_capacity"]) if ai_eff>0 else 0
    n_ao = ceil(ao_eff / ASSUMPTIONS["ao_module_capacity"]) if ao_eff>0 else 0

    tb_total = ceil((di + do + ai + ao) * (1 + ASSUMPTIONS["tb_spare_pct"]))

    dev_map = ASSUMPTIONS["device_types"]
    device_counts = {"DPM":dpm, "VAV":vav, "BTU":btu, "IAQ":iaq}

    total_modbus_devices = mb_base
    pwr_w_devices = 0.0
    for name, cnt in device_counts.items():
        cfg = dev_map.get(name, {})
        if cfg.get("protocol") == "modbus":
            total_modbus_devices += cnt
            pwr_w_devices += (cfg.get("typ_power_W", 0.0) * cnt)

    V = ASSUMPTIONS["psu_voltage_V"]
    I_io = (ASSUMPTIONS["typ_di_mA"]*di_eff + ASSUMPTIONS["typ_do_mA"]*do_eff +
            ASSUMPTIONS["typ_ai_mA"]*ai_eff + ASSUMPTIONS["typ_ao_mA"]*ao_eff) / 1000.0
    P_modbus_total = ASSUMPTIONS["modbus_device_typ_W"] * mb_base + pwr_w_devices
    I_modbus = P_modbus_total / V if V>0 else 0.0
    I_ctrl_misc = (ASSUMPTIONS["controller_base_W"] + ASSUMPTIONS["misc_W"]) / V if V>0 else 0.0
    I_total = (I_io + I_modbus + I_ctrl_misc) * (1 + ASSUMPTIONS["psu_headroom_pct"])
    psu_A = choose_psu_size(I_total)

    glands = ceil((di+do+ai+ao) / ASSUMPTIONS["signals_per_gland"]) + \
             ceil(total_modbus_devices / ASSUMPTIONS["modbus_devices_per_gland"]) + \
             ASSUMPTIONS["min_power_glands"]

    io_units = ceil((di_eff + do_eff + ai_eff + ao_eff) / 16)
    din_for_io = io_units * ASSUMPTIONS["din_mm_per_io_module_unit"]
    din_for_tb = tb_total * ASSUMPTIONS["din_mm_per_terminal"]
    total_din = din_for_io + din_for_tb

    check_msg, suggested_enc = capacity_check_and_suggest(W, H, D, total_din)
    current_enc = pick_enclosure_desc(W, H, D)

    ic = ASSUMPTIONS["item_codes"]
    rows = [
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["controller"],      Description="Base Controller",           UOM="No", Qty=1),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["di_module"],       Description=f"DI Module {ASSUMPTIONS['di_module_capacity']}ch", UOM="No", Qty=n_di),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["do_module"],       Description=f"DO Module {ASSUMPTIONS['do_module_capacity']}ch", UOM="No", Qty=n_do),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["ai_module"],       Description=f"AI Module {ASSUMPTIONS['ai_module_capacity']}ch", UOM="No", Qty=n_ai),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["ao_module"],       Description=f"AO Module {ASSUMPTIONS['ao_module_capacity']}ch", UOM="No", Qty=n_ao),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["psu"],             Description=f"24VDC PSU ≥ {psu_A}A",     UOM="No", Qty=1),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["terminal_block"],  Description="Terminal Block 2.5 mm²",   UOM="No", Qty=tb_total),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["gland"],           Description="Cable Glands (signal/power/MB)", UOM="No", Qty=glands),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["din_rail"],        Description="DIN Rail TH35 (total used mm)", UOM="mm", Qty=int(total_din)),
        dict(Project=proj, Panel_ID=pid, Item_Code=ic["enclosure"],       Description=current_enc,                 UOM="No", Qty=1),
    ]
    if total_modbus_devices > 0:
        rows.append(dict(Project=proj, Panel_ID=pid, Item_Code=ic["modbus_gateway"], Description="Modbus RTU/RS485 Gateway", UOM="No", Qty=1))

    # High-level devices as separate BQ lines
    for name, cnt in device_counts.items():
        cfg = dev_map.get(name, {})
        if cnt > 0:
            rows.append(dict(Project=proj, Panel_ID=pid,
                             Item_Code=cfg.get("item_code", f"DEV-{name}"),
                             Description=cfg.get("desc", name), UOM="No", Qty=int(cnt)))

    rows.append(dict(Project=proj, Panel_ID=pid, Item_Code="", Description="Check_Result", UOM="", Qty="", Remarks=check_msg))
    rows.append(dict(Project=proj, Panel_ID=pid, Item_Code="", Description="Suggested_Enclosure", UOM="", Qty="", Remarks=suggested_enc))
    return rows

def build_workbook(df_in: pd.DataFrame) -> io.BytesIO:
    # Ensure columns
    for c in INPUT_COLS:
        if c not in df_in.columns:
            df_in[c] = np.nan
    df_in = df_in[INPUT_COLS].copy()

    all_rows, checks, per_panel = [], [], {}
    for rec in df_in.to_dict(orient="records"):
        rows = build_panel_rows(rec)
        all_rows.extend(rows)
        pid = str(rec.get("Panel_ID", "PANEL"))
        per_panel[pid] = pd.DataFrame([r for r in rows if r["Description"] not in ("Check_Result","Suggested_Enclosure")])
        for cr in [r for r in rows if r["Description"] in ("Check_Result","Suggested_Enclosure")]:
            checks.append({
                "Project": cr.get("Project",""),
                "Panel_ID": cr.get("Panel_ID",""),
                "Type": cr["Description"],
                "Details": cr.get("Remarks","")
            })

    bq_df = pd.DataFrame(all_rows, columns=["Project","Panel_ID","Item_Code","Description","UOM","Qty","Remarks"])
    summ = bq_df[bq_df["Item_Code"] != ""].copy()
    summ["Qty_num"] = pd.to_numeric(summ["Qty"], errors="coerce")
    summary_df = (summ.groupby(["Item_Code","Description","UOM"], dropna=False)["Qty_num"]
                     .sum(min_count=1).reset_index()
                     .rename(columns={"Qty_num":"Total_Qty"})
                     .sort_values(["Item_Code","Description"]))
    checks_df = pd.DataFrame(checks, columns=["Project","Panel_ID","Type","Details"])

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df_in.to_excel(w, sheet_name="Input (as used)", index=False)
        bq_df.to_excel(w, sheet_name="Panel_BQ", index=False)
        summary_df.to_excel(w, sheet_name="Summary", index=False)
        checks_df.to_excel(w, sheet_name="Checks", index=False)
        for pid, pdf in per_panel.items():
            pdf.to_excel(w, sheet_name=sanitize_sheet_name(pid), index=False)
    out.seek(0)
    return out

# ---------------- Streamlit UI ----------------
def main():
    st.set_page_config(page_title="Panel BQ Generator", layout="wide")
    st.title("🧰 Panel BQ Auto-Generator (25% spare • DIN check • devices)")

    with st.sidebar:
        st.header("Settings")
        ASSUMPTIONS["io_spare_pct"] = st.slider("I/O spare (%)", 0.0, 0.5, 0.25, 0.05)
        ASSUMPTIONS["tb_spare_pct"] = st.slider("Terminal spare (%)", 0.0, 0.5, 0.25, 0.05)
        ASSUMPTIONS["panel_fill_factor"] = st.slider("Panel fill factor", 0.5, 0.95, 0.75, 0.01)
        st.caption("Module capacities & device powers are in-code defaults.")

    mode = st.radio("Choose input mode:",
                    ["Option A — Manual rows", "Option B — Upload Excel", "Option C — Totals → Auto-generate panels"])

    st.divider()

    # -------- Option A
    if mode == "Option A — Manual rows":
        st.subheader("Option A — Enter per-panel rows")
        st.caption("Edit the table below. Add/remove rows as needed.")
        sample = pd.DataFrame([{
            "Project_Name":"Zealcorps","Panel_ID":"ASP-01",
            "DI_Count":120,"DO_Count":64,"AI_Count":24,"AO_Count":16,
            "Modbus_Device_Count":8,"DPM_Count":2,"VAV_Count":6,"BTU_Count":1,"IAQ_Count":4,
            "Panel_Width_mm":800,"Panel_Height_mm":1200,"Panel_Depth_mm":300,
        }])
        df_edit = st.data_editor(sample, num_rows="dynamic", width="stretch")
        if st.button("Generate BQ (Option A)"):
            out = build_workbook(df_edit)
            out.seek(0)
            st.session_state["xlsx_bytes_A"] = out.getvalue()
            st.session_state["xlsx_name_A"] = f"Panel_BQ_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            st.success("Workbook generated. Use the Download button below.")

        if "xlsx_bytes_A" in st.session_state:
            st.download_button("⬇ Download Excel",
                               data=st.session_state["xlsx_bytes_A"],
                               file_name=st.session_state.get("xlsx_name_A","Panel_BQ.xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_optA")

    # -------- Option B
    elif mode == "Option B — Upload Excel":
        st.subheader("Option B — Upload Excel")
        st.caption("Upload a workbook with sheet **Input** (or first sheet) containing required columns.")
        f = st.file_uploader("Choose .xlsx/.xls", type=["xlsx","xls"])

        if st.button("Generate BQ (Option B)"):
            if not f:
                st.error("Please upload an Excel file.")
            else:
                try:
                    try:
                        df_in = pd.read_excel(f, sheet_name="Input")
                    except Exception:
                        f.seek(0)
                        df_in = pd.read_excel(f)
                    out = build_workbook(df_in)
                    out.seek(0)
                    st.session_state["xlsx_bytes_B"] = out.getvalue()  # persist bytes
                    st.session_state["xlsx_name_B"] = f"Panel_BQ_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                    st.success("Workbook generated. Use the Download button below.")
                except Exception as e:
                    st.exception(e)

        if "xlsx_bytes_B" in st.session_state:
            st.download_button(
                "⬇ Download Excel",
                data=st.session_state["xlsx_bytes_B"],
                file_name=st.session_state.get("xlsx_name_B", "Panel_BQ.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_optB",
            )

    # -------- Option C
    else:
        st.subheader("Option C — Totals → Auto-generate panels")
        with st.expander("Column descriptions", expanded=False):
            st.markdown("""
            - **Project_Name**: Label used in BQ rows  
            - **Panel Prefix**: Auto IDs (e.g., `ASP` → `ASP-01`, `ASP-02`, …)  
            - **Totals**: RAW DI/DO/AI/AO points (tool adds **+25% spare** automatically)  
            - **Total Modbus (generic)**: Other field devices beyond DPM/VAV/BTU/IAQ  
            - **DPM/VAV/BTU/IAQ**: High-level devices (BQ lines, PSU load, comms glands)  
            - **Panel size (mm)**: Same size applied to every generated panel  
            """)

        c1,c2,c3,c4 = st.columns(4)
        project = c1.text_input("Project_Name", "Zealcorps")
        prefix  = c2.text_input("Panel Prefix", "ASP")
        width   = c3.number_input("Panel Width (mm)", value=800, min_value=200, step=50)
        height  = c4.number_input("Panel Height (mm)", value=1200, min_value=400, step=50)
        c5,c6,c7 = st.columns(3)
        depth   = c5.number_input("Panel Depth (mm)", value=300, min_value=150, step=50)
        total_di= c6.number_input("Total DI", value=240, min_value=0, step=1)
        total_do= c7.number_input("Total DO", value=128, min_value=0, step=1)
        c8,c9,c10 = st.columns(3)
        total_ai= c8.number_input("Total AI", value=48, min_value=0, step=1)
        total_ao= c9.number_input("Total AO", value=24, min_value=0, step=1)
        total_mb= c10.number_input("Total Modbus (generic)", value=12, min_value=0, step=1)
        c11,c12,c13,c14 = st.columns(4)
        total_dpm = c11.number_input("DPM", value=4, min_value=0, step=1)
        total_vav = c12.number_input("VAV", value=12, min_value=0, step=1)
        total_btu = c13.number_input("BTU", value=2, min_value=0, step=1)
        total_iaq = c14.number_input("IAQ", value=8, min_value=0, step=1)

    if st.button("Generate from Totals (Option C)"):
    try:
        # ... your existing totals/panels/prev_df code ...

        out = build_workbook(prev_df)
        out.seek(0)
        file_bytes = out.getvalue()                           # <-- bytes now
        st.session_state["xlsx_bytes_C"] = file_bytes
        st.session_state["xlsx_name_C"] = f"Panel_BQ_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

        # show a tiny debug line so we know we actually have bytes
        st.caption(f"Excel ready: {bytes_size_and_sha1(file_bytes)}")
        st.success("Workbook generated. Use the buttons below.")
    except Exception as e:
        st.exception(e)

# render the download button & fallback link OUTSIDE the button branch
if "xlsx_bytes_C" in st.session_state:
    st.download_button(
        "⬇ Download Excel (Option C)",
        data=st.session_state["xlsx_bytes_C"],
        file_name=st.session_state.get("xlsx_name_C", "Panel_BQ.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_optC",
    )

    # Fallback base64 link (very reliable in Edge/Chrome/Firefox)
    st.markdown(
        make_base64_xlsx_link(
            st.session_state["xlsx_bytes_C"],
            st.session_state.get("xlsx_name_C", "Panel_BQ.xlsx")
        ),
        unsafe_allow_html=True
    )    

# --------------- Auto-launch Streamlit when run directly ---------------
if __name__ == "__main__":
    # If it's already inside Streamlit runner, just run main()
    if os.environ.get("STREAMLIT_SERVER_RUNNING") == "1":
        main()
    else:
        try:
            # Newer Streamlit entrypoint
            from streamlit.web.cli import main as st_main
            sys.argv = ["streamlit", "run", os.path.abspath(__file__)]
            os.environ["STREAMLIT_SERVER_RUNNING"] = "1"
            sys.exit(st_main())
        except Exception:
            # Fallback for some older versions
            from streamlit.web import bootstrap
            os.environ["STREAMLIT_SERVER_RUNNING"] = "1"
            bootstrap.run(os.path.abspath(__file__), "", [], {})
else:
    main()



