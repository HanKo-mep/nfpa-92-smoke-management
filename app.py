"""
NFPA 92 / 92B Smoke Management Systems – Compliance Checklist & Calculator
===========================================================================
Algebraic calculation methods for smoke management in malls, atria and
large-volume spaces (SI units).

NFPA 92B was withdrawn in 2011 and incorporated into NFPA 92
(Standard for Smoke Control Systems). The equations and design approaches
implemented here follow the classic NFPA 92B algebraic methods retained
in NFPA 92 Chapter 5, expressed entirely in SI units.

Educational / preliminary-design aid only.
Always obtain and apply the official NFPA standard and local code requirements.
"""

from __future__ import annotations

import json
import math
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Paths & data
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
EXPORTS_DIR = ROOT / "exports"
EXPORTS_DIR.mkdir(exist_ok=True)

with open(DATA_DIR / "reference_data.json", encoding="utf-8") as f:
    REF = json.load(f)

# ---------------------------------------------------------------------------
# Physical helpers (SI)
# ---------------------------------------------------------------------------
def rho_air(T_C: float, P_Pa: float = 101325.0) -> float:
    """Dry air density (kg/m³) from ideal-gas law."""
    T_K = T_C + 273.15
    return P_Pa / (287.058 * T_K)


def smoke_density(T_smoke_C: float, T_amb_C: float = 20.0, rho_amb: float = 1.2) -> float:
    """Approximate smoke density assuming ideal gas and constant pressure."""
    return rho_amb * (T_amb_C + 273.15) / (T_smoke_C + 273.15)


def flame_height_axisym(Qc_kW: float) -> float:
    """
    Limiting elevation / approx. mean flame height (m) for axisymmetric plume.
    NFPA 92 SI form: zl = 0.166 * Qc**(2/5)  (Qc in kW, zl in m).
    """
    if Qc_kW <= 0:
        return 0.0
    return 0.166 * (Qc_kW ** 0.4)


def mass_flow_axisym(Qc_kW: float, z_m: float) -> Tuple[float, str]:
    """
    Axisymmetric plume mass flow rate (kg/s) at height z above fuel base.
    Returns (m_dot, equation_used).
    SI form retained from NFPA 92B / NFPA 92:
      above flame:  m = 0.071 * Qc**(1/3) * z**(5/3) + 0.0018 * Qc
      in flame:     m = 0.032 * Qc**(3/5) * z**(1/5)
    """
    if Qc_kW <= 0 or z_m <= 0:
        return 0.0, "n/a"
    zl = flame_height_axisym(Qc_kW)
    if z_m > zl:
        m = 0.071 * (Qc_kW ** (1.0 / 3.0)) * (z_m ** (5.0 / 3.0)) + 0.0018 * Qc_kW
        return m, "above flame (0.071 Qc^{1/3} z^{5/3} + 0.0018 Qc)"
    else:
        m = 0.032 * (Qc_kW ** 0.6) * (z_m ** 0.2)
        return m, "in flame (0.032 Qc^{3/5} z^{1/5})"


def volumetric_flow(m_dot_kg_s: float, rho_smoke: float) -> float:
    """Volumetric flow rate (m³/s)."""
    if rho_smoke <= 0:
        return 0.0
    return m_dot_kg_s / rho_smoke


def avg_plume_temp_rise(Qc_kW: float, m_dot_kg_s: float, cp: float = 1.0) -> float:
    """
    Average temperature rise of plume gases (°C) from energy balance:
    ΔT = Qc / (m * cp)   with Qc in kW, m in kg/s, cp in kJ/kg·K.
    """
    if m_dot_kg_s <= 0:
        return 0.0
    return Qc_kW / (m_dot_kg_s * cp)


def filling_time_steady(
    H_m: float,
    A_m2: float,
    Q_kW: float,
    z_target_m: float,
) -> float:
    """
    Approximate time (s) for first indication of smoke to descend to z
    for a steady fire in an unvented space (NFPA 92B-style filling equation).
    Simplified SI form of the classic relation (order-of-magnitude / screening).
    """
    if H_m <= 0 or A_m2 <= 0 or Q_kW <= 0 or z_target_m >= H_m:
        return 0.0
    # Classic form rearranged for SI (approximate coefficients from literature)
    # z/H = 0.67 - 0.28 ln( (t Q^{1/3} / H^{4/3}) / (A/H²) )
    # Solve for t (screening only)
    try:
        ratio = z_target_m / H_m
        if ratio >= 0.67:
            return 0.0
        arg = math.exp((0.67 - ratio) / 0.28)
        # t ≈ arg * (A/H²) * H^{4/3} / Q^{1/3}
        t = arg * (A_m2 / (H_m ** 2)) * (H_m ** (4.0 / 3.0)) / (Q_kW ** (1.0 / 3.0))
        return max(0.0, t)
    except (ValueError, ZeroDivisionError):
        return 0.0


def max_vol_per_inlet_no_plugholing(
    d_m: float,
    T_smoke_C: float,
    T_amb_C: float = 20.0,
    gamma: float = 1.0,
) -> float:
    """
    Approximate maximum volumetric flow (m³/s) per exhaust inlet without
    plugholing. Simplified form of the NFPA 92B / 92 relation.
    d = depth of smoke layer below the inlet (m).
    """
    if d_m <= 0:
        return 0.0
    T_s = T_smoke_C + 273.15
    T_o = T_amb_C + 273.15
    # Engineering approximation consistent with published SI guidance
    # V_max ≈ γ * d^{5/2} * ((T_s - T_o)/T_o)^{1/2} * constant
    # Using a practical coefficient that yields realistic order of magnitude
    # (users should verify against the exact edition they are applying).
    factor = ((T_s - T_o) / T_o) ** 0.5 if T_s > T_o else 0.1
    V_max = gamma * 0.54 * (d_m ** 2.5) * factor   # m³/s, calibrated order-of-magnitude
    return max(0.0, V_max)


def number_of_inlets_required(V_total_m3s: float, V_max_per_inlet: float) -> int:
    if V_max_per_inlet <= 0:
        return 1
    return max(1, math.ceil(V_total_m3s / V_max_per_inlet))


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------
thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
SECTION_FILL = PatternFill("solid", fgColor="D6EAF8")
PASS_FILL = PatternFill("solid", fgColor="C6EFCE")
FAIL_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")


def style_header(ws, row: int, cols: int):
    font = Font(bold=True, color="FFFFFF", size=11)
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def style_section(ws, row: int, cols: int):
    font = Font(bold=True, size=11)
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = SECTION_FILL
        cell.font = font


def auto_width(ws, min_w: int = 12, max_w: int = 55):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_w), max_w)


def build_excel_report(
    project: Dict[str, Any],
    geometry: Dict[str, Any],
    fire: Dict[str, Any],
    results: Dict[str, Any],
    checklist_status: Dict[str, List[bool]],
) -> bytes:
    wb = Workbook()

    # --- Cover ---
    ws = wb.active
    ws.title = "Cover"
    ws["A1"] = "NFPA 92 / 92B Smoke Management – Calculation & Compliance Report"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E79")
    ws["A3"] = "Project"
    ws["B3"] = project.get("name", "")
    ws["A4"] = "Location"
    ws["B4"] = project.get("location", "")
    ws["A5"] = "Prepared by"
    ws["B5"] = project.get("prepared_by", "")
    ws["A6"] = "Date"
    ws["B6"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    ws["A8"] = "Standard basis"
    ws["B8"] = REF["standard"]["name"]
    ws["A9"] = "Note"
    ws["B9"] = REF["standard"]["note"]
    ws["A11"] = "DISCLAIMER"
    ws["A11"].font = Font(bold=True, color="C00000")
    ws["A12"] = (
        "This report is generated by an educational / preliminary-design tool. "
        "It does not replace a formal engineered design. Final designs must be "
        "produced by competent professionals, verified, and approved by the AHJ. "
        "Always consult the official NFPA 92 standard in force for the project."
    )
    ws.merge_cells("A12:F15")
    auto_width(ws)

    # --- Inputs ---
    ws = wb.create_sheet("Inputs")
    ws["A1"] = "Design Inputs (SI units)"
    ws["A1"].font = Font(bold=True, size=14)
    row = 3
    for section, data in [
        ("Project", project),
        ("Geometry", geometry),
        ("Design Fire", fire),
    ]:
        ws.cell(row=row, column=1, value=section)
        style_section(ws, row, 2)
        row += 1
        for k, v in data.items():
            ws.cell(row=row, column=1, value=k)
            ws.cell(row=row, column=2, value=str(v))
            row += 1
        row += 1
    auto_width(ws)

    # --- Calculation Summary ---
    ws = wb.create_sheet("Calculation Summary")
    ws["A1"] = "Calculation Results (SI)"
    ws["A1"].font = Font(bold=True, size=14)
    headers = ["Parameter", "Value", "Unit", "Notes"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=3, column=c, value=h)
    style_header(ws, 3, 4)
    row = 4
    for item in results.get("summary_rows", []):
        for c, val in enumerate(item, 1):
            ws.cell(row=row, column=c, value=val)
        row += 1
    auto_width(ws)

    # --- Checklist ---
    ws = wb.create_sheet("Data Checklist")
    ws["A1"] = "Compliance Data Checklist Status"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Tick = information provided / requirement addressed"
    row = 4
    for cat in REF["checklist_categories"]:
        ws.cell(row=row, column=1, value=cat["title"])
        style_section(ws, row, 3)
        row += 1
        statuses = checklist_status.get(cat["id"], [])
        for i, item in enumerate(cat["items"]):
            done = statuses[i] if i < len(statuses) else False
            ws.cell(row=row, column=1, value="✓" if done else "☐")
            ws.cell(row=row, column=2, value=item)
            if done:
                ws.cell(row=row, column=1).fill = PASS_FILL
            row += 1
        row += 1
    auto_width(ws)

    # --- Reference ---
    ws = wb.create_sheet("Reference")
    ws["A1"] = "Key Reference Values (SI)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = "Plume equations (axisymmetric)"
    ws["A4"] = REF["plume_equations_SI"]["axisymmetric_above_flame"]
    ws["A5"] = REF["plume_equations_SI"]["axisymmetric_in_flame"]
    ws["A6"] = REF["plume_equations_SI"]["flame_height_approx"]
    ws["A8"] = "Minimum clear height above highest walking surface"
    ws["B8"] = f"{REF['smoke_layer']['min_clear_height_m']} m"
    ws["A9"] = "Minimum smoke-layer depth (fraction of H)"
    ws["B9"] = f"{REF['smoke_layer']['min_layer_depth_fraction_of_H'] * 100:.0f} %"
    ws["A10"] = "Typical makeup-air fraction of exhaust"
    ws["B10"] = "85 % – 95 %"
    ws["A11"] = "Max recommended makeup-air velocity"
    ws["B11"] = f"{REF['makeup_air']['max_velocity_m_s']} m/s"
    auto_width(ws)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="NFPA 92/92B Smoke Management",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🔥 NFPA 92 / 92B Smoke Management Systems")
st.caption(
    "Compliance data checklist & algebraic calculations for malls, atria and large-volume spaces · SI units"
)

with st.expander("About this tool & standard status", expanded=False):
    st.markdown(
        f"""
**{REF['standard']['name']}**  
{REF['standard']['title']}

{REF['standard']['note']}

{REF['standard']['edition_note']}

**Scope of calculations**
- Design-fire definition (steady or t² growth)
- Axisymmetric plume mass-flow & required exhaust rate
- Approximate smoke-layer filling time (unvented)
- Minimum layer depth & clear-height checks
- Exhaust-inlet count (plugholing screening)
- Makeup-air quantity guidance
- Full compliance data checklist aligned with NFPA 92 design documentation expectations

**Limitations**  
This is an educational / preliminary-design aid. Algebraic methods have geometric and fire-size limitations. Complex geometries, balcony-spill / window plumes, stratification, and tenability analyses often require zone models or CFD. Always apply the official NFPA 92 edition adopted by the project and obtain AHJ approval.
"""
    )

# Sidebar – project info
st.sidebar.header("Project")
proj_name = st.sidebar.text_input("Project name", value="Atrium Smoke Control – Example")
proj_loc = st.sidebar.text_input("Location", value="")
prep_by = st.sidebar.text_input("Prepared by", value="")
st.sidebar.markdown("---")
st.sidebar.info("All inputs and results are in **SI units** (m, m², m³, kW, kg/s, m³/s, °C).")

# Tabs
tab_calc, tab_check, tab_test, tab_ref, tab_export = st.tabs(
    ["📐 Calculations", "📋 Data Checklist", "✅ Testing & Acceptance", "📚 Reference Data", "📄 Export Report"]
)

# ---------------------------------------------------------------------------
# TAB 1 – Calculations
# ---------------------------------------------------------------------------
with tab_calc:
    st.subheader("1. Geometry of large-volume space")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        length_m = st.number_input("Length (m)", min_value=1.0, value=40.0, step=1.0)
    with c2:
        width_m = st.number_input("Width (m)", min_value=1.0, value=30.0, step=1.0)
    with c3:
        height_m = st.number_input("Ceiling height H above fire base (m)", min_value=3.0, value=25.0, step=0.5)
    with c4:
        highest_walk_m = st.number_input(
            "Highest walking surface above fire base (m)",
            min_value=0.0,
            value=18.0,
            step=0.5,
            help="Used to set minimum clear height = walking surface + 1.83 m",
        )

    floor_area_m2 = length_m * width_m
    volume_m3 = floor_area_m2 * height_m
    st.metric("Floor area A", f"{floor_area_m2:,.0f} m²")
    st.metric("Approximate volume", f"{volume_m3:,.0f} m³")

    st.subheader("2. Design fire")
    fire_mode = st.radio("Fire type", ["Steady-state", "t-squared growth"], horizontal=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        if fire_mode == "Steady-state":
            Q_kW = st.number_input("Total heat release rate Q (kW)", min_value=100.0, value=2500.0, step=100.0)
        else:
            alpha = st.selectbox(
                "Growth rate α (kW/s²)",
                options=list(REF["design_fires"]["t_squared_alpha_kW_s2"].keys()),
                format_func=lambda k: f"{k} ({REF['design_fires']['t_squared_alpha_kW_s2'][k]} kW/s²)",
            )
            alpha_val = REF["design_fires"]["t_squared_alpha_kW_s2"][alpha]
            t_growth_s = st.number_input("Time to reach steady HRR (s)", min_value=10.0, value=150.0, step=10.0)
            Q_kW = alpha_val * (t_growth_s ** 2)
            st.write(f"Steady HRR after growth: **{Q_kW:.0f} kW**")
    with c2:
        chi_c = st.slider(
            "Convective fraction χc",
            min_value=0.50,
            max_value=0.80,
            value=float(REF["design_fires"]["convective_fraction_default"]),
            step=0.05,
            help="Typical range 0.6–0.7. Qc = χc · Q",
        )
        Qc_kW = chi_c * Q_kW
        st.write(f"Convective HRR Qc = **{Qc_kW:.0f} kW**")
    with c3:
        T_amb_C = st.number_input("Ambient temperature (°C)", min_value=-10.0, value=20.0, step=1.0)
        rho_amb = rho_air(T_amb_C)
        st.write(f"Ambient density ≈ **{rho_amb:.3f} kg/m³**")

    st.subheader("3. Smoke-layer design targets")
    min_clear = REF["smoke_layer"]["min_clear_height_m"]
    z_min_from_walk = highest_walk_m + min_clear
    z_min_depth = height_m * (1.0 - REF["smoke_layer"]["min_layer_depth_fraction_of_H"])
    z_default = max(z_min_from_walk, z_min_depth * 0.5)  # practical starting point

    c1, c2 = st.columns(2)
    with c1:
        z_interface = st.number_input(
            "Design smoke-layer interface height z (m) above fire base",
            min_value=1.0,
            max_value=float(height_m - 0.5),
            value=min(z_default, height_m - 1.0),
            step=0.5,
            help="Must keep clear height ≥ 1.83 m above highest walking surface and layer depth ≥ 20 % of H (or justified).",
        )
    with c2:
        layer_depth = height_m - z_interface
        clear_above_walk = z_interface - highest_walk_m
        st.metric("Smoke layer depth", f"{layer_depth:.2f} m")
        st.metric("Clear height above highest walking surface", f"{clear_above_walk:.2f} m")

    # Compliance flags
    clear_ok = clear_above_walk >= min_clear - 1e-6
    depth_ok = layer_depth >= height_m * REF["smoke_layer"]["min_layer_depth_fraction_of_H"] - 1e-6
    if not clear_ok:
        st.error(f"Clear height {clear_above_walk:.2f} m < required {min_clear} m above highest walking surface.")
    if not depth_ok:
        st.warning(
            f"Layer depth {layer_depth:.2f} m < 20 % of H ({height_m * 0.2:.2f} m). "
            "NFPA allows engineering justification of a shallower layer."
        )
    if clear_ok and depth_ok:
        st.success("Basic geometric smoke-layer criteria satisfied.")

    st.subheader("4. Plume & exhaust calculation (axisymmetric)")
    m_dot, eq_used = mass_flow_axisym(Qc_kW, z_interface)
    zl = flame_height_axisym(Qc_kW)
    dT = avg_plume_temp_rise(Qc_kW, m_dot)
    T_smoke = T_amb_C + dT
    rho_s = smoke_density(T_smoke, T_amb_C, rho_amb)
    V_exh = volumetric_flow(m_dot, rho_s)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Flame height zl (approx)", f"{zl:.2f} m")
    c2.metric("Mass flow ṁ at z", f"{m_dot:.1f} kg/s")
    c3.metric("Avg smoke ΔT", f"{dT:.0f} °C")
    c4.metric("Required exhaust V̇", f"{V_exh:.1f} m³/s")

    st.caption(f"Equation used: {eq_used}")
    st.caption(f"Smoke temperature ≈ {T_smoke:.0f} °C · Density ≈ {rho_s:.3f} kg/m³")

    # Makeup air
    makeup_frac = st.slider("Makeup-air fraction of exhaust", 0.85, 0.95, 0.90, 0.01)
    V_makeup = V_exh * makeup_frac
    st.write(f"Recommended makeup-air supply ≈ **{V_makeup:.1f} m³/s** ({makeup_frac*100:.0f} % of exhaust)")

    c1, c2 = st.columns(2)
    with c1:
        free_area_m2 = st.number_input(
            "Total free area of makeup-air openings (m²)",
            min_value=0.1,
            value=max(1.0, V_makeup / 1.02),
            step=0.5,
            help="Doors, louvres, vents that open automatically or are open during smoke control.",
        )
    with c2:
        v_makeup = V_makeup / free_area_m2 if free_area_m2 > 0 else 0.0
        st.metric("Makeup-air velocity", f"{v_makeup:.2f} m/s")
        limit = REF["makeup_air"]["max_velocity_m_s"]
        if v_makeup <= limit + 1e-6:
            st.success(f"≤ {limit} m/s guideline")
        else:
            st.warning(
                f"> {limit} m/s. NFPA 92 / common practice prefers ≤ 1.02 m/s near the fire. "
                "Higher velocity may need engineering justification and AHJ acceptance."
            )
    st.caption(REF["makeup_air"].get("max_velocity_note", ""))

    # Plugholing / number of inlets
    st.subheader("5. Exhaust inlets (plugholing screen)")
    gamma = st.number_input("γ factor (inlet configuration)", min_value=0.5, max_value=2.0, value=1.0, step=0.1)
    V_max_inlet = max_vol_per_inlet_no_plugholing(layer_depth, T_smoke, T_amb_C, gamma)
    n_inlets = number_of_inlets_required(V_exh, V_max_inlet)
    c1, c2 = st.columns(2)
    c1.metric("Max V̇ per inlet (approx, no plugholing)", f"{V_max_inlet:.2f} m³/s")
    c2.metric("Minimum number of inlets (screen)", f"{n_inlets}")
    st.caption(
        "Plugholing limit is approximate. Verify with the exact equation and factors "
        "in the NFPA 92 edition applicable to the project. Distribute inlets to avoid short-circuiting."
    )

    # Filling time (unvented) – screening
    st.subheader("6. Unvented filling time (screening)")
    t_fill = filling_time_steady(height_m, floor_area_m2, Q_kW, z_interface)
    st.write(
        f"Approximate time for first indication of smoke to reach the design interface "
        f"(steady fire, no exhaust): **{t_fill:.0f} s** ({t_fill/60:.1f} min)"
    )
    st.caption(
        "This is a simplified screening estimate of the classic NFPA 92B filling relation. "
        "For design, use the full algebraic procedure or a validated zone/CFD model, "
        "especially when exhaust is operating or the fire is unsteady."
    )

    # Store results for export
    summary_rows = [
        ["Floor area A", f"{floor_area_m2:.1f}", "m²", ""],
        ["Ceiling height H", f"{height_m:.2f}", "m", "above fire base"],
        ["Highest walking surface", f"{highest_walk_m:.2f}", "m", ""],
        ["Design interface height z", f"{z_interface:.2f}", "m", ""],
        ["Smoke layer depth", f"{layer_depth:.2f}", "m", ""],
        ["Clear height above walk", f"{clear_above_walk:.2f}", "m", f"min {min_clear} m"],
        ["Total HRR Q", f"{Q_kW:.0f}", "kW", fire_mode],
        ["Convective HRR Qc", f"{Qc_kW:.0f}", "kW", f"χc = {chi_c}"],
        ["Flame height zl", f"{zl:.2f}", "m", "approx"],
        ["Mass flow at z", f"{m_dot:.2f}", "kg/s", eq_used],
        ["Smoke temperature", f"{T_smoke:.0f}", "°C", ""],
        ["Smoke density", f"{rho_s:.3f}", "kg/m³", ""],
        ["Required exhaust V̇", f"{V_exh:.2f}", "m³/s", ""],
        ["Makeup air (selected)", f"{V_makeup:.2f}", "m³/s", f"{makeup_frac*100:.0f} %"],
        ["Makeup free area", f"{free_area_m2:.2f}", "m²", ""],
        ["Makeup velocity", f"{v_makeup:.2f}", "m/s", f"limit {REF['makeup_air']['max_velocity_m_s']} m/s"],
        ["Max V̇ per inlet (screen)", f"{V_max_inlet:.2f}", "m³/s", f"γ = {gamma}"],
        ["Min number of inlets", f"{n_inlets}", "–", "plugholing screen"],
        ["Unvented filling time (screen)", f"{t_fill:.0f}", "s", "steady fire"],
    ]

    st.session_state["calc_results"] = {
        "project": {"name": proj_name, "location": proj_loc, "prepared_by": prep_by},
        "geometry": {
            "length_m": length_m,
            "width_m": width_m,
            "height_m": height_m,
            "highest_walk_m": highest_walk_m,
            "floor_area_m2": floor_area_m2,
            "volume_m3": volume_m3,
            "z_interface_m": z_interface,
            "layer_depth_m": layer_depth,
        },
        "fire": {
            "mode": fire_mode,
            "Q_kW": Q_kW,
            "Qc_kW": Qc_kW,
            "chi_c": chi_c,
            "T_amb_C": T_amb_C,
        },
        "results": {
            "m_dot_kg_s": m_dot,
            "V_exh_m3s": V_exh,
            "T_smoke_C": T_smoke,
            "n_inlets": n_inlets,
            "t_fill_s": t_fill,
            "summary_rows": summary_rows,
            "clear_ok": clear_ok,
            "depth_ok": depth_ok,
        },
    }

# ---------------------------------------------------------------------------
# TAB 2 – Data Checklist
# ---------------------------------------------------------------------------
with tab_check:
    st.subheader("Compliance Data Checklist")
    st.markdown(
        "Use this checklist to track the information required for a complete NFPA 92 "
        "smoke-management design submission. Progress is saved for the current session."
    )

    if "checklist" not in st.session_state:
        st.session_state["checklist"] = {
            cat["id"]: [False] * len(cat["items"]) for cat in REF["checklist_categories"]
        }

    total_items = 0
    done_items = 0
    for cat in REF["checklist_categories"]:
        with st.expander(cat["title"], expanded=False):
            statuses = st.session_state["checklist"][cat["id"]]
            for i, item in enumerate(cat["items"]):
                key = f"chk_{cat['id']}_{i}"
                checked = st.checkbox(item, value=statuses[i], key=key)
                st.session_state["checklist"][cat["id"]][i] = checked
                total_items += 1
                if checked:
                    done_items += 1

    pct = 100.0 * done_items / total_items if total_items else 0.0
    st.progress(pct / 100.0)
    st.write(f"**Progress: {done_items} / {total_items} items ({pct:.0f} %)**")

# ---------------------------------------------------------------------------
# TAB 3 – Reference Data
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# TAB – Testing & Acceptance
# ---------------------------------------------------------------------------
with tab_test:
    st.subheader("Acceptance Testing & Periodic ITM (NFPA 92 Chapter 8)")
    st.markdown(
        """
NFPA 92 requires that testing confirm the **design objectives** are achieved.
Design documents must include acceptance procedures and pass/fail criteria.
Responsibility for each testing phase should be defined before work starts.
"""
    )

    t = REF.get("testing", {})

    st.markdown("### Acceptance Testing")
    acc = t.get("acceptance", {})
    st.caption(acc.get("title", "Acceptance Testing"))
    for item in acc.get("items", []):
        st.markdown(f"- {item}")

    st.markdown("### Periodic Inspection, Testing & Maintenance")
    per = t.get("periodic", {})
    st.caption(per.get("title", "Periodic ITM"))
    st.markdown(
        """
| System type | Minimum frequency |
|-------------|-------------------|
| **Dedicated** smoke control systems | **Semiannually** |
| **Non-dedicated** systems | **Annually** |
"""
    )
    for item in per.get("items", []):
        st.markdown(f"- {item}")

    st.markdown("### Required Documentation (Chapter 7)")
    doc = t.get("documentation", {})
    for item in doc.get("items", []):
        st.markdown(f"- {item}")

    st.markdown("### Typical acceptance measurements (atrium exhaust)")
    st.markdown(
        """
For an atrium **smoke-management (exhaust)** system, acceptance testing typically verifies:

1. **Exhaust airflow** – measured volumetric flow at each exhaust fan / inlet group ≥ design value (often with a small tolerance agreed with the AHJ).
2. **Makeup-air airflow** – approximately 85–95 % of exhaust; velocity at critical openings ≤ **1.02 m/s** (or higher only with supporting analysis).
3. **Activation sequence** – smoke detection (or other automatic means) starts exhaust, opens makeup paths, shuts down normal HVAC as designed; FSCS manual control works.
4. **Standby power** – same sequence on emergency power if provided.
5. **No adverse interactions** – stair/elevator pressurization (if present) still meets its pressure/door-force criteria while atrium exhaust is running.

Hot-smoke tests or cold-smoke visualisation may be used where required by the AHJ or design basis, but NFPA 92 does not mandate a hot-smoke test for every project.
"""
    )

    st.info(
        "The **Data Checklist** tab includes the full acceptance and periodic/ITM item lists so you can track project compliance status."
    )

with tab_ref:
    st.subheader("Reference data (SI)")

    st.markdown("### Design-fire guidance")
    df_fire = pd.DataFrame(
        [
            {"Scenario": k, "Typical steady HRR (kW)": v}
            for k, v in REF["design_fires"]["typical_steady_kW"].items()
        ]
    )
    st.dataframe(df_fire, use_container_width=True, hide_index=True)

    st.markdown("### t-squared growth coefficients")
    df_alpha = pd.DataFrame(
        [
            {"Growth": k, "α (kW/s²)": v}
            for k, v in REF["design_fires"]["t_squared_alpha_kW_s2"].items()
        ]
    )
    st.dataframe(df_alpha, use_container_width=True, hide_index=True)

    st.markdown("### Smoke-layer & tenability criteria")
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"- Minimum clear height above highest walking surface: **{REF['smoke_layer']['min_clear_height_m']} m**")
        st.write(f"- Minimum design smoke-layer depth: **{REF['smoke_layer']['min_layer_depth_fraction_of_H']*100:.0f} % of H** (or engineering analysis)")
        st.write(f"- Default ambient temperature: **{REF['smoke_layer']['ambient_temp_C_default']} °C**")
    with c2:
        ten = REF["tenability"]
        st.write(f"- Temperature: ≤ **{ten['temperature_C']} °C** at {ten['measurement_height_m']} m")
        st.write(f"- Visibility: ≥ **{ten['visibility_m']} m**")
        st.write(f"- CO: ≤ **{ten['CO_ppm']} ppm** (time-averaged)")
        st.write(f"- CO₂: ≤ **{ten['CO2_percent']} %** · O₂: ≥ **{ten['O2_min_percent']} %**")

    st.markdown("### Axisymmetric plume equations (SI)")
    st.code(REF["plume_equations_SI"]["axisymmetric_above_flame"], language="text")
    st.code(REF["plume_equations_SI"]["axisymmetric_in_flame"], language="text")
    st.code(REF["plume_equations_SI"].get("flame_height_SI", "zl = 0.166 * Qc**(2/5)"), language="text")
    st.caption(REF["plume_equations_SI"]["units"])
    st.caption(REF["plume_equations_SI"].get("flame_height_note", ""))

    st.markdown("### Makeup air")
    st.write(
        f"- Typical makeup-air quantity: **{REF['makeup_air']['fraction_of_exhaust'][0]*100:.0f}–"
        f"{REF['makeup_air']['fraction_of_exhaust'][1]*100:.0f} %** of exhaust"
    )
    st.write(f"- Maximum recommended velocity: **{REF['makeup_air']['max_velocity_m_s']} m/s**")
    st.caption(REF["makeup_air"]["recommended_location"])

# ---------------------------------------------------------------------------
# TAB 4 – Export
# ---------------------------------------------------------------------------
with tab_export:
    st.subheader("Export Excel report")
    if "calc_results" not in st.session_state:
        st.warning("Run the Calculations tab first to generate results.")
    else:
        res = st.session_state["calc_results"]
        checklist = st.session_state.get("checklist", {})
        if st.button("Generate Excel report", type="primary"):
            xlsx_bytes = build_excel_report(
                res["project"],
                res["geometry"],
                res["fire"],
                res["results"],
                checklist,
            )
            fname = f"NFPA92_Smoke_Management_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label="⬇️ Download report",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("Report ready.")

st.markdown("---")
st.caption(
    "NFPA 92 / 92B Smoke Management Calculator · SI units · Educational aid only · "
    "Not a substitute for professional engineering design or the official standard."
)
