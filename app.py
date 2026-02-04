# -*- coding: utf-8 -*-
# Bolt Torque Calculator (Streamlit) – loads private Excel from st.secrets

from typing import Optional
import io, base64
import pandas as pd
import streamlit as st

# ---------- helpers to find tables in your workbook ----------
def _find_sheet_containing(xls: pd.ExcelFile, phrase: str) -> Optional[str]:
    for sn in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sn, header=None, engine="openpyxl")
        except Exception:
            continue
        if df.astype(str).apply(lambda c: c.str.contains(phrase, case=False, na=False)).any().any():
            return sn
    return None

def _extract_bolt_size_area(xls: pd.ExcelFile) -> pd.DataFrame:
    sn = _find_sheet_containing(xls, "Bolts Size")
    frames = []
    for name in (xls.sheet_names if sn is None else [sn]):
        try:
            raw = pd.read_excel(xls, sheet_name=name, header=None, engine="openpyxl")
        except Exception:
            continue
        for _, row in raw.iterrows():
            vals = row.tolist()
            nums = [v for v in vals if isinstance(v, (int, float)) and not pd.isna(v)]
            if len(nums) >= 2:
                size = float(nums[0])
                area_candidates = [n for n in nums[1:] if 0.05 <= n <= 6]
                if 0.5 <= size <= 4 and area_candidates:
                    frames.append(pd.DataFrame({"bolt_size_in":[size], "area_in2":[area_candidates[0]]}))
    if not frames:
        raise RuntimeError("Could not parse bolt size/area table.")
    return (pd.concat(frames, ignore_index=True)
              .drop_duplicates()
              .sort_values("bolt_size_in"))

def _extract_materials(xls: pd.ExcelFile) -> pd.DataFrame:
    sn = _find_sheet_containing(xls, "Bolt Material")
    if sn is None:
        raise RuntimeError("Missing 'Bolt Material' section.")
    raw = pd.read_excel(xls, sheet_name=sn, header=None, engine="openpyxl")
    recs = []
    for _, row in raw.iterrows():
        vals = [None if pd.isna(v) else v for v in row.tolist()]
        if len(vals) < 3:
            continue
        try:
            y = float(vals[-1])
            mat = str(vals[0]).strip()
            size = str(vals[1]).strip()
            if mat and y > 0:
                recs.append((mat, size, y))
        except Exception:
            pass
    if not recs:
        raise RuntimeError("No material records found.")
    return pd.DataFrame(recs, columns=["material","size_rule","yield_ksi"]).drop_duplicates()

def _extract_lubricants(xls: pd.ExcelFile) -> pd.DataFrame:
    sn = _find_sheet_containing(xls, "Lubricant")
    if sn is None:
        # sensible defaults if the sheet is absent
        return pd.DataFrame({"lubricant":["Molykote 1000","Molykote P37"], "mu":[0.13, 0.142]})
    raw = pd.read_excel(xls, sheet_name=sn, header=None, engine="openpyxl")
    recs = []
    for _, row in raw.iterrows():
        vals = [None if pd.isna(v) else v for v in row.tolist()]
        if not vals:
            continue
        name, mu = None, None
        for v in vals:
            s = str(v).strip()
            if any(ch.isalpha() for ch in s) and len(s) < 60 and name is None:
                name = s
            try:
                fv = float(s)
                if 0.05 <= fv <= 0.3:
                    mu = fv
            except Exception:
                pass
        if name and (mu is not None):
            recs.append((name, mu))
    if not recs:
        return pd.DataFrame({"lubricant":["Molykote 1000","Molykote P37"], "mu":[0.13, 0.142]})
    return pd.DataFrame(recs, columns=["lubricant","mu"]).drop_duplicates()

# ---------- engineering logic ----------
def default_fraction_for_gasket(gasket: str) -> float:
    g = (gasket or "").lower()
    if "spiral" in g:   return 0.50
    if "rtj" in g:      return 0.40
    if "rubber" in g or "cnaf" in g: return 0.30
    if "teflon" in g or "ptfe" in g: return 0.40
    return 0.50

def compute_bolt_load(area_in2: float, yield_ksi: float, fraction: float) -> float:
    return area_in2 * (fraction * yield_ksi * 1000.0)  # lbf

def derive_nut_factor(mu: Optional[float], override_K: Optional[float]) -> float:
    if override_K and override_K > 0:
        return override_K
    mu = 0.13 if mu is None else float(mu)
    return 0.04 + mu

def torque_lbft(F_lbf: float, K: float, bolt_diameter_in: float) -> float:
    return (F_lbf * K * bolt_diameter_in) / 12.0

def lbft_to_Nm(T_lbft: float) -> float:
    return T_lbft * 1.3558179483314004

# ---------- UI ----------
st.set_page_config(page_title="Bolt Torque Calculator", layout="wide")
st.title("Bolt Torque Calculator")

# --- Load workbook bytes from secrets (server-side only) ---
try:
    b64 = st.secrets["private_files"]["bfji_excel_b64"]  # set in App settings → Secrets
    data_bytes = base64.b64decode(b64)
    xls = pd.ExcelFile(io.BytesIO(data_bytes), engine="openpyxl")
except Exception as e:
    st.error("The private Excel file could not be loaded. Please check App settings → Secrets.")
    st.stop()

# Parse once
try:
    bolts = _extract_bolt_size_area(xls)
    materials = _extract_materials(xls)
    lubes = _extract_lubricants(xls)
except Exception as e:
    st.error(f"Failed to parse workbook: {e}")
    st.stop()

# Main content (no sidebar)
colL, colR = st.columns([1,2])

with colL:
    st.subheader("Inputs")
    bolt_size = st.selectbox("Bolt size (inch)", options=sorted(bolts["bolt_size_in"].unique()))
    area_in2 = float(bolts.query("bolt_size_in == @bolt_size")["area_in2"].iloc[0])

    material = st.selectbox("Bolt material", options=sorted(materials["material"].unique()))
    yield_ksi = float(materials.query("material == @material")["yield_ksi"].median())

    gasket = st.selectbox("Gasket type", ["Spiral Wound", "RTJ", "Rubber/CNAF", "Teflon"])
    fraction = st.slider("Fraction of yield (F = fraction × Sy × A)",
                         min_value=0.2, max_value=1.0, value=float(default_fraction_for_gasket(gasket)), step=0.05)

    lubricant = st.selectbox("Lubricant", options=sorted(lubes["lubricant"].unique()))
    mu = float(lubes.query("lubricant == @lubricant")["mu"].median())
    st.caption(f"Coefficient of friction μ = {mu:.3f}. Default K ≈ 0.04 + μ")
    override_K = st.number_input("Override K (optional)", value=0.0, min_value=0.0, step=0.001)

with colR:
    st.subheader("Results")
    F = compute_bolt_load(area_in2, yield_ksi, fraction)
    K = derive_nut_factor(mu, override_K if override_K > 0 else None)
    T_lbft = torque_lbft(F, K, bolt_size)
    T_Nm = lbft_to_Nm(T_lbft)

    m1, m2, m3 = st.columns(3)
    m1.metric("Bolt load, F (lbf)", f"{F:,.0f}")
    m2.metric("Torque (lb·ft)", f"{T_lbft:,.2f}")
    m3.metric("Torque (N·m)", f"{T_Nm:,.2f}")

    st.divider()
    st.markdown("**Quick table** (at 30%, 60%, 100% of yield; same μ and K rule)")
    rows = []
    for frac, label in [(0.30,"30%"),(0.60,"60%"),(1.00,"100%")]:
        Fi = compute_bolt_load(area_in2, yield_ksi, frac)
        Tlbf = torque_lbft(Fi, K, bolt_size)
        TNm = lbft_to_Nm(Tlbf)
        rows.append({"Fraction": label, "F (lbf)": round(Fi,2), "Torque (lb·ft)": round(Tlbf,2), "Torque (N·m)": round(TNm,2)})
    st.dataframe(pd.DataFrame(rows), use_container_width=True)

with st.expander("Details & formulae"):
    st.markdown(
        f"""
* Selected area A: **{area_in2:.6f} in²** (from private workbook)  
* Yield strength Sy: **{yield_ksi:.1f} ksi**  
* Fraction: **{fraction:.2f}** → **F = A × (fraction × Sy_psi)** = A × (fraction × Sy_ksi × 1000)  
* Lubricant μ: **{mu:.3f}**; Nut factor **K ≈ 0.04 + μ = {K:.3f}**  
* Torque: **T(lb·ft) = F × K × D(in) / 12**  
* Convert: **1 lb·ft = 1.3558179483 N·m**
"""
    )
