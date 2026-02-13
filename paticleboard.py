import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import io
import json
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

matplotlib.rcParams['font.family'] = 'DejaVu Sans'

# =============================================================
# Page Config
# =============================================================
st.set_page_config(
    page_title="Particleboard Bending Test",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================
# Custom CSS — Light theme, high contrast
# =============================================================
st.markdown("""
<style>
    .stApp { background: #e8f5e9; }

    /* Sidebar — เขียวเข้มอ่อน */
    section[data-testid="stSidebar"] {
        background: #c8e6c9;
        border-right: 2px solid #81c784;
    }
    section[data-testid="stSidebar"] .stMarkdown h1,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #1b5e20 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px; background: #ffffff; border-radius: 10px;
        padding: 5px; box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px; padding: 10px 20px; font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: #2e7d32 !important; color: white !important;
    }

    /* Expander — ขอบเขียว ชัดเจน ไม่กลืน */
    div[data-testid="stExpander"] {
        background: #ffffff;
        border: 2px solid #81c784;
        border-radius: 12px;
        margin-bottom: 10px;
        box-shadow: 0 2px 6px rgba(46,125,50,0.1);
    }
    div[data-testid="stExpander"] summary {
        font-weight: 600;
        color: #1b5e20;
    }

    /* Cards */
    .info-card {
        background: #ffffff; border-radius: 12px; padding: 18px 22px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        border-left: 4px solid #2e7d32; margin-bottom: 16px; color: #1e293b;
    }
    .result-card {
        background: #f1f8e9; border-radius: 12px; padding: 18px 22px;
        border-left: 4px solid #43a047; margin-bottom: 12px; color: #1b5e20;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .result-card b { color: #2e7d32; }
    .result-card .big-val { font-size: 24px; font-weight: 700; color: #2e7d32; }

    .warn-card {
        background: #fff8e1; border-radius: 12px; padding: 14px 18px;
        border-left: 4px solid #f9a825; margin-bottom: 12px;
        color: #5d4037; font-size: 14px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .summary-card {
        background: #e3f2fd; border-radius: 12px; padding: 18px 22px;
        border-left: 4px solid #1565c0; margin-bottom: 12px; color: #0d47a1;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .stored-card {
        background: #ffffff; border-radius: 10px; padding: 14px 18px;
        border: 2px solid #a5d6a7; margin-bottom: 8px; color: #334155;
        font-size: 14px;
    }

    /* Metric boxes */
    .metric-box {
        background: #ffffff; border-radius: 10px; padding: 16px; text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06); border: 2px solid #a5d6a7;
    }
    .metric-box .label { font-size: 13px; color: #558b2f; font-weight: 500; }
    .metric-box .value { font-size: 24px; font-weight: 700; color: #1b5e20; }
    .metric-box .unit { font-size: 12px; color: #7cb342; }

    /* Header */
    .header-banner {
        background: linear-gradient(135deg, #1b5e20 0%, #2e7d32 50%, #43a047 100%);
        color: white; padding: 24px 28px; border-radius: 14px; margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(46,125,50,0.25);
    }
    .header-banner h1 { margin: 0; font-size: 26px; font-weight: 700; color: white !important; }
    .header-banner p { margin: 6px 0 0 0; font-size: 14px; opacity: 0.9; }

    .footer {
        text-align: center; padding: 18px; margin-top: 28px;
        color: #2e7d32; font-size: 13px; border-top: 2px solid #a5d6a7;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================
# Calculation Functions
# =============================================================

def calc_MOR(Fmax_N, L, b, t):
    if b <= 0 or t <= 0 or L <= 0:
        raise ValueError("L, b, t ต้องมากกว่าศูนย์")
    return (3.0 * Fmax_N * L) / (2.0 * b * t**2)

def calc_slope_elastic(load_N, defl_mm, idx_lo, idx_hi):
    x = defl_mm[idx_lo:idx_hi+1]
    y = load_N[idx_lo:idx_hi+1]
    if hasattr(x, 'values'): x = x.values
    if hasattr(y, 'values'): y = y.values
    if len(x) < 2:
        raise ValueError("ต้องเลือกอย่างน้อย 2 จุด")
    coeffs = np.polyfit(x, y, 1)
    slope, intercept = coeffs[0], coeffs[1]
    y_pred = np.polyval(coeffs, x)
    ss_res = np.sum((y - y_pred)**2)
    ss_tot = np.sum((y - np.mean(y))**2)
    r_sq = 1.0 - ss_res / ss_tot if ss_tot > 0 else 0.0
    return slope, intercept, r_sq

def slope_to_MOE(slope, L, b, t):
    if b <= 0 or t <= 0 or L <= 0:
        raise ValueError("L, b, t ต้องมากกว่าศูนย์")
    return slope * L**3 / (4.0 * b * t**3)

def calc_TS(before, after):
    avg_b, avg_a = np.mean(before), np.mean(after)
    if avg_b <= 0:
        raise ValueError("ค่าเฉลี่ยก่อนแช่น้ำต้องมากกว่าศูนย์")
    return avg_b, avg_a, ((avg_a - avg_b) / avg_b) * 100.0

def calc_statistics(values):
    arr = np.array([v for v in values if v is not None and not np.isnan(v)])
    if len(arr) == 0:
        return {"n": 0, "mean": 0, "sd": 0, "cv": 0, "min": 0, "max": 0}
    m = np.mean(arr)
    s = np.std(arr, ddof=1) if len(arr) > 1 else 0.0
    return {"n": len(arr), "mean": m, "sd": s,
            "cv": (s / m * 100) if m != 0 else 0,
            "min": np.min(arr), "max": np.max(arr)}

# Helpers
def ss_get(key, default=0.0):
    return st.session_state.get(key, default)
def store(key, value):
    st.session_state[key] = value

def render_metrics(cols, items):
    for col, (lbl, val, unt) in zip(cols, items):
        with col:
            st.markdown(f'<div class="metric-box">'
                        f'<div class="label">{lbl}</div>'
                        f'<div class="value">{val}</div>'
                        f'<div class="unit">{unt}</div></div>',
                        unsafe_allow_html=True)

# =============================================================
# Sidebar
# =============================================================
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("---")

    st.markdown("#### 📂 Save / Load")
    uploaded_json = st.file_uploader("Upload JSON", type=["json"],
                                     label_visibility="collapsed")
    if uploaded_json is not None:
        try:
            loaded_data = json.load(uploaded_json)
            fid = f"{uploaded_json.name}_{uploaded_json.size}"
            if ss_get("last_uploaded_file", "") != fid:
                store("last_uploaded_file", fid)
                for k, v in loaded_data.items():
                    st.session_state[k] = v
                st.success("✅ โหลดสำเร็จ")
                st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")

    def build_export():
        ns = int(ss_get("num_samples", 1))
        d = {"num_samples": ns}
        for i in range(ns):
            for k in ["L","b","t","Fmax"]:
                d[f"{k}_{i}"] = ss_get(f"{k}_{i}", 0.0)
        for k in ["L_moe","b_moe","t_moe"]:
            d[k] = ss_get(k, 0.0)
        nts = int(ss_get("num_ts_samples", 1))
        d["num_ts_samples"] = nts
        for j in range(nts):
            for s in range(1, 5):
                d[f"ts_b{s}_{j}"] = ss_get(f"ts_b{s}_{j}", 0.0)
                d[f"ts_a{s}_{j}"] = ss_get(f"ts_a{s}_{j}", 0.0)
        return d

    st.download_button("💾 Download JSON",
                       data=json.dumps(build_export(), ensure_ascii=False, indent=2),
                       file_name="particleboard_params.json",
                       mime="application/json")

    st.markdown("---")
    st.markdown("#### 📋 Standard")
    standard = st.selectbox("มาตรฐานอ้างอิง", [
        "TIS 876-2547 (มอก.)", "JIS A 5908:2003",
        "EN 310:1993", "ASTM D1037", "Other"
    ], key="standard_ref")

    st.markdown("---")
    st.caption("Particleboard Test v2.1")

# =============================================================
# Header
# =============================================================
st.markdown("""
<div class="header-banner">
    <h1>🧪 Particleboard Bending Test</h1>
    <p>MOR • MOE (Elastic Region) • Thickness Swelling • Summary Report</p>
</div>
""", unsafe_allow_html=True)

# =============================================================
# Tabs
# =============================================================
tab_mor, tab_moe, tab_ts, tab_summary, tab_export = st.tabs([
    "① MOR", "② MOE + Graph", "③ Thickness Swelling",
    "④ Summary", "⑤ Export"
])

# =============================================================
# TAB 1: MOR
# =============================================================
with tab_mor:
    st.markdown(f'<div class="info-card">'
                f'<b>Modulus of Rupture (MOR)</b> — Three-point bending<br>'
                f'MOR = 3·F<sub>max</sub>·L / (2·b·t²) [MPa]<br>'
                f'<small>📋 {standard}</small></div>', unsafe_allow_html=True)

    num_samples = st.selectbox("จำนวนตัวอย่าง MOR", [1,2,3,4],
                               index=[1,2,3,4].index(int(ss_get("num_samples", 1))),
                               key="num_samples")

    # Upload Excel for auto Fmax
    st.markdown('<div class="warn-card">'
                '📥 อัปโหลด Excel เพื่อดึง F<sub>max</sub> อัตโนมัติ (optional) — '
                'ดึง Load สูงสุดจากแต่ละ sheet หรือพิมพ์เองก็ได้'
                '</div>', unsafe_allow_html=True)

    mor_excel = st.file_uploader("Excel สำหรับ MOR (optional)",
                                 type=["xlsx"], key="mor_excel_upload")
    auto_fmax = {}
    if mor_excel:
        try:
            xls = pd.ExcelFile(mor_excel)
            for si, sn in enumerate(xls.sheet_names):
                dfs = pd.read_excel(xls, sheet_name=sn)
                lcol = None
                for c in dfs.columns:
                    if 'load' in str(c).lower() or 'kg' in str(c).lower():
                        lcol = c; break
                if lcol is None and len(dfs.columns) >= 1:
                    lcol = dfs.columns[0]
                if lcol:
                    try:
                        mv = float(pd.to_numeric(dfs[lcol], errors='coerce').max())
                        if not np.isnan(mv): auto_fmax[si] = mv
                    except: pass
            if auto_fmax:
                st.success(f"✅ ดึง Fmax จาก {len(auto_fmax)} sheet: " +
                          ", ".join([f"'{xls.sheet_names[k]}'={v:.2f} kg"
                                     for k, v in auto_fmax.items()]))
        except Exception as e:
            st.warning(f"⚠️ {e}")

    mor_results = []
    for i in range(num_samples):
        with st.expander(f"📐 ตัวอย่างที่ {i+1}", expanded=(i == 0)):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                L = st.number_input("L (mm)", value=ss_get(f"L_{i}"),
                                    key=f"L_{i}", min_value=0.0, format="%.2f")
            with c2:
                b = st.number_input("b (mm)", value=ss_get(f"b_{i}"),
                                    key=f"b_{i}", min_value=0.0, format="%.2f")
            with c3:
                t = st.number_input("t (mm)", value=ss_get(f"t_{i}"),
                                    key=f"t_{i}", min_value=0.0, format="%.2f")
            with c4:
                def_fmax = auto_fmax.get(i, ss_get(f"Fmax_{i}"))
                lbl = f"F_max (kg)" + (" ← Excel" if i in auto_fmax else "")
                Fmax_kg = st.number_input(lbl, value=float(def_fmax),
                                          key=f"Fmax_{i}", min_value=0.0,
                                          format="%.3f")

            if L > 0 and b > 0 and t > 0 and Fmax_kg > 0:
                Fmax_N = Fmax_kg * 9.80665
                mor = calc_MOR(Fmax_N, L, b, t)
                mor_results.append(mor)
                store(f"mor_result_{i}", mor)
                st.markdown(f'<div class="result-card">'
                            f'<b>MOR #{i+1}</b> = '
                            f'<span class="big-val">{mor:.2f}</span> MPa '
                            f'<small style="color:#64748b;">'
                            f'(F={Fmax_N:.2f} N)</small></div>',
                            unsafe_allow_html=True)
            else:
                mor_results.append(None)
                if any(v > 0 for v in [L, b, t, Fmax_kg]):
                    st.warning("ใส่ค่าให้ครบ")

    valid_mor = [v for v in mor_results if v is not None]
    if len(valid_mor) >= 2:
        stats = calc_statistics(valid_mor)
        st.markdown("---")
        st.markdown("##### 📊 MOR Summary")
        render_metrics(st.columns(4), [
            ("Mean", f'{stats["mean"]:.2f}', "MPa"),
            ("Std Dev", f'{stats["sd"]:.2f}', "MPa"),
            ("CV", f'{stats["cv"]:.1f}', "%"),
            ("N", f'{stats["n"]}', "samples"),
        ])
        store("mor_stats", stats)
    elif len(valid_mor) == 1:
        store("mor_stats", {"n":1,"mean":valid_mor[0],"sd":0,"cv":0,
                            "min":valid_mor[0],"max":valid_mor[0]})

# =============================================================
# TAB 2: MOE (multi-sample, one at a time)
# =============================================================
with tab_moe:
    st.markdown('<div class="info-card">'
                '<b>MOE — Load-Deflection Curve</b><br>'
                'อัปโหลดข้อมูลทีละตัวอย่าง → เลือกช่วง elastic → ระบบ fit เส้นตรง<br>'
                '<small>Slope (N/mm) = linear regression | '
                'MOE (MPa) = Slope × L³/(4bt³) — optional</small>'
                '</div>', unsafe_allow_html=True)

    num_moe = st.selectbox("จำนวนตัวอย่าง MOE", [1,2,3,4],
                           index=[1,2,3,4].index(int(ss_get("num_moe_samples", 1))),
                           key="num_moe_samples")

    # Radio to select sample
    current_moe = st.radio("เลือกตัวอย่าง",
                           [f"ตัวอย่างที่ {i+1}" for i in range(num_moe)],
                           horizontal=True, key="current_moe_radio")
    mi = int(current_moe.split()[-1]) - 1

    # Status cards
    stored_any = False
    cols_st = st.columns(num_moe)
    for si in range(num_moe):
        sv = ss_get(f"moe_slope_{si}", None)
        with cols_st[si]:
            if sv and sv > 0:
                stored_any = True
                st.markdown(f'<div class="stored-card">'
                            f'✅ <b>#{si+1}</b> &nbsp; '
                            f'Slope={sv:.1f} N/mm &nbsp; '
                            f'R²={ss_get(f"moe_r2_{si}",0):.4f}</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="stored-card">'
                            f'⏳ <b>#{si+1}</b> — ยังไม่มี</div>',
                            unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(f"### 📈 ตัวอย่างที่ {mi+1}")

    # Template
    tpl = pd.DataFrame({"Load (kg)": [0,5,10,15,20,25],
                        "Deflection (mm)": [0.0,0.8,1.5,2.3,3.5,5.0]})
    buf_t = io.BytesIO(); tpl.to_excel(buf_t, index=False); buf_t.seek(0)
    st.download_button("📥 Template Excel", data=buf_t,
                       file_name="load_deflection_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    uploaded_xl = st.file_uploader(f"Upload Excel — ตัวอย่างที่ {mi+1}",
                                   type=["xlsx"], key=f"moe_xl_{mi}")

    if uploaded_xl:
        try:
            df = pd.read_excel(uploaded_xl)
            if "Load (kg)" not in df.columns or "Deflection (mm)" not in df.columns:
                st.error("❌ ต้องมี 'Load (kg)' และ 'Deflection (mm)'")
            else:
                df["Load (N)"] = df["Load (kg)"] * 9.80665
                df_show = df.copy()
                df_show.insert(0, "Pt#", range(len(df)))
                st.dataframe(df_show, use_container_width=True, height=200)

                load_N = df["Load (N)"].values
                defl = df["Deflection (mm)"].values
                npts = len(df)

                st.markdown("##### 🎯 เลือกช่วง Elastic Region")
                pc1, pc2 = st.columns(2)
                with pc1:
                    dlo = 1 if npts > 2 else 0
                    idx_lo = st.number_input("Pt# เริ่มต้น",
                                             min_value=0, max_value=npts-2,
                                             value=int(ss_get(f"el_lo_{mi}", dlo)),
                                             step=1, key=f"el_lo_{mi}")
                with pc2:
                    dhi = min(max(npts//2, idx_lo+1), npts-1)
                    idx_hi = st.number_input("Pt# สิ้นสุด",
                                             min_value=idx_lo+1, max_value=npts-1,
                                             value=int(ss_get(f"el_hi_{mi}", dhi)),
                                             step=1, key=f"el_hi_{mi}")

                try:
                    slope, intercept, r_sq = calc_slope_elastic(
                        load_N, defl, idx_lo, idx_hi)
                    store(f"moe_slope_{mi}", slope)
                    store(f"moe_intercept_{mi}", intercept)
                    store(f"moe_r2_{mi}", r_sq)

                    st.markdown(f'<div class="result-card">'
                                f'<b>#{mi+1} — Slope</b> '
                                f'(Pt#{idx_lo}→#{idx_hi}, '
                                f'{idx_hi-idx_lo+1} pts)<br>'
                                f'<span class="big-val">{slope:.2f}</span> N/mm'
                                f'&nbsp;&nbsp; '
                                f'<span style="color:#64748b;">R²={r_sq:.4f}</span>'
                                f'</div>', unsafe_allow_html=True)

                    # Optional MOE
                    with st.expander("📐 คำนวณ MOE (MPa) — ใส่ Geometry",
                                     expanded=False):
                        st.markdown("MOE = Slope × L³ / (4·b·t³)")
                        gc1, gc2, gc3 = st.columns(3)
                        with gc1:
                            L_moe = st.number_input("L (mm)",
                                value=ss_get("L_moe", ss_get("L_0")),
                                key="L_moe", min_value=0.0, format="%.2f")
                        with gc2:
                            b_moe = st.number_input("b (mm)",
                                value=ss_get("b_moe", ss_get("b_0")),
                                key="b_moe", min_value=0.0, format="%.2f")
                        with gc3:
                            t_moe = st.number_input("t (mm)",
                                value=ss_get("t_moe", ss_get("t_0")),
                                key="t_moe", min_value=0.0, format="%.2f")
                        if L_moe > 0 and b_moe > 0 and t_moe > 0:
                            moe_v = slope_to_MOE(slope, L_moe, b_moe, t_moe)
                            store(f"moe_mpa_{mi}", moe_v)
                            st.markdown(f'<div class="result-card">'
                                        f'<b>MOE</b> = '
                                        f'<span class="big-val">{moe_v:.2f}</span>'
                                        f' MPa</div>', unsafe_allow_html=True)
                        else:
                            st.info("ใส่ L, b, t ให้ครบ")

                    # --- Graph ---
                    fig, ax = plt.subplots(figsize=(9, 5.5))
                    fig.patch.set_facecolor('#fafbfe')
                    ax.set_facecolor('#fafbfe')

                    ax.plot(defl, load_N, 'o-', color='#94a3b8', lw=1.5,
                            markersize=5, label='Test Data', zorder=2)

                    el_d = defl[idx_lo:idx_hi+1]
                    el_l = load_N[idx_lo:idx_hi+1]
                    ax.plot(el_d, el_l, 'o', color='#16a34a', markersize=10,
                            markeredgecolor='white', markeredgewidth=1.5,
                            label=f'Elastic (#{idx_lo}–#{idx_hi})', zorder=5)

                    pad = (defl[idx_hi] - defl[idx_lo]) * 0.4
                    xf = np.linspace(max(0, defl[idx_lo]-pad),
                                     defl[idx_hi]+pad, 100)
                    yf = slope * xf + intercept
                    ax.plot(xf, yf, '--', color='#dc2626', lw=2,
                            label=f'Fit (slope={slope:.1f}, R²={r_sq:.3f})',
                            zorder=3)

                    for k in [idx_lo, idx_hi]:
                        ax.annotate(f'#{k}', xy=(defl[k], load_N[k]),
                                    xytext=(8, 10), textcoords='offset points',
                                    fontsize=10, fontweight='bold', color='#16a34a',
                                    bbox=dict(boxstyle='round,pad=0.3',
                                              fc='white', ec='#16a34a', alpha=0.9))

                    ax.set_xlabel("Deflection (mm)", fontsize=13, fontweight='bold')
                    ax.set_ylabel("Load (N)", fontsize=13, fontweight='bold')
                    ax.set_title(f"Load–Deflection — Sample #{mi+1}",
                                 fontsize=15, fontweight='bold')
                    ax.legend(fontsize=9, loc='best', framealpha=0.9)
                    ax.grid(True, alpha=0.25, ls='--')
                    ax.spines[['top','right']].set_visible(False)
                    plt.tight_layout()
                    st.pyplot(fig)

                except Exception as e:
                    st.error(f"❌ {e}")
        except Exception as e:
            st.error(f"❌ อ่านไฟล์ไม่ได้: {e}")

    # MOE Summary
    all_slopes = [ss_get(f"moe_slope_{si}", 0) for si in range(num_moe)
                  if ss_get(f"moe_slope_{si}", 0) > 0]
    if len(all_slopes) >= 2:
        mss = calc_statistics(all_slopes)
        st.markdown("---")
        st.markdown("##### 📊 Slope Summary (ทุกตัวอย่าง)")
        render_metrics(st.columns(4), [
            ("Mean", f'{mss["mean"]:.2f}', "N/mm"),
            ("SD", f'{mss["sd"]:.2f}', "N/mm"),
            ("CV", f'{mss["cv"]:.1f}', "%"),
            ("N", f'{mss["n"]}', "samples"),
        ])
        store("moe_slope_stats", mss)

        all_mpa = [ss_get(f"moe_mpa_{si}", 0) for si in range(num_moe)
                   if ss_get(f"moe_mpa_{si}", 0) > 0]
        if len(all_mpa) >= 2:
            mms = calc_statistics(all_mpa)
            st.markdown(f'<div class="summary-card">'
                        f'<b>MOE:</b> Mean={mms["mean"]:.2f} MPa | '
                        f'SD={mms["sd"]:.2f} | CV={mms["cv"]:.1f}%</div>',
                        unsafe_allow_html=True)
            store("moe_mpa_stats", mms)

# =============================================================
# TAB 3: Thickness Swelling
# =============================================================
with tab_ts:
    st.markdown('<div class="info-card">'
                '<b>Thickness Swelling (TS)</b> — วัด 4 ด้าน ก่อน/หลังแช่น้ำ<br>'
                'TS = (t<sub>after</sub>−t<sub>before</sub>)/t<sub>before</sub>×100 [%]'
                '</div>', unsafe_allow_html=True)

    num_ts = st.selectbox("จำนวนตัวอย่าง TS", [1,2,3,4],
                          index=[1,2,3,4].index(int(ss_get("num_ts_samples", 1))),
                          key="num_ts_samples")
    ts_results = []

    for j in range(num_ts):
        with st.expander(f"📐 TS #{j+1}", expanded=(j == 0)):
            st.markdown("**ก่อนแช่น้ำ (mm)**")
            bc = st.columns(4)
            before = [bc[s].number_input(f"ด้าน {s+1}",
                      value=ss_get(f"ts_b{s+1}_{j}"),
                      key=f"ts_b{s+1}_{j}", min_value=0.0, format="%.3f")
                      for s in range(4)]
            st.markdown("**หลังแช่น้ำ (mm)**")
            ac = st.columns(4)
            after = [ac[s].number_input(f"ด้าน {s+1} ",
                     value=ss_get(f"ts_a{s+1}_{j}"),
                     key=f"ts_a{s+1}_{j}", min_value=0.0, format="%.3f")
                     for s in range(4)]

            if all(v > 0 for v in before) and all(v > 0 for v in after):
                avg_b, avg_a, ts = calc_TS(before, after)
                ts_results.append(ts)
                store(f"ts_result_{j}", ts)
                store(f"ts_avg_before_{j}", avg_b)
                store(f"ts_avg_after_{j}", avg_a)
                st.markdown(f'<div class="result-card">'
                            f'<b>TS #{j+1}</b> = '
                            f'<span class="big-val">{ts:.2f}</span> % '
                            f'<small style="color:#64748b;">'
                            f'(before={avg_b:.3f} | after={avg_a:.3f})</small>'
                            f'</div>', unsafe_allow_html=True)
            else:
                ts_results.append(None)

    valid_ts = [v for v in ts_results if v is not None]
    if len(valid_ts) >= 2:
        tss = calc_statistics(valid_ts)
        st.markdown("---")
        st.markdown("##### 📊 TS Summary")
        render_metrics(st.columns(4), [
            ("Mean", f'{tss["mean"]:.2f}', "%"),
            ("SD", f'{tss["sd"]:.2f}', "%"),
            ("CV", f'{tss["cv"]:.1f}', "%"),
            ("N", f'{tss["n"]}', "samples"),
        ])
        store("ts_stats", tss)

# =============================================================
# TAB 4: Summary
# =============================================================
with tab_summary:
    st.markdown("### 📊 Summary")

    mor_data = []
    for i in range(int(ss_get("num_samples", 1))):
        v = ss_get(f"mor_result_{i}", None)
        if v and v > 0:
            mor_data.append({"Sample": f"#{i+1}",
                             "L (mm)": ss_get(f"L_{i}"),
                             "b (mm)": ss_get(f"b_{i}"),
                             "t (mm)": ss_get(f"t_{i}"),
                             "Fmax (kg)": ss_get(f"Fmax_{i}"),
                             "MOR (MPa)": round(v, 2)})
    if mor_data:
        st.markdown("##### MOR")
        st.dataframe(pd.DataFrame(mor_data), use_container_width=True,
                     hide_index=True)
        ms = ss_get("mor_stats", {})
        if ms and ms.get("n", 0) >= 2:
            st.markdown(f'<div class="summary-card"><b>MOR:</b> '
                        f'Mean={ms["mean"]:.2f} | SD={ms["sd"]:.2f} | '
                        f'CV={ms["cv"]:.1f}%</div>', unsafe_allow_html=True)

    moe_data = []
    for si in range(int(ss_get("num_moe_samples", 1))):
        sv = ss_get(f"moe_slope_{si}", 0)
        if sv > 0:
            row = {"Sample": f"#{si+1}",
                   "Slope (N/mm)": round(sv, 2),
                   "R²": round(ss_get(f"moe_r2_{si}", 0), 4)}
            mv = ss_get(f"moe_mpa_{si}", 0)
            if mv > 0: row["MOE (MPa)"] = round(mv, 2)
            moe_data.append(row)
    if moe_data:
        st.markdown("##### MOE")
        st.dataframe(pd.DataFrame(moe_data), use_container_width=True,
                     hide_index=True)
        mss = ss_get("moe_slope_stats", {})
        if mss and mss.get("n", 0) >= 2:
            st.markdown(f'<div class="summary-card"><b>Slope:</b> '
                        f'Mean={mss["mean"]:.2f} N/mm | SD={mss["sd"]:.2f} | '
                        f'CV={mss["cv"]:.1f}%</div>', unsafe_allow_html=True)

    ts_data = []
    for j in range(int(ss_get("num_ts_samples", 1))):
        v = ss_get(f"ts_result_{j}", None)
        if v is not None:
            ts_data.append({"Sample": f"#{j+1}",
                            "Before (mm)": round(ss_get(f"ts_avg_before_{j}",0), 3),
                            "After (mm)": round(ss_get(f"ts_avg_after_{j}",0), 3),
                            "TS (%)": round(v, 2)})
    if ts_data:
        st.markdown("##### Thickness Swelling")
        st.dataframe(pd.DataFrame(ts_data), use_container_width=True,
                     hide_index=True)
        tss = ss_get("ts_stats", {})
        if tss and tss.get("n", 0) >= 2:
            st.markdown(f'<div class="summary-card"><b>TS:</b> '
                        f'Mean={tss["mean"]:.2f}% | SD={tss["sd"]:.2f}% | '
                        f'CV={tss["cv"]:.1f}%</div>', unsafe_allow_html=True)

    if not mor_data and not moe_data and not ts_data:
        st.info("ยังไม่มีผลลัพธ์ — ใส่ข้อมูลใน Tab ① ② ③ ก่อน")

# =============================================================
# TAB 5: Export
# =============================================================
with tab_export:
    st.markdown("### 📤 Export")
    ec1, ec2 = st.columns(2)

    with ec1:
        st.markdown("##### 📊 Excel")
        if st.button("สร้าง Excel", use_container_width=True):
            with io.BytesIO() as buf:
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    if mor_data:
                        pd.DataFrame(mor_data).to_excel(w, "MOR", index=False)
                    if moe_data:
                        pd.DataFrame(moe_data).to_excel(w, "MOE", index=False)
                    if ts_data:
                        pd.DataFrame(ts_data).to_excel(w, "TS", index=False)
                buf.seek(0)
                st.download_button("📥 Download Excel", buf.getvalue(),
                    "particleboard_report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with ec2:
        st.markdown("##### 📄 Word")
        if st.button("สร้าง Word", use_container_width=True):
            try:
                doc = Document()
                doc.add_heading("Particleboard Test Report", level=1)
                doc.add_paragraph(f"Standard: {ss_get('standard_ref','N/A')}")

                if mor_data:
                    doc.add_heading("1. MOR", level=2)
                    tbl = doc.add_table(rows=1, cols=6,
                                        style='Light Shading Accent 1')
                    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for ci, h in enumerate(["#","L","b","t","Fmax","MOR"]):
                        tbl.rows[0].cells[ci].text = h
                    for rd in mor_data:
                        r = tbl.add_row()
                        r.cells[0].text = rd["Sample"]
                        r.cells[1].text = f'{rd["L (mm)"]:.2f}'
                        r.cells[2].text = f'{rd["b (mm)"]:.2f}'
                        r.cells[3].text = f'{rd["t (mm)"]:.2f}'
                        r.cells[4].text = f'{rd["Fmax (kg)"]:.3f}'
                        r.cells[5].text = f'{rd["MOR (MPa)"]:.2f}'
                    ms = ss_get("mor_stats", {})
                    if ms and ms.get("n",0) >= 2:
                        doc.add_paragraph(f'Mean={ms["mean"]:.2f} SD={ms["sd"]:.2f} CV={ms["cv"]:.1f}%')

                if moe_data:
                    doc.add_heading("2. MOE", level=2)
                    headers = list(moe_data[0].keys())
                    tbl2 = doc.add_table(rows=1, cols=len(headers),
                                         style='Light Shading Accent 1')
                    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for ci, h in enumerate(headers):
                        tbl2.rows[0].cells[ci].text = h
                    for rd in moe_data:
                        r = tbl2.add_row()
                        for ci, h in enumerate(headers):
                            r.cells[ci].text = str(rd.get(h, ""))

                if ts_data:
                    doc.add_heading("3. Thickness Swelling", level=2)
                    tbl3 = doc.add_table(rows=1, cols=4,
                                         style='Light Shading Accent 1')
                    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for ci, h in enumerate(["#","Before","After","TS(%)"]):
                        tbl3.rows[0].cells[ci].text = h
                    for rd in ts_data:
                        r = tbl3.add_row()
                        r.cells[0].text = rd["Sample"]
                        r.cells[1].text = f'{rd["Before (mm)"]:.3f}'
                        r.cells[2].text = f'{rd["After (mm)"]:.3f}'
                        r.cells[3].text = f'{rd["TS (%)"]:.2f}'

                doc.add_paragraph("")
                fp = doc.add_paragraph()
                fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = fp.add_run("พัฒนาโดย: รศ.ดร.อิทธิพล มีผล\nภาควิชาครุศาสตร์โยธา มจพ.")
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(100, 100, 100)

                buf_d = io.BytesIO(); doc.save(buf_d); buf_d.seek(0)
                st.download_button("📥 Download Word", buf_d.getvalue(),
                    "particleboard_report.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e:
                st.error(f"❌ {e}")

# Footer
st.markdown("""
<div class="footer">
    <b>พัฒนาโดย:</b> รศ.ดร.อิทธิพล มีผล<br>
    ภาควิชาครุศาสตร์โยธา มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (มจพ.)
</div>
""", unsafe_allow_html=True)
