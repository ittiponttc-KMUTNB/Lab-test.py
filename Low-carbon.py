import streamlit as st
import pandas as pd
import numpy as np
from supabase import create_client

st.set_page_config(
    page_title="Low Carbon Concrete – กรรมการ",
    page_icon="🏆",
    layout="wide"
)

# =============================================
# Supabase
# =============================================
@st.cache_resource
def get_supabase():
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])

supabase = get_supabase()

# =============================================
# Emission Factor
# =============================================
EF = {
    "Cement": 0.81, "Micro silica": 0.0416, "Gypsum": 0.002536,
    "Limestone": 0.01577, "Fly ash": 0.004, "Aggregates": 0.00747,
    "Water": 0.000541, "Superplasticizer": 1.88
}

# =============================================
# ขนาดก้อนลูกบาศก์ 150 mm
# =============================================
SIDE_MM  = 150.0
AREA_MM2 = SIDE_MM ** 2   # 22,500 mm²

# =============================================
# Scoring Functions
# =============================================
def get_strength_score(fc_MPa):
    if fc_MPa >= 30:   return 35
    elif fc_MPa >= 27: return 32
    elif fc_MPa >= 24: return 28
    elif fc_MPa >= 21: return 23
    elif fc_MPa >= 18: return 18
    elif fc_MPa >= 15: return 8
    else:              return 0

def get_carbon_score(co2):
    if co2 <= 240:   return 35
    elif co2 <= 270: return 32
    elif co2 <= 310: return 28
    elif co2 <= 350: return 23
    elif co2 <= 400: return 16
    else:            return 8

def get_efficiency_score(index):
    if index >= 0.16:   return 20
    elif index >= 0.13: return 16
    elif index >= 0.10: return 12
    elif index >= 0.07: return 8
    else:               return 4

def get_workability_score(slump_mm):
    s = slump_mm / 10
    if 7 <= s <= 20:                      return 10
    elif (5 <= s < 7) or (20 < s <= 22): return 6
    elif (3 <= s < 5) or (22 < s <= 25): return 3
    else:                                 return 0

def calc_scores(fc_MPa, co2, slump_mm):
    sc1 = get_strength_score(fc_MPa)
    sc2 = get_carbon_score(co2)
    idx = fc_MPa / co2 if co2 > 0 else 0
    sc3 = get_efficiency_score(idx)
    sc4 = get_workability_score(slump_mm)
    return sc1, sc2, sc3, sc4, sc1+sc2+sc3+sc4, idx

# =============================================
# Supabase helpers
# =============================================
def load_teams():
    res = supabase.table("competition_scores").select("*").order("score_total", desc=True).execute()
    return res.data or []

def upsert_team(data: dict):
    existing = [t["team_name"] for t in load_teams()]
    if data["team_name"] in existing:
        supabase.table("competition_scores").update(data).eq("team_name", data["team_name"]).execute()
    else:
        supabase.table("competition_scores").insert(data).execute()

def delete_team(team_name: str):
    supabase.table("competition_scores").delete().eq("team_name", team_name).execute()

def reset_all():
    supabase.table("competition_scores").delete().neq("id", 0).execute()

# =============================================
# Session state
# =============================================
for key, val in {
    "active_tab": "mix",       # mix | strength | result
    "teams_mix": {},           # {team_name: {mix, carbon, slump}}
    "teams_strength": {},      # {team_name: {p1,p2,p3, fc1,fc2,fc3, fc_avg}}
    "team_list": [],           # ลำดับทีม
    "current_spec": 1,         # กำลังกรอกก้อนที่ 1/2/3
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# =============================================
# TABS
# =============================================
st.title("🏗️ Low Carbon Concrete – หน้ากรรมการ")
tab_mix, tab_str, tab_res = st.tabs([
    "📋 ขั้นตอนที่ 1 — ส่วนผสม & Slump",
    "💪 ขั้นตอนที่ 2 — กำลังอัดคอนกรีต",
    "🏆 ขั้นตอนที่ 3 — สรุปคะแนน"
])

# ══════════════════════════════════════════
# TAB 1: ส่วนผสม + Slump
# ══════════════════════════════════════════
with tab_mix:
    st.subheader("📋 บันทึกส่วนผสมและ Slump แต่ละทีม")
    st.caption("กรอกทีละทีม กด **บันทึก** แล้วเปลี่ยนทีมต่อไป")
    st.divider()

    c1, c2 = st.columns([1, 2])
    with c1:
        team_name_mix = st.text_input("ชื่อทีม", placeholder="เช่น ทีม A", key="tnm")
        slump = st.number_input("Slump (mm)", min_value=0, value=100, step=5, key="slump_in")
        st.caption(f"= {slump/10:.1f} cm")

    with c2:
        st.write("**ส่วนผสม (kg/m³)**")
        cl, cr = st.columns(2)
        mix = {}
        for i, mat in enumerate(EF.keys()):
            with (cl if i % 2 == 0 else cr):
                mix[mat] = st.number_input(mat, min_value=0.0, value=0.0, step=1.0, key=f"m1_{mat}")

    carbon = sum(mix[m] * EF[m] for m in mix)
    if carbon > 0:
        st.success(f"🌱 Embodied Carbon = **{carbon:.2f} kgCO₂e/m³**")
        with st.expander("รายละเอียด Embodied Carbon"):
            rows = [{"วัสดุ": m, "ปริมาณ (kg/m³)": mix[m], "EF": EF[m],
                     "Carbon (kgCO₂e/m³)": round(mix[m]*EF[m],4)} for m in mix if mix[m] > 0]
            if rows:
                st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    else:
        st.info("กรุณากรอกส่วนผสม")

    st.divider()
    ba, bb = st.columns([2, 1])
    with ba:
        if st.button("✅ บันทึกส่วนผสมทีมนี้", type="primary", use_container_width=True):
            if not team_name_mix:
                st.error("กรุณาใส่ชื่อทีม")
            elif carbon <= 0:
                st.error("กรุณากรอกส่วนผสม")
            else:
                st.session_state.teams_mix[team_name_mix] = {
                    "mix": mix.copy(), "carbon": round(carbon, 2), "slump": slump
                }
                if team_name_mix not in st.session_state.team_list:
                    st.session_state.team_list.append(team_name_mix)
                st.success(f"✅ บันทึกทีม **{team_name_mix}** แล้ว (CO₂ = {carbon:.2f} kgCO₂e/m³)")
                st.rerun()
    with bb:
        if st.button("🗑️ ลบทีมที่เลือก", use_container_width=True):
            del_t = st.session_state.get("del_mix_target", "—")
            if del_t != "—":
                st.session_state.teams_mix.pop(del_t, None)
                st.session_state.teams_strength.pop(del_t, None)
                if del_t in st.session_state.team_list:
                    st.session_state.team_list.remove(del_t)
                st.rerun()

    # ตารางทีมที่บันทึกแล้ว
    if st.session_state.teams_mix:
        st.write(f"**ทีมที่บันทึกแล้ว ({len(st.session_state.teams_mix)} ทีม):**")
        del_target = st.selectbox("เลือกทีมที่ต้องการลบ",
                                  ["—"] + list(st.session_state.teams_mix.keys()),
                                  key="del_mix_target")
        rows2 = []
        for tn, d in st.session_state.teams_mix.items():
            str_done = "✅" if tn in st.session_state.teams_strength else "⏳"
            rows2.append({"ทีม": tn, "CO₂ (kgCO₂e/m³)": d["carbon"],
                          "Slump (mm)": d["slump"], "กำลังอัด": str_done})
        st.dataframe(pd.DataFrame(rows2), hide_index=True, use_container_width=True)

# ══════════════════════════════════════════
# TAB 2: กำลังอัดคอนกรีต (ทีละ 1 ก้อน วนทุกทีม)
# ══════════════════════════════════════════
with tab_str:
    st.subheader("💪 กรอกแรงกด P — ทีละ 1 ก้อน วนครบทุกทีม")

    teams_with_mix = st.session_state.team_list
    if not teams_with_mix:
        st.warning("⚠️ ยังไม่มีทีม — กรุณาบันทึกส่วนผสมใน Tab 1 ก่อนครับ")
    else:
        spec = st.session_state.current_spec   # 1, 2, หรือ 3
        st.info(f"🔢 กำลังกรอก **ก้อนที่ {spec}** — วนครบทุก {len(teams_with_mix)} ทีม")
        st.caption(f"ลูกบาศก์ {SIDE_MM:.0f} mm | หน้าตัด A = {AREA_MM2:,.0f} mm²")
        st.divider()

        # แสดงฟอร์มทุกทีมพร้อมกันในหน้าเดียว (ก้อนที่ spec)
        st.write(f"### กรอกแรงกด P (kN) ก้อนที่ {spec} — ทุกทีม")

        p_inputs = {}
        cols_teams = st.columns(min(len(teams_with_mix), 4))
        for i, tn in enumerate(teams_with_mix):
            with cols_teams[i % 4]:
                # ดึงค่าเดิมถ้ามี
                prev = st.session_state.teams_strength.get(tn, {})
                prev_p = prev.get(f"p{spec}", 0.0)
                p_val = st.number_input(
                    f"**{tn}** — P{spec} (kN)",
                    min_value=0.0, value=float(prev_p), step=1.0,
                    key=f"p{spec}_{tn}"
                )
                fc_val = (p_val * 1000) / AREA_MM2  # MPa
                p_inputs[tn] = {"p": p_val, "fc": round(fc_val, 2)}
                if p_val > 0:
                    st.caption(f"f'c = **{fc_val:.2f} MPa** ({fc_val*10:.0f} ksc)")

        st.divider()
        col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
        with col_s1:
            if st.button(f"✅ บันทึกก้อนที่ {spec} ครบทุกทีม", type="primary", use_container_width=True):
                saved_count = 0
                for tn, vals in p_inputs.items():
                    if vals["p"] > 0:
                        if tn not in st.session_state.teams_strength:
                            st.session_state.teams_strength[tn] = {}
                        st.session_state.teams_strength[tn][f"p{spec}"] = vals["p"]
                        st.session_state.teams_strength[tn][f"fc{spec}"] = vals["fc"]
                        saved_count += 1
                st.success(f"✅ บันทึกก้อนที่ {spec} ครบ {saved_count} ทีมแล้ว")
                st.rerun()

        with col_s2:
            if spec < 3:
                if st.button(f"▶️ ไปก้อนที่ {spec+1}", use_container_width=True):
                    st.session_state.current_spec = spec + 1
                    st.rerun()
        with col_s3:
            if spec > 1:
                if st.button(f"◀️ ย้อนก้อนที่ {spec-1}", use_container_width=True):
                    st.session_state.current_spec = spec - 1
                    st.rerun()

        # ── ตารางสรุปกำลังอัดทุกทีม ──
        st.divider()
        st.write("#### 📊 ตารางสรุปกำลังอัดทุกทีม")
        sum_rows = []
        for tn in teams_with_mix:
            d = st.session_state.teams_strength.get(tn, {})
            fc1 = d.get("fc1", None)
            fc2 = d.get("fc2", None)
            fc3 = d.get("fc3", None)
            fc_vals = [v for v in [fc1, fc2, fc3] if v is not None and v > 0]
            fc_avg  = round(np.mean(fc_vals), 2) if fc_vals else None
            sum_rows.append({
                "ทีม": tn,
                "P1 (kN)": d.get("p1", "—"), "f'c1 (MPa)": fc1 or "—",
                "P2 (kN)": d.get("p2", "—"), "f'c2 (MPa)": fc2 or "—",
                "P3 (kN)": d.get("p3", "—"), "f'c3 (MPa)": fc3 or "—",
                "f'c เฉลี่ย (MPa)": fc_avg or "—",
                "f'c เฉลี่ย (ksc)": round(fc_avg*10, 0) if fc_avg else "—",
            })
        st.dataframe(pd.DataFrame(sum_rows), hide_index=True, use_container_width=True)

# ══════════════════════════════════════════
# TAB 3: สรุปคะแนนและบันทึก Supabase
# ══════════════════════════════════════════
with tab_res:
    st.subheader("🏆 สรุปคะแนนและบันทึกผล")

    teams_ready = [
        tn for tn in st.session_state.team_list
        if tn in st.session_state.teams_mix
        and tn in st.session_state.teams_strength
        and st.session_state.teams_strength[tn].get("fc1")
        and st.session_state.teams_strength[tn].get("fc2")
        and st.session_state.teams_strength[tn].get("fc3")
    ]
    teams_not_ready = [tn for tn in st.session_state.team_list if tn not in teams_ready]

    if teams_not_ready:
        st.warning(f"⏳ ยังรอข้อมูลครบ 3 ก้อน: {', '.join(teams_not_ready)}")

    # ── ปุ่ม Reset — แสดงเสมอ ──
    with st.expander("🗑️ Reset ข้อมูลทั้งหมด (ล้าง Session + Supabase)"):
        st.warning("⚠️ การ Reset จะลบข้อมูลทุกทีมออกจากทั้ง Session และ Supabase ไม่สามารถกู้คืนได้")
        if st.button("🗑️ ยืนยัน Reset ทั้งหมด", type="secondary", use_container_width=True):
            st.session_state.teams_mix = {}
            st.session_state.teams_strength = {}
            st.session_state.team_list = []
            st.session_state.current_spec = 1
            reset_all()
            st.toast("ล้างข้อมูลทั้งหมดแล้ว", icon="🗑️")
            st.rerun()

    st.divider()

    if not teams_ready:
        st.info("กรอกกำลังอัดครบ 3 ก้อนทุกทีมใน Tab 2 ก่อนครับ")
    else:
        # คำนวณคะแนนทุกทีม
        result_rows = []
        for tn in teams_ready:
            mix_d = st.session_state.teams_mix[tn]
            str_d = st.session_state.teams_strength[tn]
            fc1, fc2, fc3 = str_d["fc1"], str_d["fc2"], str_d["fc3"]
            fc_avg = round(np.mean([fc1, fc2, fc3]), 2)
            co2    = mix_d["carbon"]
            slump  = mix_d["slump"]
            sc1, sc2, sc3, sc4, total, idx = calc_scores(fc_avg, co2, slump)
            result_rows.append({
                "ทีม": tn,
                "f'c1 (MPa)": fc1, "f'c2 (MPa)": fc2, "f'c3 (MPa)": fc3,
                "f'c เฉลี่ย (MPa)": fc_avg,
                "f'c เฉลี่ย (ksc)": round(fc_avg*10, 0),
                "CO₂ (kgCO₂e/m³)": co2,
                "Slump (mm)": slump,
                "Index": round(idx, 4),
                "กำลังอัด /35": sc1, "CO₂ /35": sc2,
                "Index /20": sc3, "Workability /10": sc4,
                "รวม /100": total,
            })

        df_res = pd.DataFrame(result_rows).sort_values("รวม /100", ascending=False).reset_index(drop=True)
        df_res.index += 1
        df_res.index.name = "อันดับ"
        st.dataframe(df_res, use_container_width=True)

        st.divider()
        col_r1, col_r2 = st.columns([2, 1])
        with col_r1:
            if st.button("☁️ บันทึกผลทั้งหมดขึ้น Supabase (Leaderboard)", type="primary", use_container_width=True):
                errors = []
                for row in result_rows:
                    try:
                        payload = {
                            "team_name":        row["ทีม"],
                            "fc_mpa":           row["f'c เฉลี่ย (MPa)"],
                            "fc_ksc":           int(row["f'c เฉลี่ย (ksc)"]),
                            "co2":              row["CO₂ (kgCO₂e/m³)"],
                            "slump_cm":         round(row["Slump (mm)"] / 10, 1),
                            "idx":              row["Index"],
                            "score_strength":   row["กำลังอัด /35"],
                            "score_carbon":     row["CO₂ /35"],
                            "score_index":      row["Index /20"],
                            "score_workability":row["Workability /10"],
                            "score_total":      row["รวม /100"],
                        }
                        upsert_team(payload)
                    except Exception as e:
                        errors.append(f"{row['ทีม']}: {e}")
                if errors:
                    st.error("เกิดข้อผิดพลาด:\n" + "\n".join(errors))
                else:
                    st.success(f"✅ บันทึก {len(result_rows)} ทีมขึ้น Supabase แล้ว! Leaderboard อัปเดตทันที 🎉")
                    st.balloons()

        with col_r2:
            pass

        # ── แก้ไขรายทีม ──
        st.divider()
        st.write("#### ✏️ แก้ไขข้อมูลรายทีม")
        edit_t = st.selectbox("เลือกทีมที่ต้องการแก้ไข", ["— เลือกทีม —"] + teams_ready, key="edit_sel")
        if edit_t != "— เลือกทีม —":
            mix_d = st.session_state.teams_mix[edit_t]
            str_d = st.session_state.teams_strength[edit_t]
            ea, eb, ec_ = st.columns(3)
            with ea:
                new_p1 = st.number_input("P1 (kN)", value=float(str_d.get("p1",0)), step=1.0, key="ep1")
                st.caption(f"f'c1 = {(new_p1*1000/AREA_MM2):.2f} MPa")
            with eb:
                new_p2 = st.number_input("P2 (kN)", value=float(str_d.get("p2",0)), step=1.0, key="ep2")
                st.caption(f"f'c2 = {(new_p2*1000/AREA_MM2):.2f} MPa")
            with ec_:
                new_p3 = st.number_input("P3 (kN)", value=float(str_d.get("p3",0)), step=1.0, key="ep3")
                st.caption(f"f'c3 = {(new_p3*1000/AREA_MM2):.2f} MPa")

            new_slump = st.number_input("Slump (mm)", value=int(mix_d["slump"]), step=5, key="eslump")

            if st.button(f"💾 บันทึกการแก้ไขทีม {edit_t}", type="primary"):
                st.session_state.teams_strength[edit_t]["p1"] = new_p1
                st.session_state.teams_strength[edit_t]["fc1"] = round(new_p1*1000/AREA_MM2, 2)
                st.session_state.teams_strength[edit_t]["p2"] = new_p2
                st.session_state.teams_strength[edit_t]["fc2"] = round(new_p2*1000/AREA_MM2, 2)
                st.session_state.teams_strength[edit_t]["p3"] = new_p3
                st.session_state.teams_strength[edit_t]["fc3"] = round(new_p3*1000/AREA_MM2, 2)
                st.session_state.teams_mix[edit_t]["slump"] = new_slump
                st.success(f"✅ แก้ไขทีม {edit_t} แล้ว กด 'บันทึกขึ้น Supabase' เพื่ออัปเดต Leaderboard")
                st.rerun()

    st.divider()
    with st.expander("📖 ตารางอ้างอิงเกณฑ์การให้คะแนน"):
        ca, cb = st.columns(2)
        with ca:
            st.write("**หมวด 1: กำลังอัด (35 คะแนน)** — f'c เฉลี่ย 3 ก้อน")
            st.dataframe(pd.DataFrame({
                "f'c เฉลี่ย": ["≥ 30 MPa (≥ 300 ksc)","27–29 MPa","24–26 MPa",
                               "21–23 MPa","18–20 MPa","15–17 MPa","< 15 MPa"],
                "คะแนน": [35,32,28,23,18,8,0]
            }), hide_index=True, use_container_width=True)
            st.write("**หมวด 3: Index (20 คะแนน)**")
            st.dataframe(pd.DataFrame({
                "Index = f'c/CO₂": ["≥ 0.16","0.13–0.159","0.10–0.129","0.07–0.099","< 0.07"],
                "คะแนน": [20,16,12,8,4]
            }), hide_index=True, use_container_width=True)
        with cb:
            st.write("**หมวด 2: CO₂ Emission (35 คะแนน)**")
            st.dataframe(pd.DataFrame({
                "CO₂ (kgCO₂e/m³)": ["≤ 240","241–270","271–310","311–350","351–400","> 400"],
                "คะแนน": [35,32,28,23,16,8]
            }), hide_index=True, use_container_width=True)
            st.write("**หมวด 4: Workability (10 คะแนน)**")
            st.dataframe(pd.DataFrame({
                "Slump": ["7–20 cm","5–6 หรือ 21–22 cm","3–4 หรือ 23–25 cm","< 3 หรือ > 25 cm"],
                "คะแนน": [10,6,3,0]
            }), hide_index=True, use_container_width=True)
