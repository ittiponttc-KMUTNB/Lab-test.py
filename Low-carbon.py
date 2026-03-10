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
# Constants
# =============================================
EF = {
    "Cement": 0.81, "Micro silica": 0.0416, "Gypsum": 0.002536,
    "Limestone": 0.01577, "Fly ash": 0.004, "Aggregates": 0.00747,
    "Water": 0.000541, "Superplasticizer": 1.88
}
SIDE_MM  = 150.0
AREA_MM2 = SIDE_MM ** 2  # 22,500 mm²

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

def p_to_fc(p_kN):
    """แปลง P (kN) → f'c (MPa)"""
    return round((p_kN * 1000) / AREA_MM2, 2)

# =============================================
# Supabase helpers
# =============================================
def load_teams():
    res = supabase.table("competition_scores").select("*").order("score_total", desc=True).execute()
    return res.data or []

def upsert_team(payload: dict):
    existing = [t["team_name"] for t in load_teams()]
    if payload["team_name"] in existing:
        supabase.table("competition_scores").update(payload).eq("team_name", payload["team_name"]).execute()
    else:
        supabase.table("competition_scores").insert(payload).execute()

def delete_team(team_name: str):
    supabase.table("competition_scores").delete().eq("team_name", team_name).execute()

def reset_all():
    supabase.table("competition_scores").delete().neq("id", 0).execute()

# =============================================
# UI
# =============================================
st.title("🏗️ Low Carbon Concrete – หน้ากรรมการ")
st.divider()

# ─── เลือกโหมด: กรอกใหม่ หรือ แก้ไขทีมเดิม ───
teams_db = load_teams()
team_names_db = [t["team_name"] for t in teams_db]

mode = st.radio("โหมด", ["➕ กรอกทีมใหม่", "✏️ แก้ไขทีมที่มีอยู่"], horizontal=True)

selected_team_data = None
if mode == "✏️ แก้ไขทีมที่มีอยู่":
    if not team_names_db:
        st.warning("ยังไม่มีทีมในระบบ")
        st.stop()
    edit_name = st.selectbox("เลือกทีมที่ต้องการแก้ไข", team_names_db)
    selected_team_data = next(t for t in teams_db if t["team_name"] == edit_name)

st.divider()

# ─── ฟอร์มหลัก ───
col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("1️⃣ ข้อมูลทีม & ส่วนผสม")

    # ชื่อทีม
    default_name = selected_team_data["team_name"] if selected_team_data else ""
    team_name = st.text_input("ชื่อทีม", value=default_name,
                               disabled=(mode == "✏️ แก้ไขทีมที่มีอยู่"),
                               placeholder="เช่น ทีม A")

    # Slump
    default_slump = int(selected_team_data["slump_cm"] * 10) if selected_team_data else 100
    slump = st.number_input("Slump (mm)", min_value=0, value=default_slump, step=5)
    st.caption(f"= {slump/10:.1f} cm")

    st.write("**ส่วนผสม (kg/m³)**")
    mix = {}
    cl, cr = st.columns(2)
    for i, mat in enumerate(EF.keys()):
        with (cl if i % 2 == 0 else cr):
            mix[mat] = st.number_input(mat, min_value=0.0, value=0.0, step=1.0, key=f"mix_{mat}")

    carbon = sum(mix[m] * EF[m] for m in mix)
    if carbon > 0:
        st.success(f"🌱 Embodied Carbon = **{carbon:.2f} kgCO₂e/m³**")
        with st.expander("รายละเอียด Embodied Carbon"):
            rows = [{"วัสดุ": m, "ปริมาณ (kg/m³)": mix[m], "EF": EF[m],
                     "Carbon": round(mix[m]*EF[m], 4)} for m in mix if mix[m] > 0]
            if rows:
                st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    else:
        # โหมดแก้ไข: ใช้ค่า co2 เดิมจาก DB ถ้ายังไม่ได้กรอกส่วนผสมใหม่
        if selected_team_data:
            carbon = selected_team_data["co2"]
            st.info(f"🌱 Embodied Carbon (เดิม) = **{carbon:.2f} kgCO₂e/m³**  \nกรอกส่วนผสมใหม่เพื่ออัปเดต")

with col_right:
    st.subheader("2️⃣ กำลังอัดคอนกรีต 3 ก้อน")
    st.caption(f"ลูกบาศก์ {SIDE_MM:.0f} mm | A = {AREA_MM2:,.0f} mm²")

    # ค่าเริ่มต้นจาก DB (ถ้าแก้ไข) — แปลงกลับจาก fc เป็น P โดยประมาณ
    def fc_to_p(fc_mpa):
        return round(fc_mpa * AREA_MM2 / 1000, 1) if fc_mpa else 0.0

    prev_fc_avg = selected_team_data["fc_mpa"] if selected_team_data else None

    p1 = st.number_input("P₁ (kN) — ก้อนที่ 1", min_value=0.0,
                          value=fc_to_p(prev_fc_avg) if selected_team_data else 0.0,
                          step=1.0, key="p1")
    p2 = st.number_input("P₂ (kN) — ก้อนที่ 2", min_value=0.0,
                          value=fc_to_p(prev_fc_avg) if selected_team_data else 0.0,
                          step=1.0, key="p2")
    p3 = st.number_input("P₃ (kN) — ก้อนที่ 3", min_value=0.0,
                          value=fc_to_p(prev_fc_avg) if selected_team_data else 0.0,
                          step=1.0, key="p3")

    # คำนวณ fc แต่ละก้อน
    fc_list = []
    fc_vals_show = []
    for i, p in enumerate([p1, p2, p3], 1):
        if p > 0:
            fc = p_to_fc(p)
            fc_list.append(fc)
            fc_vals_show.append(f"f'c{i} = {fc:.2f} MPa ({fc*10:.0f} ksc)")
        else:
            fc_vals_show.append(f"f'c{i} = —")

    for txt in fc_vals_show:
        st.caption(txt)

    # คำนวณ f'c เฉลี่ยจากก้อนที่กรอกแล้ว
    if fc_list:
        fc_avg = round(np.mean(fc_list), 2)
        n_spec = len(fc_list)
        st.info(f"**f'c เฉลี่ย ({n_spec} ก้อน) = {fc_avg:.2f} MPa ({fc_avg*10:.0f} ksc)**")
    else:
        fc_avg = None
        st.warning("กรุณากรอก P อย่างน้อย 1 ก้อน")

# ─── Preview คะแนน ───
st.divider()
if fc_avg and carbon > 0 and team_name:
    sc1, sc2, sc3, sc4, total, idx = calc_scores(fc_avg, carbon, slump)
    st.write("**📊 Preview คะแนน:**")
    pa, pb, pc, pd_, pe = st.columns(5)
    pa.metric("กำลังอัด /35", sc1)
    pb.metric("CO₂ /35", sc2)
    pc.metric("Index /20", sc3)
    pd_.metric("Workability /10", sc4)
    pe.metric("🏆 รวม /100", total)

# ─── ปุ่มบันทึก ───
st.divider()
btn_col1, btn_col2, btn_col3 = st.columns([2, 1, 1])

with btn_col1:
    save_btn = st.button("✅ บันทึก / อัปเดตคะแนน → Leaderboard",
                          type="primary", use_container_width=True)
with btn_col2:
    del_target = st.selectbox("เลือกทีมที่ต้องการลบ",
                               ["— เลือกทีม —"] + team_names_db, key="del_sel")
    if st.button("🗑️ ลบทีมที่เลือก", use_container_width=True):
        if del_target != "— เลือกทีม —":
            delete_team(del_target)
            st.toast(f"ลบทีม '{del_target}' แล้ว", icon="🗑️")
            st.rerun()

with btn_col3:
    st.write("")
    with st.expander("⚠️ Reset ทั้งหมด"):
        if st.button("🗑️ ยืนยัน Reset", type="secondary", use_container_width=True):
            reset_all()
            st.toast("ล้างข้อมูลทั้งหมดแล้ว", icon="🗑️")
            st.rerun()

# ─── บันทึก ───
if save_btn:
    if not team_name:
        st.error("กรุณาใส่ชื่อทีม")
    elif carbon <= 0:
        st.error("กรุณากรอกส่วนผสม หรือตรวจสอบข้อมูล CO₂")
    elif not fc_avg:
        st.error("กรุณากรอก P อย่างน้อย 1 ก้อน")
    else:
        sc1, sc2, sc3, sc4, total, idx = calc_scores(fc_avg, carbon, slump)
        payload = {
            "team_name":         team_name,
            "fc_mpa":            fc_avg,
            "fc_ksc":            int(fc_avg * 10),
            "co2":               round(carbon, 2),
            "slump_cm":          round(slump / 10, 1),
            "idx":               round(idx, 4),
            "score_strength":    sc1,
            "score_carbon":      sc2,
            "score_index":       sc3,
            "score_workability": sc4,
            "score_total":       total,
        }
        try:
            upsert_team(payload)
            n_spec = len(fc_list)
            st.success(f"✅ บันทึกทีม **{team_name}** ({n_spec} ก้อน) — f'c เฉลี่ย = **{fc_avg:.2f} MPa** | รวม **{total}** คะแนน")
            st.balloons()
            st.rerun()
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

# ─── ตารางสรุป ───
st.divider()
teams_data = load_teams()
st.subheader(f"📋 ข้อมูลที่บันทึกแล้ว ({len(teams_data)} ทีม)")
if teams_data:
    df = pd.DataFrame(teams_data).rename(columns={
        "team_name": "ทีม", "fc_mpa": "f'c เฉลี่ย (MPa)", "fc_ksc": "f'c (ksc)",
        "co2": "CO₂ (kgCO₂e/m³)", "slump_cm": "Slump (cm)", "idx": "Index",
        "score_strength": "กำลังอัด /35", "score_carbon": "CO₂ /35",
        "score_index": "Index /20", "score_workability": "Workability /10",
        "score_total": "รวม /100"
    })
    cols = ["ทีม","f'c เฉลี่ย (MPa)","f'c (ksc)","CO₂ (kgCO₂e/m³)","Slump (cm)",
            "Index","กำลังอัด /35","CO₂ /35","Index /20","Workability /10","รวม /100"]
    df_show = df[cols].sort_values("รวม /100", ascending=False).reset_index(drop=True)
    df_show.index += 1
    df_show.index.name = "อันดับ"
    st.dataframe(df_show, use_container_width=True)
    st.caption("👆 ไปที่หน้า **Leaderboard** (เมนูซ้าย) เพื่อฉาย projector")
else:
    st.info("ยังไม่มีข้อมูล — กรอกทีมแรกได้เลยครับ")

# ─── ตารางอ้างอิง ───
st.divider()
with st.expander("📖 ตารางอ้างอิงเกณฑ์การให้คะแนน"):
    ca, cb = st.columns(2)
    with ca:
        st.write("**หมวด 1: กำลังอัด (35 คะแนน)** — f'c เฉลี่ย")
        st.caption("< 15 MPa = 0 คะแนน (ไม่ตัดสิทธิ์)")
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
