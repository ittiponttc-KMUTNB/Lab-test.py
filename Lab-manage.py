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
# Supabase client
# =============================================
@st.cache_resource
def get_supabase():
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase = get_supabase()

# =============================================
# Emission Factor
# =============================================
EF = {
    "Cement": 0.81,
    "Micro silica": 0.0416,
    "Gypsum": 0.002536,
    "Limestone": 0.01577,
    "Fly ash": 0.004,
    "Aggregates": 0.00747,
    "Water": 0.000541,
    "Superplasticizer": 1.88
}

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
    slump_cm = slump_mm / 10
    if 7 <= slump_cm <= 20:                              return 10
    elif (5 <= slump_cm < 7) or (20 < slump_cm <= 22):  return 6
    elif (3 <= slump_cm < 5) or (22 < slump_cm <= 25):  return 3
    else:                                                return 0

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

def save_team(data: dict):
    supabase.table("competition_scores").insert(data).execute()

def update_team(team_name: str, data: dict):
    supabase.table("competition_scores").update(data).eq("team_name", team_name).execute()

def delete_team(team_name: str):
    supabase.table("competition_scores").delete().eq("team_name", team_name).execute()

def reset_all():
    supabase.table("competition_scores").delete().neq("id", 0).execute()

# =============================================
# UI
# =============================================
st.title("🏗️ Low Carbon Concrete – หน้ากรรมการ")
st.caption("กรอกข้อมูลแต่ละทีม แล้วกด **บันทึกคะแนน** — Leaderboard จะอัปเดตทันที")
st.divider()

# ───── ส่วนที่ 1: ส่วนผสม ─────
st.subheader("1️⃣ ส่วนผสมคอนกรีต (kg/m³)")
col_left, col_right = st.columns(2)
mix = {}
for i, mat in enumerate(EF.keys()):
    with (col_left if i % 2 == 0 else col_right):
        mix[mat] = st.number_input(mat, min_value=0.0, value=0.0, step=1.0, key=f"mix_{mat}")

carbon = sum(mix[m] * EF[m] for m in mix)

if carbon > 0:
    st.success(f"🌱 Embodied Carbon = **{carbon:.2f} kgCO₂e/m³**")
    with st.expander("ดูรายละเอียด Embodied Carbon"):
        rows = [{"วัสดุ": m, "ปริมาณ (kg/m³)": mix[m], "EF (kgCO₂e/kg)": EF[m],
                 "Carbon (kgCO₂e/m³)": round(mix[m]*EF[m], 4)} for m in mix if mix[m] > 0]
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
else:
    st.info("กรุณากรอกส่วนผสมเพื่อคำนวณ Embodied Carbon")

st.divider()

# ───── ส่วนที่ 2: ผลทดสอบ ─────
st.subheader("2️⃣ ผลทดสอบคอนกรีต")
c1, c2, c3 = st.columns(3)
with c1:
    team_name = st.text_input("ชื่อทีม", placeholder="เช่น ทีม A")
with c2:
    fc_MPa = st.number_input("f'c ที่ 1 วัน (MPa)", min_value=0.0, value=0.0, step=0.5)
    st.caption(f"= {fc_MPa*10:.0f} ksc")
with c3:
    slump = st.number_input("Slump (mm)", min_value=0, value=100, step=5)
    st.caption(f"= {slump/10:.1f} cm")

# preview คะแนน
if carbon > 0 and fc_MPa > 0 and team_name:
    sc1, sc2, sc3_v, sc4, total, idx = calc_scores(fc_MPa, carbon, slump)
    st.write("**Preview คะแนน:**")
    pa, pb, pc, pd_, pe = st.columns(5)
    pa.metric("กำลังอัด /35", sc1)
    pb.metric("CO₂ /35", sc2)
    pc.metric("Index /20", sc3_v)
    pd_.metric("Workability /10", sc4)
    pe.metric("🏆 รวม /100", total)

st.divider()

# ───── ปุ่ม ─────
col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])

with col_btn1:
    save_btn = st.button("✅ บันทึกคะแนนทีมนี้", type="primary", use_container_width=True)

with col_btn2:
    # เลือกทีมที่ต้องการลบ/แก้ไข
    teams_now = load_teams()
    team_names = [t["team_name"] for t in teams_now]
    del_target = st.selectbox("เลือกทีมที่ต้องการลบ", ["—"] + team_names, key="del_target")
    if st.button("🗑️ ลบทีมที่เลือก", use_container_width=True):
        if del_target != "—":
            delete_team(del_target)
            st.toast(f"ลบทีม '{del_target}' แล้ว", icon="🗑️")
            st.rerun()

with col_btn3:
    st.write("")
    st.write("")
    if st.button("⚠️ Reset ทั้งหมด", use_container_width=True):
        reset_all()
        st.toast("ล้างข้อมูลทั้งหมดแล้ว", icon="🗑️")
        st.rerun()

# ───── บันทึก ─────
if save_btn:
    if not team_name:
        st.error("กรุณาใส่ชื่อทีม")
    elif carbon <= 0:
        st.error("กรุณากรอกส่วนผสมให้ครบ")
    elif fc_MPa <= 0:
        st.error("กรุณากรอก f'c")
    else:
        sc1, sc2, sc3_v, sc4, total, idx = calc_scores(fc_MPa, carbon, slump)
        existing = [t["team_name"] for t in load_teams()]
        payload = {
            "team_name": team_name,
            "fc_mpa": fc_MPa,
            "fc_ksc": int(fc_MPa * 10),
            "co2": round(carbon, 2),
            "slump_cm": round(slump / 10, 1),
            "idx": round(idx, 4),
            "score_strength": sc1,
            "score_carbon": sc2,
            "score_index": sc3_v,
            "score_workability": sc4,
            "score_total": total,
        }
        try:
            if team_name in existing:
                update_team(team_name, payload)
                st.success(f"✅ อัปเดตทีม **{team_name}** แล้ว! คะแนนรวม = **{total}** คะแนน")
            else:
                save_team(payload)
                st.success(f"✅ บันทึกทีม **{team_name}** แล้ว! คะแนนรวม = **{total}** คะแนน")
            st.balloons()
            st.rerun()
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

st.divider()

# ───── ตารางสรุป ─────
teams_data = load_teams()
st.subheader(f"📋 ข้อมูลที่บันทึกแล้ว ({len(teams_data)} ทีม)")

if teams_data:
    df = pd.DataFrame(teams_data)
    df = df.rename(columns={
        "team_name": "ทีม", "fc_mpa": "f'c (MPa)", "fc_ksc": "f'c (ksc)",
        "co2": "CO₂ (kgCO₂e/m³)", "slump_cm": "Slump (cm)", "idx": "Index",
        "score_strength": "กำลังอัด /35", "score_carbon": "CO₂ /35",
        "score_index": "Index /20", "score_workability": "Workability /10",
        "score_total": "รวม /100"
    })
    show_cols = ["ทีม","f'c (MPa)","f'c (ksc)","CO₂ (kgCO₂e/m³)","Slump (cm)","Index",
                 "กำลังอัด /35","CO₂ /35","Index /20","Workability /10","รวม /100"]
    df_show = df[show_cols].sort_values("รวม /100", ascending=False).reset_index(drop=True)
    df_show.index += 1
    df_show.index.name = "อันดับ"
    st.dataframe(df_show, use_container_width=True)
    st.caption("👆 ไปที่หน้า **Leaderboard** (เมนูซ้าย) เพื่อฉาย projector ครับ")
else:
    st.info("ยังไม่มีข้อมูล — กรอกทีมแรกได้เลยครับ")

st.divider()

# ───── ตารางอ้างอิง ─────
with st.expander("📖 ตารางอ้างอิงเกณฑ์การให้คะแนน"):
    ca, cb = st.columns(2)
    with ca:
        st.write("**หมวด 1: กำลังอัด (35 คะแนน)**")
        st.caption("< 15 MPa = 0 คะแนน (ไม่ตัดสิทธิ์)")
        st.dataframe(pd.DataFrame({
            "f'c ที่ 1 วัน": ["≥ 30 MPa (≥ 300 ksc)","27–29 MPa (270–290 ksc)",
                              "24–26 MPa (240–260 ksc)","21–23 MPa (210–230 ksc)",
                              "18–20 MPa (180–200 ksc)","15–17 MPa (150–170 ksc)","< 15 MPa (< 150 ksc)"],
            "คะแนน": [35,32,28,23,18,8,0]
        }), hide_index=True, use_container_width=True)
        st.write("**หมวด 3: Index (20 คะแนน)**")
        st.dataframe(pd.DataFrame({
            "Index": ["≥ 0.16","0.13–0.159","0.10–0.129","0.07–0.099","< 0.07"],
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
            "Slump": ["7–20 cm","5–6 cm หรือ 21–22 cm","3–4 cm หรือ 23–25 cm","< 3 หรือ > 25 cm"],
            "คะแนน": [10,6,3,0]
        }), hide_index=True, use_container_width=True)
