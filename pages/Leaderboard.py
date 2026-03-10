import streamlit as st
import pandas as pd
import time
from supabase import create_client

st.set_page_config(
    page_title="🏆 Leaderboard – Low Carbon Concrete",
    page_icon="🏆",
    layout="wide"
)

@st.cache_resource
def get_supabase():
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])

supabase = get_supabase()

def load_teams():
    res = supabase.table("competition_scores").select("*").order("score_total", desc=True).execute()
    return res.data or []

# ─── CSS ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;600;800&display=swap');
html, body, [class*="css"] { font-family: 'Prompt', sans-serif !important; background-color: #0a0e1a !important; color: #e8eaf6 !important; }
.stApp { background: linear-gradient(135deg, #0a0e1a 0%, #0d1b2a 50%, #0a1628 100%); }
.lb-title { text-align:center; font-size:2.8rem; font-weight:800; background:linear-gradient(90deg,#ffd700,#ff8c00,#ffd700); background-size:200% auto; -webkit-background-clip:text; -webkit-text-fill-color:transparent; animation:shimmer 3s linear infinite; letter-spacing:2px; padding: 20px 0 4px 0; }
.lb-subtitle { text-align:center; font-size:1.1rem; color:#7986cb; margin-bottom:16px; }
@keyframes shimmer { to { background-position:200% center; } }
.team-card { border-radius:16px; padding:18px 24px; margin:10px 0; display:flex; align-items:center; gap:20px; animation:slideIn 0.5s ease-out; border:1px solid rgba(255,255,255,0.08); }
.team-card.rank-1 { background:linear-gradient(135deg,rgba(255,215,0,0.18),rgba(255,140,0,0.10)); border-color:rgba(255,215,0,0.4); box-shadow:0 0 30px rgba(255,215,0,0.2); }
.team-card.rank-2 { background:linear-gradient(135deg,rgba(192,192,192,0.15),rgba(169,169,169,0.08)); border-color:rgba(192,192,192,0.3); }
.team-card.rank-3 { background:linear-gradient(135deg,rgba(205,127,50,0.15),rgba(184,115,51,0.08)); border-color:rgba(205,127,50,0.3); }
.team-card.rank-other { background:rgba(255,255,255,0.04); }
@keyframes slideIn { from{opacity:0;transform:translateX(-20px)} to{opacity:1;transform:translateX(0)} }
.rank-num { font-size:2rem; font-weight:800; min-width:60px; text-align:center; }
.rank-1 .rank-num{color:#ffd700} .rank-2 .rank-num{color:#c0c0c0} .rank-3 .rank-num{color:#cd7f32} .rank-other .rank-num{color:#546e7a}
.team-name { font-size:1.6rem; font-weight:700; flex:1; color:#e8eaf6; }
.score-bar-wrap { flex:2; }
.score-bar-bg { background:rgba(255,255,255,0.08); border-radius:8px; height:14px; overflow:hidden; }
.score-bar-fill { height:100%; border-radius:8px; transition:width 1s ease; }
.rank-1 .score-bar-fill{background:linear-gradient(90deg,#ffd700,#ff8c00)}
.rank-2 .score-bar-fill{background:linear-gradient(90deg,#a0aec0,#e2e8f0)}
.rank-3 .score-bar-fill{background:linear-gradient(90deg,#c97b38,#e8a87c)}
.rank-other .score-bar-fill{background:linear-gradient(90deg,#3949ab,#5c6bc0)}
.badges { display:flex; gap:8px; flex-wrap:wrap; justify-content:flex-end; }
.badge { background:rgba(255,255,255,0.08); border-radius:8px; padding:4px 10px; font-size:0.78rem; color:#b0bec5; text-align:center; min-width:64px; }
.badge span { display:block; font-size:1.05rem; font-weight:700; color:#e8eaf6; }
.total-score { font-size:2.2rem; font-weight:800; min-width:80px; text-align:right; }
.rank-1 .total-score{color:#ffd700} .rank-2 .total-score{color:#c0c0c0} .rank-3 .total-score{color:#cd7f32} .rank-other .total-score{color:#5c6bc0}
.status-bar { display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.04); border-radius:10px; padding:8px 20px; margin-bottom:16px; font-size:0.9rem; color:#546e7a; }
.dot-live { display:inline-block; width:10px; height:10px; background:#4caf50; border-radius:50%; animation:pulse 1.5s ease-in-out infinite; margin-right:6px; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:0.5;transform:scale(1.3)} }
.no-data { text-align:center; padding:80px 0; color:#546e7a; font-size:1.3rem; }
</style>
""", unsafe_allow_html=True)

# ─── Header ────────────────────────────────────────────────────────────────
col_title, col_ctrl = st.columns([5, 1])
with col_title:
    st.markdown('<div class="lb-title">🏆 LOW CARBON CONCRETE</div>', unsafe_allow_html=True)
    st.markdown('<div class="lb-subtitle">LEADERBOARD – REAL TIME</div>', unsafe_allow_html=True)
with col_ctrl:
    st.write("")
    auto = st.toggle("🔄 Auto-refresh", value=True)
    refresh_sec = st.select_slider("ทุก (วิ)", options=[3,5,10,15,30], value=5) if auto else 5

# ─── Status bar ────────────────────────────────────────────────────────────
teams = load_teams()
n_teams = len(teams)
now_str = time.strftime("%H:%M:%S")

st.markdown(f"""
<div class="status-bar">
    <div><span class="dot-live"></span> LIVE &nbsp;|&nbsp; {n_teams} ทีม</div>
    <div>อัปเดตล่าสุด: {now_str}</div>
    <div>{'🔄 Auto ' + str(refresh_sec) + 's' if auto else '⏸ Manual'}</div>
</div>
""", unsafe_allow_html=True)

# ─── Leaderboard ────────────────────────────────────────────────────────────
if not teams:
    st.markdown('<div class="no-data">⏳ รอกรรมการกรอกข้อมูล...<br><span style="font-size:0.9rem;color:#37474f;">ไปที่หน้า <b>Low-carbon</b> เพื่อเริ่มบันทึกคะแนน</span></div>', unsafe_allow_html=True)
else:
    medals = {1:"🥇", 2:"🥈", 3:"🥉"}
    for i, row in enumerate(teams):
        rank = i + 1
        rank_cls = f"rank-{rank}" if rank <= 3 else "rank-other"
        rank_label = medals.get(rank, f"#{rank}")
        pct = int(row["score_total"] / 100 * 100)
        st.markdown(f"""
        <div class="team-card {rank_cls}">
            <div class="rank-num">{rank_label}</div>
            <div class="team-name">{row['team_name']}</div>
            <div class="score-bar-wrap">
                <div style="font-size:0.75rem;color:#546e7a;margin-bottom:4px;">
                    f'c {row['fc_mpa']} MPa &nbsp;|&nbsp; CO₂ {row['co2']} kgCO₂e/m³ &nbsp;|&nbsp; Slump {row['slump_cm']} cm
                </div>
                <div class="score-bar-bg">
                    <div class="score-bar-fill" style="width:{pct}%;"></div>
                </div>
            </div>
            <div class="badges">
                <div class="badge">กำลังอัด<span>{row['score_strength']}/35</span></div>
                <div class="badge">CO₂<span>{row['score_carbon']}/35</span></div>
                <div class="badge">Index<span>{row['score_index']}/20</span></div>
                <div class="badge">Work.<span>{row['score_workability']}/10</span></div>
            </div>
            <div class="total-score">{row['score_total']}</div>
        </div>
        """, unsafe_allow_html=True)

    # ─── Bar chart ──────────────────────────────────────────────────────────
    st.write("")
    with st.expander("📊 แผนภูมิเปรียบเทียบคะแนนแต่ละหมวด", expanded=True):
        df = pd.DataFrame(teams)
        df = df.rename(columns={
            "team_name":"ทีม","score_strength":"กำลังอัด /35",
            "score_carbon":"CO₂ /35","score_index":"Index /20",
            "score_workability":"Workability /10"
        })
        chart_df = df[["ทีม","กำลังอัด /35","CO₂ /35","Index /20","Workability /10"]].set_index("ทีม")
        st.bar_chart(chart_df, color=["#ffd700","#5c6bc0","#4caf50","#ef5350"])

# ─── Auto-refresh ────────────────────────────────────────────────────────────
if auto:
    time.sleep(refresh_sec)
    st.rerun()
