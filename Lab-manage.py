import streamlit as st
import pandas as pd
import json
import time
from datetime import datetime, date
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from supabase import create_client
import cloudinary
import cloudinary.uploader

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ระบบเบิก-คืนอุปกรณ์ TTC",
    page_icon="🔬",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ─── Secrets ──────────────────────────────────────────────────────────────────
# Streamlit Cloud → Settings → Secrets:
#
# [supabase]
# url = "https://xxxxx.supabase.co"
# key = "eyJhbG..."
#
# [cloudinary]
# cloud_name = "your_cloud_name"
# api_key    = "123456789"
# api_secret = "abcdef..."
#
# [app]
# admin_password = "admin1234"
# logo_url       = "https://res.cloudinary.com/xxx/image/upload/lab_equipment/ttc_logo.png"
# ──────────────────────────────────────────────────────────────────────────────

SUPABASE_URL   = st.secrets["supabase"]["url"]
SUPABASE_KEY   = st.secrets["supabase"]["key"]
ADMIN_PASSWORD = st.secrets["app"]["admin_password"]
LOGO_URL       = st.secrets.get("app", {}).get("logo_url", "")

cloudinary.config(
    cloud_name=st.secrets["cloudinary"]["cloud_name"],
    api_key=st.secrets["cloudinary"]["api_key"],
    api_secret=st.secrets["cloudinary"]["api_secret"],
    secure=True
)

MAX_UPLOAD_MB = 5  # จำกัดขนาดรูปสูงสุด

# ─── Supabase Client ──────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)

sb = get_supabase()

# ─── MOBILE CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-size: 16px; }
div.stButton > button { height: 3rem; font-size: 1rem; border-radius: 10px; }
input, textarea { font-size: 16px !important; min-height: 2.8rem !important; }
div[data-baseweb="select"] { font-size: 16px; }

div[data-testid="metric-container"] {
    background: #f8f9fa; border-radius: 12px;
    padding: 12px; border: 1px solid #e0e0e0;
}
.eq-card {
    background: white; border-radius: 12px; padding: 14px;
    margin-bottom: 10px; border: 1px solid #e0e0e0;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); line-height: 1.7;
}
.overdue-alert {
    background: #fff3cd; border-left: 4px solid #ff6b6b;
    border-radius: 8px; padding: 10px 14px; margin: 6px 0; line-height: 1.7;
}
.section-header {
    font-size: 1.05rem; font-weight: 700; color: #1F4E79;
    margin: 16px 0 8px 0; padding-bottom: 4px; border-bottom: 2px solid #e0e0e0;
}
.badge {
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 0.78rem; font-weight: 600; color: white;
}
</style>
""", unsafe_allow_html=True)

# ─── ADMIN AUTH ───────────────────────────────────────────────────────────────
def is_admin():
    return st.session_state.get("is_admin", False)

def admin_login_widget():
    if is_admin():
        st.sidebar.success("🔓 Admin Mode")
        if st.sidebar.button("🔒 ออกจาก Admin", use_container_width=True):
            st.session_state.is_admin = False
            st.rerun()
    else:
        with st.sidebar.expander("🔒 Admin Login"):
            pwd = st.text_input("รหัสผ่าน", type="password", key="admin_pwd_input")
            if st.button("เข้าสู่ระบบ Admin", use_container_width=True):
                if pwd == ADMIN_PASSWORD:
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    st.error("รหัสผ่านไม่ถูกต้อง")

# ═════════════════════════════════════════════════════════════════════════════
# DATABASE LAYER — Supabase with retry & error handling          [FIX #5, #7]
# ═════════════════════════════════════════════════════════════════════════════
def _sb_retry(func, retries=2, delay=1):
    """Retry wrapper สำหรับ Supabase calls"""
    for attempt in range(retries + 1):
        try:
            return func()
        except Exception as e:
            if attempt == retries:
                raise e
            time.sleep(delay)

def query_table(table, select="*", filters=None, order=None, limit=None):
    """Query Supabase table → DataFrame (with retry)"""
    def _do():
        q = sb.table(table).select(select)
        if filters:
            for col, op, val in filters:
                if op == "in_":
                    q = q.in_(col, val)
                else:
                    q = getattr(q, op if op != "is" else "is_")(col, val)
        if order:
            for col_name, opts in order:
                q = q.order(col_name, **opts)
        if limit:
            q = q.limit(limit)
        resp = q.execute()
        return pd.DataFrame(resp.data) if resp.data else pd.DataFrame()
    return _sb_retry(_do)

def insert_row(table, data):
    """Insert → return inserted row dict"""
    def _do():
        resp = sb.table(table).insert(data).execute()
        return resp.data[0] if resp.data else None
    return _sb_retry(_do)

def update_rows(table, data, match_col, match_val):
    """Update row(s) by single column match"""
    def _do():
        resp = sb.table(table).update(data).eq(match_col, match_val).execute()
        return resp.data
    return _sb_retry(_do)

def delete_rows(table, match_col=None, match_val=None, delete_all=False):
    """Delete row(s) — delete_all ใช้ gt id 0"""
    def _do():
        if delete_all:
            resp = sb.table(table).delete().gt("id", 0).execute()
        else:
            resp = sb.table(table).delete().eq(match_col, match_val).execute()
        return resp.data
    return _sb_retry(_do)

def batch_reset_equipment():
    """รีเซ็ตทุกอุปกรณ์ — ดึงมาแล้ว update เฉพาะตัวที่ต้องเปลี่ยน [FIX #4]"""
    df = query_table("equipment", select="id,total_qty,available_qty,status")
    if df.empty:
        return
    need_update = df[(df["available_qty"] != df["total_qty"]) | (df["status"] != "พร้อมใช้")]
    for _, r in need_update.iterrows():
        update_rows("equipment",
                     {"available_qty": int(r["total_qty"]), "status": "พร้อมใช้"},
                     "id", int(r["id"]))

# ═════════════════════════════════════════════════════════════════════════════
# DATA LOADING — Batch fetch + merge (eliminates N+1)            [FIX #1, #6]
# ═════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=30, show_spinner=False)
def load_sidebar_stats():
    """[FIX #3] Cache sidebar stats 30 วินาที"""
    df_eq = query_table("equipment", select="id,available_qty")
    n_eq = len(df_eq)
    avail = int(df_eq["available_qty"].sum()) if not df_eq.empty else 0

    today_str = str(date.today())
    df_active = query_table("transactions", select="id,due_date",
                            filters=[("status", "eq", "ยืมอยู่")])
    n_borr = len(df_active)
    # [FIX #6] คำนวณ overdue โดย filter due_date < today
    n_over = len(df_active[df_active["due_date"] < today_str]) if not df_active.empty else 0
    return n_eq, avail, n_borr, n_over

def load_active_transactions_enriched():
    """[FIX #1] ดึง transactions + equipment + borrowers ทีเดียว แล้ว merge"""
    df_tx = query_table("transactions",
                        select="id,equipment_id,borrower_id,qty,borrow_date,due_date,condition_out,note,return_date,condition_in,status",
                        filters=[("status", "eq", "ยืมอยู่")],
                        order=[("due_date", {"desc": False})])
    if df_tx.empty:
        return pd.DataFrame()
    return _enrich_transactions(df_tx)

def load_pending_transactions_enriched():
    """[FIX #1] ดึง transactions รอตรวจสอบ + merge"""
    df_tx = query_table("transactions",
                        select="id,equipment_id,borrower_id,qty,borrow_date,due_date,return_date,condition_out,condition_in,note,status",
                        filters=[("status", "eq", "รอตรวจสอบ")],
                        order=[("return_date", {"desc": False})])
    if df_tx.empty:
        return pd.DataFrame()
    return _enrich_transactions(df_tx)

def _enrich_transactions(df_tx):
    """Batch fetch equipment + borrowers เฉพาะ id ที่ต้องการ แล้ว merge"""
    eq_ids = df_tx["equipment_id"].unique().tolist()
    br_ids = df_tx["borrower_id"].unique().tolist()

    # ดึงเฉพาะ id ที่เกี่ยวข้อง (ไม่ดึงทั้ง table)
    df_eq = query_table("equipment", select="id,code,name,image_url,category",
                        filters=[("id", "in_", eq_ids)]) if eq_ids else pd.DataFrame()

    df_br = query_table("borrowers", select="id,name,type,phone,student_id,department",
                        filters=[("id", "in_", br_ids)]) if br_ids else pd.DataFrame()

    # Merge: transactions ← equipment ← borrowers
    merged = df_tx.merge(
        df_eq.rename(columns={"id": "eq_id", "name": "eq_name", "image_url": "eq_image_url",
                               "code": "eq_code", "category": "eq_category"}),
        left_on="equipment_id", right_on="eq_id", how="left"
    ).merge(
        df_br.rename(columns={"id": "br_id", "name": "br_name", "type": "br_type",
                               "phone": "br_phone", "student_id": "br_student_id",
                               "department": "br_department"}),
        left_on="borrower_id", right_on="br_id", how="left"
    )
    return merged

def load_report_data(date_from, date_to, status_filter):
    """[FIX #1] รายงาน — batch fetch + merge"""
    filters = [("borrow_date", "gte", str(date_from)), ("borrow_date", "lte", str(date_to))]
    if status_filter != "ทั้งหมด":
        filters.append(("status", "eq", status_filter))

    df_tx = query_table("transactions", filters=filters, order=[("id", {"desc": True})])
    if df_tx.empty:
        return pd.DataFrame()

    df_eq = query_table("equipment", select="id,code,name",
                        filters=[("id", "in_", df_tx["equipment_id"].unique().tolist())])
    df_br = query_table("borrowers", select="id,name,type,student_id,department,phone",
                        filters=[("id", "in_", df_tx["borrower_id"].unique().tolist())])

    merged = df_tx.merge(
        df_eq.rename(columns={"id": "eq_id", "name": "eq_name", "code": "eq_code"}),
        left_on="equipment_id", right_on="eq_id", how="left"
    ).merge(
        df_br.rename(columns={"id": "br_id", "name": "br_name", "type": "br_type",
                               "student_id": "br_student_id", "department": "br_department",
                               "phone": "br_phone"}),
        left_on="borrower_id", right_on="br_id", how="left"
    )

    report = pd.DataFrame({
        "TX#": merged["id"],
        "รหัส": merged["eq_code"],
        "อุปกรณ์": merged["eq_name"],
        "ผู้เบิก": merged["br_name"],
        "ประเภท": merged["br_type"],
        "รหัส/ID": merged["br_student_id"],
        "ภาควิชา": merged["br_department"],
        "โทรศัพท์": merged["br_phone"],
        "จำนวน": merged["qty"],
        "วันเบิก": merged["borrow_date"],
        "กำหนดคืน": merged["due_date"],
        "วันคืน": merged["return_date"],
        "สภาพตอนเบิก": merged["condition_out"],
        "สภาพตอนคืน": merged["condition_in"],
        "สถานะ": merged["status"],
        "หมายเหตุ": merged["note"]
    })
    return report

def load_equipment_summary():
    """[FIX #1] สรุปอุปกรณ์ — นับ tx_count ด้วย batch"""
    df_eq = query_table("equipment",
                        select="id,code,name,category,total_qty,available_qty,status",
                        order=[("code", {"desc": False})])
    if df_eq.empty:
        return pd.DataFrame()

    # นับจำนวนครั้งที่เบิกทั้งหมดใน 1 query
    df_tx = query_table("transactions", select="id,equipment_id")
    if not df_tx.empty:
        tx_counts = df_tx.groupby("equipment_id").size().reset_index(name="ครั้งที่เบิก")
    else:
        tx_counts = pd.DataFrame(columns=["equipment_id", "ครั้งที่เบิก"])

    df_eq = df_eq.merge(tx_counts, left_on="id", right_on="equipment_id", how="left")
    df_eq["ครั้งที่เบิก"] = df_eq["ครั้งที่เบิก"].fillna(0).astype(int)
    df_eq["กำลังยืม"] = df_eq["total_qty"].astype(int) - df_eq["available_qty"].astype(int)

    return df_eq[["code", "name", "category", "total_qty", "available_qty",
                   "กำลังยืม", "status", "ครั้งที่เบิก"]].rename(columns={
        "code": "รหัส", "name": "ชื่ออุปกรณ์", "category": "หมวดหมู่",
        "total_qty": "ทั้งหมด", "available_qty": "พร้อมใช้", "status": "สถานะ"
    })

# ═════════════════════════════════════════════════════════════════════════════
# CLOUDINARY — Upload with validation + Optimized URL            [FIX #2, #8]
# ═════════════════════════════════════════════════════════════════════════════
def upload_image(file_obj, public_id):
    """[FIX #8] Upload with size validation + auto optimization"""
    # ตรวจขนาดไฟล์
    file_size_mb = len(file_obj.getvalue()) / (1024 * 1024)
    if file_size_mb > MAX_UPLOAD_MB:
        st.error(f"❌ ไฟล์ใหญ่เกิน {MAX_UPLOAD_MB} MB (ไฟล์นี้ {file_size_mb:.1f} MB) กรุณาลดขนาดก่อนอัพโหลด")
        return None
    try:
        result = cloudinary.uploader.upload(
            file_obj,
            public_id=public_id,
            overwrite=True,
            folder="lab_equipment",
            resource_type="image",
            transformation=[
                {"width": 1200, "crop": "limit"},          # จำกัดสูงสุด 1200px
                {"quality": "auto", "fetch_format": "auto"} # optimize อัตโนมัติ
            ]
        )
        return result.get("secure_url")
    except Exception as e:
        st.error(f"❌ อัพโหลดรูปไม่สำเร็จ: {e}")
        return None

def optimized_url(original_url, width=400, height=300, crop="fill"):
    """[FIX #2] สร้าง URL ที่ optimize แล้วจาก Cloudinary URL ต้นฉบับ"""
    if not original_url or not isinstance(original_url, str) or "cloudinary" not in original_url:
        return original_url
    return original_url.replace(
        "/upload/", f"/upload/w_{width},h_{height},c_{crop},q_auto,f_auto/"
    )

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def show_image(image_url, width=100, size="preview"):
    """[FIX #2] แสดงรูปพร้อม Cloudinary optimization"""
    SIZES = {"thumb": (200, 150), "preview": (400, 300), "full": (800, 600)}
    w_img, h_img = SIZES.get(size, (400, 300))
    opt_url = optimized_url(image_url, w_img, h_img)

    w_style = f"{width}px" if isinstance(width, int) else width
    if opt_url and isinstance(opt_url, str) and opt_url.startswith("http"):
        st.markdown(
            f'<img src="{opt_url}" '
            f'style="width:{w_style};max-width:100%;border-radius:8px;'
            f'border:1px solid #ddd;object-fit:contain;" loading="lazy">',
            unsafe_allow_html=True)
    else:
        st.markdown(
            f'<div style="width:{w_style};height:120px;background:#f0f2f6;'
            f'border-radius:8px;display:flex;align-items:center;justify-content:center;'
            f'font-size:2rem;border:1px solid #ddd;">📦</div>',
            unsafe_allow_html=True)

STATUS_COLOR = {
    "พร้อมใช้": "#28a745", "ยืมออก": "#ffc107",
    "ชำรุด":    "#dc3545", "สูญหาย": "#6c757d"
}

def badge(text, color):
    return f'<span class="badge" style="background:{color};">{text}</span>'

def overdue_days(due_str):
    try:
        d = datetime.strptime(str(due_str), "%Y-%m-%d").date()
        delta = (date.today() - d).days
        return max(delta, 0)
    except (ValueError, TypeError):
        return 0

# ═════════════════════════════════════════════════════════════════════════════
# NAVIGATION + HEADER                                           [FIX #3, #9]
# ═════════════════════════════════════════════════════════════════════════════
PAGES = [("🏠","Dashboard"), ("📦","อุปกรณ์"), ("➕","เบิก"), ("✅","คืน"), ("📋","รายงาน"), ("⚙️","ตั้งค่า")]

def nav():
    if "page" not in st.session_state:
        st.session_state.page = "Dashboard"

    # Sidebar — Admin Login + สรุปข้อมูล
    with st.sidebar:
        st.markdown("## 🔬 ระบบอุปกรณ์ Lab")
        st.markdown("---")
        admin_login_widget()
        st.markdown("---")
        try:
            n_eq, avail, n_borr, n_over = load_sidebar_stats()
            st.metric("📦 อุปกรณ์", n_eq)
            st.metric("🔄 กำลังยืม", n_borr)
            if n_over > 0:
                st.error(f"⚠️ เกินกำหนด {n_over} รายการ")
        except Exception:
            st.caption("⏳ กำลังโหลด...")

    # ── [FIX #9] Header พร้อมโลโก้ TTC ─────────────────────────────────────
    logo_html = ""
    if LOGO_URL:
        logo_opt = optimized_url(LOGO_URL, 140, 140, crop="fit") if "cloudinary" in LOGO_URL else LOGO_URL
        logo_html = f'<img src="{logo_opt}" style="width:70px;height:auto;margin-bottom:4px;" loading="lazy"><br>'

    st.markdown(
        f'<div style="text-align:center; margin-bottom:10px;">'
        f'{logo_html}'
        f'<span style="font-size:1.05rem; font-weight:700; color:#1F4E79; white-space:nowrap;">'
        f'ระบบบริหารจัดการห้องปฏิบัติการ TTC</span><br>'
        f'<span style="font-size:0.82rem; color:#888;">'
        f'ภาควิชาครุศาสตร์โยธา — มจพ.</span></div>',
        unsafe_allow_html=True
    )

    # ── Top nav bar ────────────────────────────────────────────────────────
    NAV_ITEMS = [
        ("Dashboard", "🏠", "หน้าหลัก"), ("อุปกรณ์", "📦", "อุปกรณ์"),
        ("เบิก", "➕", "เบิก"), ("คืน", "✅", "คืน"),
        ("รายงาน", "📋", "รายงาน"), ("ตั้งค่า", "⚙️", "ตั้งค่า"),
    ]
    cols = st.columns(len(NAV_ITEMS))
    for i, (name, icon, label) in enumerate(NAV_ITEMS):
        active = st.session_state.page == name
        with cols[i]:
            if st.button(f"{icon}\n{label}", key=f"tnav_{name}",
                         type="primary" if active else "secondary",
                         use_container_width=True):
                st.session_state.page = name
                st.rerun()

# ═════════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD                                               [FIX #1, #6]
# ═════════════════════════════════════════════════════════════════════════════
def page_dashboard():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">🏠 ภาพรวม</p>', unsafe_allow_html=True)

    # Stats — ใช้ cached sidebar stats
    try:
        n_eq, available, active_tx, overdue_count = load_sidebar_stats()
    except Exception:
        st.error("❌ ไม่สามารถเชื่อมต่อ Supabase ได้ กรุณาตรวจสอบการตั้งค่า")
        return

    c1, c2 = st.columns(2)
    c1.metric("📦 อุปกรณ์ทั้งหมด", n_eq)
    c2.metric("✅ พร้อมใช้", available)
    c1.metric("🔄 กำลังยืม", active_tx)
    c2.metric("⚠️ เกินกำหนด", overdue_count)

    # [FIX #1] ดึง enriched transactions ทีเดียว
    df = load_active_transactions_enriched()

    if overdue_count > 0 and not df.empty:
        st.error(f"⚠️ มีอุปกรณ์เกินกำหนดคืน {overdue_count} รายการ!")
        today_str = str(date.today())
        df_od = df[df["due_date"] < today_str].copy()
        df_od["days_over"] = df_od["due_date"].apply(overdue_days)
        df_od = df_od.sort_values("days_over", ascending=False)
        for _, r in df_od.iterrows():
            phone_str = f"📞 {r['br_phone']}" if pd.notna(r.get('br_phone')) and r.get('br_phone') else ""
            st.markdown(f"""
            <div class="overdue-alert">
                🔴 <b>{r['eq_code']} — {r['eq_name']}</b><br>
                👤 {r['br_name']} {phone_str}<br>
                📅 กำหนดคืน {r['due_date']} &nbsp;
                <b style="color:red;">เกิน {r['days_over']} วัน</b>
            </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-header">📋 รายการที่กำลังยืมอยู่</div>', unsafe_allow_html=True)
    if df.empty:
        st.info("ไม่มีรายการยืมในขณะนี้ ✅")
    else:
        for _, r in df.iterrows():
            od = overdue_days(r["due_date"])
            bc = "#ff6b6b" if od > 0 else "#28a745"
            img_url = optimized_url(r.get("eq_image_url"), 100, 75)
            st.markdown(f"""
            <div class="eq-card" style="border-left:4px solid {bc};">
                <b>{r['eq_code']}</b> — {r['eq_name']}<br>
                👤 {r['br_name']} ({r['br_type']}) &nbsp; 📦 {r['qty']} ชิ้น<br>
                📅 เบิก {r['borrow_date']} | คืน <b>{r['due_date']}</b>
                {"&nbsp;<b style='color:red;'>⚠️ เกิน "+str(od)+" วัน</b>" if od > 0 else ""}
            </div>""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# PAGE: EQUIPMENT                                               [FIX #1, #2]
# ═════════════════════════════════════════════════════════════════════════════
def page_equipment():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">📦 รายการอุปกรณ์</p>', unsafe_allow_html=True)

    search = st.text_input("🔍 ค้นหา ชื่อ / รหัส / หมวดหมู่", placeholder="พิมพ์เพื่อค้นหา...")

    df_all_eq = query_table("equipment",
                            select="id,code,name,category,total_qty,available_qty,status,image_url,description",
                            order=[("code", {"desc": False})])
    cats = sorted(df_all_eq["category"].dropna().unique().tolist()) if not df_all_eq.empty else []
    cat_filter = st.selectbox("📂 หมวดหมู่", ["ทั้งหมด"] + cats)

    df = df_all_eq.copy()
    if search:
        mask = (df["name"].str.contains(search, case=False, na=False) |
                df["code"].str.contains(search, case=False, na=False) |
                df["category"].str.contains(search, case=False, na=False))
        df = df[mask]
    if cat_filter != "ทั้งหมด":
        df = df[df["category"] == cat_filter]

    st.caption(f"พบ {len(df)} รายการ")

    # [FIX #1] batch fetch ผู้ยืมล่าสุดทั้งหมด
    df_last_tx = pd.DataFrame()
    df_br_all = pd.DataFrame()
    if not df.empty:
        df_all_tx = query_table("transactions", select="id,equipment_id,borrower_id,borrow_date,created_at",
                                order=[("created_at", {"desc": True})])
        if not df_all_tx.empty:
            df_last_tx = df_all_tx.drop_duplicates(subset=["equipment_id"], keep="first")
            br_ids = df_last_tx["borrower_id"].unique().tolist()
            df_br_all = query_table("borrowers", select="id,name")
            if not df_br_all.empty:
                df_br_all = df_br_all[df_br_all["id"].isin(br_ids)]

    for _, r in df.iterrows():
        with st.expander(f"{r['code']} — {r['name']}"):
            c1, c2 = st.columns([1, 2])
            with c1:
                show_image(r.get("image_url"), width="100%", size="preview")
            with c2:
                st.markdown(
                    f"**หมวด:** {r['category'] or '-'}<br>"
                    f"**ทั้งหมด:** {r['total_qty']} | **พร้อมใช้:** {r['available_qty']}<br>"
                    f"**สถานะ:** {badge(r['status'], STATUS_COLOR.get(r['status'],'#888'))}",
                    unsafe_allow_html=True)
                if r.get("description"):
                    st.caption(r["description"])

            # ผู้ยืมล่าสุด — ใช้ข้อมูลที่ fetch มาแล้ว (ไม่ query ใน loop)
            if not df_last_tx.empty:
                lt = df_last_tx[df_last_tx["equipment_id"] == r["id"]]
                if not lt.empty and not df_br_all.empty:
                    br = df_br_all[df_br_all["id"] == lt.iloc[0]["borrower_id"]]
                    if not br.empty:
                        st.caption(f"👤 ผู้ยืมล่าสุด: **{br.iloc[0]['name']}** ({lt.iloc[0]['borrow_date']})")

            if is_admin():
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("✏️ แก้ไข", key=f"edit_{r['id']}", use_container_width=True):
                        opts_tmp = ["➕ เพิ่มใหม่"] + [f"{x['code']} — {x['name']}" for _, x in df_all_eq.iterrows()]
                        matched_tmp = [o for o in opts_tmp if o.startswith(r["code"] + " —")]
                        if matched_tmp:
                            st.session_state["_next_sel"] = matched_tmp[0]
                        st.rerun()
                with btn_col2:
                    if st.button("🗑️ ลบ", key=f"del_{r['id']}", use_container_width=True):
                        active = query_table("transactions", select="id",
                                             filters=[("equipment_id","eq",r["id"]),("status","eq","ยืมอยู่")])
                        if len(active) > 0:
                            st.error("ไม่สามารถลบได้ มีการยืมอยู่")
                        else:
                            delete_rows("equipment", "id", r["id"])
                            st.success("ลบแล้ว")
                            load_sidebar_stats.clear()
                            st.rerun()

    # ── Admin: เพิ่ม/แก้ไขอุปกรณ์ ──────────────────────────────────────────
    if is_admin():
        st.markdown('<div class="section-header">➕ เพิ่ม / แก้ไขอุปกรณ์</div>', unsafe_allow_html=True)

        options = ["➕ เพิ่มใหม่"] + [f"{r['code']} — {r['name']}" for _, r in df_all_eq.iterrows()]

        if "_next_sel" in st.session_state:
            st.session_state["eq_edit_sel"] = st.session_state.pop("_next_sel")

        choice = st.selectbox("เลือกรายการ", options, key="eq_edit_sel")

        existing, eq_id = None, None
        if choice != "➕ เพิ่มใหม่":
            try:
                sel_code = choice.split(" — ")[0].strip()
                eq_data = df_all_eq[df_all_eq["code"] == sel_code]
                if not eq_data.empty:
                    existing = eq_data.iloc[0]
                    eq_id = int(existing["id"])
                else:
                    st.warning("⚠️ ไม่พบอุปกรณ์นี้ กรุณาเลือกใหม่")
            except Exception:
                st.warning("⚠️ เกิดข้อผิดพลาด กรุณาเลือกใหม่")

        if existing is not None:
            borrowed_now = int(existing["total_qty"]) - int(existing["available_qty"])
            st.markdown(
                f'<div class="eq-card" style="border-left:4px solid #1F4E79;">'
                f'📊 <b>ข้อมูลปัจจุบัน:</b> '
                f'ทั้งหมด <b>{existing["total_qty"]}</b> | '
                f'พร้อมใช้ <b>{existing["available_qty"]}</b> | '
                f'ยืมออก <b>{borrowed_now}</b>'
                f'</div>', unsafe_allow_html=True)

        form_key = f"frm_{eq_id if eq_id else 'new'}"
        with st.form(key=form_key):
            sv_code = st.text_input("รหัสอุปกรณ์ *",
                value=str(existing["code"]) if existing is not None else "")
            sv_name = st.text_input("ชื่ออุปกรณ์ *",
                value=str(existing["name"]) if existing is not None else "")
            sv_cat = st.text_input("หมวดหมู่",
                value=str(existing["category"] or "") if existing is not None else "")
            sv_qty = st.number_input("จำนวนทั้งหมด", min_value=1,
                value=int(existing["total_qty"]) if existing is not None else 1)
            sv_stat = st.selectbox("สถานะ", ["พร้อมใช้","ชำรุด","สูญหาย"],
                index=["พร้อมใช้","ชำรุด","สูญหาย"].index(str(existing["status"]))
                      if existing is not None else 0)
            sv_desc = st.text_area("รายละเอียด",
                value=str(existing["description"] or "") if existing is not None else "")
            sv_img = st.file_uploader(f"📷 รูปอุปกรณ์ (สูงสุด {MAX_UPLOAD_MB} MB)",
                                       type=["jpg","jpeg","png","gif"])

            if existing is not None and existing.get("image_url"):
                st.caption("รูปปัจจุบัน:")
                show_image(existing["image_url"], width="100%", size="preview")

            submitted = st.form_submit_button("💾 บันทึก", type="primary", use_container_width=True)

        if submitted:
            sv_code = sv_code.strip()
            sv_name = sv_name.strip()
            if not sv_code or not sv_name:
                st.error("กรุณากรอกรหัสและชื่ออุปกรณ์")
            else:
                dup = query_table("equipment", select="id", filters=[("code","eq",sv_code)])
                cur_id = eq_id if eq_id else -1
                if not dup.empty and (eq_id is None or int(dup.iloc[0]["id"]) != cur_id):
                    st.error(f"❌ รหัส '{sv_code}' มีอยู่แล้ว")
                else:
                    try:
                        img_url = existing.get("image_url") if existing is not None else None
                        if sv_img:
                            img_url = upload_image(sv_img, sv_code)
                            if img_url is None and sv_img:
                                st.stop()  # upload failed — อย่าบันทึก

                        if existing is None:
                            insert_row("equipment", {
                                "code": sv_code, "name": sv_name,
                                "category": sv_cat or None,
                                "total_qty": sv_qty, "available_qty": sv_qty,
                                "status": sv_stat, "image_url": img_url,
                                "description": sv_desc or None
                            })
                            st.success(f"✅ บันทึกการเพิ่มอุปกรณ์ '{sv_code} — {sv_name}' เรียบร้อยแล้ว")
                            st.session_state["_next_sel"] = "➕ เพิ่มใหม่"
                        else:
                            old_total = int(existing["total_qty"])
                            old_available = int(existing["available_qty"])
                            borrowed = old_total - old_available
                            diff = int(sv_qty) - old_total
                            new_available = old_available + diff
                            if new_available < 0:
                                st.error(f"❌ ลดจำนวนไม่ได้! ยืมออกอยู่ {borrowed} ชิ้น")
                                st.stop()
                            update_rows("equipment", {
                                "code": sv_code, "name": sv_name,
                                "category": sv_cat or None,
                                "total_qty": sv_qty, "available_qty": new_available,
                                "status": sv_stat, "image_url": img_url,
                                "description": sv_desc or None
                            }, "id", eq_id)
                            msg = f"✅ บันทึกการแก้ไข '{sv_code} — {sv_name}' เรียบร้อยแล้ว"
                            if diff > 0:
                                msg += f" (เพิ่มจำนวน +{diff} พร้อมใช้: {new_available})"
                            elif diff < 0:
                                msg += f" (ลดจำนวน {diff} พร้อมใช้: {new_available})"
                            st.success(msg)
                            st.session_state["_next_sel"] = f"{sv_code} — {sv_name}"
                        load_sidebar_stats.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
    else:
        st.info("🔒 การเพิ่ม/แก้ไขอุปกรณ์ สำหรับ Admin เท่านั้น")

# ═════════════════════════════════════════════════════════════════════════════
# PAGE: BORROW                                                  [FIX #2, #8]
# ═════════════════════════════════════════════════════════════════════════════
def page_borrow():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">➕ เบิกอุปกรณ์</p>', unsafe_allow_html=True)

    avail = query_table("equipment",
                        select="id,code,name,category,available_qty,image_url,description",
                        filters=[("status","eq","พร้อมใช้")],
                        order=[("category",{"desc":False}),("code",{"desc":False})])
    avail = avail[avail["available_qty"] > 0] if not avail.empty else avail

    if avail.empty:
        st.warning("⚠️ ไม่มีอุปกรณ์พร้อมใช้งานในขณะนี้")
        return

    st.markdown('<div class="section-header">📦 ขั้นตอนที่ 1 — เลือกอุปกรณ์</div>', unsafe_allow_html=True)

    cats_avail = ["ทั้งหมด"] + sorted(avail["category"].dropna().unique().tolist())
    cat_sel = st.selectbox("📂 กรองหมวดหมู่", cats_avail, key="bcat")
    search_eq = st.text_input("🔍 ค้นหาชื่อ / รหัสอุปกรณ์",
                               placeholder="พิมพ์เพื่อกรองรายการ...", key="bsearch")

    filtered = avail.copy()
    if cat_sel != "ทั้งหมด":
        filtered = filtered[filtered["category"] == cat_sel]
    if search_eq:
        mask = (filtered["name"].str.contains(search_eq, case=False, na=False) |
                filtered["code"].str.contains(search_eq, case=False, na=False))
        filtered = filtered[mask]

    if filtered.empty:
        st.info("ไม่พบอุปกรณ์ที่ค้นหา")
        return

    st.caption(f"แสดง {len(filtered)} รายการ")

    eq_opts = {f"{r['code']} — {r['name']}  (คงเหลือ {r['available_qty']})": r["id"]
               for _, r in filtered.iterrows()}
    selected_label = st.selectbox("เลือกอุปกรณ์ *", list(eq_opts.keys()), key="beq")
    eq_id = eq_opts[selected_label]
    eq_row = avail[avail["id"] == eq_id].iloc[0]

    c_img, c_info = st.columns([1, 2])
    with c_img:
        show_image(eq_row.get("image_url"), width="100%", size="preview")
    with c_info:
        st.markdown(f"**{eq_row['code']}** — {eq_row['name']}")
        st.markdown(f"หมวด: {eq_row['category'] or '-'}")
        st.markdown(f"คงเหลือ: **{eq_row['available_qty']}** ชิ้น")
        if eq_row.get("description"):
            st.caption(eq_row["description"])

    qty = st.number_input("จำนวนที่ต้องการเบิก *", min_value=1,
                          max_value=int(eq_row["available_qty"]), value=1)
    st.divider()

    st.markdown('<div class="section-header">👤 ขั้นตอนที่ 2 — ข้อมูลผู้เบิก</div>', unsafe_allow_html=True)
    borrower_type = st.radio("ประเภทผู้เบิก", ["นักศึกษา", "บุคลากร/อาจารย์"], horizontal=True)
    borrower_name = st.text_input("ชื่อ-นามสกุล *", placeholder="กรอกชื่อ-นามสกุล")
    student_id = st.text_input("รหัสนักศึกษา / รหัสพนักงาน", placeholder="เช่น 6601234567")
    department = st.text_input("ภาควิชา / หน่วยงาน")
    phone = st.text_input("เบอร์โทรศัพท์ *", placeholder="เช่น 081-234-5678")
    st.divider()

    st.markdown('<div class="section-header">📅 ขั้นตอนที่ 3 — วันที่</div>', unsafe_allow_html=True)
    borrow_date = st.date_input("วันที่เบิก *", value=date.today())
    due_date = st.date_input("วันกำหนดคืน *", value=date.today())
    condition_out = st.selectbox("สภาพอุปกรณ์ขณะเบิก", ["ปกติ", "มีรอยขีดข่วน", "ชำรุดบางส่วน"])
    note = st.text_area("หมายเหตุ (ถ้ามี)")
    st.divider()

    if st.button("✅ ยืนยันการเบิก", type="primary", use_container_width=True):
        if not borrower_name.strip():
            st.error("❌ กรุณากรอกชื่อผู้เบิก")
        elif not phone.strip():
            st.error("❌ กรุณากรอกเบอร์โทรศัพท์")
        elif due_date < borrow_date:
            st.error("❌ วันกำหนดคืนต้องไม่ก่อนวันที่เบิก")
        else:
            try:
                # Fix 7: หา borrower ที่มีอยู่แล้ว (by phone) ไม่สร้างซ้ำ
                existing_borr = query_table("borrowers", select="id",
                                            filters=[("phone", "eq", phone.strip())])
                if not existing_borr.empty:
                    borr_id = int(existing_borr.iloc[0]["id"])
                    # อัพเดทข้อมูลล่าสุด
                    update_rows("borrowers", {
                        "name": borrower_name.strip(), "type": borrower_type,
                        "student_id": student_id or None, "department": department or None
                    }, "id", borr_id)
                else:
                    borr = insert_row("borrowers", {
                        "name": borrower_name.strip(), "type": borrower_type,
                        "student_id": student_id or None, "department": department or None,
                        "phone": phone.strip()
                    })
                    borr_id = borr["id"]

                insert_row("transactions", {
                    "equipment_id": eq_id, "borrower_id": borr_id,
                    "qty": qty, "borrow_date": str(borrow_date),
                    "due_date": str(due_date), "condition_out": condition_out,
                    "note": note or None, "status": "ยืมอยู่"
                })
                new_avail = int(eq_row["available_qty"]) - qty
                update_rows("equipment", {"available_qty": new_avail}, "id", eq_id)
                load_sidebar_stats.clear()

                st.success(
                    f"✅ บันทึกสำเร็จ!\n\n"
                    f"👤 **{borrower_name}**\n"
                    f"📦 {eq_row['name']} จำนวน {qty} ชิ้น\n"
                    f"📅 กำหนดคืน: {due_date}"
                )
                st.balloons()
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")

# ═════════════════════════════════════════════════════════════════════════════
# PAGE: RETURN                                                  [FIX #1, #2]
# ═════════════════════════════════════════════════════════════════════════════
def page_return():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">✅ คืนอุปกรณ์</p>', unsafe_allow_html=True)

    # นับจำนวนรอตรวจสอบ
    df_pending = load_pending_transactions_enriched()
    n_pending = len(df_pending)
    tab2_label = f"🔍 รอตรวจสอบ ({n_pending})"

    tab1, tab2 = st.tabs(["📬 แจ้งคืนอุปกรณ์", tab2_label])

    # ── TAB 1: ผู้เบิกแจ้งคืน ────────────────────────────────────────────
    with tab1:
        st.info("👤 ผู้ยืมกรอกข้อมูลและแจ้งคืน จากนั้น Admin จะตรวจสอบและยืนยัน")
        search = st.text_input("🔍 ค้นหาชื่อผู้ยืม / รหัส / ชื่ออุปกรณ์",
                                placeholder="พิมพ์เพื่อค้นหา...", key="ret_search")

        # [FIX #1] batch fetch
        df = load_active_transactions_enriched()

        if search and not df.empty:
            s = search.lower()
            mask = (df["br_name"].str.lower().str.contains(s, na=False) |
                    df["eq_code"].str.lower().str.contains(s, na=False) |
                    df["eq_name"].str.lower().str.contains(s, na=False))
            df = df[mask]

        if df.empty:
            st.info("✅ ไม่มีรายการยืม" if not search else "ไม่พบรายการที่ค้นหา")
        else:
            st.caption(f"พบ {len(df)} รายการ")
            for _, r in df.iterrows():
                od = overdue_days(r["due_date"])
                label = f"TX#{r['id']} | {r['eq_code']} {r['eq_name']} | {r['br_name']}"
                if od > 0:
                    label += f" ⚠️ เกิน {od} วัน"

                with st.expander(label):
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        show_image(r.get("eq_image_url"), width="100%", size="preview")
                    with c2:
                        st.markdown(
                            f"📦 **{r['eq_code']}** — {r['eq_name']} ({r['qty']} ชิ้น)<br>"
                            f"👤 {r['br_name']} ({r['br_type']})<br>"
                            f"📅 เบิก {r['borrow_date']} | คืน <b>{r['due_date']}</b>"
                            + (f"<br><b style='color:red;'>⚠️ เกิน {od} วัน</b>" if od > 0 else ""),
                            unsafe_allow_html=True)
                        st.caption(f"สภาพตอนเบิก: {r['condition_out']}")

                    return_date = st.date_input("วันที่นำมาคืน", value=date.today(), key=f"rd_{r['id']}")
                    condition_in = st.selectbox("สภาพอุปกรณ์ที่นำมาคืน",
                                                ["ปกติ", "มีรอยขีดข่วน", "ชำรุด", "สูญหาย"],
                                                key=f"ci_{r['id']}")
                    return_note = st.text_input("หมายเหตุ", key=f"rn_{r['id']}")

                    if st.button("📬 แจ้งคืนอุปกรณ์", key=f"notify_{r['id']}",
                                 type="primary", use_container_width=True):
                        update_rows("transactions", {
                            "return_date": str(return_date), "condition_in": condition_in,
                            "note": return_note or None, "status": "รอตรวจสอบ"
                        }, "id", r["id"])
                        st.success("📬 แจ้งคืนเรียบร้อยแล้ว! กรุณารอ Admin ตรวจสอบและยืนยัน")
                        load_sidebar_stats.clear()
                        st.rerun()

    # ── TAB 2: Admin ตรวจสอบ ─────────────────────────────────────────────
    with tab2:
        df_wait = df_pending

        if df_wait.empty:
            st.info("✅ ไม่มีรายการรอตรวจสอบ")
        else:
            st.warning(f"🔍 มี {len(df_wait)} รายการรอ Admin ตรวจสอบ")
            for _, r in df_wait.iterrows():
                od = overdue_days(r["due_date"])
                label = f"TX#{r['id']} | {r['eq_code']} {r['eq_name']} | {r['br_name']}"

                with st.expander(label):
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        show_image(r.get("eq_image_url"), width="100%", size="preview")
                    with c2:
                        phone_str = f"📞 {r['br_phone']}" if pd.notna(r.get('br_phone')) and r.get('br_phone') else ""
                        st.markdown(
                            f"📦 **{r['eq_code']}** — {r['eq_name']} ({r['qty']} ชิ้น)<br>"
                            f"👤 {r['br_name']} ({r['br_type']}) {phone_str}<br>"
                            f"📅 เบิก {r['borrow_date']} | กำหนดคืน <b>{r['due_date']}</b><br>"
                            f"📅 แจ้งคืนวันที่: <b>{r.get('return_date','')}</b>"
                            + (f"<br><b style='color:red;'>⚠️ เกินกำหนด {od} วัน</b>" if od > 0 else ""),
                            unsafe_allow_html=True)
                        st.markdown(f"**สภาพตอนเบิก:** {r['condition_out']}")
                        st.markdown(f"**สภาพที่แจ้งคืน:** {r.get('condition_in') or '-'}")
                        if r.get("note"):
                            st.caption(f"หมายเหตุ: {r['note']}")

                    if is_admin():
                        st.markdown("**Admin ตรวจสอบจริง:**")
                        admin_condition = st.selectbox("สภาพอุปกรณ์จริงที่ตรวจสอบ",
                                                       ["ปกติ", "มีรอยขีดข่วน", "ชำรุด", "สูญหาย"],
                                                       key=f"ac_{r['id']}")
                        admin_note = st.text_input("หมายเหตุ Admin", key=f"an_{r['id']}")

                        col_ok, col_rej = st.columns(2)
                        with col_ok:
                            if st.button("✅ ยืนยันรับคืน", key=f"confirm_{r['id']}",
                                         type="primary", use_container_width=True):
                                note_final = f"[Admin: {admin_note}]" if admin_note else r.get("note")
                                update_rows("transactions", {
                                    "condition_in": admin_condition, "note": note_final,
                                    "status": "คืนแล้ว"
                                }, "id", r["id"])

                                cur_eq = query_table("equipment", select="available_qty",
                                                     filters=[("id","eq",r["equipment_id"])])
                                if not cur_eq.empty:
                                    cur_avail = int(cur_eq.iloc[0]["available_qty"])
                                    qty_returned = int(r["qty"])

                                    # Fix 8: logic ชัดเจน — สูญหาย = ไม่ได้คืนจริง
                                    if admin_condition == "สูญหาย":
                                        available_delta = 0  # ไม่เพิ่ม available
                                        new_status = "สูญหาย"
                                    elif admin_condition == "ชำรุด":
                                        available_delta = qty_returned  # คืนแต่ชำรุด
                                        new_status = "ชำรุด"
                                    else:
                                        available_delta = qty_returned  # คืนปกติ
                                        new_status = "พร้อมใช้"

                                    update_rows("equipment", {
                                        "available_qty": cur_avail + available_delta,
                                        "status": new_status
                                    }, "id", r["equipment_id"])

                                load_sidebar_stats.clear()
                                st.success(f"✅ ยืนยันรับคืนแล้ว สภาพ: {admin_condition}")
                                st.rerun()
                        with col_rej:
                            if st.button("↩️ ส่งกลับ (ยังไม่คืน)", key=f"reject_{r['id']}",
                                         use_container_width=True):
                                update_rows("transactions", {
                                    "status": "ยืมอยู่", "return_date": None, "condition_in": None
                                }, "id", r["id"])
                                load_sidebar_stats.clear()
                                st.warning("↩️ ส่งกลับเป็นสถานะ ยืมอยู่ แล้ว")
                                st.rerun()
                    else:
                        st.info("🔒 กรุณา Login Admin เพื่อยืนยันรับคืน")

# ═════════════════════════════════════════════════════════════════════════════
# PAGE: REPORT                                                  [FIX #1]
# ═════════════════════════════════════════════════════════════════════════════
def page_report():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">📋 รายงาน</p>', unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["ประวัติการเบิก-คืน", "สรุปอุปกรณ์"])

    with tab1:
        date_from = st.date_input("ตั้งแต่", value=date(date.today().year, 1, 1))
        date_to = st.date_input("ถึงวันที่", value=date.today())
        status_filter = st.selectbox("สถานะ", ["ทั้งหมด", "ยืมอยู่", "คืนแล้ว"])

        # [FIX #1] batch fetch + merge
        df_report = load_report_data(date_from, date_to, status_filter)

        st.caption(f"พบ {len(df_report)} รายการ")
        st.dataframe(df_report, use_container_width=True, hide_index=True)

        if not df_report.empty:
            if is_admin():
                st.download_button("📥 Export Excel", data=export_excel(df_report, "ประวัติการเบิก-คืน"),
                                   file_name=f"borrow_history_{date.today()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            else:
                st.info("🔒 Export Excel สำหรับ Admin เท่านั้น")

    with tab2:
        # [FIX #1] batch fetch + groupby
        df2 = load_equipment_summary()

        if not df2.empty:
            st.dataframe(df2, use_container_width=True, hide_index=True)
            if is_admin():
                st.download_button("📥 Export Excel สรุปอุปกรณ์",
                                   data=export_excel(df2, "สรุปอุปกรณ์"),
                                   file_name=f"equipment_summary_{date.today()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        else:
            st.info("ยังไม่มีข้อมูลอุปกรณ์")

# ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
def export_excel(df, sheet_name):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    hfill  = PatternFill("solid", fgColor="1F4E79")
    hfont  = Font(color="FFFFFF", bold=True, size=11)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
    for ci, col in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill, cell.font, cell.border = hfill, hfont, border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for ri, row in enumerate(df.itertuples(index=False), 2):
        fill = PatternFill("solid", fgColor="EBF3FB" if ri % 2 == 0 else "FFFFFF")
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.border, cell.fill = border, fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(w + 4, 30)
    ws.row_dimensions[1].height = 25
    wb.save(output)
    return output.getvalue()

# ═════════════════════════════════════════════════════════════════════════════
# PAGE: SETTINGS                                                [FIX #4, #7]
# ═════════════════════════════════════════════════════════════════════════════
def page_settings():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">⚙️ ตั้งค่าระบบ</p>', unsafe_allow_html=True)

    if not is_admin():
        st.warning("🔒 หน้านี้สำหรับ Admin เท่านั้น กรุณา Login ที่ Sidebar")
        return

    tab1, tab2, tab3 = st.tabs(["💾 สำรองข้อมูล", "📂 นำเข้าข้อมูล", "🗑️ ล้างข้อมูล"])

    with tab1:
        st.markdown('<div class="section-header">💾 Export — สำรองข้อมูลเป็น JSON</div>', unsafe_allow_html=True)
        st.info("Export ข้อมูลทั้งหมดเป็นไฟล์ JSON สำหรับสำรองหรือย้ายระบบ")

        if st.button("📦 สร้างไฟล์ Backup JSON", type="primary", use_container_width=True):
            eq = query_table("equipment").to_dict(orient="records")
            tx = query_table("transactions").to_dict(orient="records")
            borr = query_table("borrowers").to_dict(orient="records")
            backup = {
                "backup_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "version": "2.1-optimized",
                "equipment": eq, "transactions": tx, "borrowers": borr
            }
            json_str = json.dumps(backup, ensure_ascii=False, indent=2, default=str)
            fname = f"lab_backup_{date.today()}.json"
            st.download_button(
                label=f"⬇️ ดาวน์โหลด {fname}",
                data=json_str.encode("utf-8"),
                file_name=fname, mime="application/json",
                use_container_width=True
            )
            st.success(f"✅ สร้าง Backup สำเร็จ — อุปกรณ์ {len(eq)} | ประวัติ {len(tx)}")

    with tab2:
        st.markdown('<div class="section-header">📂 Import — นำเข้าข้อมูลจาก JSON</div>', unsafe_allow_html=True)
        st.warning("⚠️ การนำเข้าจะ **เพิ่ม** ข้อมูลเข้าระบบ ไม่ได้ลบข้อมูลเดิม")

        uploaded_json = st.file_uploader("เลือกไฟล์ JSON", type=["json"], key="import_json")
        if uploaded_json:
            try:
                data = json.loads(uploaded_json.read().decode("utf-8"))
                eq_count = len(data.get("equipment", []))
                tx_count = len(data.get("transactions", []))
                borr_count = len(data.get("borrowers", []))
                backup_date = data.get("backup_date", "ไม่ทราบ")

                st.info(f"📋 Backup วันที่: {backup_date} | อุปกรณ์: {eq_count} | ประวัติ: {tx_count} | ผู้เบิก: {borr_count}")

                import_mode = st.radio("โหมดนำเข้า",
                    ["เฉพาะอุปกรณ์ (equipment)", "ทั้งหมด (equipment + transactions + borrowers)"])

                if st.button("📂 ยืนยันนำเข้าข้อมูล", type="primary", use_container_width=True):
                    imported = 0
                    for eq in data.get("equipment", []):
                        try:
                            existing = query_table("equipment", select="id",
                                                   filters=[("code","eq",eq["code"])])
                            if existing.empty:
                                img_url = eq.get("image_url") or eq.get("image_path")
                                if img_url and not img_url.startswith("http"):
                                    img_url = None
                                insert_row("equipment", {
                                    "code": eq["code"], "name": eq["name"],
                                    "category": eq.get("category"),
                                    "total_qty": eq.get("total_qty", 1),
                                    "available_qty": eq.get("available_qty", 1),
                                    "status": eq.get("status", "พร้อมใช้"),
                                    "image_url": img_url,
                                    "description": eq.get("description")
                                })
                                imported += 1
                        except Exception:
                            pass

                    if "ทั้งหมด" in import_mode:
                        id_map = {}
                        for b in data.get("borrowers", []):
                            old_id = b["id"]
                            result = insert_row("borrowers", {
                                "name": b["name"], "type": b["type"],
                                "student_id": b.get("student_id"),
                                "department": b.get("department"),
                                "phone": b.get("phone")
                            })
                            if result:
                                id_map[old_id] = result["id"]

                        for tx in data.get("transactions", []):
                            try:
                                new_borr_id = id_map.get(tx["borrower_id"])
                                eq_id = tx.get("equipment_id")
                                if new_borr_id and eq_id:
                                    insert_row("transactions", {
                                        "equipment_id": eq_id, "borrower_id": new_borr_id,
                                        "qty": tx.get("qty", 1),
                                        "borrow_date": tx.get("borrow_date"),
                                        "due_date": tx.get("due_date"),
                                        "return_date": tx.get("return_date"),
                                        "condition_out": tx.get("condition_out", "ปกติ"),
                                        "condition_in": tx.get("condition_in"),
                                        "note": tx.get("note"),
                                        "status": tx.get("status", "คืนแล้ว")
                                    })
                            except Exception:
                                pass

                    load_sidebar_stats.clear()
                    st.success("✅ นำเข้าข้อมูลสำเร็จ!")
                    st.rerun()

            except Exception as e:
                st.error(f"❌ ไฟล์ไม่ถูกต้อง: {e}")

    # ── TAB 3: ล้างข้อมูล ─────────────────────────────────────────────────
    with tab3:
        st.markdown('<div class="section-header">🗑️ ล้างข้อมูล</div>', unsafe_allow_html=True)
        st.error("⚠️ การล้างข้อมูลไม่สามารถกู้คืนได้ แนะนำให้ **Export JSON ก่อน** ทุกครั้ง!")

        clear_mode = st.selectbox("เลือกประเภทการล้าง", [
            "เลือก...",
            "🔄 รีเซ็ตจำนวนอุปกรณ์ (available = total)",
            "📋 ล้างประวัติการเบิก-คืนทั้งหมด",
            "💥 ล้างทุกอย่าง (เริ่มระบบใหม่)",
        ])

        if clear_mode != "เลือก...":
            desc = {
                "🔄 รีเซ็ตจำนวนอุปกรณ์ (available = total)": "รีเซ็ต available_qty = total_qty ทุกอุปกรณ์ เปลี่ยนสถานะเป็น พร้อมใช้",
                "📋 ล้างประวัติการเบิก-คืนทั้งหมด":          "ลบ transactions + borrowers ข้อมูลอุปกรณ์ยังอยู่",
                "💥 ล้างทุกอย่าง (เริ่มระบบใหม่)":           "ลบทุกตาราง เริ่มใหม่ทั้งหมด",
            }
            st.info(f"ℹ️ {desc.get(clear_mode,'')}")

            st.markdown("**พิมพ์ CONFIRM เพื่อยืนยัน:**")
            confirm_text = st.text_input("", placeholder="พิมพ์ CONFIRM", key="confirm_clear")

            if st.button("🗑️ ดำเนินการล้างข้อมูล", type="primary", use_container_width=True):
                if confirm_text != "CONFIRM":
                    st.error("❌ กรุณาพิมพ์ CONFIRM ให้ถูกต้อง (ตัวพิมพ์ใหญ่)")
                else:
                    try:
                        if "รีเซ็ตจำนวน" in clear_mode:
                            # [FIX #4] batch reset
                            batch_reset_equipment()
                            st.success("✅ รีเซ็ตจำนวนอุปกรณ์เรียบร้อย")

                        elif "ล้างประวัติ" in clear_mode:
                            # [FIX #7] delete all
                            delete_rows("transactions", delete_all=True)
                            delete_rows("borrowers", delete_all=True)
                            batch_reset_equipment()
                            st.success("✅ ล้างประวัติการเบิก-คืนเรียบร้อย")

                        elif "ล้างทุกอย่าง" in clear_mode:
                            delete_rows("transactions", delete_all=True)
                            delete_rows("borrowers", delete_all=True)
                            delete_rows("equipment", delete_all=True)
                            st.success("✅ ล้างข้อมูลทั้งหมดเรียบร้อย เริ่มระบบใหม่ได้เลย")

                        load_sidebar_stats.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")

# ─── FOOTER ───────────────────────────────────────────────────────────────────
def footer():
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; color:#888; font-size:0.82rem; padding:8px 0 16px 0; line-height:1.8;">
        🔬 ระบบบริหารจัดการเบิก-คืนอุปกรณ์ห้องปฏิบัติการ TTC<br>
        <b style="color:#4CAF50;">☁️ Cloud Edition v2.1</b> — Supabase + Cloudinary<br>
        พัฒนาโดย <b style="color:#1F4E79;">รศ.ดร.อิทธิพล มีผล</b><br>
        ภาควิชาครุศาสตร์โยธา &nbsp;|&nbsp; คณะครุศาสตร์อุตสาหกรรม<br>
        มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (KMUTNB)
    </div>
    """, unsafe_allow_html=True)

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    nav()
    page = st.session_state.get("page", "Dashboard")
    if   page == "Dashboard": page_dashboard()
    elif page == "อุปกรณ์":   page_equipment()
    elif page == "เบิก":      page_borrow()
    elif page == "คืน":       page_return()
    elif page == "รายงาน":    page_report()
    elif page == "ตั้งค่า":   page_settings()
    footer()

if __name__ == "__main__":
    main()
