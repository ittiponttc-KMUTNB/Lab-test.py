import streamlit as st
import sqlite3
import pandas as pd
import os
import base64
from datetime import datetime, date
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ระบบเบิก-คืนอุปกรณ์",
    page_icon="🔬",
    layout="centered",
    initial_sidebar_state="collapsed"
)

DB_PATH = "lab_equipment.db"
IMG_DIR = "equipment_images"
ADMIN_PASSWORD = "admin1234"   # ← เปลี่ยน password ได้ที่นี่
os.makedirs(IMG_DIR, exist_ok=True)

# ─── MOBILE CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-size: 16px; }

div.stButton > button {
    height: 3rem;
    font-size: 1rem;
    border-radius: 10px;
}

input, textarea { font-size: 16px !important; min-height: 2.8rem !important; }
div[data-baseweb="select"] { font-size: 16px; }

div[data-testid="metric-container"] {
    background: #f8f9fa;
    border-radius: 12px;
    padding: 12px;
    border: 1px solid #e0e0e0;
}

.eq-card {
    background: white;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 10px;
    border: 1px solid #e0e0e0;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    line-height: 1.7;
}

.overdue-alert {
    background: #fff3cd;
    border-left: 4px solid #ff6b6b;
    border-radius: 8px;
    padding: 10px 14px;
    margin: 6px 0;
    line-height: 1.7;
}

.section-header {
    font-size: 1.05rem;
    font-weight: 700;
    color: #1F4E79;
    margin: 16px 0 8px 0;
    padding-bottom: 4px;
    border-bottom: 2px solid #e0e0e0;
}

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.78rem;
    font-weight: 600;
    color: white;
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

# ─── DATABASE ─────────────────────────────────────────────────────────────────
def get_conn():
    return sqlite3.connect(DB_PATH)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS equipment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            category TEXT,
            total_qty INTEGER DEFAULT 1,
            available_qty INTEGER DEFAULT 1,
            status TEXT DEFAULT 'พร้อมใช้',
            image_path TEXT,
            description TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS borrowers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            type TEXT NOT NULL,
            student_id TEXT,
            department TEXT,
            phone TEXT
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_id INTEGER,
            borrower_id INTEGER,
            qty INTEGER DEFAULT 1,
            borrow_date TEXT,
            due_date TEXT,
            return_date TEXT,
            condition_out TEXT DEFAULT 'ปกติ',
            condition_in TEXT,
            note TEXT,
            status TEXT DEFAULT 'ยืมอยู่',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            FOREIGN KEY(equipment_id) REFERENCES equipment(id),
            FOREIGN KEY(borrower_id) REFERENCES borrowers(id)
        );
    """)
    conn.commit()
    conn.close()

def query(sql, params=()):
    conn = get_conn()
    df = pd.read_sql_query(sql, conn, params=params)
    conn.close()
    return df

def execute(sql, params=()):
    conn = get_conn()
    c = conn.cursor()
    c.execute(sql, params)
    conn.commit()
    last_id = c.lastrowid
    conn.close()
    return last_id

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def img_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

def show_image(image_path, width=100):
    if image_path and os.path.exists(image_path):
        b64 = img_b64(image_path)
        st.markdown(
            f'<img src="data:image/png;base64,{b64}" '
            f'style="width:{width}px;max-width:100%;border-radius:8px;border:1px solid #ddd;">',
            unsafe_allow_html=True)
    else:
        st.markdown(
            f'<div style="width:{width}px;height:{width}px;background:#f0f2f6;'
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
        d = datetime.strptime(due_str, "%Y-%m-%d").date()
        delta = (date.today() - d).days
        return max(delta, 0)
    except:
        return 0

# ─── NAVIGATION ───────────────────────────────────────────────────────────────
PAGES = [("🏠","Dashboard"), ("📦","อุปกรณ์"), ("➕","เบิก"), ("✅","คืน"), ("📋","รายงาน")]

def nav():
    if "page" not in st.session_state:
        st.session_state.page = "Dashboard"

    with st.sidebar:
        st.markdown("## 🔬 ระบบอุปกรณ์ Lab")
        st.markdown("---")
        for icon, name in PAGES:
            active = st.session_state.page == name
            if st.button(f"{icon} {name}", use_container_width=True,
                         type="primary" if active else "secondary", key=f"snav_{name}"):
                st.session_state.page = name
                st.rerun()
        st.markdown("---")
        admin_login_widget()
        st.markdown("---")
        n_eq   = query("SELECT COUNT(*) as n FROM equipment").iloc[0]["n"]
        n_borr = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่'").iloc[0]["n"]
        n_over = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่' AND due_date < date('now')").iloc[0]["n"]
        st.metric("📦 อุปกรณ์", n_eq)
        st.metric("🔄 กำลังยืม", n_borr)
        if n_over > 0:
            st.error(f"⚠️ เกินกำหนด {n_over} รายการ")

    # ── ชื่อโปรแกรม ─────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="text-align:center; margin-bottom:10px;">
        <span style="font-size:1.4rem; font-weight:700; color:#1F4E79;">
            🔬 ระบบบริหารจัดการห้องปฏิบัติการ TTC
        </span>
    </div>
    """, unsafe_allow_html=True)

    # ── Top nav bar — icon + ชื่อ ────────────────────────────────────────────
    NAV_ITEMS = [
        ("Dashboard", "🏠", "หน้าหลัก"),
        ("อุปกรณ์",   "📦", "อุปกรณ์"),
        ("เบิก",      "➕", "เบิก"),
        ("คืน",       "✅", "คืน"),
        ("รายงาน",    "📋", "รายงาน"),
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

# ─── PAGE: DASHBOARD ──────────────────────────────────────────────────────────
def page_dashboard():
    st.markdown("## 🏠 ภาพรวม")

    total_eq  = query("SELECT COUNT(*) as n FROM equipment").iloc[0]["n"]
    available = query("SELECT COALESCE(SUM(available_qty),0) as n FROM equipment").iloc[0]["n"]
    active_tx = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่'").iloc[0]["n"]
    overdue   = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่' AND due_date < date('now')").iloc[0]["n"]

    c1, c2 = st.columns(2)
    c1.metric("📦 อุปกรณ์ทั้งหมด", total_eq)
    c2.metric("✅ พร้อมใช้", int(available))
    c1.metric("🔄 กำลังยืม", active_tx)
    c2.metric("⚠️ เกินกำหนด", overdue)

    if overdue > 0:
        st.error(f"⚠️ มีอุปกรณ์เกินกำหนดคืน {overdue} รายการ!")
        df_od = query("""
            SELECT e.code, e.name, b.name as borrower, b.phone, t.due_date,
                   CAST(julianday('now')-julianday(t.due_date) AS INTEGER) as days_over
            FROM transactions t
            JOIN equipment e ON t.equipment_id=e.id
            JOIN borrowers b ON t.borrower_id=b.id
            WHERE t.status='ยืมอยู่' AND t.due_date < date('now')
            ORDER BY days_over DESC
        """)
        for _, r in df_od.iterrows():
            st.markdown(f"""
            <div class="overdue-alert">
                🔴 <b>{r['code']} — {r['name']}</b><br>
                👤 {r['borrower']} {"📞 "+r['phone'] if r['phone'] else ""}<br>
                📅 กำหนดคืน {r['due_date']} &nbsp;
                <b style="color:red;">เกิน {r['days_over']} วัน</b>
            </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-header">📋 รายการที่กำลังยืมอยู่</div>', unsafe_allow_html=True)
    df = query("""
        SELECT e.code, e.name, e.image_path, b.name as borrower,
               b.type, t.qty, t.borrow_date, t.due_date
        FROM transactions t
        JOIN equipment e ON t.equipment_id=e.id
        JOIN borrowers b ON t.borrower_id=b.id
        WHERE t.status='ยืมอยู่'
        ORDER BY t.due_date ASC
    """)
    if df.empty:
        st.info("ไม่มีรายการยืมในขณะนี้ ✅")
    else:
        for _, r in df.iterrows():
            od = overdue_days(r["due_date"])
            bc = "#ff6b6b" if od > 0 else "#28a745"
            st.markdown(f"""
            <div class="eq-card" style="border-left:4px solid {bc};">
                <b>{r['code']}</b> — {r['name']}<br>
                👤 {r['borrower']} ({r['type']}) &nbsp; 📦 {r['qty']} ชิ้น<br>
                📅 เบิก {r['borrow_date']} | คืน <b>{r['due_date']}</b>
                {"&nbsp;<b style='color:red;'>⚠️ เกิน "+str(od)+" วัน</b>" if od > 0 else ""}
            </div>""", unsafe_allow_html=True)

# ─── PAGE: EQUIPMENT ──────────────────────────────────────────────────────────
def page_equipment():
    st.markdown("## 📦 รายการอุปกรณ์")

    search = st.text_input("🔍 ค้นหา ชื่อ / รหัส / หมวดหมู่", placeholder="พิมพ์เพื่อค้นหา...")
    cats   = query("SELECT DISTINCT category FROM equipment WHERE category IS NOT NULL ORDER BY category")
    cat_filter = st.selectbox("📂 หมวดหมู่", ["ทั้งหมด"] + cats["category"].tolist())

    sql = """SELECT e.id, e.code, e.name, e.category, e.total_qty, e.available_qty,
                    e.status, e.image_path, e.description
             FROM equipment e WHERE 1=1"""
    params = []
    if search:
        sql += " AND (e.name LIKE ? OR e.code LIKE ? OR e.category LIKE ?)"
        params += [f"%{search}%"] * 3
    if cat_filter != "ทั้งหมด":
        sql += " AND e.category=?"
        params.append(cat_filter)
    sql += " ORDER BY e.code"

    df = query(sql, params)
    st.caption(f"พบ {len(df)} รายการ")

    for _, r in df.iterrows():
        with st.expander(f"{r['code']} — {r['name']}"):
            c1, c2 = st.columns([1, 2])
            with c1:
                show_image(r["image_path"], width=90)
            with c2:
                st.markdown(
                    f"**หมวด:** {r['category'] or '-'}<br>"
                    f"**ทั้งหมด:** {r['total_qty']} | **พร้อมใช้:** {r['available_qty']}<br>"
                    f"**สถานะ:** {badge(r['status'], STATUS_COLOR.get(r['status'],'#888'))}",
                    unsafe_allow_html=True)
                if r["description"]:
                    st.caption(r["description"])
            lt = query("""SELECT b.name, t.borrow_date FROM transactions t
                          JOIN borrowers b ON t.borrower_id=b.id
                          WHERE t.equipment_id=? ORDER BY t.created_at DESC LIMIT 1""", (r["id"],))
            if not lt.empty:
                st.caption(f"👤 ผู้ยืมล่าสุด: **{lt.iloc[0]['name']}** ({lt.iloc[0]['borrow_date']})")
            if is_admin():
                if st.button("🗑️ ลบ", key=f"del_{r['id']}"):
                    active = query("SELECT COUNT(*) as n FROM transactions WHERE equipment_id=? AND status='ยืมอยู่'",
                                   (r["id"],)).iloc[0]["n"]
                    if active > 0:
                        st.error("ไม่สามารถลบได้ มีการยืมอยู่")
                    else:
                        execute("DELETE FROM equipment WHERE id=?", (r["id"],))
                        st.success("ลบแล้ว")
                        st.rerun()

    if is_admin():
        st.markdown('<div class="section-header">➕ เพิ่ม / แก้ไขอุปกรณ์</div>', unsafe_allow_html=True)
        eq_list = query("SELECT id, code, name FROM equipment ORDER BY code")
        options = ["➕ เพิ่มใหม่"] + [f"{r['code']} — {r['name']}" for _, r in eq_list.iterrows()]
        choice  = st.selectbox("เลือกรายการ", options, key="eq_edit_sel")

        existing, eq_id = None, None
        if choice != "➕ เพิ่มใหม่":
            eq_id    = eq_list.iloc[options.index(choice) - 1]["id"]
            existing = query("SELECT * FROM equipment WHERE id=?", (eq_id,)).iloc[0]

        with st.form("eq_form", clear_on_submit=True):
            code        = st.text_input("รหัสอุปกรณ์ *", value=existing["code"]        if existing is not None else "")
            name        = st.text_input("ชื่ออุปกรณ์ *",  value=existing["name"]        if existing is not None else "")
            category    = st.text_input("หมวดหมู่",        value=existing["category"]    if existing is not None else "")
            total_qty   = st.number_input("จำนวนทั้งหมด", min_value=1,
                                          value=int(existing["total_qty"]) if existing is not None else 1)
            status_sel  = st.selectbox("สถานะ", ["พร้อมใช้","ชำรุด","สูญหาย"],
                                       index=["พร้อมใช้","ชำรุด","สูญหาย"].index(existing["status"])
                                       if existing is not None else 0)
            description = st.text_area("รายละเอียด", value=existing["description"] if existing is not None else "")
            uploaded    = st.file_uploader("📷 รูปอุปกรณ์ (optional)", type=["jpg","jpeg","png"])

            if existing is not None and existing["image_path"] and os.path.exists(existing["image_path"]):
                st.caption("รูปปัจจุบัน:")
                show_image(existing["image_path"], width=80)

            if st.form_submit_button("💾 บันทึก", type="primary", use_container_width=True):
                if not code or not name:
                    st.error("กรุณากรอกรหัสและชื่ออุปกรณ์")
                else:
                    dup = query("SELECT id FROM equipment WHERE code=?", (code,))
                    cur_id = int(eq_id) if existing is not None else -1
                    if not dup.empty and (existing is None or int(dup.iloc[0]["id"]) != cur_id):
                        st.error(f"❌ รหัส '{code}' มีอยู่แล้ว กรุณาใช้รหัสอื่น")
                    else:
                        try:
                            img_path = existing["image_path"] if existing is not None else None
                            if uploaded:
                                ext = uploaded.name.split(".")[-1]
                                img_path = os.path.join(IMG_DIR, f"{code}.{ext}")
                                with open(img_path, "wb") as fp:
                                    fp.write(uploaded.getbuffer())
                            if existing is None:
                                execute("""INSERT INTO equipment
                                    (code,name,category,total_qty,available_qty,status,image_path,description)
                                    VALUES (?,?,?,?,?,?,?,?)""",
                                    (code, name, category, total_qty, total_qty, status_sel, img_path, description))
                                st.success("✅ เพิ่มอุปกรณ์เรียบร้อย")
                            else:
                                execute("""UPDATE equipment SET code=?,name=?,category=?,total_qty=?,
                                    status=?,image_path=?,description=? WHERE id=?""",
                                    (code, name, category, total_qty, status_sel, img_path, description, eq_id))
                                st.success("✅ อัปเดตเรียบร้อย")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
    else:
        st.info("🔒 การเพิ่ม/แก้ไขอุปกรณ์ สำหรับ Admin เท่านั้น")

# ─── PAGE: BORROW ─────────────────────────────────────────────────────────────
def page_borrow():
    st.markdown("## ➕ เบิกอุปกรณ์")

    avail = query("""SELECT id, code, name, category, available_qty, image_path, description
                     FROM equipment WHERE available_qty > 0 AND status='พร้อมใช้'
                     ORDER BY category, code""")
    if avail.empty:
        st.warning("⚠️ ไม่มีอุปกรณ์พร้อมใช้งานในขณะนี้")
        return

    # ── เลือกอุปกรณ์ ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📦 ขั้นตอนที่ 1 — เลือกอุปกรณ์</div>', unsafe_allow_html=True)

    cats_avail = ["ทั้งหมด"] + sorted(avail["category"].dropna().unique().tolist())
    cat_sel    = st.selectbox("📂 กรองหมวดหมู่", cats_avail, key="bcat")
    search_eq  = st.text_input("🔍 ค้นหาชื่อ / รหัสอุปกรณ์",
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
    eq_id  = eq_opts[selected_label]
    eq_row = avail[avail["id"] == eq_id].iloc[0]

    # Preview อุปกรณ์ที่เลือก
    c_img, c_info = st.columns([1, 2])
    with c_img:
        show_image(eq_row["image_path"], width=100)
    with c_info:
        st.markdown(f"**{eq_row['code']}** — {eq_row['name']}")
        st.markdown(f"หมวด: {eq_row['category'] or '-'}")
        st.markdown(f"คงเหลือ: **{eq_row['available_qty']}** ชิ้น")
        if eq_row["description"]:
            st.caption(eq_row["description"])

    qty = st.number_input("จำนวนที่ต้องการเบิก *", min_value=1,
                          max_value=int(eq_row["available_qty"]), value=1)
    st.divider()

    # ── ข้อมูลผู้เบิก ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">👤 ขั้นตอนที่ 2 — ข้อมูลผู้เบิก</div>', unsafe_allow_html=True)

    borrower_type = st.radio("ประเภทผู้เบิก", ["นักศึกษา", "บุคลากร/อาจารย์"], horizontal=True)
    borrower_name = st.text_input("ชื่อ-นามสกุล *", placeholder="กรอกชื่อ-นามสกุล")
    student_id    = st.text_input("รหัสนักศึกษา / รหัสพนักงาน", placeholder="เช่น 6601234567")
    department    = st.text_input("ภาควิชา / หน่วยงาน")
    phone         = st.text_input("เบอร์โทรศัพท์", placeholder="เช่น 081-234-5678")
    st.divider()

    # ── วันที่ ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📅 ขั้นตอนที่ 3 — วันที่</div>', unsafe_allow_html=True)

    borrow_date   = st.date_input("วันที่เบิก *", value=date.today())
    due_date      = st.date_input("วันกำหนดคืน *", value=date.today())
    condition_out = st.selectbox("สภาพอุปกรณ์ขณะเบิก", ["ปกติ", "มีรอยขีดข่วน", "ชำรุดบางส่วน"])
    note          = st.text_area("หมายเหตุ (ถ้ามี)")
    st.divider()

    if st.button("✅ ยืนยันการเบิก", type="primary", use_container_width=True):
        if not borrower_name.strip():
            st.error("❌ กรุณากรอกชื่อผู้เบิก")
        elif due_date < borrow_date:
            st.error("❌ วันกำหนดคืนต้องไม่ก่อนวันที่เบิก")
        else:
            b_id = execute("""INSERT INTO borrowers (name,type,student_id,department,phone)
                               VALUES (?,?,?,?,?)""",
                           (borrower_name.strip(), borrower_type, student_id, department, phone))
            execute("""INSERT INTO transactions
                       (equipment_id,borrower_id,qty,borrow_date,due_date,condition_out,note)
                       VALUES (?,?,?,?,?,?,?)""",
                    (eq_id, b_id, qty, str(borrow_date), str(due_date), condition_out, note))
            execute("UPDATE equipment SET available_qty=available_qty-? WHERE id=?", (qty, eq_id))
            st.success(
                f"✅ บันทึกสำเร็จ!\n\n"
                f"👤 **{borrower_name}**\n"
                f"📦 {eq_row['name']} จำนวน {qty} ชิ้น\n"
                f"📅 กำหนดคืน: {due_date}"
            )
            st.balloons()

# ─── PAGE: RETURN ─────────────────────────────────────────────────────────────
def page_return():
    st.markdown("## ✅ คืนอุปกรณ์")

    search = st.text_input("🔍 ค้นหาชื่อผู้ยืม / รหัส / ชื่ออุปกรณ์",
                            placeholder="พิมพ์เพื่อค้นหา...")

    df = query("""
        SELECT t.id, e.code, e.name, e.image_path, b.name as borrower,
               b.type, t.qty, t.borrow_date, t.due_date, t.condition_out,
               t.note as borrow_note, t.equipment_id
        FROM transactions t
        JOIN equipment e ON t.equipment_id=e.id
        JOIN borrowers b ON t.borrower_id=b.id
        WHERE t.status='ยืมอยู่'
        ORDER BY t.due_date ASC
    """)
    if search:
        mask = (df["borrower"].str.contains(search, case=False, na=False) |
                df["code"].str.contains(search, case=False, na=False) |
                df["name"].str.contains(search, case=False, na=False))
        df = df[mask]

    if df.empty:
        st.info("✅ ไม่มีรายการยืม" if not search else "ไม่พบรายการที่ค้นหา")
        return

    st.caption(f"พบ {len(df)} รายการ")

    for _, r in df.iterrows():
        od     = overdue_days(r["due_date"])
        border = "#ff6b6b" if od > 0 else "#1F4E79"
        label  = f"TX#{r['id']} | {r['code']} {r['name']} | {r['borrower']}"
        if od > 0:
            label += f" ⚠️ เกิน {od} วัน"

        with st.expander(label):
            c1, c2 = st.columns([1, 2])
            with c1:
                show_image(r["image_path"], width=80)
            with c2:
                st.markdown(
                    f"📦 **{r['code']}** — {r['name']} ({r['qty']} ชิ้น)<br>"
                    f"👤 {r['borrower']} ({r['type']})<br>"
                    f"📅 เบิก {r['borrow_date']}<br>"
                    f"📅 คืน <b>{r['due_date']}</b>"
                    + (f"<br><b style='color:red;'>⚠️ เกิน {od} วัน</b>" if od > 0 else ""),
                    unsafe_allow_html=True)
                st.caption(f"สภาพตอนเบิก: {r['condition_out']}")

            return_date  = st.date_input("วันที่คืน", value=date.today(), key=f"rd_{r['id']}")
            condition_in = st.selectbox("สภาพอุปกรณ์เมื่อคืน",
                                        ["ปกติ", "มีรอยขีดข่วน", "ชำรุด", "สูญหาย"],
                                        key=f"ci_{r['id']}")
            return_note  = st.text_input("หมายเหตุ", key=f"rn_{r['id']}")

            if st.button("✅ บันทึกการคืน", key=f"ret_{r['id']}", type="primary",
                         use_container_width=True):
                execute("""UPDATE transactions SET return_date=?,condition_in=?,note=?,status='คืนแล้ว'
                           WHERE id=?""", (str(return_date), condition_in, return_note, r["id"]))
                execute("UPDATE equipment SET available_qty=available_qty+? WHERE id=?",
                        (r["qty"], r["equipment_id"]))
                if condition_in == "ชำรุด":
                    execute("UPDATE equipment SET status='ชำรุด' WHERE id=?", (r["equipment_id"],))
                elif condition_in == "สูญหาย":
                    execute("UPDATE equipment SET status='สูญหาย',available_qty=available_qty-? WHERE id=?",
                            (r["qty"], r["equipment_id"]))
                st.success(f"✅ บันทึกคืนเรียบร้อย สภาพ: {condition_in}")
                st.rerun()

# ─── PAGE: REPORT ─────────────────────────────────────────────────────────────
def page_report():
    st.markdown("## 📋 รายงาน")

    tab1, tab2 = st.tabs(["ประวัติการเบิก-คืน", "สรุปอุปกรณ์"])

    with tab1:
        date_from     = st.date_input("ตั้งแต่", value=date(date.today().year, 1, 1))
        date_to       = st.date_input("ถึงวันที่", value=date.today())
        status_filter = st.selectbox("สถานะ", ["ทั้งหมด", "ยืมอยู่", "คืนแล้ว"])

        sql = """
            SELECT t.id as 'TX#', e.code as 'รหัส', e.name as 'อุปกรณ์',
                   b.name as 'ผู้เบิก', b.type as 'ประเภท', b.student_id as 'รหัส/ID',
                   b.department as 'ภาควิชา', b.phone as 'โทรศัพท์',
                   t.qty as 'จำนวน', t.borrow_date as 'วันเบิก',
                   t.due_date as 'กำหนดคืน', t.return_date as 'วันคืน',
                   t.condition_out as 'สภาพตอนเบิก', t.condition_in as 'สภาพตอนคืน',
                   t.status as 'สถานะ', t.note as 'หมายเหตุ'
            FROM transactions t
            JOIN equipment e ON t.equipment_id=e.id
            JOIN borrowers b ON t.borrower_id=b.id
            WHERE t.borrow_date BETWEEN ? AND ?
        """
        params = [str(date_from), str(date_to)]
        if status_filter != "ทั้งหมด":
            sql += " AND t.status=?"
            params.append(status_filter)
        sql += " ORDER BY t.id DESC"

        df = query(sql, params)
        st.caption(f"พบ {len(df)} รายการ")
        st.dataframe(df, use_container_width=True, hide_index=True)

        if not df.empty:
            if is_admin():
                st.download_button("📥 Export Excel", data=export_excel(df, "ประวัติการเบิก-คืน"),
                                   file_name=f"borrow_history_{date.today()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            else:
                st.info("🔒 Export Excel สำหรับ Admin เท่านั้น")

    with tab2:
        df2 = query("""
            SELECT e.code as 'รหัส', e.name as 'ชื่ออุปกรณ์', e.category as 'หมวดหมู่',
                   e.total_qty as 'ทั้งหมด', e.available_qty as 'พร้อมใช้',
                   (e.total_qty-e.available_qty) as 'กำลังยืม', e.status as 'สถานะ',
                   (SELECT COUNT(*) FROM transactions t WHERE t.equipment_id=e.id) as 'ครั้งที่เบิก'
            FROM equipment e ORDER BY e.code
        """)
        st.dataframe(df2, use_container_width=True, hide_index=True)
        if not df2.empty and is_admin():
            st.download_button("📥 Export Excel สรุปอุปกรณ์",
                               data=export_excel(df2, "สรุปอุปกรณ์"),
                               file_name=f"equipment_summary_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

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

# ─── FOOTER ───────────────────────────────────────────────────────────────────
def footer():
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; color:#888; font-size:0.82rem; padding:8px 0 16px 0; line-height:1.8;">
        🔬 ระบบบริหารจัดการเบิก-คืนอุปกรณ์ห้องปฏิบัติการ<br>
        พัฒนาโดย <b style="color:#1F4E79;">รศ.ดร.อิทธิพล มีผล</b><br>
        ภาควิชาครุศาสตร์วิศวกรรม (วิศวกรรมโยธา) &nbsp;|&nbsp;
        คณะครุศาสตร์อุตสาหกรรม<br>
        มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (KMUTNB)
    </div>
    """, unsafe_allow_html=True)

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    init_db()
    nav()
    page = st.session_state.get("page", "Dashboard")
    if   page == "Dashboard": page_dashboard()
    elif page == "อุปกรณ์":   page_equipment()
    elif page == "เบิก":      page_borrow()
    elif page == "คืน":       page_return()
    elif page == "รายงาน":    page_report()
    footer()

if __name__ == "__main__":
    main()
