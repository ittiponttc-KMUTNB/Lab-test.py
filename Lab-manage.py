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
PAGES = [("🏠","Dashboard"), ("📦","อุปกรณ์"), ("➕","เบิก"), ("✅","คืน"), ("📋","รายงาน"), ("⚙️","ตั้งค่า")]

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
        ("ตั้งค่า",   "⚙️", "ตั้งค่า"),
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
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">🏠 ภาพรวม</p>', unsafe_allow_html=True)

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
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">📦 รายการอุปกรณ์</p>', unsafe_allow_html=True)

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
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("✏️ แก้ไข", key=f"edit_{r['id']}", use_container_width=True):
                        # set เป็น string value ของ option (ไม่ใช่ integer index)
                        eq_list_tmp = query("SELECT code, name FROM equipment ORDER BY code")
                        opts_tmp = ["➕ เพิ่มใหม่"] + [f"{x['code']} — {x['name']}" for _, x in eq_list_tmp.iterrows()]
                        matched_tmp = [o for o in opts_tmp if o.startswith(r["code"] + " —")]
                        if matched_tmp:
                            st.session_state["eq_edit_sel"] = matched_tmp[0]
                        st.rerun()
                with btn_col2:
                    if st.button("🗑️ ลบ", key=f"del_{r['id']}", use_container_width=True):
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

        # eq_edit_sel ถูก set จากปุ่ม ✏️ แก้ไข ก่อน rerun แล้ว
        if "eq_edit_sel" not in st.session_state:
            st.session_state["eq_edit_sel"] = 0

        choice = st.selectbox("เลือกรายการ", options, key="eq_edit_sel")

        existing, eq_id = None, None
        if choice != "➕ เพิ่มใหม่":
            # ดึง code จากชื่อที่เลือก แล้วค้นหาจาก DB โดยตรง ไม่ใช้ index
            try:
                selected_code = choice.split(" — ")[0].strip()
                eq_data = query("SELECT * FROM equipment WHERE code=?", (selected_code,))
                if not eq_data.empty:
                    existing = eq_data.iloc[0]
                    eq_id    = existing["id"]
                else:
                    st.warning("⚠️ ไม่พบอุปกรณ์นี้ในระบบ กรุณาเลือกรายการใหม่")
                    existing = None
            except Exception:
                st.warning("⚠️ เกิดข้อผิดพลาด กรุณาเลือกรายการใหม่อีกครั้ง")
                existing = None

        # ── แสดงข้อมูลปัจจุบัน ─────────────────────────────────────────────────
        if existing is not None:
            borrowed_now = int(existing["total_qty"]) - int(existing["available_qty"])
            st.markdown(
                f'<div class="eq-card" style="border-left:4px solid #1F4E79;">'
                f'📊 <b>ข้อมูลปัจจุบัน:</b> &nbsp;'
                f'ทั้งหมด <b>{existing["total_qty"]}</b> | '
                f'พร้อมใช้ <b>{existing["available_qty"]}</b> | '
                f'ยืมออก <b>{borrowed_now}</b>'
                f'</div>', unsafe_allow_html=True)

        # ── ฟอร์มโดยไม่ใช้ st.form (หลีกเลี่ยงปัญหา cache) ─────────────────
        # ใช้ key ที่ผูกกับ eq_id ทำให้ widget สร้างใหม่เมื่อเปลี่ยน item
        fk = str(eq_id) if eq_id else "new"

        inp_code  = st.text_input("รหัสอุปกรณ์ *",
                        value=existing["code"] if existing is not None else "",
                        key=f"inp_code_{fk}")
        inp_name  = st.text_input("ชื่ออุปกรณ์ *",
                        value=existing["name"] if existing is not None else "",
                        key=f"inp_name_{fk}")
        inp_cat   = st.text_input("หมวดหมู่",
                        value=existing["category"] if existing is not None else "",
                        key=f"inp_cat_{fk}")
        inp_qty   = st.number_input("จำนวนทั้งหมด", min_value=1,
                        value=int(existing["total_qty"]) if existing is not None else 1,
                        key=f"inp_qty_{fk}")
        inp_stat  = st.selectbox("สถานะ", ["พร้อมใช้","ชำรุด","สูญหาย"],
                        index=["พร้อมใช้","ชำรุด","สูญหาย"].index(existing["status"])
                                if existing is not None else 0,
                        key=f"inp_stat_{fk}")
        inp_desc  = st.text_area("รายละเอียด",
                        value=existing["description"] if existing is not None else "",
                        key=f"inp_desc_{fk}")
        inp_img   = st.file_uploader("📷 รูปอุปกรณ์ (optional)",
                        type=["jpg","jpeg","png"], key=f"inp_img_{fk}")

        if existing is not None and existing["image_path"] and os.path.exists(existing["image_path"]):
            st.caption("รูปปัจจุบัน:")
            show_image(existing["image_path"], width=80)

        if st.button("💾 บันทึก", type="primary", use_container_width=True, key=f"btn_save_{fk}"):
            if not inp_code or not inp_name:
                st.error("กรุณากรอกรหัสและชื่ออุปกรณ์")
            else:
                dup = query("SELECT id FROM equipment WHERE code=?", (inp_code,))
                cur_id = int(eq_id) if existing is not None else -1
                if not dup.empty and (existing is None or int(dup.iloc[0]["id"]) != cur_id):
                    st.error(f"❌ รหัส '{inp_code}' มีอยู่แล้ว กรุณาใช้รหัสอื่น")
                else:
                    try:
                        img_path = existing["image_path"] if existing is not None else None
                        if inp_img:
                            ext = inp_img.name.split(".")[-1]
                            img_path = os.path.join(IMG_DIR, f"{inp_code}.{ext}")
                            with open(img_path, "wb") as fp:
                                fp.write(inp_img.getbuffer())
                        if existing is None:
                            execute("""INSERT INTO equipment
                                (code,name,category,total_qty,available_qty,status,image_path,description)
                                VALUES (?,?,?,?,?,?,?,?)""",
                                (inp_code, inp_name, inp_cat, inp_qty, inp_qty,
                                 inp_stat, img_path, inp_desc))
                            st.success("✅ เพิ่มอุปกรณ์เรียบร้อย")
                            st.session_state["eq_edit_sel"] = "➕ เพิ่มใหม่"
                            st.rerun()
                        else:
                            old_total     = int(existing["total_qty"])
                            old_available = int(existing["available_qty"])
                            borrowed      = old_total - old_available
                            diff          = inp_qty - old_total
                            new_available = old_available + diff
                            if new_available < 0:
                                st.error(
                                    f"❌ ลดจำนวนไม่ได้! ยืมออกอยู่ {borrowed} ชิ้น "
                                    f"ลดได้สูงสุดเหลือ {borrowed} ชิ้น"
                                )
                            else:
                                execute("""UPDATE equipment SET code=?,name=?,category=?,
                                    total_qty=?,available_qty=?,status=?,image_path=?,description=?
                                    WHERE id=?""",
                                    (inp_code, inp_name, inp_cat, inp_qty, new_available,
                                     inp_stat, img_path, inp_desc, eq_id))
                                msg = "✅ อัปเดตเรียบร้อย"
                                if diff > 0:   msg += f" (เพิ่ม +{diff} พร้อมใช้: {new_available})"
                                elif diff < 0: msg += f" (ลด {diff} พร้อมใช้: {new_available})"
                                st.success(msg)
                                # refresh ฟอร์มด้วย code ใหม่
                                st.session_state["eq_edit_sel"] = f"{inp_code} — {inp_name}"
                                st.rerun()
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
    else:
        st.info("🔒 การเพิ่ม/แก้ไขอุปกรณ์ สำหรับ Admin เท่านั้น")

# ─── PAGE: BORROW ─────────────────────────────────────────────────────────────
def page_borrow():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">➕ เบิกอุปกรณ์</p>', unsafe_allow_html=True)

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
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">✅ คืนอุปกรณ์</p>', unsafe_allow_html=True)

    # ── Tab: แจ้งคืน (ทุกคน) vs รอตรวจสอบ (Admin ยืนยัน) ──────────────────
    tab1, tab2 = st.tabs(["📬 แจ้งคืนอุปกรณ์", f"🔍 รอตรวจสอบ {'(Admin)' if is_admin() else ''}"])

    # ── TAB 1: ผู้เบิกแจ้งคืน ────────────────────────────────────────────────
    with tab1:
        st.info("👤 ผู้ยืมกรอกข้อมูลและแจ้งคืน จากนั้น Admin จะตรวจสอบและยืนยัน")
        search = st.text_input("🔍 ค้นหาชื่อผู้ยืม / รหัส / ชื่ออุปกรณ์",
                                placeholder="พิมพ์เพื่อค้นหา...", key="ret_search")

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
        else:
            st.caption(f"พบ {len(df)} รายการ")
            for _, r in df.iterrows():
                od    = overdue_days(r["due_date"])
                bc    = "#ff6b6b" if od > 0 else "#1F4E79"
                label = f"TX#{r['id']} | {r['code']} {r['name']} | {r['borrower']}"
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
                            f"📅 เบิก {r['borrow_date']} | คืน <b>{r['due_date']}</b>"
                            + (f"<br><b style='color:red;'>⚠️ เกิน {od} วัน</b>" if od > 0 else ""),
                            unsafe_allow_html=True)
                        st.caption(f"สภาพตอนเบิก: {r['condition_out']}")

                    return_date  = st.date_input("วันที่นำมาคืน", value=date.today(), key=f"rd_{r['id']}")
                    condition_in = st.selectbox("สภาพอุปกรณ์ที่นำมาคืน",
                                                ["ปกติ", "มีรอยขีดข่วน", "ชำรุด", "สูญหาย"],
                                                key=f"ci_{r['id']}")
                    return_note  = st.text_input("หมายเหตุ", key=f"rn_{r['id']}")

                    if st.button("📬 แจ้งคืนอุปกรณ์", key=f"notify_{r['id']}",
                                 type="primary", use_container_width=True):
                        execute("""UPDATE transactions SET return_date=?,condition_in=?,note=?,status='รอตรวจสอบ'
                                   WHERE id=?""", (str(return_date), condition_in, return_note, r["id"]))
                        st.success("📬 แจ้งคืนเรียบร้อยแล้ว! กรุณารอ Admin ตรวจสอบและยืนยัน")
                        st.rerun()

    # ── TAB 2: Admin ตรวจสอบและยืนยัน ───────────────────────────────────────
    with tab2:
        df_wait = query("""
            SELECT t.id, e.code, e.name, e.image_path, b.name as borrower,
                   b.type, b.phone, t.qty, t.borrow_date, t.due_date,
                   t.return_date, t.condition_out, t.condition_in,
                   t.note, t.equipment_id
            FROM transactions t
            JOIN equipment e ON t.equipment_id=e.id
            JOIN borrowers b ON t.borrower_id=b.id
            WHERE t.status='รอตรวจสอบ'
            ORDER BY t.return_date ASC
        """)

        if df_wait.empty:
            st.info("✅ ไม่มีรายการรอตรวจสอบ")
        else:
            st.warning(f"🔍 มี {len(df_wait)} รายการรอ Admin ตรวจสอบ")
            for _, r in df_wait.iterrows():
                od    = overdue_days(r["due_date"])
                label = f"TX#{r['id']} | {r['code']} {r['name']} | {r['borrower']}"

                with st.expander(label):
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        show_image(r["image_path"], width=80)
                    with c2:
                        st.markdown(
                            f"📦 **{r['code']}** — {r['name']} ({r['qty']} ชิ้น)<br>"
                            f"👤 {r['borrower']} ({r['type']}) {'📞 '+r['phone'] if r['phone'] else ''}<br>"
                            f"📅 เบิก {r['borrow_date']} | กำหนดคืน <b>{r['due_date']}</b><br>"
                            f"📅 แจ้งคืนวันที่: <b>{r['return_date']}</b>"
                            + (f"<br><b style='color:red;'>⚠️ เกินกำหนด {od} วัน</b>" if od > 0 else ""),
                            unsafe_allow_html=True)
                        st.markdown(f"**สภาพตอนเบิก:** {r['condition_out']}")
                        st.markdown(f"**สภาพที่แจ้งคืน:** {r['condition_in'] or '-'}")
                        if r["note"]:
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
                                note_final = f"[Admin: {admin_note}]" if admin_note else r["note"]
                                execute("""UPDATE transactions SET condition_in=?,note=?,status='คืนแล้ว'
                                           WHERE id=?""", (admin_condition, note_final, r["id"]))
                                execute("UPDATE equipment SET available_qty=available_qty+? WHERE id=?",
                                        (r["qty"], r["equipment_id"]))
                                if admin_condition == "ชำรุด":
                                    execute("UPDATE equipment SET status='ชำรุด' WHERE id=?", (r["equipment_id"],))
                                elif admin_condition == "สูญหาย":
                                    execute("UPDATE equipment SET status='สูญหาย',available_qty=available_qty-? WHERE id=?",
                                            (r["qty"], r["equipment_id"]))
                                else:
                                    execute("UPDATE equipment SET status='พร้อมใช้' WHERE id=?", (r["equipment_id"],))
                                st.success(f"✅ ยืนยันรับคืนแล้ว สภาพ: {admin_condition}")
                                st.rerun()
                        with col_rej:
                            if st.button("↩️ ส่งกลับ (ยังไม่คืน)", key=f"reject_{r['id']}",
                                         use_container_width=True):
                                execute("UPDATE transactions SET status='ยืมอยู่', return_date=NULL, condition_in=NULL WHERE id=?",
                                        (r["id"],))
                                st.warning("↩️ ส่งกลับเป็นสถานะ ยืมอยู่ แล้ว")
                                st.rerun()
                    else:
                        st.info("🔒 กรุณา Login Admin เพื่อยืนยันรับคืน")

# ─── PAGE: REPORT ─────────────────────────────────────────────────────────────
def page_report():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">📋 รายงาน</p>', unsafe_allow_html=True)

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


# ─── PAGE: SETTINGS ───────────────────────────────────────────────────────────
def page_settings():
    st.markdown('<p style="font-size:1.3rem;font-weight:700;color:#1F4E79;margin:4px 0 12px 0;">⚙️ ตั้งค่าระบบ</p>', unsafe_allow_html=True)

    if not is_admin():
        st.warning("🔒 หน้านี้สำหรับ Admin เท่านั้น กรุณา Login ที่ Sidebar")
        return

    tab1, tab2, tab3 = st.tabs(["💾 สำรองข้อมูล", "📂 นำเข้าข้อมูล", "🗑️ ล้างข้อมูล"])

    # ── TAB 1: Export JSON ────────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-header">💾 Export — สำรองข้อมูลเป็น JSON</div>', unsafe_allow_html=True)
        st.info("Export ข้อมูลทั้งหมด (อุปกรณ์, ประวัติการเบิก-คืน, รายชื่อผู้เบิก) เป็นไฟล์ JSON สำหรับสำรองหรือย้ายระบบ")

        if st.button("📦 สร้างไฟล์ Backup JSON", type="primary", use_container_width=True):
            import json
            eq   = query("SELECT * FROM equipment").to_dict(orient="records")
            tx   = query("SELECT * FROM transactions").to_dict(orient="records")
            borr = query("SELECT * FROM borrowers").to_dict(orient="records")
            backup = {
                "backup_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "version": "1.0",
                "equipment":    eq,
                "transactions": tx,
                "borrowers":    borr
            }
            json_str = json.dumps(backup, ensure_ascii=False, indent=2)
            fname = f"lab_backup_{date.today()}.json"
            st.download_button(
                label=f"⬇️ ดาวน์โหลด {fname}",
                data=json_str.encode("utf-8"),
                file_name=fname,
                mime="application/json",
                use_container_width=True
            )
            st.success(f"✅ สร้าง Backup สำเร็จ — อุปกรณ์ {len(eq)} รายการ, ประวัติ {len(tx)} รายการ")

    # ── TAB 2: Import JSON ────────────────────────────────────────────────────
    with tab2:
        st.markdown('<div class="section-header">📂 Import — นำเข้าข้อมูลจาก JSON</div>', unsafe_allow_html=True)
        st.warning("⚠️ การนำเข้าจะ **เพิ่ม** ข้อมูลเข้าระบบ ไม่ได้ลบข้อมูลเดิม")

        uploaded_json = st.file_uploader("เลือกไฟล์ JSON", type=["json"], key="import_json")

        if uploaded_json:
            import json
            try:
                data = json.loads(uploaded_json.read().decode("utf-8"))
                eq_count   = len(data.get("equipment", []))
                tx_count   = len(data.get("transactions", []))
                borr_count = len(data.get("borrowers", []))
                backup_date = data.get("backup_date", "ไม่ทราบ")

                st.info(f"📋 ข้อมูลใน Backup | วันที่: {backup_date} | อุปกรณ์: {eq_count} รายการ | ประวัติ: {tx_count} รายการ | ผู้เบิก: {borr_count} คน")

                import_mode = st.radio("โหมดนำเข้า", 
                    ["เฉพาะอุปกรณ์ (equipment)", "ทั้งหมด (equipment + transactions + borrowers)"])

                if st.button("📂 ยืนยันนำเข้าข้อมูล", type="primary", use_container_width=True):
                    conn = get_conn()
                    c = conn.cursor()
                    imported = 0

                    # Import equipment
                    for eq in data.get("equipment", []):
                        try:
                            c.execute("""INSERT OR IGNORE INTO equipment
                                (code,name,category,total_qty,available_qty,status,image_path,description)
                                VALUES (?,?,?,?,?,?,?,?)""",
                                (eq["code"], eq["name"], eq.get("category"), eq.get("total_qty",1),
                                 eq.get("available_qty",1), eq.get("status","พร้อมใช้"),
                                 eq.get("image_path"), eq.get("description")))
                            imported += 1
                        except:
                            pass

                    if "ทั้งหมด" in import_mode:
                        # Import borrowers
                        id_map = {}
                        for b in data.get("borrowers", []):
                            old_id = b["id"]
                            c.execute("""INSERT INTO borrowers (name,type,student_id,department,phone)
                                VALUES (?,?,?,?,?)""",
                                (b["name"], b["type"], b.get("student_id"), b.get("department"), b.get("phone")))
                            id_map[old_id] = c.lastrowid

                        # Import transactions
                        eq_code_map = {r["code"]: r["id"] for r in query("SELECT id,code FROM equipment").to_dict("records")}
                        for tx in data.get("transactions", []):
                            try:
                                new_borr_id = id_map.get(tx["borrower_id"])
                                eq_id = tx.get("equipment_id")
                                if new_borr_id and eq_id:
                                    c.execute("""INSERT INTO transactions
                                        (equipment_id,borrower_id,qty,borrow_date,due_date,return_date,
                                         condition_out,condition_in,note,status)
                                        VALUES (?,?,?,?,?,?,?,?,?,?)""",
                                        (eq_id, new_borr_id, tx.get("qty",1),
                                         tx.get("borrow_date"), tx.get("due_date"), tx.get("return_date"),
                                         tx.get("condition_out","ปกติ"), tx.get("condition_in"),
                                         tx.get("note"), tx.get("status","คืนแล้ว")))
                            except:
                                pass

                    conn.commit()
                    conn.close()
                    st.success(f"✅ นำเข้าข้อมูลสำเร็จ!")
                    st.rerun()

            except Exception as e:
                st.error(f"❌ ไฟล์ไม่ถูกต้อง: {e}")

    # ── TAB 3: ล้างข้อมูล ─────────────────────────────────────────────────────
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
            # คำอธิบายแต่ละโหมด
            desc = {
                "🔄 รีเซ็ตจำนวนอุปกรณ์ (available = total)": "รีเซ็ต available_qty ของทุกอุปกรณ์ให้เท่ากับ total_qty และเปลี่ยนสถานะเป็น พร้อมใช้ ใช้หลังตรวจนับอุปกรณ์",
                "📋 ล้างประวัติการเบิก-คืนทั้งหมด":          "ลบ transactions และ borrowers ทั้งหมด ข้อมูลอุปกรณ์ยังอยู่ ใช้เปิดเทอมใหม่",
                "💥 ล้างทุกอย่าง (เริ่มระบบใหม่)":           "ลบข้อมูลทุกตาราง เริ่มระบบใหม่ทั้งหมด",
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
                            execute("UPDATE equipment SET available_qty=total_qty, status='พร้อมใช้'")
                            st.success("✅ รีเซ็ตจำนวนอุปกรณ์เรียบร้อย")

                        elif "ล้างประวัติ" in clear_mode:
                            execute("DELETE FROM transactions")
                            execute("DELETE FROM borrowers")
                            execute("UPDATE equipment SET available_qty=total_qty, status='พร้อมใช้'")
                            st.success("✅ ล้างประวัติการเบิก-คืนเรียบร้อย")

                        elif "ล้างทุกอย่าง" in clear_mode:
                            execute("DELETE FROM transactions")
                            execute("DELETE FROM borrowers")
                            execute("DELETE FROM equipment")
                            st.success("✅ ล้างข้อมูลทั้งหมดเรียบร้อย เริ่มระบบใหม่ได้เลย")

                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")

# ─── FOOTER ───────────────────────────────────────────────────────────────────
def footer():
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; color:#888; font-size:0.82rem; padding:8px 0 16px 0; line-height:1.8;">
        🔬 ระบบบริหารจัดการเบิก-คืนอุปกรณ์ห้องปฏิบัติการ TTC<br>
        พัฒนาโดย <b style="color:#1F4E79;">รศ.ดร.อิทธิพล มีผล</b><br>
        ภาควิชาครุศาสตร์โยธา &nbsp;|&nbsp; คณะครุศาสตร์อุตสาหกรรม<br>
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
    elif page == "ตั้งค่า":   page_settings()
    footer()

if __name__ == "__main__":
    main()
