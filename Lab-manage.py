import streamlit as st
import sqlite3
import pandas as pd
import os
import base64
from datetime import datetime, date
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from pathlib import Path

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ระบบเบิก-คืนอุปกรณ์ห้องปฏิบัติการ",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_PATH = "lab_equipment.db"
IMG_DIR = "equipment_images"
ADMIN_PASSWORD = "admin1234"   # ← เปลี่ยน password ได้ที่นี่
os.makedirs(IMG_DIR, exist_ok=True)

# ─── ADMIN AUTH ───────────────────────────────────────────────────────────────
def is_admin():
    return st.session_state.get("is_admin", False)

def admin_login_widget():
    """Show lock icon + login form in sidebar"""
    if is_admin():
        st.sidebar.success("🔓 Admin Mode")
        if st.sidebar.button("🔒 ออกจาก Admin", use_container_width=True):
            st.session_state.is_admin = False
            st.rerun()
    else:
        with st.sidebar.expander("🔒 Admin Login"):
            pwd = st.text_input("รหัสผ่าน Admin", type="password", key="admin_pwd_input")
            if st.button("เข้าสู่ระบบ", use_container_width=True):
                if pwd == ADMIN_PASSWORD:
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    st.error("รหัสผ่านไม่ถูกต้อง")

def require_admin():
    """Block page and show warning if not admin"""
    if not is_admin():
        st.warning("🔒 ฟังก์ชันนี้สำหรับ Admin เท่านั้น กรุณา Login ที่ Sidebar")
        st.stop()

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
def img_to_base64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

def show_equipment_image(image_path, width=120):
    if image_path and os.path.exists(image_path):
        b64 = img_to_base64(image_path)
        st.markdown(
            f'<img src="data:image/png;base64,{b64}" style="width:{width}px;border-radius:8px;border:1px solid #ddd;">',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<div style="width:{width}px;height:{width}px;background:#f0f2f6;border-radius:8px;'
            f'display:flex;align-items:center;justify-content:center;font-size:2rem;border:1px solid #ddd;">📦</div>',
            unsafe_allow_html=True
        )

def status_badge(status):
    colors = {
        "พร้อมใช้": "#28a745", "ยืมออก": "#ffc107",
        "ชำรุด": "#dc3545", "สูญหาย": "#6c757d"
    }
    c = colors.get(status, "#6c757d")
    return f'<span style="background:{c};color:white;padding:2px 10px;border-radius:12px;font-size:0.8rem;">{status}</span>'

def overdue_badge(due_date):
    if not due_date or due_date == "-":
        return ""
    try:
        d = datetime.strptime(due_date, "%Y-%m-%d").date()
        if d < date.today():
            delta = (date.today() - d).days
            return f'<span style="background:#dc3545;color:white;padding:2px 8px;border-radius:12px;font-size:0.75rem;">เกินกำหนด {delta} วัน</span>'
    except:
        pass
    return ""

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("## 🔬 ระบบอุปกรณ์ Lab")
        st.markdown("---")
        pages = {
            "🏠 Dashboard": "dashboard",
            "📦 จัดการอุปกรณ์": "equipment",
            "➕ เบิกอุปกรณ์": "borrow",
            "✅ คืนอุปกรณ์": "return",
            "📋 รายงาน": "report",
        }
        if "page" not in st.session_state:
            st.session_state.page = "dashboard"
        for label, key in pages.items():
            if st.button(label, use_container_width=True,
                         type="primary" if st.session_state.page == key else "secondary"):
                st.session_state.page = key
                st.rerun()

        st.markdown("---")
        admin_login_widget()
        st.markdown("---")
        # Quick stats
        eq = query("SELECT COUNT(*) as n FROM equipment").iloc[0]["n"]
        active = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่'").iloc[0]["n"]
        overdue = query("""SELECT COUNT(*) as n FROM transactions 
                           WHERE status='ยืมอยู่' AND due_date < date('now')""").iloc[0]["n"]
        st.metric("อุปกรณ์ทั้งหมด", eq)
        st.metric("กำลังยืม", active)
        if overdue > 0:
            st.error(f"⚠️ เกินกำหนด {overdue} รายการ")

# ─── PAGE: DASHBOARD ──────────────────────────────────────────────────────────
def page_dashboard():
    st.title("🏠 Dashboard — ภาพรวมห้องปฏิบัติการ")

    col1, col2, col3, col4 = st.columns(4)
    total_eq = query("SELECT COUNT(*) as n FROM equipment").iloc[0]["n"]
    available = query("SELECT COALESCE(SUM(available_qty),0) as n FROM equipment").iloc[0]["n"]
    active_tx = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่'").iloc[0]["n"]
    overdue_tx = query("SELECT COUNT(*) as n FROM transactions WHERE status='ยืมอยู่' AND due_date < date('now')").iloc[0]["n"]

    col1.metric("📦 รายการอุปกรณ์", total_eq)
    col2.metric("✅ พร้อมใช้งาน", int(available))
    col3.metric("🔄 กำลังยืม", active_tx)
    col4.metric("⚠️ เกินกำหนด", overdue_tx, delta=None)
    if overdue_tx > 0:
        col4.error("มีรายการเกินกำหนด!")

    st.markdown("---")

    # Active borrows
    st.subheader("📋 รายการที่กำลังยืมอยู่")
    df = query("""
        SELECT t.id, e.code, e.name, b.name as borrower, b.type, t.qty,
               t.borrow_date, t.due_date, e.image_path
        FROM transactions t
        JOIN equipment e ON t.equipment_id = e.id
        JOIN borrowers b ON t.borrower_id = b.id
        WHERE t.status = 'ยืมอยู่'
        ORDER BY t.due_date ASC
    """)
    if df.empty:
        st.info("ไม่มีรายการยืมในขณะนี้")
    else:
        for _, row in df.iterrows():
            with st.container():
                c1, c2, c3, c4, c5 = st.columns([1, 2, 3, 2, 2])
                with c1:
                    show_equipment_image(row["image_path"], width=60)
                c2.markdown(f"**{row['code']}**<br>{row['name']}", unsafe_allow_html=True)
                c3.markdown(f"👤 {row['borrower']} ({row['type']})")
                c4.markdown(f"📅 คืน: **{row['due_date']}**<br>{overdue_badge(row['due_date'])}", unsafe_allow_html=True)
                c5.markdown(f"จำนวน: {row['qty']} ชิ้น")
                st.divider()

    # Overdue alert
    if overdue_tx > 0:
        st.error("### ⚠️ รายการเกินกำหนดคืน")
        df_od = query("""
            SELECT e.code, e.name, b.name as borrower, b.phone, t.due_date,
                   julianday('now') - julianday(t.due_date) as days_overdue
            FROM transactions t
            JOIN equipment e ON t.equipment_id = e.id
            JOIN borrowers b ON t.borrower_id = b.id
            WHERE t.status = 'ยืมอยู่' AND t.due_date < date('now')
            ORDER BY days_overdue DESC
        """)
        df_od["days_overdue"] = df_od["days_overdue"].apply(lambda x: f"{int(x)} วัน")
        st.dataframe(df_od.rename(columns={
            "code": "รหัส", "name": "อุปกรณ์", "borrower": "ผู้ยืม",
            "phone": "โทรศัพท์", "due_date": "กำหนดคืน", "days_overdue": "เกินมา"
        }), use_container_width=True, hide_index=True)

# ─── PAGE: EQUIPMENT ──────────────────────────────────────────────────────────
def page_equipment():
    st.title("📦 จัดการอุปกรณ์")

    tab_labels = ["รายการอุปกรณ์"]
    if is_admin():
        tab_labels.append("➕ เพิ่ม/แก้ไขอุปกรณ์")
    tabs = st.tabs(tab_labels)

    # ── TAB 1: รายการอุปกรณ์ (ทุกคนดูได้) ────────────────────────────────────
    with tabs[0]:
        search = st.text_input("🔍 ค้นหา (ชื่อ/รหัส/หมวดหมู่)", "")
        cats = query("SELECT DISTINCT category FROM equipment WHERE category IS NOT NULL")
        cat_filter = st.selectbox("หมวดหมู่", ["ทั้งหมด"] + cats["category"].tolist())

        sql = """SELECT e.id, e.code, e.name, e.category, e.total_qty, e.available_qty,
                        e.status, e.image_path, e.description
                 FROM equipment e WHERE 1=1"""
        params = []
        if search:
            sql += " AND (e.name LIKE ? OR e.code LIKE ? OR e.category LIKE ?)"
            params += [f"%{search}%", f"%{search}%", f"%{search}%"]
        if cat_filter != "ทั้งหมด":
            sql += " AND e.category = ?"
            params.append(cat_filter)
        sql += " ORDER BY e.code"

        df = query(sql, params)
        if df.empty:
            st.info("ไม่พบอุปกรณ์")
        else:
            for _, row in df.iterrows():
                with st.expander(f"**{row['code']}** — {row['name']}  |  {row['category'] or '-'}"):
                    c1, c2, c3 = st.columns([1, 3, 1])
                    with c1:
                        show_equipment_image(row["image_path"], width=100)
                    with c2:
                        st.markdown(f"**ชื่อ:** {row['name']}")
                        st.markdown(f"**หมวด:** {row['category'] or '-'}")
                        st.markdown(f"**จำนวนทั้งหมด:** {row['total_qty']} | **พร้อมใช้:** {row['available_qty']}")
                        st.markdown(status_badge(row["status"]), unsafe_allow_html=True)
                        if row["description"]:
                            st.caption(row["description"])
                    with c3:
                        lt = query("""SELECT b.name, t.borrow_date FROM transactions t
                                      JOIN borrowers b ON t.borrower_id=b.id
                                      WHERE t.equipment_id=? ORDER BY t.created_at DESC LIMIT 1""",
                                   (row["id"],))
                        if not lt.empty:
                            st.caption(f"ผู้ยืมล่าสุด:\n**{lt.iloc[0]['name']}**\n{lt.iloc[0]['borrow_date']}")
                        # ปุ่มลบ: เฉพาะ Admin
                        if is_admin():
                            if st.button("🗑️ ลบ", key=f"del_{row['id']}"):
                                active = query("SELECT COUNT(*) as n FROM transactions WHERE equipment_id=? AND status='ยืมอยู่'",
                                               (row["id"],)).iloc[0]["n"]
                                if active > 0:
                                    st.error("ไม่สามารถลบได้ มีการยืมอยู่")
                                else:
                                    execute("DELETE FROM equipment WHERE id=?", (row["id"],))
                                    st.success("ลบแล้ว")
                                    st.rerun()
                        else:
                            st.caption("🔒 Admin เท่านั้น")

    # ── TAB 2: เพิ่ม/แก้ไข (Admin เท่านั้น) ──────────────────────────────────
    if is_admin():
        with tabs[1]:
            eq_list = query("SELECT id, code, name FROM equipment ORDER BY code")
            options = ["➕ เพิ่มใหม่"] + [f"{r['code']} — {r['name']}" for _, r in eq_list.iterrows()]
            choice = st.selectbox("เลือกอุปกรณ์ที่ต้องการแก้ไข หรือเพิ่มใหม่", options)

            if choice == "➕ เพิ่มใหม่":
                existing = None
            else:
                eq_id = eq_list.iloc[options.index(choice) - 1]["id"]
                existing = query("SELECT * FROM equipment WHERE id=?", (eq_id,)).iloc[0]

            with st.form("eq_form"):
                col1, col2 = st.columns(2)
                code = col1.text_input("รหัสอุปกรณ์*", value=existing["code"] if existing is not None else "")
                name = col2.text_input("ชื่ออุปกรณ์*", value=existing["name"] if existing is not None else "")
                category = col1.text_input("หมวดหมู่", value=existing["category"] if existing is not None else "")
                total_qty = col2.number_input("จำนวนทั้งหมด", min_value=1, value=int(existing["total_qty"]) if existing is not None else 1)
                status = col1.selectbox("สถานะ", ["พร้อมใช้", "ชำรุด", "สูญหาย"],
                                        index=["พร้อมใช้", "ชำรุด", "สูญหาย"].index(existing["status"]) if existing is not None else 0)
                description = st.text_area("รายละเอียดเพิ่มเติม", value=existing["description"] if existing is not None else "")
                uploaded = st.file_uploader("📷 อัปโหลดรูปอุปกรณ์ (optional)", type=["jpg", "jpeg", "png"])

                submitted = st.form_submit_button("💾 บันทึก", type="primary")
                if submitted:
                    if not code or not name:
                        st.error("กรุณากรอกรหัสและชื่ออุปกรณ์")
                    else:
                        img_path = existing["image_path"] if existing is not None else None
                        if uploaded:
                            ext = uploaded.name.split(".")[-1]
                            img_path = os.path.join(IMG_DIR, f"{code}.{ext}")
                            with open(img_path, "wb") as f:
                                f.write(uploaded.getbuffer())
                        if existing is None:
                            execute("""INSERT INTO equipment (code, name, category, total_qty, available_qty, status, image_path, description)
                                       VALUES (?,?,?,?,?,?,?,?)""",
                                    (code, name, category, total_qty, total_qty, status, img_path, description))
                            st.success("✅ เพิ่มอุปกรณ์เรียบร้อย")
                        else:
                            execute("""UPDATE equipment SET code=?, name=?, category=?, total_qty=?,
                                       status=?, image_path=?, description=? WHERE id=?""",
                                    (code, name, category, total_qty, status, img_path, description, eq_id))
                            st.success("✅ อัปเดตอุปกรณ์เรียบร้อย")
                        st.rerun()

# ─── PAGE: BORROW ─────────────────────────────────────────────────────────────
def page_borrow():
    st.title("➕ เบิกอุปกรณ์")

    avail = query("SELECT id, code, name, available_qty FROM equipment WHERE available_qty > 0 AND status='พร้อมใช้' ORDER BY code")
    if avail.empty:
        st.warning("ไม่มีอุปกรณ์พร้อมใช้งานในขณะนี้")
        return

    st.subheader("ข้อมูลผู้เบิก")
    borrower_type = st.radio("ประเภทผู้เบิก", ["นักศึกษา", "บุคลากร/อาจารย์"], horizontal=True)

    col1, col2 = st.columns(2)
    borrower_name = col1.text_input("ชื่อ-นามสกุล*")
    student_id = col2.text_input("รหัสนักศึกษา/รหัสพนักงาน")
    department = col1.text_input("ภาควิชา/หน่วยงาน")
    phone = col2.text_input("เบอร์โทรศัพท์")

    st.subheader("รายการอุปกรณ์")
    eq_options = {f"{r['code']} — {r['name']} (คงเหลือ {r['available_qty']})": r["id"] for _, r in avail.iterrows()}
    selected_eq = st.selectbox("เลือกอุปกรณ์*", list(eq_options.keys()))
    eq_id = eq_options[selected_eq]

    max_qty = int(avail[avail["id"] == eq_id]["available_qty"].values[0])
    qty = st.number_input("จำนวน*", min_value=1, max_value=max_qty, value=1)

    # Show image
    eq_data = query("SELECT image_path, description FROM equipment WHERE id=?", (eq_id,)).iloc[0]
    if eq_data["image_path"] and os.path.exists(eq_data["image_path"]):
        c1, c2 = st.columns([1, 3])
        with c1:
            show_equipment_image(eq_data["image_path"], width=120)
        with c2:
            if eq_data["description"]:
                st.info(eq_data["description"])
    
    col1, col2 = st.columns(2)
    borrow_date = col1.date_input("วันที่เบิก*", value=date.today())
    due_date = col2.date_input("วันกำหนดคืน*", value=date.today())
    condition_out = st.selectbox("สภาพอุปกรณ์ขณะเบิก", ["ปกติ", "มีรอยขีดข่วน", "ชำรุดบางส่วน"])
    note = st.text_area("หมายเหตุ")

    if st.button("✅ ยืนยันการเบิก", type="primary", use_container_width=True):
        if not borrower_name:
            st.error("กรุณากรอกชื่อผู้เบิก")
        elif due_date < borrow_date:
            st.error("วันกำหนดคืนต้องไม่ก่อนวันที่เบิก")
        else:
            # Save borrower
            b_id = execute("""INSERT INTO borrowers (name, type, student_id, department, phone)
                               VALUES (?,?,?,?,?)""",
                           (borrower_name, borrower_type, student_id, department, phone))
            # Save transaction
            execute("""INSERT INTO transactions (equipment_id, borrower_id, qty, borrow_date, due_date, condition_out, note)
                       VALUES (?,?,?,?,?,?,?)""",
                    (eq_id, b_id, qty, str(borrow_date), str(due_date), condition_out, note))
            # Update available qty
            execute("UPDATE equipment SET available_qty = available_qty - ? WHERE id=?", (qty, eq_id))
            st.success(f"✅ บันทึกการเบิกเรียบร้อย — {borrower_name} เบิก {qty} ชิ้น กำหนดคืน {due_date}")
            st.balloons()

# ─── PAGE: RETURN ─────────────────────────────────────────────────────────────
def page_return():
    st.title("✅ คืนอุปกรณ์")

    df = query("""
        SELECT t.id, e.code, e.name, e.image_path, b.name as borrower, b.type,
               t.qty, t.borrow_date, t.due_date, t.condition_out, t.note as borrow_note
        FROM transactions t
        JOIN equipment e ON t.equipment_id = e.id
        JOIN borrowers b ON t.borrower_id = b.id
        WHERE t.status = 'ยืมอยู่'
        ORDER BY t.due_date ASC
    """)

    if df.empty:
        st.info("ไม่มีรายการยืมในขณะนี้")
        return

    # Search
    search = st.text_input("🔍 ค้นหาชื่อผู้ยืม หรือรหัสอุปกรณ์")
    if search:
        df = df[df["borrower"].str.contains(search, case=False) | 
                df["code"].str.contains(search, case=False) |
                df["name"].str.contains(search, case=False)]

    if df.empty:
        st.info("ไม่พบรายการ")
        return

    for _, row in df.iterrows():
        overdue_html = overdue_badge(row["due_date"])
        with st.expander(f"🔑 TX#{row['id']} | **{row['code']}** {row['name']} — {row['borrower']} | กำหนดคืน: {row['due_date']}"):
            c1, c2 = st.columns([1, 4])
            with c1:
                show_equipment_image(row["image_path"], width=100)
            with c2:
                st.markdown(f"**ผู้ยืม:** {row['borrower']} ({row['type']})")
                st.markdown(f"**อุปกรณ์:** {row['code']} — {row['name']} จำนวน {row['qty']} ชิ้น")
                st.markdown(f"**วันที่เบิก:** {row['borrow_date']} | **กำหนดคืน:** {row['due_date']} {overdue_html}", unsafe_allow_html=True)
                st.markdown(f"**สภาพตอนเบิก:** {row['condition_out']}")
                if row["borrow_note"]:
                    st.caption(f"หมายเหตุ: {row['borrow_note']}")

            col1, col2 = st.columns(2)
            return_date = col1.date_input("วันที่คืน", value=date.today(), key=f"rd_{row['id']}")
            condition_in = col2.selectbox("สภาพอุปกรณ์เมื่อคืน", ["ปกติ", "มีรอยขีดข่วน", "ชำรุด", "สูญหาย"], key=f"ci_{row['id']}")
            return_note = st.text_input("หมายเหตุ (ถ้ามี)", key=f"rn_{row['id']}")

            if st.button(f"✅ บันทึกการคืน", key=f"ret_{row['id']}", type="primary"):
                # Get equipment id
                tx_data = query("SELECT equipment_id, qty FROM transactions WHERE id=?", (row["id"],)).iloc[0]
                execute("""UPDATE transactions SET return_date=?, condition_in=?, note=?, status='คืนแล้ว'
                           WHERE id=?""", (str(return_date), condition_in, return_note, row["id"]))
                execute("UPDATE equipment SET available_qty = available_qty + ? WHERE id=?",
                        (tx_data["qty"], tx_data["equipment_id"]))
                if condition_in == "ชำรุด":
                    execute("UPDATE equipment SET status='ชำรุด' WHERE id=?", (tx_data["equipment_id"],))
                elif condition_in == "สูญหาย":
                    execute("UPDATE equipment SET status='สูญหาย', available_qty=available_qty-? WHERE id=?",
                            (tx_data["qty"], tx_data["equipment_id"]))
                st.success(f"✅ บันทึกการคืนเรียบร้อย สภาพ: {condition_in}")
                st.rerun()

# ─── PAGE: REPORT ─────────────────────────────────────────────────────────────
def page_report():
    st.title("📋 รายงาน")

    tab1, tab2 = st.tabs(["ประวัติการเบิก-คืน", "สรุปอุปกรณ์"])

    with tab1:
        col1, col2, col3 = st.columns(3)
        date_from = col1.date_input("ตั้งแต่วันที่", value=date(date.today().year, 1, 1))
        date_to = col2.date_input("ถึงวันที่", value=date.today())
        status_filter = col3.selectbox("สถานะ", ["ทั้งหมด", "ยืมอยู่", "คืนแล้ว"])

        sql = """
            SELECT t.id as 'TX#', e.code as 'รหัสอุปกรณ์', e.name as 'ชื่ออุปกรณ์',
                   b.name as 'ผู้เบิก', b.type as 'ประเภท', b.student_id as 'รหัส',
                   b.department as 'ภาควิชา', b.phone as 'โทรศัพท์',
                   t.qty as 'จำนวน', t.borrow_date as 'วันที่เบิก',
                   t.due_date as 'กำหนดคืน', t.return_date as 'วันที่คืน',
                   t.condition_out as 'สภาพตอนเบิก', t.condition_in as 'สภาพตอนคืน',
                   t.status as 'สถานะ', t.note as 'หมายเหตุ'
            FROM transactions t
            JOIN equipment e ON t.equipment_id = e.id
            JOIN borrowers b ON t.borrower_id = b.id
            WHERE t.borrow_date BETWEEN ? AND ?
        """
        params = [str(date_from), str(date_to)]
        if status_filter != "ทั้งหมด":
            sql += " AND t.status = ?"
            params.append(status_filter)
        sql += " ORDER BY t.id DESC"

        df = query(sql, params)
        st.markdown(f"**พบ {len(df)} รายการ**")
        st.dataframe(df, use_container_width=True, hide_index=True)

        if not df.empty:
            if is_admin():
                if st.button("📥 Export Excel", type="primary"):
                    excel_data = export_excel(df, "ประวัติการเบิก-คืนอุปกรณ์")
                    st.download_button(
                        "⬇️ ดาวน์โหลด Excel",
                        data=excel_data,
                        file_name=f"borrow_history_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("🔒 Export Excel สำหรับ Admin เท่านั้น")

    with tab2:
        df2 = query("""
            SELECT e.code as 'รหัส', e.name as 'ชื่ออุปกรณ์', e.category as 'หมวดหมู่',
                   e.total_qty as 'จำนวนทั้งหมด', e.available_qty as 'พร้อมใช้',
                   (e.total_qty - e.available_qty) as 'กำลังยืม',
                   e.status as 'สถานะ',
                   (SELECT COUNT(*) FROM transactions t WHERE t.equipment_id=e.id) as 'จำนวนครั้งที่เบิก'
            FROM equipment e ORDER BY e.code
        """)
        st.dataframe(df2, use_container_width=True, hide_index=True)

        if not df2.empty:
            if is_admin():
                if st.button("📥 Export Excel สรุปอุปกรณ์", type="primary"):
                    excel_data = export_excel(df2, "สรุปรายการอุปกรณ์")
                    st.download_button(
                        "⬇️ ดาวน์โหลด Excel",
                        data=excel_data,
                        file_name=f"equipment_summary_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("🔒 Export Excel สำหรับ Admin เท่านั้น")

def export_excel(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        fill = PatternFill("solid", fgColor="EBF3FB") if row_idx % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 30)

    ws.row_dimensions[1].height = 25
    wb.save(output)
    return output.getvalue()

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    init_db()
    sidebar()

    page = st.session_state.get("page", "dashboard")
    if page == "dashboard":
        page_dashboard()
    elif page == "equipment":
        page_equipment()
    elif page == "borrow":
        page_borrow()
    elif page == "return":
        page_return()
    elif page == "report":
        page_report()

if __name__ == "__main__":
    main()
