import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import json
from docx import Document

# ---------------------------------------------------------
# Calculation Functions
# ---------------------------------------------------------

def calc_MOR(Fmax_N, L, b, t):
    return (3 * Fmax_N * L) / (2 * b * t**2)

def calc_MOE_from_excel(Fmax_N, ymax, L, b, t):
    if ymax == 0 or b == 0 or t == 0:
        raise ZeroDivisionError
    return (Fmax_N * L**3) / (4 * b * t**3 * ymax)

def calc_TS(avg_before, avg_after):
    if avg_before == 0:
        raise ZeroDivisionError
    return ((avg_after - avg_before) / avg_before) * 100

# ---------------------------------------------------------
# Sidebar: JSON Upload
# ---------------------------------------------------------

st.sidebar.header("📂 โหลด/บันทึกค่าพารามิเตอร์")

uploaded_json = st.sidebar.file_uploader("Upload JSON", type=["json"])

if uploaded_json is not None:
    try:
        loaded_data = json.load(uploaded_json)
        file_id = f"{uploaded_json.name}_{uploaded_json.size}"
        if st.session_state.get("last_uploaded_file") != file_id:
            st.session_state["last_uploaded_file"] = file_id
            for key, value in loaded_data.items():
                st.session_state[key] = value
            st.sidebar.success("✅ โหลดข้อมูลจาก JSON สำเร็จ")
            st.rerun()
    except Exception as e:
        st.sidebar.error(f"❌ อ่านไฟล์ JSON ไม่ได้: {e}")

# ---------------------------------------------------------
# Main UI
# ---------------------------------------------------------

st.title("🧪 Particleboard Bending Test")
st.subheader("MOR • MOE • Thickness Swelling • Load–Deflection Graph")

# ---------------------------------------------------------
# 1) MOR – เลือกจำนวนตัวอย่าง 1–4
# ---------------------------------------------------------

st.write("---")
st.header("1) คำนวณ MOR (เลือกจำนวนตัวอย่าง 1–4)")

num_samples = st.selectbox(
    "เลือกจำนวนตัวอย่าง",
    [1, 2, 3, 4],
    index=[1, 2, 3, 4].index(st.session_state.get("num_samples", 1)),
    key="num_samples"
)

sample_inputs = []

for i in range(num_samples):
    st.subheader(f"ตัวอย่างที่ {i+1}")

    L = st.number_input(
        f"L (mm) – ตัวอย่าง {i+1}",
        value=st.session_state.get(f"L_{i}", 0.0),
        key=f"L_{i}"
    )
    b = st.number_input(
        f"b (mm) – ตัวอย่าง {i+1}",
        value=st.session_state.get(f"b_{i}", 0.0),
        key=f"b_{i}"
    )
    t = st.number_input(
        f"t (mm) – ตัวอย่าง {i+1}",
        value=st.session_state.get(f"t_{i}", 0.0),
        key=f"t_{i}"
    )
    Fmax_kg = st.number_input(
        f"Fmax (kg) – ตัวอย่าง {i+1}",
        value=st.session_state.get(f"Fmax_{i}", 0.0),
        key=f"Fmax_{i}"
    )

    sample_inputs.append((L, b, t, Fmax_kg))

if st.button("คำนวณ MOR ทั้งหมด"):
    for i, (L, b, t, Fmax_kg) in enumerate(sample_inputs):
        try:
            Fmax_N = Fmax_kg * 9.80665
            mor = calc_MOR(Fmax_N, L, b, t)
            st.success(f"MOR ตัวอย่างที่ {i+1} = {mor:.2f} MPa")
        except ZeroDivisionError:
            st.error(f"❌ ตัวอย่างที่ {i+1}: L, b, t ต้องไม่เป็นศูนย์")
        except Exception as e:
            st.error(f"⚠️ ตัวอย่างที่ {i+1}: เกิดข้อผิดพลาด {e}")

# ---------------------------------------------------------
# 2) MOE + Load–Deflection จาก Excel
# ---------------------------------------------------------

st.write("---")
st.header("2) Upload Excel เพื่อคำนวณ MOE และสร้างกราฟ")

st.info("Template Excel จะมีคอลัมน์: Load (kg), Deflection (mm)")

template = pd.DataFrame({
    "Load (kg)": [0, 5, 10, 15],
    "Deflection (mm)": [0, 1, 2, 3]
})

buffer_template = io.BytesIO()
template.to_excel(buffer_template, index=False)
buffer_template.seek(0)

st.download_button(
    label="📥 ดาวน์โหลด Template Excel",
    data=buffer_template,
    file_name="load_deflection_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.write("ข้อมูลที่อ่านได้:")
    st.dataframe(df)

    df["Load (N)"] = df["Load (kg)"] * 9.80665

    auto_Fmax_kg = float(df["Load (kg)"].max())
    Fmax_kg_excel = st.number_input(
        "Fmax (kg) สำหรับคำนวณ MOE (ดึงจาก Excel อัตโนมัติ แก้ไขได้)",
        value=st.session_state.get("Fmax_moe", auto_Fmax_kg),
        key="Fmax_moe"
    )
    Fmax_N_excel = Fmax_kg_excel * 9.80665

    # ใช้ geometry จากตัวอย่างที่ 1 (หรือจะแยกชุดใหม่ก็ได้)
    L_moe = st.number_input(
        "L (mm) สำหรับ MOE",
        value=st.session_state.get("L_moe", st.session_state.get("L_0", 0.0)),
        key="L_moe"
    )
    b_moe = st.number_input(
        "b (mm) สำหรับ MOE",
        value=st.session_state.get("b_moe", st.session_state.get("b_0", 0.0)),
        key="b_moe"
    )
    t_moe = st.number_input(
        "t (mm) สำหรับ MOE",
        value=st.session_state.get("t_moe", st.session_state.get("t_0", 0.0)),
        key="t_moe"
    )

    ymax = df["Deflection (mm)"].max()

    if st.button("คำนวณ MOE จาก Excel"):
        try:
            moe = calc_MOE_from_excel(Fmax_N_excel, ymax, L_moe, b_moe, t_moe)
            st.success(f"MOE (จาก Excel) = {moe:.2f} MPa")
        except ZeroDivisionError:
            st.error("❌ ไม่สามารถคำนวณ MOE ได้: L, b, t หรือ ymax ต้องไม่เป็นศูนย์")
        except Exception as e:
            st.error(f"⚠️ เกิดข้อผิดพลาด {e}")

    fig, ax = plt.subplots()
    ax.plot(df["Deflection (mm)"], df["Load (N)"], marker="o")
    ax.set_xlabel("Deflection (mm)")
    ax.set_ylabel("Load (N)")
    ax.set_title("Load–Deflection Curve")
    ax.grid(True)
    st.pyplot(fig)

# ---------------------------------------------------------
# 3) Thickness Swelling (TS) – วัด 4 ด้าน
# ---------------------------------------------------------

st.write("---")
st.header("3) Thickness Swelling (TS) – วัด 4 ด้าน")

st.subheader("ก่อนแช่น้ำ")
before = [
    st.number_input("ด้านที่ 1 ก่อนแช่น้ำ (mm)", value=st.session_state.get("b1", 0.0), key="b1"),
    st.number_input("ด้านที่ 2 ก่อนแช่น้ำ (mm)", value=st.session_state.get("b2", 0.0), key="b2"),
    st.number_input("ด้านที่ 3 ก่อนแช่น้ำ (mm)", value=st.session_state.get("b3", 0.0), key="b3"),
    st.number_input("ด้านที่ 4 ก่อนแช่น้ำ (mm)", value=st.session_state.get("b4", 0.0), key="b4"),
]

st.subheader("หลังแช่น้ำ")
after = [
    st.number_input("ด้านที่ 1 หลังแช่น้ำ (mm)", value=st.session_state.get("a1", 0.0), key="a1"),
    st.number_input("ด้านที่ 2 หลังแช่น้ำ (mm)", value=st.session_state.get("a2", 0.0), key="a2"),
    st.number_input("ด้านที่ 3 หลังแช่น้ำ (mm)", value=st.session_state.get("a3", 0.0), key="a3"),
    st.number_input("ด้านที่ 4 หลังแช่น้ำ (mm)", value=st.session_state.get("a4", 0.0), key="a4"),
]

if st.button("คำนวณ TS"):
    try:
        avg_before = sum(before) / 4
        avg_after = sum(after) / 4
        ts = calc_TS(avg_before, avg_after)
        st.success(f"TS = {ts:.2f} % (เฉลี่ย 4 ด้าน)")
    except ZeroDivisionError:
        st.error("❌ ค่าเฉลี่ยก่อนแช่น้ำต้องไม่เป็นศูนย์")
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาด {e}")

# ---------------------------------------------------------
# 4) Export JSON / Excel / Word
# ---------------------------------------------------------

st.write("---")
st.header("4) บันทึกผลการทดสอบ")

# JSON
export_data = {
    "num_samples": st.session_state.get("num_samples", 1),
    "L_0": st.session_state.get("L_0", 0),
    "b_0": st.session_state.get("b_0", 0),
    "t_0": st.session_state.get("t_0", 0),
    "Fmax_0": st.session_state.get("Fmax_0", 0),
    "L_moe": st.session_state.get("L_moe", 0),
    "b_moe": st.session_state.get("b_moe", 0),
    "t_moe": st.session_state.get("t_moe", 0),
    "Fmax_moe": st.session_state.get("Fmax_moe", 0),
    "b1": st.session_state.get("b1", 0),
    "b2": st.session_state.get("b2", 0),
    "b3": st.session_state.get("b3", 0),
    "b4": st.session_state.get("b4", 0),
    "a1": st.session_state.get("a1", 0),
    "a2": st.session_state.get("a2", 0),
    "a3": st.session_state.get("a3", 0),
    "a4": st.session_state.get("a4", 0),
}

json_str = json.dumps(export_data, ensure_ascii=False, indent=2)

st.download_button(
    label="💾 Download Input (JSON)",
    data=json_str,
    file_name="input_data.json",
    mime="application/json"
)

# Excel
if st.button("📊 เตรียมไฟล์ Excel"):
    export_excel = {
        "L_moe": [st.session_state.get("L_moe", 0)],
        "b_moe": [st.session_state.get("b_moe", 0)],
        "t_moe": [st.session_state.get("t_moe", 0)],
        "Fmax_moe": [st.session_state.get("Fmax_moe", 0)],
        "Before_1": [st.session_state.get("b1", 0)],
        "Before_2": [st.session_state.get("b2", 0)],
        "Before_3": [st.session_state.get("b3", 0)],
        "Before_4": [st.session_state.get("b4", 0)],
        "After_1": [st.session_state.get("a1", 0)],
        "After_2": [st.session_state.get("a2", 0)],
        "After_3": [st.session_state.get("a3", 0)],
        "After_4": [st.session_state.get("a4", 0)],
    }
    df_export = pd.DataFrame(export_excel)
    buffer_xlsx = io.BytesIO()
    df_export.to_excel(buffer_xlsx, index=False)
    buffer_xlsx.seek(0)

    st.download_button(
        label="📥 Download Excel (.xlsx)",
        data=buffer_xlsx,
        file_name="particleboard_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Word
if st.button("📄 Export Word Report"):
    try:
        doc = Document()
        doc.add_heading("Particleboard Test Report", level=1)

        doc.add_heading("MOR Parameters (ตัวอย่างที่ 1)", level=2)
        doc.add_paragraph(f"L = {st.session_state.get('L_0', 0)} mm")
        doc.add_paragraph(f"b = {st.session_state.get('b_0', 0)} mm")
        doc.add_paragraph(f"t = {st.session_state.get('t_0', 0)} mm")
        doc.add_paragraph(f"Fmax = {st.session_state.get('Fmax_0', 0)} kg")

        doc.add_heading("MOE Parameters", level=2)
        doc.add_paragraph(f"L_moe = {st.session_state.get('L_moe', 0)} mm")
        doc.add_paragraph(f"b_moe = {st.session_state.get('b_moe', 0)} mm")
        doc.add_paragraph(f"t_moe = {st.session_state.get('t_moe', 0)} mm")
        doc.add_paragraph(f"Fmax_moe = {st.session_state.get('Fmax_moe', 0)} kg")

        doc.add_heading("Thickness Swelling (TS)", level=2)
        doc.add_paragraph(
            f"Before (4 ด้าน) = "
            f"{st.session_state.get('b1', 0)}, "
            f"{st.session_state.get('b2', 0)}, "
            f"{st.session_state.get('b3', 0)}, "
            f"{st.session_state.get('b4', 0)}"
        )
        doc.add_paragraph(
            f"After (4 ด้าน) = "
            f"{st.session_state.get('a1', 0)}, "
            f"{st.session_state.get('a2', 0)}, "
            f"{st.session_state.get('a3', 0)}, "
            f"{st.session_state.get('a4', 0)}"
        )

        buffer_docx = io.BytesIO()
        doc.save(buffer_docx)
        buffer_docx.seek(0)

        st.download_button(
            "📥 Download Word Report",
            data=buffer_docx,
            file_name="particleboard_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"⚠️ ไม่สามารถสร้าง Word report ได้: {e}")
        # ---------------------------------------------------------
# Footer Credit
# ---------------------------------------------------------

st.write("---")
st.markdown(
    """
    <div style='text-align: center; font-size: 16px; padding-top: 10px;'>
        <b>พัฒนาโดย:</b> รศ.ดร.อิทธิพล มีผล<br>
        ภาควิชาครุศาสตร์โยธา มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (มจพ.)
    </div>
    """,
    unsafe_allow_html=True
)

