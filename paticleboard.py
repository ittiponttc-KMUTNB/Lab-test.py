import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt

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
# Streamlit UI
# ---------------------------------------------------------

st.title("🧪 Particleboard Bending Test")
st.subheader("MOR • MOE • Thickness Swelling • Load–Deflection Graph")

# ---------------------------------------------------------
# MOR MULTI-SAMPLE
# ---------------------------------------------------------

st.write("---")
st.header("1) คำนวณ MOR (เลือกจำนวนตัวอย่าง 1–4)")

num_samples = st.selectbox("เลือกจำนวนตัวอย่าง", [1, 2, 3, 4])

sample_inputs = []

for i in range(num_samples):
    st.subheader(f"ตัวอย่างที่ {i+1}")

    L = st.number_input(f"L (mm) – ตัวอย่าง {i+1}", 0.0, key=f"L_{i}")
    b = st.number_input(f"b (mm) – ตัวอย่าง {i+1}", 0.0, key=f"b_{i}")
    t = st.number_input(f"t (mm) – ตัวอย่าง {i+1}", 0.0, key=f"t_{i}")
    Fmax_kg = st.number_input(f"Fmax (kg) – ตัวอย่าง {i+1}", 0.0, key=f"Fmax_{i}")

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
# Excel Template for Load–Deflection
# ---------------------------------------------------------

st.write("---")
st.header("2) Upload Excel เพื่อคำนวณ MOE และสร้างกราฟ")

st.info("Template Excel จะมีคอลัมน์: Load (kg), Deflection (mm)")

template = pd.DataFrame({
    "Load (kg)": [0, 5, 10, 15],
    "Deflection (mm)": [0, 1, 2, 3]
})

buffer = io.BytesIO()
template.to_excel(buffer, index=False)
buffer.seek(0)

st.download_button(
    label="📥 ดาวน์โหลด Template Excel",
    data=buffer,
    file_name="load_deflection_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.write("ข้อมูลที่อ่านได้:")
    st.dataframe(df)

    df["Load (N)"] = df["Load (kg)"] * 9.80665

    # ดึง Fmax อัตโนมัติจาก Excel
    auto_Fmax_kg = df["Load (kg)"].max()

    # ให้ผู้ใช้แก้ไขได้
    Fmax_kg_excel = st.number_input(
        "Fmax (kg) สำหรับคำนวณ MOE (ดึงจาก Excel อัตโนมัติ แก้ไขได้)",
        value=float(auto_Fmax_kg)
    )
    Fmax_N_excel = Fmax_kg_excel * 9.80665

    ymax = df["Deflection (mm)"].max()

    try:
        moe = calc_MOE_from_excel(Fmax_N_excel, ymax, L, b, t)
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
# Thickness Swelling (TS)
# ---------------------------------------------------------

st.write("---")
st.header("3) Thickness Swelling (TS) – วัด 4 ด้าน")

st.subheader("ก่อนแช่น้ำ")
before = [
    st.number_input("ด้านที่ 1 ก่อนแช่น้ำ (mm)", 0.0, key="b1"),
    st.number_input("ด้านที่ 2 ก่อนแช่น้ำ (mm)", 0.0, key="b2"),
    st.number_input("ด้านที่ 3 ก่อนแช่น้ำ (mm)", 0.0, key="b3"),
    st.number_input("ด้านที่ 4 ก่อนแช่น้ำ (mm)", 0.0, key="b4"),
]

st.subheader("หลังแช่น้ำ")
after = [
    st.number_input("ด้านที่ 1 หลังแช่น้ำ (mm)", 0.0, key="a1"),
    st.number_input("ด้านที่ 2 หลังแช่น้ำ (mm)", 0.0, key="a2"),
    st.number_input("ด้านที่ 3 หลังแช่น้ำ (mm)", 0.0, key="a3"),
    st.number_input("ด้านที่ 4 หลังแช่น้ำ (mm)", 0.0, key="a4"),
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
