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
        return None
    return (Fmax_N * L**3) / (4 * b * t**3 * ymax)

def calc_TS(t_before, t_after):
    return ((t_after - t_before) / t_before) * 100

# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------

st.title("🧪 Particleboard Bending Test (Single Sample)")
st.subheader("MOR • MOE (from Excel) • Thickness Swelling • Load–Deflection Graph")

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
    results = []
    for i, (L, b, t, Fmax_kg) in enumerate(sample_inputs):
        Fmax_N = Fmax_kg * 9.80665
        mor = calc_MOR(Fmax_N, L, b, t)
        results.append(mor)
        st.success(f"MOR ตัวอย่างที่ {i+1} = {mor:.2f} MPa")

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
    ymax = df["Deflection (mm)"].max()

    moe = calc_MOE_from_excel(Fmax_N, ymax, L, b, t)

    if moe is None:
        st.error("ไม่สามารถคำนวณ MOE ได้ (ymax หรือค่าบางตัวเป็นศูนย์)")
    else:
        st.success(f"MOE (จาก Excel) = {moe:.2f} MPa")

    fig, ax = plt.subplots()
    ax.plot(df["Deflection (mm)"], df["Load (N)"], marker="o")
    ax.set_xlabel("Deflection (mm)")
    ax.set_ylabel("Load (N)")
    ax.set_title("Load–Deflection Curve")
    ax.grid(True)

    st.pyplot(fig)

# ---------------------------------------------------------
# Thickness Swelling
# ---------------------------------------------------------

st.write("---")
st.header("3) Thickness Swelling (TS)")

t_before = st.number_input("ความหนาก่อนแช่น้ำ (mm)", 0.0)
t_after = st.number_input("ความหนาหลังแช่น้ำ (mm)", 0.0)

if st.button("คำนวณ TS"):
    if t_before > 0:
        ts = calc_TS(t_before, t_after)
        st.success(f"Thickness Swelling = {ts:.2f} %")
    else:
        st.error("ความหนาก่อนแช่น้ำต้องมากกว่า 0")
