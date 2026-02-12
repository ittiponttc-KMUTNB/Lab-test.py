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
# MOR Section
# ---------------------------------------------------------

st.write("---")
st.header("1) คำนวณ MOR (จากค่าที่กรอก)")

L = st.number_input("ระยะห่างแท่นรองรับ L (mm)", 0.0)
b = st.number_input("ความกว้าง b (mm)", 0.0)
t = st.number_input("ความหนา t (mm)", 0.0)

Fmax_kg = st.number_input("แรงสูงสุด Fmax (kg)", 0.0)
Fmax_N = Fmax_kg * 9.80665

if st.button("คำนวณ MOR"):
    mor = calc_MOR(Fmax_N, L, b, t)
    st.success(f"MOR = {mor:.2f} MPa")

# ---------------------------------------------------------
# Excel Upload for MOE + Graph
# ---------------------------------------------------------

st.write("---")
st.header("2) Upload Excel เพื่อคำนวณ MOE และสร้างกราฟ")

st.info("Template Excel จะมีคอลัมน์: Load (kg), Deflection (mm)")

# Create template
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

    # Convert Load kg → N
    df["Load (N)"] = df["Load (kg)"] * 9.80665

    # ymax สำหรับคำนวณ MOE (ใช้ deflection สูงสุด)
    ymax = df["Deflection (mm)"].max()

    # คำนวณ MOE
    moe = calc_MOE_from_excel(Fmax_N, ymax, L, b, t)

    if moe is None:
        st.error("ไม่สามารถคำนวณ MOE ได้ (ymax หรือค่าบางตัวเป็นศูนย์)")
    else:
        st.success(f"MOE (จาก Excel) = {moe:.2f} MPa")

    # ---------------------------------------------------------
    # Plot Graph (Line Chart แบบ Excel)
    # ---------------------------------------------------------

    fig, ax = plt.subplots()

    ax.plot(
        df["Deflection (mm)"],
        df["Load (N)"],
        color="red",
        linewidth=2,
        label="Load–Deflection Curve"
    )

    ax.set_xlabel("Deflection (mm)")
    ax.set_ylabel("Load (N)")
    ax.set_title("Load–Deflection Curve")
    ax.grid(True)
    ax.legend()

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
