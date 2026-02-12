import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt

# ---------------------------------------------------------
# Calculation Functions
# ---------------------------------------------------------

def calc_MOR(Fmax_N, L, b, t):
    return (3 * Fmax_N * L) / (2 * b * t**2)

def calc_MOE(F1_N, F2_N, y1, y2, L, b, t):
    dF = F2_N - F1_N
    dy = y2 - y1
    return (dF * L**3) / (4 * b * t**3 * dy)

def calc_TS(t_before, t_after):
    return ((t_after - t_before) / t_before) * 100

# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------

st.title("🧪 Particleboard Testing System")
st.subheader("MOR • MOE • Thickness Swelling • Load–Deflection Graph")

st.write("---")
st.header("1) จำนวนตัวอย่างทดสอบ")

num_samples = st.selectbox("เลือกจำนวนตัวอย่าง (1–5 ชิ้น)", [1, 2, 3, 4, 5])

st.write("---")
st.header("2) กรอกข้อมูลการทดสอบดัด (Bending Test)")

L = st.number_input("ระยะห่างแท่นรองรับ L (mm)", 0.0)
b = st.number_input("ความกว้าง b (mm)", 0.0)
t = st.number_input("ความหนา t (mm)", 0.0)

sample_data = []

for i in range(num_samples):
    st.subheader(f"ตัวอย่างที่ {i+1}")

    label_Fmax = f"แรงสูงสุด Fmax (kg) – ตัวอย่าง {i+1}"
    label_F1 = f"แรงจุดที่ 1 F1 (kg) – ตัวอย่าง {i+1}"
    label_F2 = f"แรงจุดที่ 2 F2 (kg) – ตัวอย่าง {i+1}"
    label_y1 = f"การโก่งตัว y1 (mm) – ตัวอย่าง {i+1}"
    label_y2 = f"การโก่งตัว y2 (mm) – ตัวอย่าง {i+1}"

    Fmax_kg = st.number_input(label_Fmax, 0.0)
    F1_kg = st.number_input(label_F1, 0.0)
    F2_kg = st.number_input(label_F2, 0.0)
    y1 = st.number_input(label_y1, 0.0)
    y2 = st.number_input(label_y2, 0.0)

    # Convert kg → Newton
    Fmax_N = Fmax_kg * 9.80665
    F1_N = F1_kg * 9.80665
    F2_N = F2_kg * 9.80665

    sample_data.append((Fmax_N, F1_N, F2_N, y1, y2))

if st.button("คำนวณ MOR & MOE"):
    results = []
    for i, (Fmax_N, F1_N, F2_N, y1, y2) in enumerate(sample_data):
        mor = calc_MOR(Fmax_N, L, b, t)
        moe = calc_MOE(F1_N, F2_N, y1, y2, L, b, t)
        results.append([i+1, mor, moe])

    df_results = pd.DataFrame(results, columns=["Sample", "MOR (MPa)", "MOE (MPa)"])
    st.success("ผลการคำนวณ")
    st.dataframe(df_results)

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

# ---------------------------------------------------------
# Excel Template for Load–Deflection
# ---------------------------------------------------------

st.write("---")
st.header("4) กราฟ Load–Deflection (Upload Excel)")

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

    # Plot graph
    fig, ax = plt.subplots()
    ax.plot(df["Deflection (mm)"], df["Load (N)"], marker="o")
    ax.set_xlabel("Deflection (mm)")
    ax.set_ylabel("Load (N)")
    ax.set_title("Load–Deflection Curve")
    ax.grid(True)

    st.pyplot(fig)
