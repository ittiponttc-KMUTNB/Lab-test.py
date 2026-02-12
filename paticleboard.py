import streamlit as st

# ---------------------------------------------------------
# Functions
# ---------------------------------------------------------

def calc_MOR(Fmax, L, b, t):
    return (3 * Fmax * L) / (2 * b * t**2)

def calc_MOE(F1, F2, y1, y2, L, b, t):
    dF = F2 - F1
    dy = y2 - y1
    return (dF * L**3) / (4 * b * t**3 * dy)

def calc_TS(t_before, t_after):
    return ((t_after - t_before) / t_before) * 100


# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------

st.title("🧪 Particleboard Testing App")
st.subheader("Modulus of Rupture (MOR), Modulus of Elasticity (MOE), Thickness Swelling (TS)")

st.write("---")
st.header("1) MOR – Modulus of Rupture")

Fmax = st.number_input("แรงกดสูงสุด Fmax (N)", 0.0)
L = st.number_input("ระยะห่างแท่นรองรับ L (mm)", 0.0)
b = st.number_input("ความกว้าง b (mm)", 0.0)
t = st.number_input("ความหนา t (mm)", 0.0)

if st.button("คำนวณ MOR"):
    if L > 0 and b > 0 and t > 0:
        mor = calc_MOR(Fmax, L, b, t)
        st.success(f"MOR = {mor:.2f} MPa")
    else:
        st.error("กรุณากรอกค่าที่มากกว่า 0")


st.write("---")
st.header("2) MOE – Modulus of Elasticity")

F1 = st.number_input("แรงจุดที่ 1 F1 (N)", 0.0)
F2 = st.number_input("แรงจุดที่ 2 F2 (N)", 0.0)
y1 = st.number_input("การโก่งตัวจุดที่ 1 y1 (mm)", 0.0)
y2 = st.number_input("การโก่งตัวจุดที่ 2 y2 (mm)", 0.0)

if st.button("คำนวณ MOE"):
    if L > 0 and b > 0 and t > 0 and (y2 - y1) != 0:
        moe = calc_MOE(F1, F2, y1, y2, L, b, t)
        st.success(f"MOE = {moe:.2f} MPa")
    else:
        st.error("ตรวจสอบค่าที่กรอกอีกครั้ง")


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
