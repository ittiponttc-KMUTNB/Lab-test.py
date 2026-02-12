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
    Fmax_kg = st.number_input(f"แรงสูงสุด Fmax (kg) – ตัวอย่าง {i+1}", 0.0)
    F1_kg = st.number_input(f"แรงจุดที่ 1 F1 (kg) – ตัวอย่าง {i+1}", 0.0)
    F2_kg = st.number_input(f"แรงจุดที่ 2 F2 (kg) – ตัวอย่าง {i+1}", 0.0)
    y1 = st.number_input(f"การโก่งตัว y1 (mm) – ตัวอย่าง {i+1}", 0.0)
    y2 = st.number_input(f"การโก่งตัว y2 (mm) – ตัวอย่าง
