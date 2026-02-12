import math

# ---------------------------------------------------------
# 1) Modulus of Rupture (MOR)
#    f_m = (3 * Fmax * L) / (2 * b * t^2)
# ---------------------------------------------------------
def calc_MOR(Fmax, L, b, t):
    """
    Fmax : แรงกดสูงสุด (N)
    L    : ระยะห่างแท่นรองรับ (mm)
    b    : ความกว้างชิ้นทดสอบ (mm)
    t    : ความหนาชิ้นทดสอบ (mm)
    return: MOR (MPa)
    """
    fm = (3 * Fmax * L) / (2 * b * t**2)  # N/mm^2 = MPa
    return fm


# ---------------------------------------------------------
# 2) Modulus of Elasticity (MOE)
#    MOE = (ΔF * L^3) / (4 * b * t^3 * Δy)
# ---------------------------------------------------------
def calc_MOE(F1, F2, y1, y2, L, b, t):
    """
    F1, F2 : แรงในช่วงเชิงเส้น (N)
    y1, y2 : การโก่งตัวที่สอดคล้องกับ F1, F2 (mm)
    L      : ระยะห่างแท่นรองรับ (mm)
    b      : ความกว้าง (mm)
    t      : ความหนา (mm)
    return: MOE (MPa)
    """
    dF = F2 - F1
    dy = y2 - y1
    MOE = (dF * L**3) / (4 * b * t**3 * dy)  # N/mm^2 = MPa
    return MOE


# ---------------------------------------------------------
# 3) Thickness Swelling (TS)
#    TS = ((t2 - t1) / t1) * 100
# ---------------------------------------------------------
def calc_TS(t_before, t_after):
    """
    t_before : ความหนาก่อนแช่น้ำ (mm)
    t_after  : ความหนาหลังแช่น้ำ (mm)
    return: TS (%)
    """
    TS = ((t_after - t_before) / t_before) * 100
    return TS


# ---------------------------------------------------------
# 4) ตัวอย่างการใช้งาน (ใส่ค่าจริงจาก Lab)
# ---------------------------------------------------------
if __name__ == "__main__":
    # ===== ตัวอย่างชิ้นทดสอบดัด =====
    Fmax = 1500   # N
    L = 300       # mm
    b = 50        # mm
    t = 12        # mm

    mor = calc_MOR(Fmax, L, b, t)
    print(f"MOR = {mor:.2f} MPa")

    # สมมติเลือกจุดเชิงเส้นสองจุดจากกราฟ Load-Deflection
    F1, F2 = 300, 600   # N
    y1, y2 = 1.0, 2.0   # mm

    moe = calc_MOE(F1, F2, y1, y2, L, b, t)
    print(f"MOE = {moe:.2f} MPa")

    # ===== ตัวอย่างการบวมตัว =====
    t_before = 12.0   # mm ก่อนแช่น้ำ
    t_after = 12.9    # mm หลังแช่น้ำ 24 ชม.

    ts = calc_TS(t_before, t_after)
    print(f"Thickness Swelling = {ts:.2f} %")
