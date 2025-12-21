import streamlit as st
import pandas as pd
import random
from ortools.sat.python import cp_model

st.set_page_config(page_title="System Check", layout="centered")

st.title("🛠️ System Diagnostic Check")
st.write("กดปุ่มด้านล่างเพื่อทดสอบระบบ (ไม่ต้องอัปโหลดไฟล์)")

if st.button("🔴 เริ่มการทดสอบ (Run Test)", type="primary"):
    st.write("1. ✅ กำลังโหลด Library... (ผ่าน)")
    
    try:
        # 1. จำลองข้อมูล (Mock Data)
        rooms = [{'name': 'R101', 'cap': 50}, {'name': 'R102', 'cap': 100}]
        courses = [
            {'name': 'Math', 'students': 40, 'duration': 2},
            {'name': 'Physics', 'students': 80, 'duration': 3},
            {'name': 'Chem', 'students': 30, 'duration': 2}
        ]
        st.write("2. ✅ สร้างข้อมูลจำลอง... (ผ่าน)")

        # 2. เรียก Solver
        model = cp_model.CpModel()
        solver = cp_model.CpSolver()
        
        # ตัวแปรโง่ๆ เพื่อเช็คว่า Solver ทำงานไหม
        x = model.NewIntVar(0, 10, 'x')
        model.Add(x >= 5)
        
        st.write("3. ⏳ กำลังเรียก OR-Tools Solver...")
        status = solver.Solve(model)
        
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            st.success(f"4. 🎉 สำเร็จ! Solver ทำงานปกติ (Value: {solver.Value(x)})")
            st.info("สรุป: ระบบของคุณทำงานได้ปกติ ปัญหาอยู่ที่ไฟล์ CSV หรือข้อมูลนำเข้า")
        else:
            st.error("4. ❌ Solver ทำงานแต่หาคำตอบไม่ได้")

    except Exception as e:
        st.error(f"❌ ระบบพัง (Crash): {e}")
        st.write("คำแนะนำ: กรุณาเช็คไฟล์ requirements.txt อีกครั้ง")
