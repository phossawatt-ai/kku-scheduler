import streamlit as st
# ------------------------------------------------------------------
# ⚡️ PART 0: SYSTEM CONFIG (ต้องอยู่บรรทัดแรกสุด ห้ามย้าย!)
# ------------------------------------------------------------------
st.set_page_config(page_title="KKU Ultimate Scheduler", layout="wide")

import pandas as pd
import io
import math
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ==========================================
# ⚙️ PART 1: CONSTANTS & CONFIGURATION
# ==========================================
class Config:
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    TIME_SLOTS = range(8, 20)  # 08:00 - 20:00
    LUNCH_BREAK_HOUR = 12      # ห้ามเรียนช่วง 12:00 - 13:00
    
    # 🔧 ตั้งค่าเพดานความจุ (ถ้า นศ. เกินค่านี้ ระบบจะแบ่ง Sec ให้อัตโนมัติ)
    MAX_CAPACITY_LAB = 45     # ถ้าเกิน 45 คน -> แบ่ง Sec Lab
    MAX_CAPACITY_LEC = 100    # ถ้าเกิน 100 คน -> แบ่ง Sec Lecture

    # 🎨 สีสำหรับ Excel
    COLOR_MAP = {
        'SC': 'FFF59D', 'CP': 'B3E5FC', 'LI': 'C8E6C9', 
        'GE': 'FFE0B2', 'EN': 'E1BEE7', 'DEFAULT': 'F5F5F5'
    }

# ==========================================
# 📦 PART 2: DATA STRUCTURES
# ==========================================
class Course:
    def __init__(self, data):
        self.code = str(data.get('course_code', 'N/A')).strip()
        self.name = str(data.get('course_name', 'Unknown')).strip()
        self.major = data.get('For_Major', 'Gen')
        self.year = int(data.get('year', 1))
        self.program = data.get('program', 'Regular')
        
        self.type = data.get('type', 'Lec')
        self.duration = int(data.get('duration', 3))
        self.instructor = str(data.get('instructor', 'TBA'))
        
        self.students = int(data.get('student_count', 40))
        self.section_idx = int(data.get('section_idx', 1)) # เลข Sec
        
        # Unique ID สำหรับ Solver (ห้ามซ้ำ)
        self.uid = f"{self.major}_{self.year}_{self.program}_{self.code}_{self.type}_S{self.section_idx}_{id(self)}"

    def __repr__(self):
        return f"[{self.major} Y{self.year} {self.program}] {self.code} Sec.{self.section_idx} ({self.students} คน)"

# ==========================================
# 🧠 PART 3: SOLVER ENGINE (OPTIMIZED)
# ==========================================
class UniversityScheduler:
    def __init__(self, courses, rooms, fixed_data):
        self.courses = courses
        # เรียงห้องจากเล็กไปใหญ่ เพื่อให้เลือกห้องที่พอดีตัวที่สุดก่อน (Save Resource)
        self.rooms = sorted(rooms, key=lambda x: x['capacity'])
        self.fixed_data = fixed_data
        self.assignments = {}
        self.failed_courses = []
        self.debug_logs = []
        
    def solve(self):
        model = cp_model.CpModel()
        shifts = {} 
        scheduled_vars = [] # เก็บตัวแปรเพื่อใช้ Maximize
        
        print(f"🧩 เริ่มประมวลผล: {len(self.courses)} รายการวิชา (รวม Sections)")

        # --- A. สร้างตัวแปร (Variables) ---
        for c in self.courses:
            valid_rooms = []
            
            # 1. กรองห้องที่จุคนพอ และ ประเภทถูกต้อง
            for r in self.rooms:
                if r['capacity'] >= c.students:
                    if c.type.lower() == 'lab':
                        if 'lab' in r['type'].lower(): valid_rooms.append(r)
                    else:
                        valid_rooms.append(r) # Lec ลงได้หมด

            if not valid_rooms:
                self.debug_logs.append(f"❌ {c} -> ไม่มีห้องรองรับ (ต้องการ {c.students} ที่นั่ง)")
                continue
            
            # 🔥 OPTIMIZATION: เลือกห้องที่ดีที่สุดแค่ 5 ห้องพอ (Room Pruning)
            # เพื่อลดจำนวนตัวแปรให้ AI คิดทัน (จากเป็นพันตัวเลือกเหลือแค่ 5)
            valid_rooms = valid_rooms[:5] 

            c_moves = []
            for d in Config.DAYS:
                for h in Config.TIME_SLOTS:
                    if h + c.duration > 20: continue
                    
                    # ห้ามเรียนพักเที่ยง
                    if Config.LUNCH_BREAK_HOUR in range(h, h + c.duration): continue
                    
                    # เช็คตาราง Fixed
                    for r in valid_rooms:
                        is_busy = False
                        for t in range(h, h + c.duration):
                            if (d, t, r['name']) in self.fixed_data:
                                is_busy = True; break
                        if is_busy: continue

                        # สร้าง Boolean Variable
                        var = model.NewBoolVar(f'shift_{c.uid}_{d}_{h}_{r["name"]}')
                        shifts[(c.uid, d, h, r["name"])] = var
                        c_moves.append(var)
            
            # สร้างตัวแปรสถานะว่า "วิชานี้ถูกจัดหรือไม่" (0 หรือ 1)
            if c_moves:
                is_scheduled = model.NewBoolVar(f'sched_{c.uid}')
                model.Add(sum(c_moves) == is_scheduled)
                scheduled_vars.append(is_scheduled)
            else:
                self.debug_logs.append(f"❌ {c} -> เวลาเต็มหมดแล้ว หรือติด Fixed")

        # --- B. สร้างเงื่อนไข (Constraints) ---
        time_map = {}
        
        # 1. Map Fixed Data
        for (d, t, r_name), info in self.fixed_data.items():
            key = (d, t)
            if key not in time_map: time_map[key] = []
            time_map[key].append({'type': 'fixed', 'room': r_name})

        # 2. Map Variables
        for (uid, d, h, r_name), var in shifts.items():
            c = next(x for x in self.courses if x.uid == uid)
            for i in range(c.duration):
                key = (d, h + i)
                if key not in time_map: time_map[key] = []
                
                # Group Key: (Major, Year, Program) -> ห้ามเรียนชนกันเอง
                # หมายความว่า CS Y1 Reg จะเรียน 2 วิชาพร้อมกันไม่ได้ (แม้จะคนละ Sec)
                # เป็นการบังคับให้ตารางเรียนของเด็กกลุ่มนี้ไม่ชนกันแน่นอน
                grp_key = f"{c.major}_{c.year}_{c.program}"
                
                time_map[key].append({
                    'type': 'var', 
                    'var': var, 
                    'room': r_name, 
                    'grp': grp_key, 
                    'instr': c.instructor
                })

        # Loop Check Conflicts
        for slot, items in time_map.items():
            # 2.1 ห้องห้ามซ้อน
            room_usage = {}
            for item in items:
                r = item['room']
                if r not in room_usage: room_usage[r] = []
                if item['type'] == 'var': room_usage[r].append(item['var'])
                else: room_usage[r].append(1) # Fixed

            for r, vars in room_usage.items():
                if any(isinstance(v, int) for v in vars):
                    for v in vars: 
                        if not isinstance(v, int): model.Add(v == 0) # ถ้าห้องติด Fixed, ตัวแปรเราต้องเป็น 0
                elif len(vars) > 1:
                    model.Add(sum(vars) <= 1)

            # 2.2 อาจารย์ห้ามซ้อน (สำคัญมากสำหรับการแยก Sec)
            instr_usage = {}
            for item in items:
                instr = item.get('instr')
                if not instr or instr == 'TBA': continue
                if item['type'] == 'var':
                    if instr not in instr_usage: instr_usage[instr] = []
                    instr_usage[instr].append(item['var'])
            
            for instr, vars in instr_usage.items():
                if len(vars) > 1: model.Add(sum(vars) <= 1)

            # 2.3 นักเรียนกลุ่มเดียวกัน ห้ามเรียนซ้อน
            grp_usage = {}
            for item in items:
                grp = item.get('grp')
                if not grp: continue
                if item['type'] == 'var':
                    if grp not in grp_usage: grp_usage[grp] = []
                    grp_usage[grp].append(item['var'])

            for grp, vars in grp_usage.items():
                if len(vars) > 1: model.Add(sum(vars) <= 1)

        # --- C. Objective: จัดให้ได้เยอะที่สุด (Maximize) ---
        model.Maximize(sum(scheduled_vars))

        # --- D. Run Solver (ปรับจูนค่าตรงนี้เพื่อแก้ Error) ---
        solver = cp_model.CpSolver()
        
        # 🔥 แก้ปัญหา RAM หมด / จอขาว
        solver.parameters.num_search_workers = 1  # ใช้ CPU 1 Core (ช้าหน่อยแต่ไม่กิน RAM จนพัง)
        
        # 🔥 แก้ปัญหา Solver Error / Timeout
        solver.parameters.max_time_in_seconds = 600.0 # ให้เวลา 10 นาที (ถ้าเกิน ให้เอาเท่าที่ได้)
        solver.parameters.log_search_progress = True    # แสดง Log ใน Terminal
        
        status = solver.Solve(model)

        # ถ้าเจอคำตอบ (ไม่ว่าจะดีที่สุดหรือไม่)
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            # เก็บผลลัพธ์
            assigned_uids = set()
            for (uid, d, h, r_name), var in shifts.items():
                if solver.Value(var) == 1:
                    c = next(course for course in self.courses if course.uid == uid)
                    self.assignments[c] = (d, h, r_name)
                    assigned_uids.add(uid)
            
            # เช็ควิชาที่ตกหล่น
            for c in self.courses:
                if c.uid not in assigned_uids:
                    self.failed_courses.append(c)
            return True
        else:
            return False

# ==========================================
# 📊 PART 4: EXCEL GENERATOR
# ==========================================
def generate_excel(assignments):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    # Group ข้อมูลตาม Sheet
    sheet_data = {}
    for c, (d, t, r) in assignments.items():
        key = f"{c.major}_Y{c.year}_{c.program}"
        if key not in sheet_data: sheet_data[key] = []
        sheet_data[key].append((c, d, t, r))
        
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for sheet_name in sorted(sheet_data.keys()):
        safe_name = sheet_name.replace('/', '_')[:30] # ชื่อ Sheet ห้ามยาวเกิน
        ws = wb.create_sheet(title=safe_name)
        
        ws.merge_cells('A1:N1')
        ws['A1'] = f"ตารางเรียน {sheet_name}"
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = align
        
        ws['A2'] = "Day/Time"
        for i, h in enumerate(Config.TIME_SLOTS):
            col = chr(66 + i) # B, C, D...
            ws[f'{col}2'] = f"{h}:00"
            ws.column_dimensions[col].width = 16
            
        row_map = {d: i+3 for i, d in enumerate(Config.DAYS)}
        for d, r_idx in row_map.items():
            ws[f'A{r_idx}'] = d
            ws[f'A{r_idx}'].font = Font(bold=True)
            
        for c, d, t, r in sheet_data[sheet_name]:
            if d not in row_map: continue
            r_idx = row_map[d]
            for i in range(c.duration):
                h_cur = t + i
                if h_cur in Config.TIME_SLOTS:
                    c_idx = Config.TIME_SLOTS.index(h_cur) + 2
                    cell = ws.cell(row=r_idx, column=c_idx)
                    
                    # รายละเอียดในช่อง
                    sec_info = f" (Sec {c.section_idx})" if c.section_idx > 1 else ""
                    txt = f"{c.code}{sec_info}\n{c.name}\n{r} ({c.instructor})"
                    
                    cell.value = txt
                    cell.alignment = align
                    cell.border = thin
                    color = Config.COLOR_MAP.get(c.code[:2].upper(), 'F5F5F5')
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 🖥️ PART 5: USER INTERFACE (AUTO-SPLIT LOGIC)
# ==========================================
st.title("🎓 KKU Ultimate Scheduler")
st.success(f"⚡️ ระบบ Auto-Sectioning เปิดใช้งาน: แบ่ง Sec อัตโนมัติเมื่อ Lab > {Config.MAX_CAPACITY_LAB} คน หรือ Lec > {Config.MAX_CAPACITY_LEC} คน")

with st.sidebar:
    st.header("📂 Upload Files")
    f_rooms = st.file_uploader("1. Rooms (csv)", type=['csv'])
    f_students = st.file_uploader("2. Students (csv)", type=['csv'])
    f_subjects = st.file_uploader("3. Subjects (csv)", type=['csv'], accept_multiple_files=True)
    f_fixed = st.file_uploader("4. Fixed (Optional)", type=['csv'], accept_multiple_files=True)

if st.button("🚀 Start Scheduling", type="primary"):
    if f_rooms and f_students and f_subjects:
        with st.spinner("⏳ กำลังอ่านข้อมูล และคำนวณแบ่ง Section..."):
            
            # --- 1. Load Rooms ---
            df_rooms = pd.read_csv(f_rooms)
            rooms_list = []
            for _, r in df_rooms.iterrows():
                rooms_list.append({
                    'name': str(r['room_name']),
                    'capacity': int(r['capacity']),
                    'type': str(r['type']).lower()
                })
            
            # --- 2. Load Students ---
            df_std = pd.read_csv(f_students)
            std_map = {} # (Major, Year, Program) -> Count
            prog_map = {} # (Major, Year) -> [Reg, Sp]
            
            for _, r in df_std.iterrows():
                try:
                    maj, yr, prog = r['major'], int(r['year']), r['program']
                    cnt = int(r['student_count'])
                    
                    std_map[(maj, yr, prog)] = cnt
                    
                    if (maj, yr) not in prog_map: prog_map[(maj, yr)] = []
                    if prog not in prog_map[(maj, yr)]: prog_map[(maj, yr)].append(prog)
                except: continue

            # --- 3. Load Subjects & AUTO-SPLIT ---
            all_courses = []
            
            for f in f_subjects:
                df = pd.read_csv(f)
                maj = f.name.split('_')[1].split('.')[0] # subjects_CS.csv -> CS
                
                for _, row in df.iterrows():
                    yr = int(row['year'])
                    # หาว่าปีนี้มีภาคอะไรบ้าง (Reg/Sp)
                    available_progs = prog_map.get((maj, yr), ['Regular'])
                    
                    for prog in available_progs:
                        total_students = std_map.get((maj, yr, prog), 50)
                        
                        # ฟังก์ชันย่อยสำหรับสร้างวิชาและแบ่ง Sec
                        def create_course_variants(c_type, hours):
                            if hours <= 0: return
                            
                            limit = Config.MAX_CAPACITY_LAB if c_type == 'Lab' else Config.MAX_CAPACITY_LEC
                            
                            # คำนวณจำนวน Sec
                            num_secs = 1
                            if total_students > limit:
                                num_secs = math.ceil(total_students / limit)
                            
                            # เฉลี่ยคน
                            count_per_sec = math.ceil(total_students / num_secs)
                            
                            for i in range(num_secs):
                                d = row.to_dict()
                                d.update({
                                    'For_Major': maj,
                                    'program': prog,
                                    'student_count': count_per_sec,
                                    'type': c_type,
                                    'duration': hours,
                                    'section_idx': i + 1 # Sec 1, 2...
                                })
                                all_courses.append(Course(d))

                        create_course_variants('Lec', row['lecture_hours'])
                        create_course_variants('Lab', row['lab_hours'])

            # --- 4. Load Fixed ---
            fixed_data = {}
            if f_fixed:
                for f in f_fixed:
                    dfx = pd.read_csv(f)
                    for _, r in dfx.iterrows():
                        try:
                            s = int(float(str(r['start_time']).split(':')[0]))
                            e = int(float(str(r['end_time']).split(':')[0]))
                            for t in range(s, e):
                                fixed_data[(str(r['day']), t, str(r['room']))] = 'FIXED'
                        except: pass

            st.write(f"✅ เตรียมข้อมูลสำเร็จ: {len(all_courses)} รายการ (Sec ย่อยถูกสร้างเรียบร้อย)")

            # --- 5. Run Solver ---
            scheduler = UniversityScheduler(all_courses, rooms_list, fixed_data)
            success = scheduler.solve()
            
            if success:
                st.balloons()
                st.success(f"🎉 จัดตารางเสร็จสิ้น! (ลงได้ {len(scheduler.assignments)} / {len(all_courses)})")
                
                # Excel
                excel_file = generate_excel(scheduler.assignments)
                st.download_button("📥 ดาวน์โหลดตารางเรียน (Excel)", excel_file, "Final_Schedule.xlsx")
                
                # Report Failed
                if scheduler.failed_courses:
                    st.warning(f"⚠️ มี {len(scheduler.failed_courses)} Sec ที่หาลงไม่ได้ (กดดูรายละเอียดด้านล่าง):")
                    with st.expander("🔍 ดูรายชื่อวิชาที่มีปัญหา"):
                        for c in scheduler.failed_courses:
                            st.write(f"- {c} : ต้องการ {c.students} ที่นั่ง")
                    
                    with st.expander("🛠 ดู Debug Log (สำหรับ Admin)"):
                        for log in scheduler.debug_logs:
                            st.text(log)
            else:
                st.error("❌ Solver Timeout: ระบบตัดจบก่อนเพราะใช้เวลานานเกินไป (ลองลดจำนวนวิชา หรือเพิ่มเวลาใน Code)")
    else:
        st.warning("กรุณาอัปโหลดไฟล์ให้ครบ (1-3 จำเป็น)")
