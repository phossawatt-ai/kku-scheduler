import streamlit as st
import pandas as pd
import io
import math
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ==========================================
# ⚙️ 1. CONFIGURATION
# ==========================================
class Config:
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    TIME_SLOTS = range(8, 20)
    LUNCH_BREAK_HOUR = 12
    
    # 🔧 ตั้งค่าเพดานจำนวนนักศึกษาต่อห้อง (ถ้าเกินจะแตก Sec)
    MAX_CAPACITY_LAB = 45     # ห้อง Lab ส่วนใหญ่จุได้ประมาณนี้
    MAX_CAPACITY_LEC = 90     # ห้อง Lec ใหญ่สุด (ถ้าเกินให้แตก Sec)

    COLOR_MAP = {
        'SC': 'FFF59D', 'CP': 'B3E5FC', 'LI': 'C8E6C9', 
        'GE': 'FFE0B2', 'EN': 'E1BEE7', 'DEFAULT': 'F5F5F5'
    }

# ==========================================
# 📦 2. DATA STRUCTURES
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
        self.section_idx = int(data.get('section_idx', 1)) # เลข Sec (1, 2, 3...)
        
        # Unique ID
        self.uid = f"{self.major}_{self.year}_{self.program}_{self.code}_{self.type}_S{self.section_idx}_{id(self)}"

    def __repr__(self):
        return f"[{self.major} Y{self.year} {self.program}] {self.code} (Sec {self.section_idx})"

# ==========================================
# 🧠 3. SOLVER ENGINE
# ==========================================
class UniversityScheduler:
    def __init__(self, courses, rooms, fixed_data):
        self.courses = courses
        self.rooms = sorted(rooms, key=lambda x: x['capacity'])
        self.fixed_data = fixed_data
        self.assignments = {}
        self.failed_courses = []
        self.debug_logs = []
        
    def solve(self):
        model = cp_model.CpModel()
        shifts = {} 
        scheduled_vars = []
        
        print(f"🧩 เริ่มจัดตาราง: {len(self.courses)} รายการ (รวม Sections)")

        # --- A. สร้างตัวแปร ---
        for c in self.courses:
            valid_rooms = []
            # กรองห้องที่จุคนพอ
            for r in self.rooms:
                if r['capacity'] >= c.students:
                    # Logic เลือกประเภทห้อง
                    if c.type.lower() == 'lab':
                        if 'lab' in r['type'].lower(): valid_rooms.append(r)
                    else:
                        valid_rooms.append(r) # Lec ลงได้หมด

            if not valid_rooms:
                self.debug_logs.append(f"❌ {c} ({c.students} คน) -> ไม่มีห้องรองรับ (ขนาด/ประเภท)")
                continue

            c_moves = []
            for d in Config.DAYS:
                for h in Config.TIME_SLOTS:
                    if h + c.duration > 20: continue
                    
                    # ห้ามพักเที่ยง
                    if Config.LUNCH_BREAK_HOUR in range(h, h + c.duration): continue
                    
                    # เช็ค Fixed
                    for r in valid_rooms:
                        is_busy = False
                        for t in range(h, h + c.duration):
                            if (d, t, r['name']) in self.fixed_data:
                                is_busy = True; break
                        if is_busy: continue

                        var = model.NewBoolVar(f'shift_{c.uid}_{d}_{h}_{r["name"]}')
                        shifts[(c.uid, d, h, r["name"])] = var
                        c_moves.append(var)
            
            if c_moves:
                is_scheduled = model.NewBoolVar(f'sched_{c.uid}')
                model.Add(sum(c_moves) == is_scheduled)
                scheduled_vars.append(is_scheduled)
            else:
                self.debug_logs.append(f"❌ {c} -> เวลาเต็ม/ชน Fixed")

        # --- B. Constraints ---
        time_map = {}
        
        # 1. Map Fixed
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
                # Note: Sec 1 กับ Sec 2 ของวิชาเดียวกัน เรียนชนกันได้ไหม?
                # ปกติ Sec 1 กับ Sec 2 คือเด็กคนละกลุ่มกัน ดังนั้น "เวลาชนกันได้" (ถ้าคนละห้อง)
                # แต่ "อาจารย์" ห้ามชนกัน
                
                # Logic: ถ้าเป็นเด็กกลุ่มเดียวกัน (เช่น CS ปี 1 Regular) แต่คนละ Sec ของวิชาเดียวกัน?
                # ปกติ CS ปี 1 Regular จะถูกจับแยก Sec ตั้งแต่ต้นทางแล้ว 
                # เพื่อความง่าย: เราจะถือว่า Group นี้ห้ามเรียนซ้อนวิชาอื่น
                grp_key = f"{c.major}_{c.year}_{c.program}"
                
                time_map[key].append({
                    'type': 'var', 'var': var, 'room': r_name, 
                    'grp': grp_key, 'instr': c.instructor,
                    'course_code': c.code, 'sec': c.section_idx
                })

        # Loop Constraints
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
                        if not isinstance(v, int): model.Add(v == 0)
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
            # (ยกเว้น: เรียนวิชาเดียวกันแต่คนละ Sec ถือว่าทำได้ เพราะเด็กแบ่งกลุ่มกันแล้ว)
            # แต่เพื่อความปลอดภัยของระบบ เราจะบังคับไม่ให้ชนกันไปเลย 
            # (ถ้าอยากให้ชนกันได้ต้องมี Logic Student Grouping ที่ซับซ้อนกว่านี้)
            grp_usage = {}
            for item in items:
                grp = item.get('grp')
                if not grp: continue
                
                # ถ้าวิชาเดียวกัน (Code เดียวกัน) แต่นคนละ Sec -> ยอมให้ชนได้ (เพราะเด็กคนละคน)
                # เราจึงสร้าง Key เป็น (Group + CourseCode) เพื่อเช็คการชนเฉพาะวิชาต่างกัน
                # แต่เดี๋ยวก่อน! ถ้าเด็กกลุ่ม CS_1 เรียน Math Sec 1 ตอนเช้า
                # เขาจะเรียน Eng Sec 1 ตอนเช้าพร้อมกันไม่ได้
                # ดังนั้น Key ต้องเป็น Group เฉยๆ ถูกแล้ว
                
                # แต่! ถ้า CS_1 มี 80 คน แบ่งเป็น Sec 1, Sec 2
                # ระบบต้องรู้ว่า Sec 1 กับ Sec 2 คือ Slot ของ Group นี้เหมือนกัน
                # ในเวอร์ชันนี้ เพื่อความ Simple: "ห้าม Group เดียวกันเรียนซ้อนกันทุกกรณี" 
                # ผลคือ: Sec 1 และ Sec 2 ของวิชา A จะไม่ชนกันเอง และไม่ชนวิชา B
                
                if item['type'] == 'var':
                    if grp not in grp_usage: grp_usage[grp] = []
                    grp_usage[grp].append(item['var'])

            for grp, vars in grp_usage.items():
                if len(vars) > 1: model.Add(sum(vars) <= 1)

        # --- C. Objective ---
        # 1. Maximize จำนวนวิชาที่ลงได้
        model.Maximize(sum(scheduled_vars))

        # --- D. Run Solver ---
        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 1  # บังคับใช้ CPU แค่ 1 Core (ประหยัด RAM สูงสุด)
        solver.parameters.max_time_in_seconds = 60.0 # ลดเวลาลงเหลือ 60 วิ (ถ้าเกินให้ตัดจบเลย เอาเท่าที่ได้)
        
        # log การค้นหา (เผื่อไว้ดูใน Terminal)
        solver.parameters.log_search_progress = True
        status = solver.Solve(model)

        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            for (uid, d, h, r_name), var in shifts.items():
                if solver.Value(var) == 1:
                    c = next(x for x in self.courses if x.uid == uid)
                    self.assignments[c] = (d, h, r_name)
            
            # Check Failed
            assigned_uids = set(c.uid for c in self.assignments.keys())
            for c in self.courses:
                if c.uid not in assigned_uids:
                    self.failed_courses.append(c)
            return True
        else:
            return False

# ==========================================
# 📊 4. EXCEL GENERATOR
# ==========================================
def generate_excel(assignments):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    sheet_data = {}
    for c, (d, t, r) in assignments.items():
        key = f"{c.major}_Y{c.year}_{c.program}"
        if key not in sheet_data: sheet_data[key] = []
        sheet_data[key].append((c, d, t, r))
        
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for sheet_name in sorted(sheet_data.keys()):
        safe_name = sheet_name.replace('/', '_')[:31]
        ws = wb.create_sheet(title=safe_name)
        
        ws.merge_cells('A1:N1')
        ws['A1'] = f"ตารางเรียน {sheet_name}"
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = align
        
        ws['A2'] = "Day/Time"
        for i, h in enumerate(Config.TIME_SLOTS):
            col = chr(66 + i)
            ws[f'{col}2'] = f"{h}:00"
            ws.column_dimensions[col].width = 15
            
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
                    
                    # แสดงชื่อวิชา + Section
                    txt = f"{c.code} (Sec {c.section_idx})\n{c.name}\n{r} ({c.instructor})"
                    
                    cell.value = txt
                    cell.alignment = align
                    cell.border = thin
                    color = Config.COLOR_MAP.get(c.code[:2].upper(), 'F5F5F5')
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 🖥️ 5. UI & MAIN LOGIC (AUTO-SPLIT)
# ==========================================
st.set_page_config(page_title="KKU Auto-Scheduler", layout="wide")
st.title("🎓 ระบบจัดตารางเรียน (Auto-Split Sections)")
st.info(f"ระบบจะแบ่ง Section อัตโนมัติเมื่อนักศึกษาเกิน {Config.MAX_CAPACITY_LAB} คน (สำหรับ Lab) หรือ {Config.MAX_CAPACITY_LEC} คน (สำหรับ Lec)")

with st.sidebar:
    st.header("Upload Data")
    f_rooms = st.file_uploader("1. Rooms", type=['csv'])
    f_students = st.file_uploader("2. Students", type=['csv'])
    f_subjects = st.file_uploader("3. Subjects", type=['csv'], accept_multiple_files=True)
    f_fixed = st.file_uploader("4. Fixed", type=['csv'], accept_multiple_files=True)

if st.button("🚀 Run Scheduler", type="primary"):
    if f_rooms and f_students and f_subjects:
        with st.spinner("⏳ กำลังคำนวณและแบ่ง Section..."):
            
            # 1. Load Rooms
            df_rooms = pd.read_csv(f_rooms)
            rooms_list = []
            for _, r in df_rooms.iterrows():
                rooms_list.append({
                    'name': str(r['room_name']),
                    'capacity': int(r['capacity']),
                    'type': str(r['type']).lower()
                })
            
            # 2. Load Students
            df_std = pd.read_csv(f_students)
            std_map = {}
            prog_map = {}
            for _, r in df_std.iterrows():
                k = (r['major'], int(r['year']), r['program'])
                std_map[k] = int(r['student_count'])
                
                pk = (r['major'], int(r['year']))
                if pk not in prog_map: prog_map[pk] = []
                if r['program'] not in prog_map[pk]: prog_map[pk].append(r['program'])

            # 3. Load Subjects & *** AUTO-SPLIT LOGIC ***
            all_courses = []
            for f in f_subjects:
                df = pd.read_csv(f)
                maj = f.name.split('_')[1].split('.')[0]
                for _, row in df.iterrows():
                    yr = int(row['year'])
                    progs = prog_map.get((maj, yr), ['Regular'])
                    
                    for prog in progs:
                        total_students = std_map.get((maj, yr, prog), 50)
                        
                        # Function เพื่อสร้าง Course ตามประเภท
                        def create_courses(c_type, hours):
                            if hours <= 0: return
                            
                            # กำหนดเพดาน
                            limit = Config.MAX_CAPACITY_LAB if c_type == 'Lab' else Config.MAX_CAPACITY_LEC
                            
                            # คำนวณจำนวน Sec
                            num_sections = 1
                            if total_students > limit:
                                num_sections = math.ceil(total_students / limit)
                            
                            # เฉลี่ยจำนวนคนต่อ Sec
                            students_per_sec = math.ceil(total_students / num_sections)
                            
                            for i in range(num_sections):
                                d = row.to_dict()
                                d.update({
                                    'For_Major': maj, 
                                    'program': prog, 
                                    'student_count': students_per_sec, # จำนวนคนหลังแบ่ง
                                    'type': c_type, 
                                    'duration': hours,
                                    'section_idx': i + 1 # Sec 1, 2, 3...
                                })
                                all_courses.append(Course(d))

                        # สร้าง Lec (อาจแบ่ง Sec ถ้าคนเยอะจัด)
                        create_courses('Lec', row['lecture_hours'])
                        
                        # สร้าง Lab (แบ่ง Sec แน่นอนถ้าคน > 45)
                        create_courses('Lab', row['lab_hours'])

            # 4. Load Fixed
            fixed_data = {}
            if f_fixed:
                for f in f_fixed:
                    dfx = pd.read_csv(f)
                    for _, r in dfx.iterrows():
                        try:
                            s_h = int(float(str(r['start_time']).split(':')[0]))
                            e_h = int(float(str(r['end_time']).split(':')[0]))
                            for t in range(s_h, e_h):
                                fixed_data[(str(r['day']), t, str(r['room']))] = 'FIXED'
                        except: pass

            st.write(f"✅ สร้างรายการสอนทั้งหมด (รวม Sec): {len(all_courses)} รายการ")
            
            # 5. Run Solver
            scheduler = UniversityScheduler(all_courses, rooms_list, fixed_data)
            success = scheduler.solve()
            
            if success:
                st.success(f"🎉 จัดเสร็จแล้ว! (สำเร็จ {len(scheduler.assignments)} / {len(all_courses)})")
                excel_file = generate_excel(scheduler.assignments)
                st.download_button("📥 Download Excel", excel_file, "Auto_Schedule.xlsx")
                
                if scheduler.failed_courses:
                    st.warning(f"⚠️ มี {len(scheduler.failed_courses)} Sec ที่ลงไม่ได้:")
                    for c in scheduler.failed_courses:
                        st.write(f"- {c} ({c.students} คน)")
                    with st.expander("ดูสาเหตุ"):
                        for log in scheduler.debug_logs:
                            st.text(log)
            else:
                st.error("Solver Error")

