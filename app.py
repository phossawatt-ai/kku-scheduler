import streamlit as st
import pandas as pd
import io
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ==========================================
# ⚙️ 1. CONFIGURATION & CONSTANTS
# ==========================================
class Config:
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    TIME_SLOTS = range(8, 20)  # 08:00 - 20:00
    LUNCH_BREAK_HOUR = 12      # ห้ามเรียนช่วง 12:00 - 13:00
    
    # สีสำหรับ Excel (แยกตามรหัสวิชา)
    COLOR_MAP = {
        'SC': 'FFF59D', # เหลืองอ่อน
        'CP': 'B3E5FC', # ฟ้าอ่อน
        'LI': 'C8E6C9', # เขียวอ่อน
        'GE': 'FFE0B2', # ส้มอ่อน
        'EN': 'E1BEE7', # ม่วงอ่อน
        'DEFAULT': 'F5F5F5' # เทา
    }

# ==========================================
# 📦 2. DATA STRUCTURES (CLASS)
# ==========================================
class Course:
    def __init__(self, data):
        self.code = str(data.get('course_code', 'N/A')).strip()
        self.name = str(data.get('course_name', 'Unknown')).strip()
        self.major = data.get('For_Major', 'Gen')   # สาขาเจ้าของวิชา
        self.year = int(data.get('year', 1))        # ชั้นปี
        self.program = data.get('program', 'Regular') # Regular / Special
        
        self.type = data.get('type', 'Lec')         # Lec / Lab
        self.duration = int(data.get('duration', 3))
        self.instructor = str(data.get('instructor', 'TBA'))
        
        self.students = int(data.get('student_count', 40)) # จำนวนคนเรียน
        
        # ID ที่ไม่ซ้ำกัน (ใช้เป็น Key ใน Solver)
        self.uid = f"{self.major}_{self.year}_{self.program}_{self.code}_{self.type}_{id(self)}"
        
        # ข้อมูล Fixed (ถ้ามี)
        self.is_fixed = False
        self.fix_day = None
        self.fix_time = None
        self.fix_room = None

    def __repr__(self):
        return f"[{self.major} Y{self.year} {self.program}] {self.code} ({self.type})"

# ==========================================
# 🧠 3. SOLVER ENGINE (OR-TOOLS)
# ==========================================
class UniversityScheduler:
    def __init__(self, courses, rooms, fixed_data):
        self.courses = courses
        self.rooms = sorted(rooms, key=lambda x: x['capacity']) # เรียงห้องเล็ก -> ใหญ่
        self.fixed_data = fixed_data # Map: (day, time, room) -> details
        self.assignments = {}
        self.failed_courses = []
        
    def solve(self):
        model = cp_model.CpModel()
        shifts = {} # เก็บตัวแปร Boolean
        
        # กรองวิชาที่จะนำมาจัดตาราง (ถ้า Fixed มาแล้วถือว่าจบ)
        courses_to_solve = self.courses
        
        print(f"🧩 เริ่มประมวลผล: {len(courses_to_solve)} รายการวิชา")

        # --- A. สร้างตัวแปร (Variables) ---
        for c in courses_to_solve:
            
            # กรองห้องที่ "จุคนพอ" และ "ประเภทถูก"
            # (อนุโลม: ห้อง Lec ใช้ห้อง Lab ได้ไหม? ปกติไม่ได้ แต่ Lab ใช้ Lec สอนบรรยายได้)
            valid_rooms = []
            for r in self.rooms:
                if r['capacity'] >= c.students:
                    # Logic เลือกประเภทห้อง
                    if c.type.lower() == 'lab':
                        if 'lab' in r['type'].lower(): valid_rooms.append(r)
                    else:
                        # Lecture ลงได้ทั้งห้อง Lec และห้อง Lab (ถ้าจำเป็น)
                        valid_rooms.append(r)
            
            if not valid_rooms:
                print(f"⚠️ Warning: วิชา {c} ({c.students} คน) ไม่มีห้องรองรับเลย")
                continue

            for d in Config.DAYS:
                for h in Config.TIME_SLOTS:
                    # 1. เช็คเวลาจบไม่เกิน 20:00
                    if h + c.duration > 20: continue
                    
                    # 2. 🥗 เช็คพักเที่ยง (ห้ามเรียนคาบ 12:00-13:00)
                    # ช่วงเรียนคือ [h, h+1, ..., h+duration-1]
                    # ถ้า 12 อยู่ในช่วงนี้ ถือว่าผิด
                    learning_hours = range(h, h + c.duration)
                    if Config.LUNCH_BREAK_HOUR in learning_hours: continue
                    
                    for r in valid_rooms:
                        # 3. เช็คตาราง Fixed (ห้องต้องว่าง)
                        is_room_busy_by_fixed = False
                        for t in learning_hours:
                            if (d, t, r['name']) in self.fixed_data:
                                is_room_busy_by_fixed = True
                                break
                        if is_room_busy_by_fixed: continue

                        # สร้างตัวแปร
                        var_name = f'shift_{c.uid}_{d}_{h}_{r["name"]}'
                        shifts[(c.uid, d, h, r['name'])] = model.NewBoolVar(var_name)

        # --- B. สร้างเงื่อนไข (Constraints) ---
        
        # C1: วิชาหนึ่งต้องลง 1 ครั้งเท่านั้น
        for c in courses_to_solve:
            c_shifts = [shifts[key] for key in shifts if key[0] == c.uid]
            if c_shifts:
                model.Add(sum(c_shifts) == 1)
            else:
                self.failed_courses.append(c) # จัดไม่ได้เพราะไม่มี Slot ลง

        # เตรียม Map เช็คการชน
        # Key: (Day, Hour) -> List of usage
        time_map = {}
        
        # ใส่ Fixed Data ลง Map
        for (d, t, r_name), info in self.fixed_data.items():
            key = (d, t)
            if key not in time_map: time_map[key] = []
            # Fixed ถือว่าเป็น Resource ที่ถูกใช้ไปแล้ว
            time_map[key].append({'type': 'fixed', 'room': r_name, 'grp': 'FIXED'})

        # ใส่ Variable Data ลง Map
        for (uid, d, h, r_name), var in shifts.items():
            c = next(course for course in courses_to_solve if course.uid == uid)
            for i in range(c.duration):
                key = (d, h + i)
                if key not in time_map: time_map[key] = []
                
                # Composite Key สำหรับเช็คกลุ่มเรียนชนกัน
                # เช่น CS ปี 1 ภาคปกติ (CS_1_Regular) ห้ามเรียนซ้อน
                group_key = f"{c.major}_{c.year}_{c.program}"
                
                time_map[key].append({
                    'type': 'var', 
                    'var': var, 
                    'room': r_name, 
                    'grp': group_key, 
                    'instr': c.instructor
                })

        # C2: Loop เช็ค Conflict ในแต่ละ Slot เวลา
        for slot, items in time_map.items():
            
            # 2.1 ห้ามห้องซ้อน (Room Conflict)
            room_usage = {}
            for item in items:
                r = item['room']
                if r not in room_usage: room_usage[r] = []
                if item['type'] == 'var': room_usage[r].append(item['var'])
                else: room_usage[r].append(1) # Fixed

            for r, vars in room_usage.items():
                # ถ้ามี Fixed อยู่แล้ว (1) ตัวแปรอื่นต้องเป็น 0
                if any(isinstance(v, int) for v in vars):
                    for v in vars:
                        if not isinstance(v, int): model.Add(v == 0)
                else:
                    # ถ้าไม่มี Fixed เลือกได้แค่ 1 วิชา
                    if len(vars) > 1: model.Add(sum(vars) <= 1)

            # 2.2 ห้ามกลุ่มเรียนซ้อน (Student Group Conflict)
            # 2.3 ห้ามอาจารย์ซ้อน (Instructor Conflict)
            for check_key in ['grp', 'instr']:
                res_usage = {}
                for item in items:
                    val = item.get(check_key)
                    if not val or val == 'TBA' or val == 'FIXED': continue
                    
                    if val not in res_usage: res_usage[val] = []
                    if item['type'] == 'var': res_usage[val].append(item['var'])
                
                for res, vars in res_usage.items():
                    if len(vars) > 1: model.Add(sum(vars) <= 1)

        # --- C. เป้าหมาย (Objectives) ---
        # พยายามไม่ให้เลิกดึก (หลัง 17:00)
        penalties = []
        for (uid, d, h, r_name), var in shifts.items():
            if h >= 17:
                penalties.append(var * 10) 
        if penalties:
            model.Minimize(sum(penalties))

        # --- D. Run Solver ---
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 120.0 # ให้เวลาคิด 2 นาที
        status = solver.Solve(model)

        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            for (uid, d, h, r_name), var in shifts.items():
                if solver.Value(var) == 1:
                    c = next(course for course in courses_to_solve if course.uid == uid)
                    self.assignments[c] = (d, h, r_name)
            return True
        else:
            return False

# ==========================================
# 📊 4. EXCEL GENERATOR
# ==========================================
def generate_excel(assignments):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # ลบ Sheet default
    
    # Group ข้อมูลตาม Sheet: Major_Year_Program
    # ตัวอย่าง Key: "CS_Y1_Regular"
    sheet_data = {}
    
    for c, (d, t, r) in assignments.items():
        key = f"{c.major}_Y{c.year}_{c.program}"
        if key not in sheet_data: sheet_data[key] = []
        sheet_data[key].append((c, d, t, r))
        
    # Style
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Loop สร้าง Sheet
    for sheet_name in sorted(sheet_data.keys()):
        ws = wb.create_sheet(title=sheet_name[:31]) # Excel จำกัดชื่อ Sheet 31 ตัวอักษร
        
        # Header
        ws.merge_cells('A1:N1')
        ws['A1'] = f"ตารางเรียน {sheet_name}"
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = center_align
        
        # Table Header (Times)
        ws['A2'] = "Day/Time"
        for i, h in enumerate(Config.TIME_SLOTS):
            col_char = chr(66 + i) # B, C, D...
            ws[f'{col_char}2'] = f"{h}:00"
            ws.column_dimensions[col_char].width = 15
            
        # Fill Days
        row_map = {d: i+3 for i, d in enumerate(Config.DAYS)}
        for d, r_idx in row_map.items():
            ws[f'A{r_idx}'] = d
            ws[f'A{r_idx}'].font = Font(bold=True)
            
        # Fill Data
        for c, d, t, r in sheet_data[sheet_name]:
            if d not in row_map: continue
            row_idx = row_map[d]
            
            # Fill ตามจำนวนชั่วโมง
            for i in range(c.duration):
                h_current = t + i
                if h_current in Config.TIME_SLOTS:
                    col_idx = Config.TIME_SLOTS.index(h_current) + 2 # +2 เพราะเริ่ม col B
                    cell = ws.cell(row=row_idx, column=col_idx)
                    
                    # ข้อมูลในช่อง
                    txt = f"{c.code}\n{c.name}\n{r} ({c.instructor})"
                    cell.value = txt
                    cell.alignment = center_align
                    cell.border = thin_border
                    
                    # สีพื้นหลัง
                    color_code = Config.COLOR_MAP.get(c.code[:2].upper(), Config.COLOR_MAP['DEFAULT'])
                    cell.fill = PatternFill(start_color=color_code, end_color=color_code, fill_type='solid')

    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 🖥️ 5. USER INTERFACE (STREAMLIT)
# ==========================================
st.set_page_config(page_title="KKU Scheduler Pro", layout="wide")

st.title("🎓 ระบบจัดตารางเรียนอัตโนมัติ (แยกภาค/ชั้นปี)")
st.info("ระบบจะอ่านไฟล์นักศึกษา เพื่อแยกตารางเรียนภาคปกติ (Regular) และภาคพิเศษ (Special) ให้โดยอัตโนมัติ")

with st.sidebar:
    st.header("📂 1. Upload Data")
    f_rooms = st.file_uploader("1. ไฟล์ห้องเรียน (rooms_data_flat.csv)", type=['csv'])
    f_students = st.file_uploader("2. ไฟล์นักศึกษา (students_cleaned.csv)", type=['csv'])
    f_subjects = st.file_uploader("3. ไฟล์รายวิชา (subjects_XX.csv)", type=['csv'], accept_multiple_files=True)
    f_fixed = st.file_uploader("4. ไฟล์ Fixed (Optional)", type=['csv'], accept_multiple_files=True)

if st.button("🚀 เริ่มจัดตาราง (Start Scheduling)", type="primary"):
    if f_rooms and f_students and f_subjects:
        with st.spinner("⏳ กำลังอ่านข้อมูลและประมวลผล..."):
            
            # --- STEP 1: LOAD ROOMS ---
            df_rooms = pd.read_csv(f_rooms)
            rooms_list = []
            for _, r in df_rooms.iterrows():
                rooms_list.append({
                    'name': str(r['room_name']),
                    'capacity': int(r['capacity']),
                    'type': str(r['type']).lower()
                })
            
            # --- STEP 2: LOAD STUDENTS & PREPARE GROUPS ---
            # Map: (Major, Year, Program) -> Student Count
            df_students = pd.read_csv(f_students)
            student_map = {}
            programs_map = {} # (Major, Year) -> [Regular, Special]
            
            for _, row in df_students.iterrows():
                maj = row['major']
                yr = int(row['year'])
                prog = row['program'] # 'Regular', 'Special'
                cnt = int(row['student_count'])
                
                if cnt > 0:
                    student_map[(maj, yr, prog)] = cnt
                    
                    if (maj, yr) not in programs_map: programs_map[(maj, yr)] = []
                    if prog not in programs_map[(maj, yr)]: programs_map[(maj, yr)].append(prog)
            
            # --- STEP 3: LOAD SUBJECTS & EXPLODE ---
            all_courses = []
            for f in f_subjects:
                df_sub = pd.read_csv(f)
                # ดึงชื่อสาขาจากชื่อไฟล์ (subjects_CS.csv -> CS)
                major_file = f.name.split('_')[1].split('.')[0]
                
                for _, row in df_sub.iterrows():
                    yr = int(row['year'])
                    
                    # หากลุ่มเรียนที่มีอยู่จริงของสาขานี้ ปีนี้
                    available_programs = programs_map.get((major_file, yr), ['Regular']) # Default Regular if not found
                    
                    # *** Loop สร้างวิชาตามจำนวนภาคที่มี ***
                    for prog in available_programs:
                        # ดึงจำนวนนักศึกษา
                        count = student_map.get((major_file, yr, prog), 50)
                        
                        # สร้าง Lec
                        if row['lecture_hours'] > 0:
                            data = row.to_dict()
                            data['For_Major'] = major_file
                            data['program'] = prog # <--- Assign Program
                            data['student_count'] = count
                            data['type'] = 'Lec'
                            data['duration'] = row['lecture_hours']
                            all_courses.append(Course(data))
                        
                        # สร้าง Lab
                        if row['lab_hours'] > 0:
                            data = row.to_dict()
                            data['For_Major'] = major_file
                            data['program'] = prog # <--- Assign Program
                            data['student_count'] = count
                            data['type'] = 'Lab'
                            data['duration'] = row['lab_hours']
                            all_courses.append(Course(data))
            
            # --- STEP 4: LOAD FIXED SCHEDULE ---
            # Map: (day, time, room_name) -> Info
            fixed_data = {}
            if f_fixed:
                for f in f_fixed:
                    df_fix = pd.read_csv(f)
                    for _, row in df_fix.iterrows():
                        try:
                            # Parse Time (09:00 -> 9)
                            start_str = str(row['start_time']).split(':')[0]
                            end_str = str(row['end_time']).split(':')[0]
                            start_h = int(float(start_str))
                            end_h = int(float(end_str))
                            
                            r_name = str(row['room'])
                            day_val = str(row['day']) # Mon, Tue...
                            
                            for t in range(start_h, end_h):
                                fixed_data[(day_val, t, r_name)] = row['course_code']
                                
                                # Note: ถ้าจะให้ดีต้องไป Mark ที่ตัว Course Object ด้วยว่า Fixed
                                # แต่ใน version นี้เอาแค่บล็อคห้องไม่ให้คนอื่นใช้ก่อน
                        except Exception as e:
                            pass # ข้าม row ที่ error

            st.write(f"✅ เตรียมข้อมูลเสร็จสิ้น: {len(all_courses)} วิชาที่ต้องจัดตาราง")

            # --- STEP 5: RUN SOLVER ---
            scheduler = UniversityScheduler(all_courses, rooms_list, fixed_data)
            success = scheduler.solve()
            
            if success:
                st.success(f"🎉 จัดตารางสำเร็จ! ({len(scheduler.assignments)} วิชา)")
                
                # --- STEP 6: EXPORT EXCEL ---
                excel_file = generate_excel(scheduler.assignments)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel (แยก Sheet ตามภาค)",
                    data=excel_file,
                    file_name="University_Timetable_Complete.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Show Failed
                if scheduler.failed_courses:
                    st.error(f"⚠️ มี {len(scheduler.failed_courses)} วิชาที่ลงไม่ได้ (ห้องเต็ม/เวลาเต็ม):")
                    for c in scheduler.failed_courses:
                        st.write(f"- {c} ({c.students} คน)")
            else:
                st.error("❌ ไม่สามารถจัดตารางได้ (เงื่อนไขแน่นเกินไป หรือห้องไม่พอ)")

    else:
        st.warning("กรุณาอัปโหลดไฟล์ 1, 2, 3 ให้ครบถ้วน")
