import streamlit as st
import pandas as pd
import io
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ==========================================
# ⚙️ CONFIGURATION
# ==========================================
class Config:
    COLOR_MAP = {
        'SC': 'FFF59D', 'CP': 'B3E5FC', 'LI': 'C8E6C9',
        'EN': 'C8E6C9', 'GE': 'FFE0B2', 'DEFAULT': 'F5F5F5'
    }
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'] # ตัดเสาร์อาทิตย์ออกตามโจทย์
    TIME_SLOTS = range(8, 20) # 08:00 - 20:00

# ==========================================
# 📦 DATA MODELS (ปรับให้เข้ากับ CSV ของเรา)
# ==========================================
class Course:
    def __init__(self, data):
        self.code = data['course_code']
        self.name = data['course_name']
        self.major = data.get('For_Major', 'Gen') # CS, AI, etc.
        self.year = int(data.get('year', 1))
        
        # Lec/Lab Logic
        # ถ้า lecture_hours > 0 ให้สร้างเป็น Lec
        # (ในโค้ดนี้เราจะแยก Lec/Lab ตั้งแต่ตอนโหลด)
        self.type = data['type'] # 'Lec' or 'Lab'
        self.duration = int(data['duration'])
        
        # Teacher & Student
        self.instructor = str(data.get('instructor', 'TBA'))
        self.student_group = f"{self.major}_Y{self.year}" # e.g. CS_Y1
        self.students = 50 # Default ไปก่อนถ้าไม่มีข้อมูล
        
        # Fixed Data
        self.is_fixed = False
        self.fixed_day = None
        self.fixed_time = None
        self.fixed_room = None
        
        # Unique ID
        self.uid = f"{self.code}_{self.major}_{self.type}_{id(self)}"

    def __repr__(self):
        return f"[{self.major}] {self.code} ({self.type})"

# ==========================================
# 🧠 OR-TOOLS ENGINE (คงเดิมแต่ปรับปรุง Input)
# ==========================================
class UniversityScheduler:
    def __init__(self, courses_list, rooms_df, fixed_data):
        self.rooms = sorted(rooms_df.to_dict('records'), key=lambda x: x['capacity'])
        self.courses = courses_list
        self.fixed_schedule = fixed_data # format: {(day, time, room): course_code}
        self.assignment = {}
        self.failed_courses = []

    def solve(self):
        model = cp_model.CpModel()
        shifts = {} 
        
        # 1. แบ่งวิชาเป็น Fixed กับ Variable
        courses_to_solve = []
        
        for c in self.courses:
            # เช็คว่าวิชานี้โดน Fix หรือไม่
            # (Logic นี้อาจต้องปรับถ้า Fixed File ระบุละเอียดกว่านี้)
            is_fixed_match = False
            if c.is_fixed:
                 self.assignment[c] = (c.fixed_day, c.fixed_time, c.fixed_room)
                 is_fixed_match = True
            
            if not is_fixed_match:
                courses_to_solve.append(c)

        # 2. สร้างตัวแปร (Variables)
        print(f"🧩 กำลังจัดตารางให้: {len(courses_to_solve)} วิชา")
        
        for c in courses_to_solve:
            # กรองห้องที่ประเภทตรงกันและจุคนพอ
            valid_rooms = [r for r in self.rooms if r['capacity'] >= c.students]
            pref_rooms = [r for r in valid_rooms if r['type'].lower() == c.type.lower()]
            if not pref_rooms: pref_rooms = valid_rooms # ยืดหยุ่นถ้าห้อง Lab เต็มให้ลง Lec ได้ไหม? (อาจต้องแก้ตรงนี้)

            for d in Config.DAYS:
                for h in Config.TIME_SLOTS:
                    if h + c.duration > 20: continue
                    
                    # 🥗 เงื่อนไข: ห้ามเรียนตอนพักเที่ยง (12:00-13:00)
                    # ถ้าช่วงเวลาเรียน ไปทับ 12:00 (คือเลข 12) ให้ข้าม
                    # ช่วงเรียนคือ range(h, h + c.duration)
                    if 12 in range(h, h + c.duration): continue 
                    
                    for r in pref_rooms:
                        # เช็ค Fixed Schedule (ห้องนี้เวลานี้ว่างไหม)
                        is_busy = False
                        for t in range(h, h + c.duration):
                             if (d, t, r['room_name']) in self.fixed_schedule:
                                 is_busy = True; break
                        if is_busy: continue

                        shifts[(c.uid, d, h, r['room_name'])] = model.NewBoolVar(f'shift_{c.uid}_{d}_{h}_{r["room_name"]}')

        # 3. Constraints
        # C1: วิชาหนึ่งเรียนแค่ครั้งเดียว
        for c in courses_to_solve:
            c_shifts = [shifts[key] for key in shifts if key[0] == c.uid]
            if c_shifts:
                model.Add(sum(c_shifts) == 1)
            else:
                self.failed_courses.append(c)

        # Map เพื่อเช็คการชน
        # Key: (Day, Hour) -> List of {type, room, teacher, group, var}
        time_slot_map = {} 

        # ใส่ Fixed Data ลงใน Map
        for (d, t, r_name), code in self.fixed_schedule.items():
            key = (d, t)
            if key not in time_slot_map: time_slot_map[key] = []
            time_slot_map[key].append({'type': 'fixed', 'room': r_name, 'instr': 'FIXED', 'grp': 'FIXED'})

        # ใส่ Variable Data
        for (uid, d, h, r_name), var in shifts.items():
            c = next(x for x in courses_to_solve if x.uid == uid)
            for i in range(c.duration):
                key = (d, h + i)
                if key not in time_slot_map: time_slot_map[key] = []
                time_slot_map[key].append({'type': 'var', 'var': var, 'room': r_name, 'instr': c.instructor, 'grp': c.student_group})

        # C2: Conflict Checks
        for slot, items in time_slot_map.items():
            # 2.1 Room Conflict (ห้ามห้องซ้อน)
            rooms_map = {}
            for item in items:
                r = item['room']
                if r not in rooms_map: rooms_map[r] = []
                if item['type'] == 'var': rooms_map[r].append(item['var'])
                else: rooms_map[r].append(1) # Fixed occupied
            
            for r, vars_list in rooms_map.items():
                has_fixed = any(isinstance(v, int) for v in vars_list)
                if has_fixed:
                    # ถ้ามี Fixed อยู่แล้ว ตัวแปรอื่นต้องเป็น 0 ทั้งหมด
                    for v in vars_list:
                        if not isinstance(v, int): model.Add(v == 0)
                else:
                    # ถ้าไม่มี Fixed ให้เลือกได้แค่ 1 ตัวแปร
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)

            # 2.2 Teacher Conflict & Student Group Conflict
            # (ใช้ Logic เดียวกัน: Group by Resource -> Sum <= 1)
            for key_type in ['instr', 'grp']:
                res_map = {}
                for item in items:
                    val = item[key_type]
                    if val == 'FIXED' or val == 'TBA' or val == '-': continue
                    if val not in res_map: res_map[val] = []
                    if item['type'] == 'var': res_map[val].append(item['var'])
                
                for res, vars_list in res_map.items():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)

        # 4. Objectives (Minimize Late Classes)
        penalties = []
        for (uid, d, h, r_name), var in shifts.items():
            if h >= 17: penalties.append(var * 10) # เลิกเย็นโดนหักคะแนน
        
        if penalties: model.Minimize(sum(penalties))

        # 5. Run Solver
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 60.0
        status = solver.Solve(model)

        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            for (uid, d, h, r_name), var in shifts.items():
                if solver.Value(var) == 1:
                    c = next(x for x in courses_to_solve if x.uid == uid)
                    self.assignment[c] = (d, h, r_name)
            return True
        else:
            return False

# ==========================================
# 📊 EXPORT EXCEL (ของเดิม - ดีอยู่แล้ว)
# ==========================================
def generate_excel_report(sched):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    # รวมข้อมูล Fixed และ Assigned
    # (ในตัวอย่างนี้ขอทำแบบง่ายเฉพาะที่ Assign ได้ก่อน)
    
    majors = set(c.major for c in sched.assignment.keys())
    
    thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    
    for major in sorted(list(majors)):
        sheet_name = f"{major}"
        data = [(c, d, t, r) for c, (d, t, r) in sched.assignment.items() if c.major == major]
        
        # สร้าง DataFrame ตารางเรียน
        df_disp = pd.DataFrame('', index=Config.DAYS, columns=[f"{h}:00" for h in Config.TIME_SLOTS])
        colors = {}
        
        for c, d, t, r in data:
            bg = Config.COLOR_MAP.get(c.code[:2].upper(), Config.COLOR_MAP['DEFAULT'])
            txt = f"{c.code}\n{c.name}\n{r}"
            
            for i in range(c.duration):
                h_col = t + i
                if 8 <= h_col < 20:
                    col_name = f"{h_col}:00"
                    if d in df_disp.index and col_name in df_disp.columns:
                        df_disp.at[d, col_name] = txt
                        colors[(Config.DAYS.index(d)+2, h_col - 8 + 2)] = bg # Row, Col offset

        df_disp.to_excel(writer, sheet_name=sheet_name)
        
        # จัด Format Excel
        ws = writer.sheets[sheet_name]
        for r in range(2, 2 + len(Config.DAYS)):
            for c in range(2, 2 + len(Config.TIME_SLOTS)):
                cell = ws.cell(row=r, column=c)
                cell.border = thin
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                
                if (r, c) in colors:
                    cell.fill = PatternFill(start_color=colors[(r,c)], end_color=colors[(r,c)], fill_type='solid')

    writer.close()
    output.seek(0)
    return output

# ==========================================
# 🖥️ APP INTERFACE (ปรับใหม่รับ CSV)
# ==========================================
st.set_page_config(page_title="KKU Scheduler (Hybrid)", layout="wide")
st.title("🎓 ระบบจัดตารางเรียนอัตโนมัติ (CSV + OR-Tools)")

with st.sidebar:
    st.header("1. Upload Files")
    f_rooms = st.file_uploader("ไฟล์ห้องเรียน (rooms_data_flat.csv)", type=['csv'])
    f_subjects = st.file_uploader("ไฟล์รายวิชา (subjects_XX.csv)", type=['csv'], accept_multiple_files=True)
    f_fixed = st.file_uploader("ไฟล์ Fixed (schedule_fix.csv)", type=['csv'], accept_multiple_files=True)

if st.button("🚀 เริ่มจัดตาราง", type="primary"):
    if f_rooms and f_subjects:
        with st.spinner("⏳ กำลังอ่านข้อมูลและประมวลผล..."):
            
            # 1. อ่านข้อมูลห้อง
            df_rooms = pd.read_csv(f_rooms)
            
            # 2. อ่านข้อมูลวิชา (รวมทุกไฟล์)
            all_courses = []
            for f in f_subjects:
                df = pd.read_csv(f)
                # ดึงชื่อสาขาจากชื่อไฟล์ (เช่น subjects_CS.csv -> CS)
                major_name = f.name.split('_')[1].split('.')[0]
                
                for _, row in df.iterrows():
                    # แตก Lec/Lab
                    if row['lecture_hours'] > 0:
                        data = row.to_dict()
                        data['For_Major'] = major_name
                        data['type'] = 'Lec'
                        data['duration'] = row['lecture_hours']
                        all_courses.append(Course(data))
                        
                    if row['lab_hours'] > 0:
                        data = row.to_dict()
                        data['For_Major'] = major_name
                        data['type'] = 'Lab'
                        data['duration'] = row['lab_hours']
                        # Lab ต้องเรียนแยกห้องกับ Lec ไหม? (ในที่นี้แยก object กันชัดเจน)
                        all_courses.append(Course(data))

            # 3. อ่าน Fixed Schedule
            fixed_data = {} # (day, time, room) -> code
            if f_fixed:
                for f in f_fixed:
                    df_fix = pd.read_csv(f)
                    for _, row in df_fix.iterrows():
                        # แปลงเวลาเริ่ม 09:00 -> 9
                        try:
                            start_h = int(float(row['start_time'].split(':')[0]))
                            end_h = int(float(row['end_time'].split(':')[0]))
                            duration = end_h - start_h
                            room = row['room']
                            day = row['day'] # ต้องเป็น Mon, Tue...
                            
                            for i in range(duration):
                                fixed_data[(day, start_h + i, room)] = row['course_code']
                                
                            # อัปเดต Course object ว่าวิชานี้ Fixed
                            # (ทำแบบง่าย: วนลูปหา code ที่ตรงกัน)
                            for c in all_courses:
                                if c.code == row['course_code']:
                                    c.is_fixed = True
                                    c.fixed_day = day
                                    c.fixed_time = start_h
                                    c.fixed_room = room
                        except:
                            pass

            # 4. รัน Solver
            scheduler = UniversityScheduler(all_courses, df_rooms, fixed_data)
            success = scheduler.solve()
            
            if success:
                st.success(f"✅ จัดตารางสำเร็จ! ({len(scheduler.assignment)} วิชา)")
                
                # Export Excel
                xls_data = generate_excel_report(scheduler)
                st.download_button("📥 ดาวน์โหลดตารางเรียน (Excel)", 
                                   xls_data, 
                                   "Final_Schedule.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                # แสดงผล Failed Courses
                if scheduler.failed_courses:
                    st.error(f"⚠️ มี {len(scheduler.failed_courses)} วิชาที่ลงไม่ได้:")
                    for fc in scheduler.failed_courses:
                        st.write(f"- {fc}")
            else:
                st.error("❌ ไม่สามารถจัดตารางได้ (เงื่อนไขแน่นเกินไป)")

    else:
        st.warning("กรุณาอัปโหลดไฟล์ห้องเรียนและรายวิชา")
