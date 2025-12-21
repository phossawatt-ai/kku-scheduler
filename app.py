import streamlit as st
import traceback
import sys

# ------------------------------------------------------------------
# ⚡️ PART 0: SYSTEM CONFIG (ต้องอยู่บรรทัดแรกสุดเสมอ)
# ------------------------------------------------------------------
try:
    st.set_page_config(page_title="KKU Scheduler (Ultra Debug)", layout="wide")
except:
    pass # กัน Error กรณีรันซ้ำ

import pandas as pd
import io
import math
import random
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ฟังก์ชันช่วย Log ข้อความลงหน้าจอ
def log(msg):
    st.session_state.logs.append(msg)
    # st.write(f"🔹 {msg}") # Uncomment ถ้าอยากเห็น Realtime

if 'logs' not in st.session_state:
    st.session_state.logs = []

# ==========================================
# ⚙️ PART 1: CONSTANTS
# ==========================================
class Config:
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    TIME_SLOTS = range(8, 20)
    MAX_CAPACITY_LAB = 45     
    MAX_CAPACITY_LEC = 100    
    DEFAULT_REGULAR = 40
    DEFAULT_SPECIAL = 30
    COLOR_MAP = {'SC': 'FFF59D', 'CP': 'B3E5FC', 'LI': 'C8E6C9', 'GE': 'FFE0B2', 'EN': 'E1BEE7', 'DEFAULT': 'F5F5F5'}

# ==========================================
# 📦 PART 2: CLASS DEFINITION
# ==========================================
class Course:
    def __init__(self, data):
        # ใช้ .get() ทุกตัวเพื่อกัน Key Error
        self.code = str(data.get('course_code', 'N/A')).strip()
        self.name = str(data.get('course_name', 'Unknown')).strip()
        self.major = str(data.get('For_Major', 'Gen'))
        self.year = int(data.get('year', 1))
        self.program = str(data.get('program', 'Regular'))
        self.type = str(data.get('type', 'Lec'))
        self.duration = int(data.get('duration', 3))
        self.instructor = str(data.get('instructor', 'TBA'))
        self.students = int(data.get('student_count', 40))
        self.section_idx = int(data.get('section_idx', 1))
        # Unique ID แบบปลอดภัย
        self.uid = f"{self.major}_{self.year}_{self.program}_{self.code}_{self.type}_S{self.section_idx}_{random.randint(1000,9999)}"

    def __repr__(self):
        return f"{self.code} ({self.program}) Sec.{self.section_idx}"

# ==========================================
# 🧠 PART 3: SCHEDULER ENGINE
# ==========================================
class UniversityScheduler:
    def __init__(self, courses, rooms, fixed_data, options):
        self.courses = courses
        self.rooms = sorted(rooms, key=lambda x: x['capacity'])
        self.fixed_data = fixed_data
        self.options = options 
        self.assignments = {}
        self.failed_courses = []
        
    def solve(self):
        try:
            model = cp_model.CpModel()
            shifts = {} 
            scheduled_vars = [] 
            
            active_days = Config.DAYS + (['Sat'] if self.options['allow_saturday'] else [])
            capacity_flex = 0.9 if self.options['allow_squeeze'] else 1.0 

            # --- Variables ---
            for c in self.courses:
                valid_rooms = [r for r in self.rooms if r['capacity'] >= (c.students * capacity_flex)]
                # Filter Lab/Lec
                if c.type.lower() == 'lab':
                    valid_rooms = [r for r in valid_rooms if 'lab' in r['type']]
                else:
                    valid_rooms = [r for r in valid_rooms if 'lab' not in r['type']] # เลี่ยงห้อง Lab ถ้าเป็น Lec

                if not valid_rooms:
                    self.failed_courses.append(c)
                    continue
                
                valid_rooms = valid_rooms[:5] # 🔥 ลดจำนวนห้องที่พิจารณา เพื่อประหยัดเมม

                c_moves = []
                for d in active_days:
                    for h in Config.TIME_SLOTS:
                        if h + c.duration > 20: continue
                        if not self.options['allow_lunch'] and (12 in range(h, h + c.duration)): continue
                        
                        # Check Fixed
                        for r in valid_rooms:
                            if any((d, t, r['name']) in self.fixed_data for t in range(h, h + c.duration)): continue

                            var = model.NewBoolVar(f'shift_{c.uid}_{d}_{h}_{r["name"]}')
                            shifts[(c.uid, d, h, r["name"])] = var
                            c_moves.append(var)
                
                if c_moves:
                    is_scheduled = model.NewBoolVar(f'sched_{c.uid}')
                    model.Add(sum(c_moves) == is_scheduled)
                    scheduled_vars.append(is_scheduled)
                else:
                    self.failed_courses.append(c)

            # --- Constraints ---
            time_map = {}
            for (d, t, r_name), _ in self.fixed_data.items():
                time_map.setdefault((d, t), []).append({'type': 'fixed', 'room': r_name})

            for (uid, d, h, r_name), var in shifts.items():
                c = next(x for x in self.courses if x.uid == uid)
                for i in range(c.duration):
                    key = (d, h + i)
                    grp_key = f"{c.major}_{c.year}_{c.program}"
                    time_map.setdefault(key, []).append({
                        'type': 'var', 'var': var, 'room': r_name, 
                        'grp': grp_key, 'instr': c.instructor
                    })

            for slot, items in time_map.items():
                # 1. Room
                room_usage = {}
                for item in items:
                    room_usage.setdefault(item['room'], []).append(item.get('var', 1))
                for vars_list in room_usage.values():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)

                # 2. Instr & Group
                instr_usage = {}
                grp_usage = {}
                for item in items:
                    if item['type'] == 'var':
                        if item['instr'] != 'TBA': instr_usage.setdefault(item['instr'], []).append(item['var'])
                        grp_usage.setdefault(item['grp'], []).append(item['var'])
                
                for vars_list in instr_usage.values():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)
                for vars_list in grp_usage.values():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)

            # --- Solve ---
            model.Maximize(sum(scheduled_vars))
            solver = cp_model.CpSolver()
            solver.parameters.num_search_workers = 2    # 🔥 ลด Worker ลงกันเครื่องค้าง
            solver.parameters.max_time_in_seconds = 300.0 # 🔥 ลดเวลาลง
            
            status = solver.Solve(model)

            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                for (uid, d, h, r_name), var in shifts.items():
                    if solver.Value(var) == 1:
                        c = next(course for course in self.courses if course.uid == uid)
                        self.assignments[c] = (d, h, r_name)
                
                # เก็บตกวิชาที่หลุด
                assigned = set(self.assignments.keys())
                for c in self.courses:
                    if c not in assigned: self.failed_courses.append(c)
                self.failed_courses = list(set(self.failed_courses))
                return True
            return False

        except Exception as e:
            st.error(f"💥 Solver Error: {e}")
            st.code(traceback.format_exc())
            return False

# ==========================================
# 📊 PART 4: EXCEL
# ==========================================
def generate_excel(assignments, active_days):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    sheet_data = {}
    for c, (d, t, r) in assignments.items():
        key = f"{c.major}_Y{c.year}_{c.program}"
        sheet_data.setdefault(key, []).append((c, d, t, r))
        
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for sheet_name in sorted(sheet_data.keys()):
        safe_name = sheet_name.replace('/', '_')[:30]
        ws = wb.create_sheet(title=safe_name)
        
        ws.merge_cells('A1:N1')
        ws['A1'] = f"ตารางเรียน {sheet_name}"
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = align
        
        ws['A2'] = "Day/Time"
        for i, h in enumerate(Config.TIME_SLOTS):
            col = chr(66 + i) 
            ws[f'{col}2'] = f"{h}:00"
            ws.column_dimensions[col].width = 16
            
        row_map = {d: i+3 for i, d in enumerate(active_days)}
        for d, r_idx in row_map.items():
            ws[f'A{r_idx}'] = d
            ws[f'A{r_idx}'].font = Font(bold=True)
            for col_idx in range(2, 15):
                ws.cell(row=r_idx, column=col_idx).border = thin

        for c, d, t, r in sheet_data[sheet_name]:
            if d not in row_map: continue
            r_idx = row_map[d]
            for i in range(c.duration):
                h_cur = t + i
                if h_cur in Config.TIME_SLOTS:
                    c_idx = Config.TIME_SLOTS.index(h_cur) + 2
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.value = f"{c.code} Sec {c.section_idx}\n{c.name}\n{r} ({c.instructor})"
                    cell.alignment = align
                    cell.border = thin
                    color = Config.COLOR_MAP.get(c.code[:2].upper(), 'F5F5F5')
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 🖥️ PART 5: MAIN UI (SAFE MODE)
# ==========================================
st.title("🛡️ KKU Scheduler (Debug Mode)")

# Sidebar
with st.sidebar:
    st.header("1. Upload Files")
    f_rooms = st.file_uploader("Rooms", type=['csv'])
    f_students = st.file_uploader("Students", type=['csv'])
    f_subjects = st.file_uploader("Subjects", type=['csv'], accept_multiple_files=True)
    f_fixed = st.file_uploader("Fixed", type=['csv'], accept_multiple_files=True)
    
    st.markdown("---")
    st.header("2. Options")
    opt_force_special = st.checkbox("🔥 Force Special", value=True)
    opt_saturday = st.checkbox("📅 Enable Saturday", value=False)
    opt_lunch = st.checkbox("🍱 Allow Lunch", value=False)
    opt_squeeze = st.checkbox("🪑 Squeeze 10%", value=True)

# Main Process
if st.button("🚀 Run Scheduler", type="primary"):
    st.session_state.logs = [] # Clear logs
    log("Process Started...")
    
    if not f_rooms or not f_subjects:
        st.error("❌ ขาดไฟล์ Rooms หรือ Subjects")
        st.stop()

    try:
        # 1. READ ROOMS
        log("Reading Rooms...")
        df_rooms = pd.read_csv(f_rooms)
        # Check Columns
        req_room_cols = ['room_name', 'capacity', 'type']
        if not all(col in df_rooms.columns for col in req_room_cols):
            st.error(f"❌ ไฟล์ Rooms ต้องมีหัวข้อ: {req_room_cols}")
            st.stop()
            
        df_rooms['capacity'] = pd.to_numeric(df_rooms['capacity'], errors='coerce').fillna(30)
        rooms_list = [{'name': str(r['room_name']), 'capacity': int(r['capacity']), 'type': str(r['type']).lower()} for _, r in df_rooms.iterrows()]
        log(f"✅ Loaded {len(rooms_list)} rooms")

        # 2. READ STUDENTS
        log("Reading Students...")
        std_summary = {}
        if f_students:
            df_std = pd.read_csv(f_students)
            if 'student_count' in df_std.columns:
                df_std['student_count'] = pd.to_numeric(df_std['student_count'], errors='coerce').fillna(0)
                for _, r in df_std.iterrows():
                    std_summary.setdefault((str(r.get('major')), int(r.get('year',1))), {})[str(r.get('program'))] = int(r['student_count'])
        log("✅ Loaded Students info")

        # 3. READ SUBJECTS
        log("Processing Subjects...")
        all_courses = []
        
        for f in f_subjects:
            try:
                log(f"Reading {f.name}...")
                df = pd.read_csv(f)
                
                # Check Columns
                req_sub_cols = ['course_code', 'course_name', 'year', 'lecture_hours', 'lab_hours']
                missing = [c for c in req_sub_cols if c not in df.columns]
                if missing:
                    st.warning(f"⚠️ ไฟล์ {f.name} ขาดหัวข้อ {missing} -> ข้าม")
                    continue

                # Clean Data
                df['lecture_hours'] = pd.to_numeric(df['lecture_hours'], errors='coerce').fillna(0)
                df['lab_hours'] = pd.to_numeric(df['lab_hours'], errors='coerce').fillna(0)
                df['year'] = pd.to_numeric(df['year'], errors='coerce').fillna(1)

                maj = f.name.split('_')[1].split('.')[0] if '_' in f.name else f.name.split('.')[0]

                for _, row in df.iterrows():
                    yr = int(row['year'])
                    progs = std_summary.get((maj, yr), {})
                    if not progs: progs = {'Regular': Config.DEFAULT_REGULAR}
                    if opt_force_special and 'Special' not in progs: progs['Special'] = Config.DEFAULT_SPECIAL

                    for prog, count in progs.items():
                        if count <= 0: count = Config.DEFAULT_REGULAR
                        
                        start_sec = 1 if prog == 'Regular' else 80
                        
                        # Function to add course variants
                        def add_course(type_name, hours):
                            if hours <= 0: return
                            limit = int(Config.MAX_CAPACITY_LAB if type_name=='Lab' else Config.MAX_CAPACITY_LEC * (1.1 if opt_squeeze else 1.0))
                            
                            num_secs = math.ceil(count / limit)
                            if num_secs < 1: num_secs = 1
                            count_per_sec = math.ceil(count / num_secs)
                            
                            for i in range(num_secs):
                                d = row.to_dict()
                                d.update({
                                    'For_Major': maj, 'program': prog, 'student_count': count_per_sec,
                                    'type': type_name, 'duration': hours, 'section_idx': start_sec + i
                                })
                                all_courses.append(Course(d))

                        add_course('Lec', row['lecture_hours'])
                        add_course('Lab', row['lab_hours'])

            except Exception as e:
                st.error(f"❌ Error in file {f.name}: {e}")

        log(f"✅ Total Secs to schedule: {len(all_courses)}")
        if len(all_courses) == 0:
            st.error("❌ ไม่พบวิชาที่จะจัดตารางเลย")
            st.stop()

        # 4. FIXED DATA
        fixed_data = {}
        if f_fixed:
            log("Reading Fixed Data...")
            for f in f_fixed:
                try:
                    dfx = pd.read_csv(f)
                    for _, r in dfx.iterrows():
                        s = int(float(str(r['start_time']).split(':')[0]))
                        e = int(float(str(r['end_time']).split(':')[0]))
                        for t in range(s, e): fixed_data[(str(r['day']), t, str(r['room']))] = 'FIXED'
                except: pass

        # 5. SOLVE
        log("🚀 Calling Solver...")
        options = {'allow_saturday': opt_saturday, 'allow_lunch': opt_lunch, 'allow_squeeze': opt_squeeze}
        scheduler = UniversityScheduler(all_courses, rooms_list, fixed_data, options)
        
        success = scheduler.solve()
        
        if success:
            log("✅ Solved Successfully!")
            st.balloons()
            
            # Summary Metrics
            c1, c2, c3 = st.columns(3)
            c1.metric("Scheduled", len(scheduler.assignments))
            c2.metric("Failed", len(scheduler.failed_courses))
            c3.metric("Rooms Used", len(set(r for _,_,r in scheduler.assignments.values())))
            
            # Excel
            excel_file = generate_excel(scheduler.assignments, Config.DAYS + (['Sat'] if opt_saturday else []))
            st.download_button("📥 Download Excel", excel_file, "Result.xlsx", type='primary')
            
            # Error Details
            if scheduler.failed_courses:
                st.error(f"มีวิชาหลุด {len(scheduler.failed_courses)} วิชา")
                with st.expander("ดูรายการที่หลุด"):
                    for c in scheduler.failed_courses:
                        st.write(f"❌ {c} ({c.students} คน) - หาห้อง/เวลาไม่ได้")
        else:
            st.error("❌ Solver Timeout: หาคำตอบไม่เจอในเวลาที่กำหนด (ลองลดวิชา หรือเพิ่มห้อง)")

    except Exception as e:
        st.error("💥 Critical Crash!")
        st.code(traceback.format_exc())
    
    # Show Logs
    with st.expander("📜 System Logs (ใช้สำหรับแก้ปัญหา)"):
        for l in st.session_state.logs:
            st.text(l)
