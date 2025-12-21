import streamlit as st
import traceback
import io
import math
import random
import pandas as pd
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ------------------------------------------------------------------
# ⚡️ PART 0: SYSTEM CONFIG (Flexible Mode)
# ------------------------------------------------------------------
st.set_page_config(page_title="KKU Scheduler (Flexible)", layout="wide")

# ==========================================
# ⚙️ PART 1: CONSTANTS
# ==========================================
class Config:
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    TIME_SLOTS = range(8, 20) 
    
    # ความจุห้องเริ่มต้น (ใช้กรณีไฟล์ Rooms ไม่ระบุ)
    MAX_CAPACITY_LAB = 50     
    MAX_CAPACITY_LEC = 120    
    
    DEFAULT_REGULAR = 40
    DEFAULT_SPECIAL = 30
    
    COLOR_MAP = {
        'SC': 'FFF59D', 'CP': 'B3E5FC', 'LI': 'C8E6C9', 
        'GE': 'FFE0B2', 'EN': 'E1BEE7', 'COMBINED': 'FFCCBC', 'DEFAULT': 'F5F5F5'
    }

# ==========================================
# 📦 PART 2: DATA STRUCTURES
# ==========================================
class Course:
    def __init__(self, data):
        self.code = str(data.get('course_code', 'N/A')).strip()
        self.name = str(data.get('course_name', 'Unknown')).strip()
        self.major = str(data.get('For_Major', 'Gen'))
        self.year = int(data.get('year', 1))
        
        self.program = str(data.get('program', 'Regular')).capitalize()
        self.type = str(data.get('type', 'Lec'))
        self.duration = int(data.get('duration', 3))
        self.instructor = str(data.get('instructor', 'TBA'))
        self.students = int(data.get('student_count', 40))
        self.section_idx = int(data.get('section_idx', 1))
        
        # Unique ID
        self.uid = f"{self.major}_{self.year}_{self.program}_{self.code}_S{self.section_idx}_{random.randint(100000,999999)}"

    def __repr__(self):
        return f"{self.code} ({self.program}) Sec.{self.section_idx}"

# ==========================================
# 🧠 PART 3: SOLVER ENGINE
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
            capacity_flex = 0.85 if self.options['allow_squeeze'] else 1.0 

            st.write(f"⚙️ กำลังจัดตารางเรียนให้ {len(self.courses)} กลุ่มเรียน...")

            for c in self.courses:
                # 1. กรองห้อง
                valid_rooms = []
                for r in self.rooms:
                    if r['capacity'] >= (c.students * capacity_flex):
                        if c.type.lower() == 'lab':
                            if 'lab' in r['type'].lower(): valid_rooms.append(r)
                        else:
                            valid_rooms.append(r)

                if not valid_rooms:
                    self.failed_courses.append(c)
                    continue
                
                # ไม่จำกัดห้อง (Unlock) เพื่อให้หาเจอแน่นอน
                candidate_rooms = valid_rooms 

                c_moves = []
                for d in active_days:
                    for h in Config.TIME_SLOTS:
                        if h + c.duration > 20: continue
                        if not self.options['allow_lunch'] and (12 in range(h, h + c.duration)): continue
                        
                        # Check Fixed
                        for r in candidate_rooms:
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
                     self.failed_courses.append(c)

            # --- Constraints ---
            time_map = {}
            for (d, t, r_name), _ in self.fixed_data.items():
                time_map.setdefault((d, t), []).append({'type': 'fixed', 'room': r_name})

            for (uid, d, h, r_name), var in shifts.items():
                c = next(x for x in self.courses if x.uid == uid)
                # Group Key: ถ้าเรียนรวม (Combined) ให้ถือว่าเป็นกลุ่มเดียวกันไปเลย
                grp_key = f"{c.major}_{c.year}_{c.program}" 
                
                for i in range(c.duration):
                    key = (d, h + i)
                    time_map.setdefault(key, []).append({
                        'type': 'var', 'var': var, 'room': r_name, 
                        'grp': grp_key, 'instr': c.instructor
                    })

            for slot, items in time_map.items():
                room_usage = {}
                for item in items:
                    room_usage.setdefault(item['room'], []).append(item.get('var', 1))
                for vars_list in room_usage.values():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)

                instr_usage = {}
                grp_usage = {} 
                for item in items:
                    if item['type'] == 'var':
                        ins = item.get('instr')
                        if ins and ins not in ['TBA', 'nan']:
                            instr_usage.setdefault(ins, []).append(item['var'])
                        grp = item.get('grp')
                        if grp:
                            grp_usage.setdefault(grp, []).append(item['var'])
                
                for vars_list in instr_usage.values():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)
                for vars_list in grp_usage.values():
                    if len(vars_list) > 1: model.Add(sum(vars_list) <= 1)

            # --- Solve ---
            model.Maximize(sum(scheduled_vars))
            solver = cp_model.CpSolver()
            solver.parameters.num_search_workers = 8
            solver.parameters.max_time_in_seconds = 300.0
            
            status = solver.Solve(model)

            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                for (uid, d, h, r_name), var in shifts.items():
                    if solver.Value(var) == 1:
                        c = next(course for course in self.courses if course.uid == uid)
                        self.assignments[c] = (d, h, r_name)
                
                assigned_uids = {c.uid for c in self.assignments.keys()}
                for c in self.courses:
                    if c.uid not in assigned_uids and c not in self.failed_courses:
                        self.failed_courses.append(c)
                self.failed_courses = list(set(self.failed_courses))
                return True
            return False
        except Exception as e:
            st.error(f"Solver Error: {e}")
            return False

# ==========================================
# 📊 PART 4: EXCEL GENERATOR
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
                cell = ws.cell(row=r_idx, column=col_idx)
                cell.border = thin

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
                    color_key = 'COMBINED' if 'Combined' in c.program else c.code[:2].upper()
                    color = Config.COLOR_MAP.get(color_key, 'F5F5F5')
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 🖥️ PART 5: USER INTERFACE
# ==========================================
st.title("🎓 KKU Scheduler (Smart Merge)")
st.info("แก้ไข: ปรับปรุงการจัดการภาคปกติ/พิเศษ ไม่ให้แยกวิชาแบบผิดๆ")

with st.sidebar:
    st.header("1. Upload Files")
    f_rooms = st.file_uploader("Rooms (csv)", type=['csv'])
    f_students = st.file_uploader("Students (csv)", type=['csv'])
    f_subjects = st.file_uploader("Subjects (csv)", type=['csv'], accept_multiple_files=True)
    f_fixed = st.file_uploader("Fixed Data (Optional)", type=['csv'], accept_multiple_files=True)
    
    st.markdown("---")
    st.header("2. โหมดจัดการภาคพิเศษ")
    # 🌟 ตัวเลือกสำคัญที่จะแก้ปัญหาของคุณ
    special_mode = st.radio(
        "เลือกรูปแบบการเรียน:",
        ["Separate (แยก Sec ปกติ/พิเศษ)", "Combine (เรียนรวมกันทั้งหมด)", "Strict CSV (ตามไฟล์เท่านั้น)"],
        index=0,
        help="Separate: แยก Sec 1.. กับ Sec 80.. / Combine: เอาคนมารวมกันแล้วซอย Sec 1,2,3 / Strict: ไม่สร้างอะไรเพิ่มเลย"
    )
    
    st.markdown("---")
    st.header("3. Options")
    opt_saturday = st.checkbox("📅 เปิดสอนวันเสาร์", value=True)
    opt_lunch = st.checkbox("🍱 เรียนพักเที่ยงได้", value=False)
    opt_squeeze = st.checkbox("🪑 นั่งเบียดได้ 15%", value=True)

if st.button("🚀 Start Scheduling", type="primary"):
    if f_rooms and f_subjects: 
        with st.spinner("⏳ Analyzing Data..."):
            try:
                # 1. READ ROOMS
                df_rooms = pd.read_csv(f_rooms)
                df_rooms['capacity'] = pd.to_numeric(df_rooms['capacity'], errors='coerce').fillna(30)
                rooms_list = [{'name': str(r['room_name']), 'capacity': int(r['capacity']), 'type': str(r['type']).lower()} for _, r in df_rooms.iterrows()]
                
                # 2. READ STUDENTS
                std_summary = {}
                if f_students:
                    try:
                        df_std = pd.read_csv(f_students)
                        df_std['student_count'] = pd.to_numeric(df_std['student_count'], errors='coerce').fillna(0)
                        for _, r in df_std.iterrows():
                            m = str(r.get('major')).strip()
                            y = int(r.get('year', 1))
                            p = str(r.get('program')).strip().capitalize()
                            std_summary.setdefault((m, y), {})[p] = int(r['student_count'])
                    except: pass
                
                # 3. READ SUBJECTS & CREATE COURSES
                all_courses = []
                for f in f_subjects:
                    try:
                        df = pd.read_csv(f)
                        maj = f.name.split('_')[1].split('.')[0] if '_' in f.name else f.name.split('.')[0]
                        
                        df['year'] = pd.to_numeric(df['year'], errors='coerce').fillna(1)
                        df['lecture_hours'] = pd.to_numeric(df['lecture_hours'], errors='coerce').fillna(0)
                        df['lab_hours'] = pd.to_numeric(df['lab_hours'], errors='coerce').fillna(0)

                        # เช็คว่าในไฟล์ Subject มีระบุ Program มาแล้วหรือไม่
                        has_prog_col = 'program' in df.columns or 'Program' in df.columns

                        for _, row in df.iterrows():
                            yr = int(row['year'])
                            
                            # ดึงจำนวนนักเรียนจากไฟล์ Students
                            std_info = std_summary.get((maj, yr), {})
                            count_reg = std_info.get('Regular', Config.DEFAULT_REGULAR)
                            count_spec = std_info.get('Special', Config.DEFAULT_SPECIAL)
                            
                            # ตรวจสอบ Mode ที่เลือก
                            final_groups = [] # เก็บ (ProgramName, Count, StartSec)
                            
                            if special_mode == "Combine (เรียนรวมกันทั้งหมด)":
                                # เอาคนมารวมกันเลย แล้วสร้างเป็น Sec ปกติ (1, 2, 3...)
                                total_std = count_reg + count_spec if std_info else (Config.DEFAULT_REGULAR + Config.DEFAULT_SPECIAL)
                                final_groups.append(('Combined', total_std, 1))
                                
                            elif special_mode == "Strict CSV (ตามไฟล์เท่านั้น)":
                                # ถ้าไฟล์ระบุ Program มา ให้ใช้ตามนั้น
                                if has_prog_col:
                                    p_name = str(row.get('program', 'Regular')).capitalize()
                                    cnt = count_spec if 'Special' in p_name else count_reg
                                    final_groups.append((p_name, cnt, 1))
                                else:
                                    # ถ้าไม่ระบุ ให้ถือเป็น Regular
                                    final_groups.append(('Regular', count_reg, 1))
                                    
                            else: # Default: Separate (แยก Sec)
                                # สร้าง Regular
                                final_groups.append(('Regular', count_reg, 1))
                                # สร้าง Special (Sec 80+)
                                if count_spec > 0:
                                    final_groups.append(('Special', count_spec, 80))

                            # Loop สร้าง Course ตามกลุ่มที่สรุปได้
                            for prog_name, count, start_sec in final_groups:
                                def create_variants(c_type, hours):
                                    if hours <= 0: return
                                    limit = Config.MAX_CAPACITY_LAB if c_type == 'Lab' else Config.MAX_CAPACITY_LEC
                                    effective_limit = int(limit * 1.15) if opt_squeeze else limit
                                    
                                    num_secs = math.ceil(count / effective_limit)
                                    if num_secs < 1: num_secs = 1
                                    count_per_sec = math.ceil(count / num_secs)
                                    
                                    for i in range(num_secs):
                                        d = row.to_dict()
                                        d.update({
                                            'For_Major': maj,
                                            'program': prog_name,
                                            'student_count': count_per_sec,
                                            'type': c_type,
                                            'duration': hours,
                                            'section_idx': start_sec + i
                                        })
                                        all_courses.append(Course(d))

                                create_variants('Lec', row['lecture_hours'])
                                create_variants('Lab', row['lab_hours'])
                    except Exception as e:
                        st.warning(f"File Error {f.name}: {e}")

                # 4. SOLVE
                fixed_data = {}
                if f_fixed:
                    for f in f_fixed:
                        try:
                            dfx = pd.read_csv(f)
                            for _, r in dfx.iterrows():
                                s = int(float(str(r['start_time']).split(':')[0]))
                                e = int(float(str(r['end_time']).split(':')[0]))
                                for t in range(s, e): fixed_data[(str(r['day']), t, str(r['room']))] = 'FIXED'
                        except: pass

                if not all_courses:
                    st.error("❌ No courses generated.")
                    st.stop()

                options = {'allow_saturday': opt_saturday, 'allow_lunch': opt_lunch, 'allow_squeeze': opt_squeeze}
                scheduler = UniversityScheduler(all_courses, rooms_list, fixed_data, options)
                success = scheduler.solve()
                
                if success:
                    st.balloons()
                    c1, c2 = st.columns(2)
                    c1.metric("✅ Scheduled", len(scheduler.assignments))
                    c2.metric("❌ Failed", len(scheduler.failed_courses))

                    active_days = Config.DAYS + (['Sat'] if opt_saturday else [])
                    excel_file = generate_excel(scheduler.assignments, active_days)
                    st.download_button("📥 Download Schedule", excel_file, "Smart_Schedule.xlsx", type='primary')
                    
                    if scheduler.failed_courses:
                        st.error(f"Failed {len(scheduler.failed_courses)} items")
                        with st.expander("Details"):
                            for c in scheduler.failed_courses:
                                st.write(f"- {c} ({c.students})")
                else:
                    st.error("❌ Solver Timeout")

            except Exception as e:
                st.error("💥 Error")
                st.code(traceback.format_exc())
    else:
        st.info("Upload files to start.")
