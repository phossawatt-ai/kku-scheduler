import streamlit as st
import pandas as pd
import re
import io
import time
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font

# ==========================================
# ⚙️ CONFIGURATION
# ==========================================
class Config:
    MAJOR_MAP = {
        'CS': 'วิทยาการคอมพิวเตอร์', 
        'IT': 'เทคโนโลยีสารสนเทศ',
        'AI': 'ปัญญาประดิษฐ์', 
        'GIS': 'ภูมิสารสนเทศศาสตร์',
        'CYBER': 'ความมั่นคงปลอดภัยไซเบอร์'
    }
    
    COLOR_MAP = {
        'SC': 'FFF59D', # เหลือง
        'CP': 'B3E5FC', # ฟ้า
        'LI': 'C8E6C9', # เขียว
        'EN': 'C8E6C9', 
        'GE': 'FFE0B2', # ส้ม
        'DEFAULT': 'F5F5F5' # เทา
    }
    
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    
    # *** (แก้บั๊ก) เติมตัวแปรนี้กลับมาครับ ***
    TIME_SLOTS = range(8, 20) 
    
    # คำที่เจอแล้วให้หยุดอ่านทันที (ป้องกันวิชาสรุปท้ายตารางหลุดมา)
    STOP_KEYWORDS = ['วิชาเลือก', 'สรุปจำนวน', 'รวมหน่วยกิต', 'Elective', 'Total', 'สรุป']

# ==========================================
# 📦 DATA MODELS
# ==========================================
class Course:
    def __init__(self, row_data):
        self.major = row_data['major']
        self.semester = row_data['semester']
        self.program = row_data['program']
        self.code = row_data['code']
        self.name = row_data['name']
        self.instructor = str(row_data.get('instructor', 'TBA'))
        self.type = 'Lab' if row_data.get('lab_hours', 0) > 0 else 'Lec'
        self.year = int(row_data.get('year', 1))
        
        # Student Count
        base = int(row_data.get('student_count', 0))
        self.students = base if base > 0 else (50 if self.program == 'Reg' else 30)
        
        # Duration
        raw_dur = int(row_data['lec_hours']) if self.type == 'Lec' else int(row_data['lab_hours'])
        self.duration = min(raw_dur, 4) if raw_dur > 0 else 2
        
        # Fixed Schedule Flags
        self.is_fixed = False
        self.fixed_day = None
        self.fixed_time = None
        self.fixed_room = None
        
        self.related_course = None 

    def __repr__(self):
        return f"{self.code} ({self.program})"

# ==========================================
# 🧠 SCHEDULER ENGINE
# ==========================================
class UniversityScheduler:
    def __init__(self, courses_list, rooms_df, busy_df):
        self.rooms = sorted(rooms_df.to_dict('records'), key=lambda x: x['capacity'])
        self.busy_slots = set((str(row['room']), row['day'], row['hour']) for _, row in busy_df.iterrows())
        self.assignment = {} 
        self.failed_courses = [] 
        
        self.courses_to_schedule = courses_list
        self._link_courses()
        self._apply_fixed_schedules()
        
        # Sort: Fixed first, then Hardest to schedule
        self.courses_to_schedule.sort(key=lambda x: (not x.is_fixed, -x.students, -x.duration))

    def _link_courses(self):
        course_map = {}
        for c in self.courses_to_schedule:
            if c.type == 'Lec': course_map[(c.major, c.program, c.code)] = c
        for c in self.courses_to_schedule:
            if c.type == 'Lab':
                key = (c.major, c.program, c.code)
                if key in course_map: c.related_course = course_map[key]

    def _apply_fixed_schedules(self):
        for c in self.courses_to_schedule:
            if c.is_fixed and c.fixed_day and c.fixed_time is not None:
                self.assignment[c] = (c.fixed_day, c.fixed_time, c.fixed_room)

    def is_valid(self, course, day, start, room, squeeze=1.25, strict=True, lunch=False):
        if course.is_fixed: return False 
        
        end = start + course.duration
        if end > 20: return False
        if not lunch and any(t == 12 for t in range(start, end)): return False
        if (room['capacity'] * squeeze) < course.students: return False
        
        if strict:
            if (course.type == 'Lab' and room['type'] != 'Lab') or (course.type == 'Lec' and room['type'] == 'Lab'): return False

        for t in range(start, end):
            if (str(room['room_name']), day, t) in self.busy_slots: return False
            
        for c, (d, t, r) in self.assignment.items():
            if d == day:
                if max(start, t) < min(end, t + c.duration):
                    if r == room['room_name']: return False
                    
                    inst_a = set(course.instructor.split(','))
                    inst_b = set(c.instructor.split(','))
                    if 'TBA' not in inst_a and 'TBA' not in inst_b:
                        if not inst_a.isdisjoint(inst_b): return False
                    
                    # Student Conflict (Major+Year+Prog must match)
                    if (course.major == c.major) and (course.year == c.year) and (course.program == c.program):
                        return False

        if course.type == 'Lab' and course.related_course:
            lec = course.related_course
            if lec not in self.assignment: return False
            lec_day, _, _ = self.assignment[lec]
            d_idx = {d: i for i, d in enumerate(Config.DAYS)}
            if d_idx[day] < d_idx[lec_day]: return False
            if d_idx[day] == d_idx[lec_day] and start < (self.assignment[lec][1] + lec.duration): return False
            
        return True

    def solve(self):
        phases = [
            {'days': Config.DAYS[:5], 'hours': range(8, 16), 'squeeze': 1.25, 'strict': True, 'lunch': False},
            {'days': Config.DAYS[:5], 'hours': range(8, 16), 'squeeze': 1.50, 'strict': False, 'lunch': False},
            {'days': Config.DAYS[:5], 'hours': range(16, 18), 'squeeze': 1.50, 'strict': False, 'lunch': False},
            {'days': Config.DAYS[:5], 'hours': range(18, 20), 'squeeze': 1.50, 'strict': False, 'lunch': False},
            {'days': Config.DAYS[5:], 'hours': range(8, 20), 'squeeze': 1.50, 'strict': False, 'lunch': True},
        ]
        
        for phase in phases:
            remaining = [c for c in self.courses_to_schedule if c not in self.assignment]
            rooms_try = self.rooms if phase['strict'] else sorted(self.rooms, key=lambda x: x['capacity'], reverse=True)
            
            for course in remaining:
                assigned = False
                for day in phase['days']:
                    if assigned: break
                    for hour in phase['hours']:
                        if assigned: break
                        if hour + course.duration > 20: continue
                        for room in rooms_try:
                            if self.is_valid(course, day, hour, room, phase['squeeze'], phase['strict'], phase['lunch']):
                                self.assignment[course] = (day, hour, room['room_name'])
                                assigned = True; break
        
        self.failed_courses = [c for c in self.courses_to_schedule if c not in self.assignment]

# ==========================================
# 🛠️ HELPER: EXTRACT DATA (Robust V19)
# ==========================================
def extract_data_v19(course_file, room_file, fixed_file, blacklist_codes=[]):
    # 1. Read Fixed Schedule (Optional)
    fixed_map = {} 
    if fixed_file:
        try:
            xls_fix = pd.ExcelFile(fixed_file)
            for sheet in xls_fix.sheet_names:
                df_fix = pd.read_excel(fixed_file, sheet_name=sheet)
                df_fix.columns = df_fix.columns.astype(str).str.lower()
                
                col_code = next((c for c in df_fix.columns if 'รหัส' in c or 'code' in c), None)
                col_day = next((c for c in df_fix.columns if 'วัน' in c or 'day' in c), None)
                col_time = next((c for c in df_fix.columns if 'เวลา' in c or 'time' in c), None)
                col_room = next((c for c in df_fix.columns if 'ห้อง' in c or 'room' in c), None)
                
                if col_code and col_day and col_time:
                    for _, row in df_fix.iterrows():
                        c_code = str(row[col_code]).replace(" ", "").strip()
                        c_day = str(row[col_day]).strip()
                        c_time = str(row[col_time]).strip()
                        c_room = str(row[col_room]) if col_room else "External"
                        
                        d_map = {'จันทร์':'Mon', 'อังคาร':'Tue', 'พุธ':'Wed', 'พฤหัส':'Thu', 'ศุกร์':'Fri', 'เสาร์':'Sat', 'อาทิตย์':'Sun'}
                        day_en = next((en for th, en in d_map.items() if th in c_day), None)
                        time_match = re.search(r'(\d+)', c_time)
                        start_h = int(time_match.group(1)) if time_match else None
                        
                        if day_en and start_h:
                            fixed_map[c_code] = {'day': day_en, 'time': start_h, 'room': c_room}
        except: pass

    all_courses = []
    
    # 2. Course File
    try:
        xls = pd.ExcelFile(course_file)
        for sheet in xls.sheet_names:
            major = next((m for m in Config.MAJOR_MAP if m in sheet.upper()), None)
            if not major: continue
            
            df = pd.read_excel(course_file, sheet_name=sheet, header=None)
            
            # --- Aggressive Split Search (แก้ปัญหา GIS โล่ง) ---
            split_idx = len(df.columns) // 2 
            found_split = False
            for r in range(min(20, len(df))):
                row_txt = "".join([str(x) for x in df.iloc[r].values])
                if "ภาคปลาย" in row_txt or "การศึกษาที่ 2" in row_txt or "Semester 2" in row_txt:
                    for c in range(len(df.columns)):
                        val = str(df.iloc[r, c])
                        if "ภาคปลาย" in val or "การศึกษาที่ 2" in val or "Semester 2" in val:
                            split_idx = c; found_split = True; break
                    if found_split: break
            
            current_year = 1
            stop_reading = False 
            
            for _, row in df.iterrows():
                row_str = " ".join([str(x) for x in row.values if str(x) != 'nan']).strip()
                
                # --- Stop Word (แก้ปัญหาวิชาผีท้ายตาราง) ---
                if any(kw in row_str for kw in Config.STOP_KEYWORDS):
                    stop_reading = True
                if stop_reading: continue 

                if len(row_str) < 100 and "หน่วยกิต" not in row_str:
                    ym = re.search(r'ปี.*?(\d)', row_str)
                    if ym and 1 <= int(ym.group(1)) <= 4: current_year = int(ym.group(1))

                def parse_segment(seg, semester):
                    res = []
                    matches = list(re.finditer(r'\b([A-Z]{2}\s?\d{3}\s?\d{3})\b', seg))
                    for i, m in enumerate(matches):
                        code = m.group(1).replace(" ", "")
                        
                        # --- Blacklist Filter (กรองรหัสวิชาที่ไม่ต้องการ) ---
                        if code in blacklist_codes: continue 

                        y = current_year
                        if len(code) >= 5 and code[4].isdigit():
                            dy = int(code[4])
                            if dy > y and dy <= 4: y = dy 
                        
                        start, end = m.end(), matches[i+1].start() if i+1 < len(matches) else len(seg)
                        sub = seg[start:end]
                        
                        lec, lab, cnt, instr = 3, 0, 0, "TBA"
                        cr = re.search(r'(\d+)\s*\(\s*(\d+)\s*-\s*(\d+)\s*-\s*(\d+)\s*\)', sub)
                        name_raw = sub
                        
                        if cr:
                            name_raw = sub[:cr.start()]
                            lec, lab = int(cr.group(2)), int(cr.group(3))
                            meta = sub[cr.end():]
                            instrs = re.findall(r'\b[A-Z]{1,3}\d\b', meta)
                            if instrs: instr = ",".join(instrs)
                            nums = [int(n) for n in re.findall(r'\b(\d{2,3})\b', meta) if 20 <= int(n) <= 200]
                            if nums: cnt = max(nums)
                        
                        name_clean = re.sub(r'\(.*?\)', '', name_raw).strip()
                        if len(name_clean) > 2:
                            for prog in ['Reg', 'Sp']:
                                c_obj = Course({
                                    'major': major, 'program': prog, 'year': y, 'semester': semester,
                                    'code': code, 'name': name_clean, 'lec_hours': lec, 'lab_hours': lab,
                                    'student_count': cnt if cnt > 0 else (50 if prog=='Reg' else 30), 
                                    'instructor': instr
                                })
                                
                                # Apply Fixed
                                if code in fixed_map:
                                    fix = fixed_map[code]
                                    c_obj.is_fixed = True
                                    c_obj.fixed_day = fix['day']
                                    c_obj.fixed_time = fix['time']
                                    c_obj.fixed_room = fix['room']
                                
                                res.append(c_obj)
                    return res

                s1 = " ".join([str(x) for x in row.iloc[:split_idx].values if str(x) != 'nan'])
                s2 = " ".join([str(x) for x in row.iloc[split_idx:].values if str(x) != 'nan'])
                all_courses.extend(parse_segment(s1, 1))
                all_courses.extend(parse_segment(s2, 2))
    except Exception as e: st.error(f"Course Read Error: {e}")

    # 3. Room File
    rooms, busy = [], []
    try:
        xls = pd.ExcelFile(room_file)
        for sheet in xls.sheet_names:
            if not re.search(r'\d{3,4}', sheet): continue
            df = pd.read_excel(room_file, sheet_name=sheet, header=None)
            cap = 40
            head = "".join([str(df.iloc[i].values) for i in range(min(5, len(df)))])
            sm = re.search(r'(\d+)\s*ที่นั่ง', head)
            if sm: cap = int(sm.group(1))
            else:
                caps = [int(c) for c in re.findall(r'\((\d+)\)', head) if int(c) < 500]
                if caps: cap = caps[0]
            
            rooms.append({'room_name': sheet.strip(), 'type': 'Lab' if 'LAB' in sheet.upper() else 'Lec', 'capacity': cap})
            
            s_row = next((i for i, r in df.iterrows() if "จันทร์" in str(r.values)), -1)
            if s_row != -1:
                sch = df.iloc[s_row:].reset_index(drop=True)
                d_map = {'จันทร์': 'Mon', 'อังคาร': 'Tue', 'พุธ': 'Wed', 'พฤหัส': 'Thu', 'ศุกร์': 'Fri', 'เสาร์': 'Sat', 'อาทิตย์': 'Sun'}
                for r_idx, row in sch.iterrows():
                    d = next((en for th, en in d_map.items() if th in str(row.iloc[0])), None)
                    if d:
                        for c in range(1, len(row)):
                            if str(row.iloc[c]).lower() != 'nan' and len(str(row.iloc[c])) > 2:
                                if 8+(c-1) <= 20: busy.append({'room': sheet.strip(), 'day': d, 'hour': 8+(c-1)})
    except Exception as e: st.error(f"Room Read Error: {e}")
    
    return all_courses, pd.DataFrame(rooms), pd.DataFrame(busy)

# ==========================================
# 📊 EXPORT EXCEL
# ==========================================
def generate_excel_report(sched1, sched2):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    all_scheds = [sched1, sched2]
    majors = set()
    years = set()
    for s in all_scheds:
        for c in s.assignment:
            majors.add(c.major)
            years.add(c.year)
            
    thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    
    for major in sorted(list(majors)):
        for year in sorted(list(years)):
            for program in ['Reg', 'Sp']:
                prog_name = 'ภาคปกติ' if program == 'Reg' else 'โครงการพิเศษ'
                sheet_name = f"{major}-{year}-{program}"
                
                def get_data(sched):
                    return [
                        (c, d, t, r) for c, (d, t, r) in sched.assignment.items()
                        if c.major == major and c.year == year and c.program == program
                    ]

                def draw_term(data, start_row):
                    df = pd.DataFrame('', index=Config.DAYS, columns=[f"{h}:00-{h+1}:00" for h in Config.TIME_SLOTS])
                    colors = {}
                    for c, d, t, r in data:
                        bg = Config.COLOR_MAP.get(c.code[:2].upper(), Config.COLOR_MAP['DEFAULT'])
                        txt = f"{c.code}\n{c.name}\n{r} ({c.instructor})"
                        for i in range(c.duration):
                            if 8 <= t+i < 20:
                                col = f"{t+i}:00-{t+i+1}:00"
                                if d in df.index:
                                    prev = df.at[d, col]
                                    df.at[d, col] = (prev + "\n---\n" + txt).strip() if prev else txt
                                    colors[(Config.DAYS.index(d)+start_row+1, (t+i-8)+2)] = bg
                    return df, colors

                df1, c1 = draw_term(get_data(sched1), 4)
                df2, c2 = draw_term(get_data(sched2), 15)
                
                df1.to_excel(writer, sheet_name=sheet_name, startrow=3)
                df2.to_excel(writer, sheet_name=sheet_name, startrow=14)
                
                ws = writer.sheets[sheet_name]
                for r, txt in [(1, "ภาคต้น"), (13, "ภาคปลาย")]:
                    ws[f'A{r}'] = f"ตารางเรียน {Config.MAJOR_MAP.get(major, major)} ปี {year} ({prog_name}) - {txt}"
                    ws.merge_cells(f'A{r}:M{r}')
                    ws[f'A{r}'].font = Font(size=14, bold=True); ws[f'A{r}'].alignment = Alignment(horizontal='center')
                
                ws.column_dimensions['A'].width = 12
                for i in range(2, 15): ws.column_dimensions[chr(64+i)].width = 18
                
                for start, cmap in [(4, c1), (15, c2)]:
                    for r in range(start, start+8):
                        for c in range(1, 14):
                            cell = ws.cell(row=r, column=c)
                            cell.border = thin
                            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                            if (r, c) in cmap: cell.fill = PatternFill(start_color=cmap[(r, c)], end_color=cmap[(r, c)], fill_type='solid')
                            elif r == start or c == 1: cell.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type='solid'); cell.font = Font(bold=True)

                failed = [c for s in all_scheds for c in s.failed_courses if c.major == major and c.year == year and c.program == program]
                if failed:
                    ws.cell(row=24, column=1, value="⚠️ วิชาที่ลงไม่ได้:").font = Font(color="FF0000", bold=True)
                    for i, f in enumerate(set(failed)): 
                        ws.cell(row=25+i, column=1, value=f"{f.code} {f.name}")

    writer.close()
    output.seek(0)
    return output

# ==========================================
# 🖥️ APP INTERFACE
# ==========================================
st.set_page_config(page_title="KKU Scheduler Pro", layout="wide")
st.title("🎓 ระบบจัดตารางเรียนอัตโนมัติ (KKU AI Scheduler)")
st.info("ระบบจัดการภาคปกติ/พิเศษ แยกอิสระ และกรองวิชาผี (Ghost Courses) ออกอัตโนมัติ")

c1, c2, c3 = st.columns(3)
f_course = c1.file_uploader("1. ไฟล์หลักสูตร (kku30...)", type=['xlsx'])
f_room = c2.file_uploader("2. ไฟล์ห้องเรียน (LAB...)", type=['xlsx'])
f_fixed = c3.file_uploader("3. ไฟล์วิชาต่างคณะ (Optional)", type=['xlsx'])

# *** NEW: ช่องกรอกวิชาผี ***
blacklist_input = st.text_area("🚫 รายวิชาที่ต้องการลบออก (พิมพ์รหัสวิชา คั่นด้วยจุลภาค)", 
                               placeholder="เช่น SC402101, GE101001 (วิชาเหล่านี้จะไม่ถูกนำมาจัดตาราง)")

if f_course and f_room:
    if st.button("🚀 เริ่มจัดตาราง", type="primary"):
        with st.spinner("⏳ กำลังประมวลผล..."):
            try:
                # แปลง Blacklist เป็น List
                blacklist = [x.strip() for x in blacklist_input.split(',') if x.strip()]
                
                courses, rooms, busy = extract_data_v19(f_course, f_room, f_fixed, blacklist)
                
                if not courses or rooms.empty: st.error("❌ ไม่พบข้อมูลในไฟล์")
                else:
                    scheds = {}
                    for t in [1, 2]:
                        term_courses = [c for c in courses if c.semester == t]
                        s = UniversityScheduler(term_courses, rooms, busy)
                        s.solve()
                        scheds[t] = s
                    
                    xls = generate_excel_report(scheds[1], scheds[2])
                    st.success("✅ เสร็จสมบูรณ์!")
                    st.download_button("📥 ดาวน์โหลดไฟล์ Excel", xls, "Final_Schedule_Fixed.xlsx", 
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                     use_container_width=True)
            except Exception as e: st.error(f"Error: {e}")
