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
        'SC': 'FFF59D', 'CP': 'B3E5FC', 'LI': 'C8E6C9',
        'EN': 'C8E6C9', 'GE': 'FFE0B2', 'DEFAULT': 'F5F5F5'
    }
    DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    TIME_SLOTS = range(8, 20)

# ==========================================
# 📦 DATA MODELS
# ==========================================
class Course:
    def __init__(self, row_data):
        self.major = row_data['major']
        self.program = row_data['program'] # Reg / Sp
        self.code = row_data['code']
        self.name = row_data['name']
        self.instructor = str(row_data.get('instructor', 'TBA'))
        self.type = 'Lab' if row_data['lab_hours'] > 0 else 'Lec'
        self.year = int(row_data.get('year', 1))
        
        # Logic: จำนวนนักศึกษา (แยกตามภาค)
        # ถ้าไม่มีข้อมูลระบุชัดเจน ให้ใช้ค่า Default ต่างกัน
        base_count = int(row_data.get('student_count', 0))
        if base_count == 0:
            self.students = 50 if self.program == 'Reg' else 30
        else:
            self.students = base_count

        # Logic: เวลาเรียน
        raw_duration = int(row_data['lec_hours']) if self.type == 'Lec' else int(row_data['lab_hours'])
        self.duration = min(raw_duration, 4)
        if self.duration == 0: self.duration = 2
        
        # ID ต้องไม่ซ้ำกันระหว่างภาค
        self.id = f"{self.major}-{self.program}-{self.code}-{self.type}"
        self.related_course = None 

    def __repr__(self):
        return f"{self.code} ({self.program})"

# ==========================================
# 🧠 SCHEDULER ENGINE
# ==========================================
class UniversityScheduler:
    def __init__(self, courses_list, rooms_df, busy_df):
        # Rooms: Sort Small -> Large
        self.rooms = sorted(rooms_df.to_dict('records'), key=lambda x: x['capacity'])
        self.busy_slots = set((str(row['room']), row['day'], row['hour']) for _, row in busy_df.iterrows())
        self.assignment = {} # {course: (day, start_time, room)}
        self.failed_courses = []
        
        # Link Lec-Lab (ภายในภาคเดียวกันเท่านั้น)
        self.courses_to_schedule = courses_list
        self._link_courses()
        
        # Sort: คนเยอะ + เรียนยาว -> ลงก่อน
        self.courses_to_schedule.sort(key=lambda x: (-x.students, -x.duration))

    def _link_courses(self):
        # จับคู่ Lec-Lab ที่เป็นวิชาเดียวกันและภาคเดียวกัน
        course_map = {}
        for c in self.courses_to_schedule:
            if c.type == 'Lec':
                key = (c.major, c.program, c.code) # Key ต้องมี Program ด้วย
                course_map[key] = c
        
        for c in self.courses_to_schedule:
            if c.type == 'Lab':
                key = (c.major, c.program, c.code)
                if key in course_map:
                    c.related_course = course_map[key]

    def is_valid(self, course, day, start, room, squeeze=1.25, strict=True, lunch=False):
        end = start + course.duration
        
        # 1. Basic Constraints
        if end > 20: return False
        if not lunch and any(t == 12 for t in range(start, end)): return False
        if (room['capacity'] * squeeze) < course.students: return False
        
        # 2. Room Type
        if strict:
            if (course.type == 'Lab' and room['type'] != 'Lab') or (course.type == 'Lec' and room['type'] == 'Lab'): return False

        # 3. Busy Slots (ห้องไม่ว่างจากไฟล์ห้อง)
        for t in range(start, end):
            if (str(room['room_name']), day, t) in self.busy_slots: return False
            
        # 4. Conflict with Assigned Courses
        for c, (d, t, r) in self.assignment.items():
            if d == day:
                # ช่วงเวลาทับกัน
                if max(start, t) < min(end, t + c.duration):
                    # A. ห้องชน
                    if r == room['room_name']: return False
                    
                    # B. อาจารย์ชน (สอนพร้อมกันไม่ได้ ไม่ว่าจะภาคไหน)
                    inst_a = set(course.instructor.split(','))
                    inst_b = set(c.instructor.split(','))
                    if 'TBA' not in inst_a and 'TBA' not in inst_b:
                        if not inst_a.isdisjoint(inst_b): return False
                    
                    # C. นักศึกษาชน (กลุ่มเดียวกัน ห้ามเรียนซ้อน)
                    # ต้องเป็น Major เดียวกัน + Year เดียวกัน + Program เดียวกัน
                    if (course.major == c.major) and (course.year == c.year) and (course.program == c.program):
                        return False

        # 5. Lec < Lab Sequence
        if course.type == 'Lab' and course.related_course:
            lec = course.related_course
            if lec not in self.assignment: return False # Lec ต้องลงก่อน
            lec_day, _, _ = self.assignment[lec]
            d_idx = {d: i for i, d in enumerate(Config.DAYS)}
            
            # Lab ต้องไม่เรียนก่อน Lec (ในสัปดาห์)
            if d_idx[day] < d_idx[lec_day]: return False
            # ถ้าวันเดียวกัน ต้องเรียนทีหลัง
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
# 🛠️ HELPER: EXTRACT DATA (แยกภาคปกติ/พิเศษ)
# ==========================================
def extract_data_from_files(course_file, room_file):
    all_courses = []
    
    # --- 1. อ่านไฟล์หลักสูตร ---
    try:
        xls = pd.ExcelFile(course_file)
        for sheet in xls.sheet_names:
            major = next((m for m in Config.MAJOR_MAP if m in sheet.upper()), None)
            if not major: continue
            
            df = pd.read_excel(course_file, sheet_name=sheet, header=None)
            
            # หาเส้นแบ่งเทอม
            split_idx = len(df.columns) // 2
            for r in range(min(15, len(df))):
                row_txt = "".join([str(x) for x in df.iloc[r].values])
                if "ภาคปลาย" in row_txt or "การศึกษาที่ 2" in row_txt:
                    for c in range(len(df.columns)):
                        if "ภาคปลาย" in str(df.iloc[r, c]) or "การศึกษาที่ 2" in str(df.iloc[r, c]):
                            split_idx = c; break
                    break

            curr_year = 1
            for _, row in df.iterrows():
                # ตรวจจับปี
                row_str = " ".join([str(x) for x in row.values if str(x) != 'nan']).strip()
                if len(row_str) < 100 and "หน่วยกิต" not in row_str:
                    ym = re.search(r'ปี.*?(\d)', row_str)
                    if ym and 1 <= int(ym.group(1)) <= 4: curr_year = int(ym.group(1))

                # ฟังก์ชันย่อยแกะวิชา
                def parse_segment(seg, semester):
                    res = []
                    matches = list(re.finditer(r'\b([A-Z]{2}\s?\d{3}\s?\d{3})\b', seg))
                    for i, m in enumerate(matches):
                        code = m.group(1).replace(" ", "")
                        # Year Logic Check
                        y = curr_year
                        if len(code) >= 5 and code[4].isdigit():
                            dy = int(code[4])
                            if dy > y and dy <= 4: y = dy # เชื่อรหัสวิชาถ้าปีมากกว่า
                        
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
                            # หาจำนวนนศ. (ถ้ามี)
                            nums = [int(n) for n in re.findall(r'\b(\d{2,3})\b', meta) if 20 <= int(n) <= 200]
                            if nums: cnt = max(nums)
                        
                        name_clean = re.sub(r'\(.*?\)', '', name_raw).strip()
                        if len(name_clean) > 2:
                            # *** สร้าง 2 ใบ (Reg/Sp) แยกกัน ***
                            # ถ้าเจอจำนวน นศ. ให้ assume ว่าเป็นของภาคปกติ ส่วนภาคพิเศษลดลงหน่อย
                            # หรือถ้าในอนาคตไฟล์มีแยกคอลัมน์ชัดเจนค่อยแก้ตรงนี้
                            
                            # 1. ภาคปกติ
                            res.append(Course({
                                'major': major, 'program': 'Reg', 'year': y, 'semester': semester,
                                'code': code, 'name': name_clean, 'lec_hours': lec, 'lab_hours': lab,
                                'student_count': cnt if cnt > 0 else 50, 'instructor': instr
                            }))
                            
                            # 2. ภาคพิเศษ (สร้างด้วย เพื่อให้จัดแยกกัน)
                            res.append(Course({
                                'major': major, 'program': 'Sp', 'year': y, 'semester': semester,
                                'code': code, 'name': name_clean, 'lec_hours': lec, 'lab_hours': lab,
                                'student_count': 30, # สมมติภาคพิเศษคนน้อยกว่า
                                'instructor': instr
                            }))
                    return res

                s1 = " ".join([str(x) for x in row.iloc[:split_idx].values if str(x) != 'nan'])
                s2 = " ".join([str(x) for x in row.iloc[split_idx:].values if str(x) != 'nan'])
                all_courses.extend(parse_segment(s1, 1))
                all_courses.extend(parse_segment(s2, 2))
    except Exception as e: st.error(f"Course Read Error: {e}")

    # --- 2. อ่านไฟล์ห้อง ---
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
                for _, row in sch.iterrows():
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
    
    # Collect Majors/Years from assignments
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
            for program in ['Reg', 'Sp']: # Loop แยกภาค
                
                prog_name = 'ภาคปกติ' if program == 'Reg' else 'โครงการพิเศษ'
                sheet_name = f"{major}-{year}-{program}"
                
                # Filter Courses
                def get_data(sched):
                    return [
                        (c, d, t, r) for c, (d, t, r) in sched.assignment.items()
                        if c.major == major and c.year == year and c.program == program
                    ]

                # Draw Grid
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
                
                # Styling
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

                # Failed List
                failed = [c for s in all_scheds for c in s.failed_courses if c.major == major and c.year == year and c.program == program]
                if failed:
                    ws.cell(row=24, column=1, value="⚠️ วิชาที่ลงไม่ได้:").font = Font(color="FF0000", bold=True)
                    for i, f in enumerate(set(failed)): # set to dedup
                        ws.cell(row=25+i, column=1, value=f"{f.code} {f.name}")

    writer.close()
    output.seek(0)
    return output

# ==========================================
# 🖥️ APP INTERFACE
# ==========================================
st.set_page_config(page_title="KKU Scheduler", layout="wide")
st.title("🎓 ระบบจัดตารางเรียนอัตโนมัติ (KKU AI Scheduler)")
st.info("ระบบจะจัดตารางเรียนแยก ภาคปกติ และ โครงการพิเศษ ให้อัตโนมัติ")

c1, c2 = st.columns(2)
f_course = c1.file_uploader("1. ไฟล์หลักสูตร (kku30...)", type=['xlsx'])
f_room = c2.file_uploader("2. ไฟล์ห้องเรียน (LAB...)", type=['xlsx'])

if f_course and f_room:
    if st.button("🚀 เริ่มจัดตาราง", type="primary"):
        with st.spinner("⏳ กำลังประมวลผล..."):
            try:
                courses, rooms, busy = extract_data_from_files(f_course, f_room)
                if not courses or rooms.empty: st.error("❌ ไม่พบข้อมูลในไฟล์")
                else:
                    scheds = {}
                    for t in [1, 2]:
                        # กรองวิชาตามเทอม
                        term_courses = [c for c in courses if c.semester == t]
                        s = UniversityScheduler(term_courses, rooms, busy)
                        s.solve()
                        scheds[t] = s
                    
                    xls = generate_excel_report(scheds[1], scheds[2])
                    st.success("✅ เสร็จสมบูรณ์! (แยกภาคปกติ/พิเศษ เรียบร้อย)")
                    st.download_button("📥 ดาวน์โหลดไฟล์ Excel", xls, "Final_Schedule_Split.xlsx", 
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                     use_container_width=True)
            except Exception as e: st.error(f"Error: {e}")
