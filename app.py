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
    TIME_SLOTS = range(8, 20)

# ==========================================
# 📦 DATA MODELS
# ==========================================
class Course:
    def __init__(self, row_data, type='Lec'):
        self.majors = {row_data['major']} 
        self.codes = {row_data['code']} 
        self.name = row_data['name']
        self.instructor = str(row_data.get('instructor', 'TBA'))
        self.type = type 
        self.year = int(row_data.get('year', 1))
        
        raw_students = int(row_data.get('student_count', 40))
        self.students = 50 if (type == 'Lab' and raw_students > 50) else raw_students
            
        raw_duration = int(row_data['lec_hours']) if type == 'Lec' else int(row_data['lab_hours'])
        self.duration = min(raw_duration, 4)
        if self.duration == 0: self.duration = 2
        
        self.related_course = None 

    def merge(self, other_course):
        self.majors.update(other_course.majors)
        self.codes.update(other_course.codes)
        self.duration = max(self.duration, other_course.duration)
        self.students = max(self.students, other_course.students)

# ==========================================
# 🧠 SCHEDULER ENGINE (แก้ไข Bug นักศึกษาแยกร่าง)
# ==========================================
class UniversityScheduler:
    def __init__(self, courses_df, rooms_df, busy_df):
        self.rooms = sorted(rooms_df.to_dict('records'), key=lambda x: x['capacity'])
        self.busy_slots = self._process_busy_slots(busy_df)
        self.assignment = {} 
        self.failed_courses = [] 
        self.courses_to_schedule = self._prepare_and_merge_courses(courses_df)

    def _process_busy_slots(self, df):
        busy = set()
        if df.empty: return busy
        for _, row in df.iterrows():
            busy.add((str(row['room']), row['day'], row['hour']))
        return busy

    def _prepare_and_merge_courses(self, df):
        temp_dict = {} 
        for _, row in df.iterrows():
            if row['lec_hours'] > 0:
                key = (row['name'], 'Lec')
                if key not in temp_dict: temp_dict[key] = Course(row, 'Lec')
                else: temp_dict[key].merge(Course(row, 'Lec'))
            if row['lab_hours'] > 0:
                key = (row['name'], 'Lab')
                if key not in temp_dict: temp_dict[key] = Course(row, 'Lab')
                else: temp_dict[key].merge(Course(row, 'Lab'))
        
        final_courses = list(temp_dict.values())
        name_map = {c.name: c for c in final_courses if c.type == 'Lec'}
        for c in final_courses:
            if c.type == 'Lab' and c.name in name_map: c.related_course = name_map[c.name]
        
        final_courses.sort(key=lambda x: (-x.students, -x.duration))
        return final_courses

    def is_valid(self, course, day, start_time, room, squeeze_factor=1.25, strict_type=True, allow_lunch=False):
        end_time = start_time + course.duration
        
        # 1. เช็คเวลา (Time Limit & Lunch)
        if end_time > 20: return False
        if not allow_lunch and any(t == 12 for t in range(start_time, end_time)): return False

        # 2. เช็คความจุ (Capacity)
        if (room['capacity'] * squeeze_factor) < course.students: return False

        # 3. เช็คประเภทห้อง (Room Type)
        if strict_type:
            if (course.type == 'Lab' and room['type'] != 'Lab') or (course.type == 'Lec' and room['type'] == 'Lab'): return False

        # 4. เช็คห้องไม่ว่าง (Busy Slots)
        for t in range(start_time, end_time):
            if (str(room['room_name']), day, t) in self.busy_slots: return False
            
        # 5. เช็คการชนกันกับวิชาที่ลงไปแล้ว (Conflicts)
        for assigned_c, (a_day, a_time, a_room) in self.assignment.items():
            if a_day == day:
                a_end = a_time + assigned_c.duration
                # ถ้าเวลาซ้อนทับกัน
                if max(start_time, a_time) < min(end_time, a_end):
                    
                    # A. ห้องชน (Room Conflict)
                    if a_room == room['room_name']: return False 
                    
                    # B. อาจารย์ชน (Instructor Conflict)
                    inst_a = set(course.instructor.split(','))
                    inst_b = set(assigned_c.instructor.split(','))
                    if 'TBA' not in inst_a and 'TBA' not in inst_b:
                        if not inst_a.isdisjoint(inst_b): return False

                    # C. นักศึกษาชน (Student Conflict) *** เพิ่มใหม่ ***
                    # ถ้าสาขาเดียวกัน และ ชั้นปีเดียวกัน -> ห้ามเรียนเวลาเดียวกัน
                    if not course.majors.isdisjoint(assigned_c.majors):
                        if course.year == assigned_c.year:
                            return False

        # 6. เช็คลำดับ Lec < Lab
        if course.type == 'Lab' and course.related_course:
            lec = course.related_course
            if lec not in self.assignment: return False
            lec_day, _, _ = self.assignment[lec]
            days_order = {d: i for i, d in enumerate(Config.DAYS)}
            if days_order[day] < days_order[lec_day]: return False
            lec_end_time = self.assignment[lec][1] + lec.duration
            if days_order[day] == days_order[lec_day] and start_time < lec_end_time: return False
            
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
            rooms_to_try = self.rooms
            if not phase['strict']:
                rooms_to_try = sorted(self.rooms, key=lambda x: x['capacity'], reverse=True)

            for course in remaining:
                assigned = False
                for day in phase['days']:
                    if assigned: break
                    for hour in phase['hours']:
                        if assigned: break
                        if hour + course.duration > 20: continue
                        
                        for room in rooms_to_try:
                            if self.is_valid(course, day, hour, room, 
                                           phase['squeeze'], phase['strict'], phase['lunch']):
                                self.assignment[course] = (day, hour, room['room_name'])
                                assigned = True
                                break
        
        self.failed_courses = [c for c in self.courses_to_schedule if c not in self.assignment]
        return True

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================
def extract_data_from_files(course_file, room_file):
    all_courses = []
    try:
        xls = pd.ExcelFile(course_file)
        for sheet in xls.sheet_names:
            major = next((m for m in Config.MAJOR_MAP if m in sheet.upper()), None)
            if not major: continue
            df = pd.read_excel(course_file, sheet_name=sheet, header=None)
            
            split_idx = len(df.columns) // 2
            for r in range(min(15, len(df))):
                row_txt = "".join([str(x) for x in df.iloc[r].values])
                if "ภาคปลาย" in row_txt or "การศึกษาที่ 2" in row_txt:
                    for c in range(len(df.columns)):
                        if "ภาคปลาย" in str(df.iloc[r, c]) or "การศึกษาที่ 2" in str(df.iloc[r, c]):
                            split_idx = c; break
                    break

            current_year = 1
            for _, row in df.iterrows():
                row_str = " ".join([str(x) for x in row.values if str(x) != 'nan']).strip()
                if len(row_str) < 100 and "หน่วยกิต" not in row_str:
                    ym = re.search(r'ปี.*?(\d)', row_str)
                    if ym and 1 <= int(ym.group(1)) <= 4: current_year = int(ym.group(1))

                def parse_segment(seg, semester):
                    res = []
                    matches = list(re.finditer(r'\b([A-Z]{2}\s?\d{3}\s?\d{3})\b', seg))
                    for i, m in enumerate(matches):
                        code = m.group(1).replace(" ", "")
                        y = current_year
                        if len(code) >= 5 and code[4].isdigit() and int(code[4]) > y and int(code[4]) <= 4:
                            y = int(code[4])
                        
                        start, end = m.end(), matches[i+1].start() if i+1 < len(matches) else len(seg)
                        sub = seg[start:end]
                        
                        lec, lab, count, instr = 3, 0, 40, "TBA"
                        cr = re.search(r'(\d+)\s*\(\s*(\d+)\s*-\s*(\d+)\s*-\s*(\d+)\s*\)', sub)
                        name_raw = sub
                        
                        if cr:
                            name_raw = sub[:cr.start()]
                            lec, lab = int(cr.group(2)), int(cr.group(3))
                            meta = sub[cr.end():]
                            instrs = re.findall(r'\b[A-Z]{1,3}\d\b', meta)
                            if instrs: instr = ",".join(instrs)
                            nums = re.findall(r'\b(\d{2,3})\b', meta)
                            v_nums = [int(n) for n in nums if 20 <= int(n) <= 200]
                            if v_nums: count = max(v_nums)
                        
                        name_clean = re.sub(r'\(.*?\)', '', name_raw).strip()
                        if len(name_clean) > 2:
                            res.append({'major': major, 'year': y, 'semester': semester, 'code': code, 
                                        'name': name_clean, 'lec_hours': lec, 'lab_hours': lab, 
                                        'student_count': count, 'instructor': instr})
                    return res

                s1 = " ".join([str(x) for x in row.iloc[:split_idx].values if str(x) != 'nan'])
                s2 = " ".join([str(x) for x in row.iloc[split_idx:].values if str(x) != 'nan'])
                all_courses.extend(parse_segment(s1, 1))
                all_courses.extend(parse_segment(s2, 2))
    except Exception as e: st.error(f"Course File Error: {e}")

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
            
            start_row = next((i for i, r in df.iterrows() if "จันทร์" in str(r.values)), -1)
            if start_row != -1:
                sch = df.iloc[start_row:].reset_index(drop=True)
                d_map = {'จันทร์': 'Mon', 'อังคาร': 'Tue', 'พุธ': 'Wed', 'พฤหัส': 'Thu', 'ศุกร์': 'Fri', 'เสาร์': 'Sat', 'อาทิตย์': 'Sun'}
                for r_idx, row in sch.iterrows():
                    day_en = next((en for th, en in d_map.items() if th in str(row.iloc[0])), None)
                    if day_en:
                        for c_idx in range(1, len(row)):
                            if str(row.iloc[c_idx]).lower() != 'nan' and len(str(row.iloc[c_idx])) > 2:
                                if 8 + (c_idx - 1) <= 20:
                                    busy.append({'room': sheet.strip(), 'day': day_en, 'hour': 8 + (c_idx - 1)})

    except Exception as e: st.error(f"Room File Error: {e}")
    
    return pd.DataFrame(all_courses), pd.DataFrame(rooms), pd.DataFrame(busy)

def generate_excel_report(sched1, sched2):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    majors, years = set(), set()
    for s in [sched1, sched2]:
        for c in s.assignment: majors.add(list(c.majors)[0]); years.add(c.year)
    
    thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    
    for major in sorted(list(majors)):
        for year in sorted(list(years)):
            for program in ['ภาคปกติ', 'โครงการพิเศษ']:
                sheet_name = f"{major}-{year}-{'Reg' if program=='ภาคปกติ' else 'Sp'}"
                
                def make_grid(sched, start_row):
                    data = []
                    for c, (d, t, r) in sched.assignment.items():
                        if major in c.majors and c.year == year: data.append({'c': c, 'd': d, 't': t, 'r': r})
                    df = pd.DataFrame('', index=Config.DAYS, columns=[f"{h}:00-{h+1}:00" for h in Config.TIME_SLOTS])
                    colors = {}
                    for item in data:
                        c = item['c']
                        bg = Config.COLOR_MAP.get(list(c.codes)[0][:2].upper(), Config.COLOR_MAP['DEFAULT'])
                        txt = f"{'/'.join(c.codes)}\n{c.name}\n{item['r']} ({c.instructor})"
                        for i in range(c.duration):
                            h = item['t'] + i
                            if 8 <= h < 20:
                                col = f"{h}:00-{h+1}:00"
                                if item['d'] in df.index:
                                    prev = df.at[item['d'], col]
                                    df.at[item['d'], col] = (prev + "\n---\n" + txt).strip() if prev else txt
                                    colors[(Config.DAYS.index(item['d']) + start_row + 1, (h - 8) + 2)] = bg
                    return df, colors

                df1, col1 = make_grid(sched1, 4)
                df2, col2 = make_grid(sched2, 15)
                df1.to_excel(writer, sheet_name=sheet_name, startrow=3)
                df2.to_excel(writer, sheet_name=sheet_name, startrow=14)
                
                ws = writer.sheets[sheet_name]
                for r_head, txt in [(1, "ภาคต้น"), (13, "ภาคปลาย")]:
                    ws[f'A{r_head}'] = f"ตารางเรียน {Config.MAJOR_MAP.get(major, major)} ปี {year} ({program}) - {txt}"
                    ws.merge_cells(f'A{r_head}:M{r_head}')
                    ws[f'A{r_head}'].font = Font(size=14, bold=True); ws[f'A{r_head}'].alignment = Alignment(horizontal='center')

                ws.column_dimensions['A'].width = 12
                for i in range(2, 15): ws.column_dimensions[chr(64+i)].width = 18
                
                for r_start, r_end, c_map in [(4, 11, col1), (15, 22, col2)]:
                    for row in ws.iter_rows(min_row=r_start, max_row=r_end, min_col=1, max_col=13):
                        for cell in row:
                            cell.border = thin
                            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                            if (cell.row, cell.column) in c_map:
                                c = c_map[(cell.row, cell.column)]
                                cell.fill = PatternFill(start_color=c, end_color=c, fill_type='solid')
                            elif cell.row == r_start or cell.column == 1:
                                cell.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type='solid')
                                cell.font = Font(bold=True)
                
                # Check Failed Courses
                failed = [c for c in sched1.failed_courses + sched2.failed_courses if major in c.majors and c.year == year]
                if failed:
                    ws.cell(row=24, column=1, value="⚠️ รายวิชาที่จัดไม่ลง (FAILED):").font = Font(color="FF0000", bold=True)
                    for i, fc in enumerate(set(failed)):
                        ws.cell(row=25+i, column=1, value=f"{'/'.join(fc.codes)} {fc.name}")

    writer.close()
    output.seek(0)
    return output

# ==========================================
# 🖥️ APP INTERFACE
# ==========================================
st.set_page_config(page_title="KKU Scheduler", layout="wide")
st.title("🎓 ระบบจัดตารางเรียนอัตโนมัติ (KKU AI Scheduler)")
st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")

c1, c2 = st.columns(2)
f_course = c1.file_uploader("1. ไฟล์หลักสูตร (kku30...)", type=['xlsx'])
f_room = c2.file_uploader("2. ไฟล์ห้องเรียน (LAB...)", type=['xlsx'])

if f_course and f_room:
    if st.button("🚀 เริ่มจัดตาราง (Start)", type="primary"):
        with st.spinner("⏳ กำลังประมวลผล... (ระบบ AI กำลังจัดตาราง 2 เทอม)"):
            try:
                df_c, df_r, df_b = extract_data_from_files(f_course, f_room)
                if df_c.empty or df_r.empty: st.error("❌ ไม่พบข้อมูลในไฟล์!")
                else:
                    schedulers = {}
                    for term in [1, 2]:
                        sched = UniversityScheduler(df_c[df_c['semester'] == term], df_r, df_b)
                        sched.solve()
                        schedulers[term] = sched
                    
                    excel_file = generate_excel_report(schedulers[1], schedulers[2])
                    st.success("✅ เสร็จสมบูรณ์!")
                    
                    m1, m2 = st.columns(2)
                    m1.metric("ภาคต้น (Sem 1)", f"{len(schedulers[1].assignment)} วิชา")
                    m2.metric("ภาคปลาย (Sem 2)", f"{len(schedulers[2].assignment)} วิชา")
                    
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel (ตารางเรียน)",
                        data=excel_file,
                        file_name="Final_Schedule.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e: st.error(f"เกิดข้อผิดพลาด: {e}")
