import sqlite3
from datetime import datetime, timedelta
import csv

def setup_timetable():
    """Initialize timetable database and load data"""
    print("\n🚀 Setting up timetable system...")
    
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    
    # CLASS SCHEDULE TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS class_schedule
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  student_group TEXT,
                  day_of_week TEXT,
                  subject_name TEXT,
                  teacher_name TEXT,
                  start_time TEXT,
                  end_time TEXT,
                  room_number TEXT,
                  urdu_type TEXT DEFAULT 'both',
                  is_active INTEGER DEFAULT 1,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    
    # SUBJECTS TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS subjects
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subject_code TEXT UNIQUE,
                  subject_name TEXT,
                  semester INTEGER,
                  department TEXT,
                  credits INTEGER)''')
    
    # TEACHERS TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS teachers
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_code TEXT UNIQUE,
                  teacher_name TEXT,
                  department TEXT,
                  email TEXT,
                  specialization TEXT)''')
    
    # CLASS UPDATES TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS class_updates
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  update_type TEXT,
                  subject_name TEXT,
                  class_type TEXT DEFAULT 'both',
                  target_date DATE,
                  original_time TEXT,
                  new_time TEXT,
                  room_change TEXT,
                  reason TEXT,
                  posted_by TEXT,
                  posted_at TIMESTAMP,
                  expires_at TIMESTAMP,
                  is_active INTEGER DEFAULT 1)''')
    
    # CR TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS class_representatives
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  student_id TEXT UNIQUE,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  appointed_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    
    # SYLLABUS TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS syllabus
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subject_code TEXT,
                  subject_name TEXT,
                  semester INTEGER,
                  unit_number INTEGER,
                  unit_name TEXT,
                  topics TEXT,
                  practical_components TEXT,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # ✅ FIX: UNIQUE constraint to prevent duplicate rows
    c.execute('''CREATE UNIQUE INDEX IF NOT EXISTS
                 idx_syllabus_unique ON syllabus(subject_code, unit_number)''')
    
    # SYLLABUS PROGRESS TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS syllabus_progress
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subject_code TEXT,
                  unit_number INTEGER,
                  is_completed INTEGER DEFAULT 0,
                  completed_at TIMESTAMP,
                  completed_by TEXT,
                  UNIQUE(subject_code, unit_number))''')
    
    # CLASS HISTORY TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS class_history
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  student_group TEXT,
                  subject_name TEXT,
                  class_type TEXT,
                  unit_covered TEXT,
                  topics_covered TEXT,
                  practical_work TEXT,
                  demonstrations TEXT,
                  key_points TEXT,
                  homework_assigned TEXT,
                  posted_by TEXT,
                  posted_at TIMESTAMP,
                  class_date DATE,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  UNIQUE(course, department, semester, student_group, subject_name, class_type, class_date))''')
    
    # HOMEWORK TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS homework
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  student_group TEXT,
                  subject_name TEXT,
                  homework_description TEXT,
                  submission_date DATE,
                  submission_time TEXT,
                  posted_by TEXT,
                  posted_at TIMESTAMP,
                  is_completed INTEGER DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    
    # HOMEWORK REMINDERS SENT TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS homework_reminders_sent
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  homework_id INTEGER,
                  sent_date DATE,
                  UNIQUE(homework_id, sent_date))''')
    
    # ATTENDANCE RECORDS TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS attendance_records
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  student_id TEXT,
                  subject_name TEXT,
                  month TEXT,
                  year INTEGER,
                  classes_held INTEGER,
                  classes_attended INTEGER,
                  attendance_percentage REAL,
                  updated_at TIMESTAMP,
                  UNIQUE(student_id, subject_name, month, year))''')
    
    # ATTENDANCE ALERTS TABLE
    c.execute('''CREATE TABLE IF NOT EXISTS attendance_alerts
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  student_id TEXT,
                  subject_name TEXT,
                  alert_type TEXT,
                  alert_message TEXT,
                  sent_at TIMESTAMP)''')
    
    conn.commit()
    print("✅ Timetable database initialized!")
    
    # Load data
    load_subjects_data(conn)
    load_teachers_data(conn)
    load_timetable_data(conn)
    load_cr_data(conn)
    load_syllabus_data(conn)
    load_attendance_data('attendance_january_2026.csv', 'January', 2026)
    load_attendance_data('attendance_february_2026.csv', 'February', 2026)
    load_attendance_data('attendance_march_2026.csv', 'March', 2026)
    
    conn.close()
    print("✅ Timetable system ready!\n")


def load_subjects_data(conn):
    """Load subjects from CSV"""
    try:
        c = conn.cursor()
        with open('subjects_data.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                c.execute('''INSERT OR REPLACE INTO subjects 
                             (subject_code, subject_name, semester, department, credits)
                             VALUES (?, ?, ?, ?, ?)''',
                          (row['subject_code'].strip(),
                           row['subject_name'].strip(),
                           int(row['semester']),
                           row['department'].strip(),
                           int(row.get('credits', 4))))
        conn.commit()
        print("✅ Subjects data loaded!")
    except FileNotFoundError:
        print("⚠️ subjects_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading subjects: {e}")


def load_teachers_data(conn):
    """Load teachers from CSV"""
    try:
        c = conn.cursor()
        with open('teachers_data.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                teacher_code = row.get('teacher_id', row.get('teacher_code', '')).strip()
                c.execute('''INSERT OR REPLACE INTO teachers 
                             (teacher_code, teacher_name, department, email, specialization)
                             VALUES (?, ?, ?, ?, ?)''',
                          (teacher_code,
                           row['teacher_name'].strip(),
                           row['department'].strip(),
                           row.get('email', '').strip(),
                           row.get('subjects_taught', row.get('specialization', '')).strip()))
        conn.commit()
        print("✅ Teachers data loaded!")
    except FileNotFoundError:
        print("⚠️ teachers_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading teachers: {e}")


def load_timetable_data(conn):
    """Load timetable from CSV"""
    try:
        c = conn.cursor()
        with open('timetable_data.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                urdu_type = row.get('urdu_type', 'both').strip().lower()
                if urdu_type not in ['regular', 'advanced', 'both']:
                    urdu_type = 'both'
                c.execute('''INSERT OR REPLACE INTO class_schedule 
                             (course, department, semester, student_group, day_of_week, 
                              subject_name, teacher_name, start_time, end_time, room_number, urdu_type)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                          (row['course'].strip(),
                           row['department'].strip(),
                           int(row['semester']),
                           row.get('group', 'A').strip().upper(),
                           row['day_of_week'].strip(),
                           row['subject_name'].strip(),
                           row['teacher_name'].strip(),
                           row['start_time'].strip(),
                           row['end_time'].strip(),
                           row.get('room_number', 'TBA').strip(),
                           urdu_type))
        conn.commit()
        print("✅ Timetable data loaded!")
    except FileNotFoundError:
        print("⚠️ timetable_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading timetable: {e}")


def load_cr_data(conn):
    """Load Class Representatives from CSV"""
    try:
        c = conn.cursor()
        with open('cr_data.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                c.execute('''INSERT OR REPLACE INTO class_representatives 
                             (student_id, course, department, semester)
                             VALUES (?, ?, ?, ?)''',
                          (row['student_id'].strip().upper(),
                           row['course'].strip(),
                           row['department'].strip(),
                           int(row['semester'])))
        conn.commit()
        print("✅ CR data loaded!")
    except FileNotFoundError:
        print("⚠️ cr_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading CR data: {e}")


def load_syllabus_data(conn):
    """Load syllabus from CSV"""
    try:
        c = conn.cursor()

        # ✅ FIX: Clear existing data before reload to prevent duplicates
        c.execute('DELETE FROM syllabus')
        print("🗑️ Cleared existing syllabus data")

        with open('syllabus_data.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                subject_code = row['subject_code'].strip()
                try:
                    semester = int(subject_code.split('-')[1][0])
                except:
                    semester = 2
                
                practical = row.get('practical_component', row.get('practical_components', '')).strip()
                
                # ✅ INSERT OR REPLACE now works correctly with UNIQUE index
                c.execute('''INSERT OR REPLACE INTO syllabus 
                             (subject_code, subject_name, semester, unit_number, 
                              unit_name, topics, practical_components)
                             VALUES (?, ?, ?, ?, ?, ?, ?)''',
                          (subject_code,
                           row['subject_name'].strip(),
                           semester,
                           int(row['unit_number']),
                           row['unit_name'].strip(),
                           row.get('topics', '').strip(),
                           practical))
        
        conn.commit()
        print("✅ Syllabus data loaded!")
    except FileNotFoundError:
        print("⚠️ syllabus_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading syllabus: {e}")


def load_attendance_data(csv_file, month, year):
    """Load attendance from CSV"""
    try:
        conn = sqlite3.connect('students.db', timeout=20)
        c = conn.cursor()
        
        with open(csv_file, 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            count = 0
            for row in csv_reader:
                student_id = row['student_id'].strip().upper()
                subject_name = row['subject_name'].strip()
                classes_held = int(row['classes_held'])
                classes_attended = int(row['classes_attended'])
                percentage = (classes_attended / classes_held * 100) if classes_held > 0 else 0
                
                c.execute('''INSERT OR REPLACE INTO attendance_records
                             (student_id, subject_name, month, year, 
                              classes_held, classes_attended, attendance_percentage, updated_at)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                          (student_id, subject_name, month, year,
                           classes_held, classes_attended, percentage, datetime.now()))
                count += 1
        
        conn.commit()
        conn.close()
        print(f"✅ Attendance data loaded for {month} {year}! ({count} records)")
        return True
    except FileNotFoundError:
        print(f"⚠️ {csv_file} not found.")
        return False
    except Exception as e:
        print(f"❌ Error loading attendance: {e}")
        return False


# ==========================================
# TEACHER TO SUBJECT MAPPING
# ==========================================

def get_subject_from_teacher(teacher_name):
    """Map teacher name to subject"""
    teacher_lower = teacher_name.lower()
    
    if any(x in teacher_lower for x in ['akash', 'chef akash', 'dr. akash', 'dr akash']):
        return 'Food Production Foundation - II'
    if any(x in teacher_lower for x in ['mohit', 'mr. mohit', 'mr mohit', 'mohit sir']):
        return 'Food & Beverage Service Foundation - II'
    if any(x in teacher_lower for x in ['aarti', 'dr. aarti', 'dr aarti', 'aarti maam']):
        return None
    if any(x in teacher_lower for x in ['jaya', 'chef jaya', 'jaya maam']):
        return 'Personality Development and Grooming'
    return None


# ==========================================
# SMART SUBJECT DETECTION
# ==========================================

def detect_subject_name(text):
    """Intelligently detect subject from text"""
    text_lower = text.lower()
    
    food_prod_keywords = [
        'food production', 'food prod', 'fp', 'bhm-201', 'bhm 201',
        'cooking', 'chef', 'kitchen', 'chicken', 'fish', 'meat', 'egg',
        'soup', 'poultry', 'pastry', 'bread', 'cuts', 'frying', 'grilling',
        'akash sir', 'chef akash', 'dr. akash', 'dr akash'
    ]
    fb_keywords = [
        'f&b', 'f and b', 'fnb', 'food and beverage', 'food & beverage',
        'beverage service', 'fb service', 'bhm-202', 'bhm 202',
        'table setting', 'service', 'menu', 'restaurant', 'buffet',
        'gueridon', 'silver service', 'french service', 'journal',
        'mohit sir', 'mr. mohit', 'mr mohit'
    ]
    housekeeping_keywords = [
        'housekeeping', 'house keeping', 'hk', 'bhm-203', 'bhm 203',
        'accommodation', 'cleaning', 'room', 'linen', 'bed making',
        'guestroom', 'bathroom', 'public area', 'turndown',
        'aarti maam housekeeping', 'aarti hk'
    ]
    front_office_keywords = [
        'front office', 'fo', 'bhm-204', 'bhm 204',
        'reception', 'reservation', 'check-in', 'checkout',
        'registration', 'guest cycle', 'telephone exchange',
        'aarti maam front', 'aarti fo'
    ]
    personality_keywords = [
        'personality', 'personality development', 'grooming', 'pd',
        'communication', 'body language', 'jaya maam', 'chef jaya'
    ]
    
    if any(kw in text_lower for kw in food_prod_keywords):
        return 'Food Production Foundation - II'
    if any(kw in text_lower for kw in fb_keywords):
        return 'Food & Beverage Service Foundation - II'
    if any(kw in text_lower for kw in housekeeping_keywords):
        return 'Housekeeping Skills - II'
    if any(kw in text_lower for kw in front_office_keywords):
        return 'Front Office Foundation - II'
    if any(kw in text_lower for kw in personality_keywords):
        return 'Personality Development and Grooming'
    
    teacher_subject = get_subject_from_teacher(text)
    if teacher_subject:
        return teacher_subject
    return None


# ==========================================
# SCHEDULE FUNCTIONS
# ==========================================

def get_today_schedule(course, department, semester, student_group):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    today = datetime.now().strftime('%A')
    c.execute('''SELECT subject_name, teacher_name, start_time, end_time, room_number, student_group
                 FROM class_schedule
                 WHERE course = ? AND department = ? AND semester = ? 
                 AND day_of_week = ? 
                 AND (student_group = ? OR student_group = 'BOTH')
                 AND is_active = 1
                 ORDER BY start_time''',
              (course, department, semester, today, student_group))
    results = c.fetchall()
    conn.close()
    return [{'subject': r[0], 'teacher': r[1], 'start_time': r[2],
             'end_time': r[3], 'room': r[4], 'group': r[5]} for r in results]


def get_next_class(course, department, semester, student_group):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    today = datetime.now().strftime('%A')
    current_time = datetime.now().strftime('%H:%M')
    c.execute('''SELECT subject_name, teacher_name, start_time, end_time, room_number
                 FROM class_schedule
                 WHERE course = ? AND department = ? AND semester = ? 
                 AND day_of_week = ? 
                 AND (student_group = ? OR student_group = 'BOTH')
                 AND start_time > ? 
                 AND is_active = 1
                 ORDER BY start_time LIMIT 1''',
              (course, department, semester, today, student_group, current_time))
    result = c.fetchone()
    conn.close()
    if result:
        return {'subject': result[0], 'teacher': result[1], 'start_time': result[2],
                'end_time': result[3], 'room': result[4]}
    return None


def get_subject_schedule(course, department, semester, subject_name):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    c.execute('''SELECT day_of_week, start_time, end_time, teacher_name, room_number
                 FROM class_schedule
                 WHERE course = ? AND department = ? AND semester = ? 
                 AND subject_name = ? AND is_active = 1
                 ORDER BY CASE day_of_week
                     WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3
                     WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5
                     WHEN 'Saturday' THEN 6 WHEN 'Sunday' THEN 7
                 END, start_time''',
              (course, department, semester, subject_name))
    results = c.fetchall()
    conn.close()
    return [{'day': r[0], 'start_time': r[1], 'end_time': r[2],
             'teacher': r[3], 'room': r[4]} for r in results]


def format_schedule_for_ai(schedule):
    if not schedule:
        return "No classes scheduled for today"
    text = ""
    for cls in schedule:
        group_text = f" (Group {cls['group']})" if cls['group'] != 'BOTH' else ""
        text += f"• {cls['start_time']}-{cls['end_time']}: {cls['subject']} ({cls['teacher']}){group_text} - {cls['room']}\n"
    return text


# ==========================================
# CR FUNCTIONS
# ==========================================

def is_cr(student_id):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    c.execute('SELECT course, department, semester FROM class_representatives WHERE student_id = ?',
              (student_id.upper(),))
    result = c.fetchone()
    conn.close()
    if result:
        return {'is_cr': True, 'course': result[0], 'department': result[1], 'semester': result[2]}
    return {'is_cr': False}


def post_class_update(cr_student_id, update_type, subject_name, original_time=None,
                     new_time=None, room_change=None, reason=None, class_type='both', target_date=None):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    cr_info = is_cr(cr_student_id)
    if not cr_info['is_cr']:
        conn.close()
        return {'success': False, 'message': 'Not authorized as CR'}
    if target_date is None:
        target_date = datetime.now().date()
    expires_at = datetime.combine(target_date, datetime.max.time())
    c.execute('''INSERT INTO class_updates 
                 (course, department, semester, update_type, subject_name, class_type,
                  target_date, original_time, new_time, room_change, reason, posted_by, posted_at, expires_at)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (cr_info['course'], cr_info['department'], cr_info['semester'],
               update_type, subject_name, class_type, target_date,
               original_time, new_time, room_change,
               reason, cr_student_id, datetime.now(), expires_at))
    conn.commit()
    conn.close()
    return {'success': True, 'course': cr_info['course'],
            'department': cr_info['department'], 'semester': cr_info['semester']}


def get_active_updates(course, department, semester):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    today = datetime.now().date()
    c.execute('''SELECT update_type, subject_name, class_type, target_date,
                        original_time, new_time, room_change, reason, posted_at
                 FROM class_updates
                 WHERE course = ? AND department = ? AND semester = ?
                 AND is_active = 1 AND target_date >= ?
                 ORDER BY target_date, posted_at DESC''',
              (course, department, semester, today))
    results = c.fetchall()
    conn.close()
    return [{'type': r[0], 'subject': r[1], 'class_type': r[2], 'target_date': r[3],
             'original_time': r[4], 'new_time': r[5], 'room_change': r[6],
             'reason': r[7], 'posted_at': r[8]} for r in results]


def format_updates_for_ai(updates):
    if not updates:
        return "No active updates"
    text = "🔥 ACTIVE CLASS UPDATES (CRITICAL - CHECK THESE FIRST!):\n"
    for update in updates:
        try:
            if isinstance(update['target_date'], str):
                date_obj = datetime.fromisoformat(update['target_date'])
            else:
                date_obj = update['target_date']
            date_str = date_obj.strftime('%A, %B %d')
        except:
            date_str = str(update['target_date'])
        if update['type'] == 'cancelled':
            text += f"❌ CANCELLED on {date_str}: {update.get('class_type','ALL').upper()} - {update.get('subject','ALL CLASSES')}\n"
        elif update['type'] == 'postponed':
            text += f"⏰ POSTPONED on {date_str}: {update['subject']} → New time: {update.get('new_time','TBD')}\n"
        elif update['type'] == 'room_change':
            text += f"🔄 ROOM CHANGE on {date_str}: {update['subject']} → {update.get('room_change','TBD')}\n"
        elif update['type'] == 'extra_class':
            text += f"➕ EXTRA CLASS on {date_str}: {update['subject']} at {update.get('new_time','TBD')}\n"
    return text


# ==========================================
# SYLLABUS FUNCTIONS
# ==========================================

def get_subject_syllabus(semester):
    """Get complete syllabus for a semester — deduped"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT subject_code, subject_name, unit_number, unit_name, 
                            topics, practical_components
                     FROM syllabus
                     WHERE semester = ?
                     ORDER BY subject_code, unit_number''', (semester,))
        results = c.fetchall()
        conn.close()

        subjects = {}
        seen_units = {}  # ✅ FIX: track seen (subject_code, unit_number) pairs
        for row in results:
            subject_code = row[0]
            unit_number = row[2]
            unit_key = (subject_code, unit_number)

            if unit_key in seen_units:
                continue  # skip duplicate
            seen_units[unit_key] = True

            if subject_code not in subjects:
                subjects[subject_code] = {
                    'subject_code': subject_code,
                    'subject_name': row[1],
                    'units': []
                }
            subjects[subject_code]['units'].append({
                'unit_number': unit_number,
                'unit_name': row[3],
                'topics': row[4],
                'practical_components': row[5]
            })
        return list(subjects.values())
    except Exception as e:
        conn.close()
        print(f"Error getting syllabus: {e}")
        return []


def get_syllabus_progress(semester):
    """Get completion status for all topics"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT s.subject_code, s.unit_number, sp.is_completed
                     FROM syllabus s
                     LEFT JOIN syllabus_progress sp ON 
                         s.subject_code = sp.subject_code AND 
                         s.unit_number = sp.unit_number
                     WHERE s.semester = ?''', (semester,))
        results = c.fetchall()
        conn.close()
        progress = {}
        for row in results:
            key = f"{row[0]}-{row[1]}"
            if key not in progress:  # ✅ FIX: don't overwrite with duplicate
                progress[key] = bool(row[2]) if row[2] is not None else False
        return progress
    except Exception as e:
        conn.close()
        print(f"Error getting progress: {e}")
        return {}


def get_next_topic(semester):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT s.subject_code, s.subject_name, s.unit_number, s.unit_name
                     FROM syllabus s
                     LEFT JOIN syllabus_progress sp ON 
                         s.subject_code = sp.subject_code AND 
                         s.unit_number = sp.unit_number
                     WHERE s.semester = ? AND (sp.is_completed IS NULL OR sp.is_completed = 0)
                     ORDER BY s.subject_code, s.unit_number''', (semester,))
        results = c.fetchall()
        conn.close()
        next_topics = []
        seen_subjects = set()
        for row in results:
            subject_code = row[0]
            if subject_code not in seen_subjects:
                next_topics.append({'subject_code': subject_code, 'subject_name': row[1],
                                    'unit_number': row[2], 'unit_name': row[3]})
                seen_subjects.add(subject_code)
        return next_topics
    except Exception as e:
        conn.close()
        print(f"Error getting next topics: {e}")
        return []


def format_syllabus_for_ai(syllabus_data, progress_data):
    if not syllabus_data:
        return "No syllabus data available"
    text = ""
    for subject in syllabus_data:
        text += f"\n{subject['subject_name']} ({subject['subject_code']}):\n"
        for unit in subject['units']:
            unit_key = f"{subject['subject_code']}-{unit['unit_number']}"
            status = "✅" if progress_data.get(unit_key, False) else "⬜"
            text += f"{status} Unit {unit['unit_number']}: {unit['unit_name']}\n"
            if unit['topics']:
                text += f"   Topics: {unit['topics']}\n"
            if unit['practical_components']:
                text += f"   Practical: {unit['practical_components']}\n"
        text += "\n"
    return text


def mark_topic_complete(semester, subject_identifier, unit_number, completed_by_student_id):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        detected_subject = detect_subject_name(subject_identifier)
        subject_code = None
        if subject_identifier.upper().startswith('BHM-'):
            subject_code = subject_identifier.upper()
        elif detected_subject:
            c.execute('''SELECT subject_code FROM syllabus
                         WHERE semester = ? AND subject_name = ? LIMIT 1''',
                      (semester, detected_subject))
            result = c.fetchone()
            if result:
                subject_code = result[0]
        else:
            c.execute('SELECT subject_code, subject_name FROM syllabus WHERE semester = ?', (semester,))
            all_subjects = c.fetchall()
            for code, name in all_subjects:
                if subject_identifier.lower() in name.lower():
                    subject_code = code
                    break
        if not subject_code:
            conn.close()
            return {'success': False, 'message': f'Subject "{subject_identifier}" not found'}
        c.execute('''SELECT subject_name, unit_name FROM syllabus
                     WHERE semester = ? AND subject_code = ? AND unit_number = ?''',
                  (semester, subject_code, unit_number))
        result = c.fetchone()
        if not result:
            conn.close()
            return {'success': False, 'message': f'Unit {unit_number} not found'}
        subject_name, unit_name = result
        c.execute('''INSERT OR REPLACE INTO syllabus_progress
                     (subject_code, unit_number, is_completed, completed_at, completed_by)
                     VALUES (?, ?, 1, ?, ?)''',
                  (subject_code, unit_number, datetime.now(), completed_by_student_id))
        conn.commit()
        conn.close()
        return {'success': True, 'subject_name': subject_name,
                'unit_name': unit_name, 'unit_number': unit_number}
    except Exception as e:
        conn.close()
        print(f"Error marking topic complete: {e}")
        return {'success': False, 'message': str(e)}


def get_exam_tips(semester):
    tips = {
        2: """Focus on:
- Food Production: Master classification systems, cooking methods, stocks preparation
- F&B Service: Table setting standards, service styles, beverage knowledge
- Front Office: Registration procedures, reservation systems, guest cycle
- Housekeeping: Cleaning procedures, linen management, room inspection standards
- Practice practical components regularly - they're 50% of your grade!"""
    }
    return tips.get(semester, "Study all units thoroughly and practice regularly!")


# ==========================================
# CLASS HISTORY FUNCTIONS
# ==========================================

def save_last_theory_class(cr_student_id, subject_identifier, unit_covered, topics_covered,
                           key_points, homework_assigned=None):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    cr_info = is_cr(cr_student_id)
    if not cr_info['is_cr']:
        conn.close()
        return {'success': False, 'message': 'Not authorized as CR'}
    try:
        subject_name = detect_subject_name(subject_identifier)
        if not subject_name:
            search_term = f"%{subject_identifier}%"
            c.execute('''SELECT subject_name FROM syllabus
                         WHERE semester = ? AND (subject_name LIKE ? OR subject_code LIKE ?) LIMIT 1''',
                      (cr_info['semester'], search_term, search_term))
            result = c.fetchone()
            subject_name = result[0] if result else subject_identifier
        class_date = datetime.now().date()
        c.execute('''INSERT OR REPLACE INTO class_history 
                     (course, department, semester, student_group, subject_name, 
                      class_type, unit_covered, topics_covered, key_points, 
                      homework_assigned, posted_by, posted_at, class_date)
                     VALUES (?, ?, ?, ?, ?, 'theory', ?, ?, ?, ?, ?, ?, ?)''',
                  (cr_info['course'], cr_info['department'], cr_info['semester'],
                   'BOTH', subject_name, unit_covered, topics_covered, key_points,
                   homework_assigned, cr_student_id, datetime.now(), class_date))
        conn.commit()
        conn.close()
        return {'success': True, 'subject': subject_name, 'unit': unit_covered,
                'date': class_date.strftime('%B %d, %Y')}
    except Exception as e:
        conn.close()
        print(f"Error saving theory class: {e}")
        return {'success': False, 'message': str(e)}


def save_last_practical_class(cr_student_id, student_group, subject_identifier, unit_covered,
                              practical_work, demonstrations, key_points):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    cr_info = is_cr(cr_student_id)
    if not cr_info['is_cr']:
        conn.close()
        return {'success': False, 'message': 'Not authorized as CR'}
    try:
        subject_name = detect_subject_name(subject_identifier)
        if not subject_name:
            if 'aarti' in subject_identifier.lower():
                subject_name = 'Housekeeping Skills - II'
            else:
                search_term = f"%{subject_identifier}%"
                c.execute('''SELECT subject_name FROM syllabus
                             WHERE semester = ? AND (subject_name LIKE ? OR subject_code LIKE ?) LIMIT 1''',
                          (cr_info['semester'], search_term, search_term))
                result = c.fetchone()
                subject_name = result[0] if result else subject_identifier
        class_date = datetime.now().date()
        c.execute('''INSERT OR REPLACE INTO class_history 
                     (course, department, semester, student_group, subject_name, 
                      class_type, unit_covered, practical_work, demonstrations, 
                      key_points, posted_by, posted_at, class_date)
                     VALUES (?, ?, ?, ?, ?, 'practical', ?, ?, ?, ?, ?, ?, ?)''',
                  (cr_info['course'], cr_info['department'], cr_info['semester'],
                   student_group, subject_name, unit_covered, practical_work,
                   demonstrations, key_points, cr_student_id, datetime.now(), class_date))
        conn.commit()
        conn.close()
        return {'success': True, 'subject': subject_name, 'group': student_group,
                'unit': unit_covered, 'date': class_date.strftime('%B %d, %Y')}
    except Exception as e:
        conn.close()
        print(f"Error saving practical class: {e}")
        return {'success': False, 'message': str(e)}


def get_class_history(course, department, semester, student_group, limit=10):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT subject_name, class_type, unit_covered, topics_covered,
                            practical_work, demonstrations, key_points, class_date,
                            student_group, homework_assigned
                     FROM class_history
                     WHERE course = ? AND department = ? AND semester = ?
                     AND (student_group = ? OR student_group = 'BOTH')
                     ORDER BY class_date DESC, posted_at DESC LIMIT ?''',
                  (course, department, semester, student_group, limit))
        results = c.fetchall()
        conn.close()
        return [{'subject': r[0], 'class_type': r[1], 'unit': r[2], 'topics': r[3],
                 'practical': r[4], 'demonstrations': r[5], 'key_points': r[6],
                 'date': r[7], 'group': r[8], 'homework': r[9]} for r in results]
    except Exception as e:
        conn.close()
        print(f"Error getting class history: {e}")
        return []


def format_class_history_for_ai(class_history):
    if not class_history:
        return "No class history available yet"
    text = "\n📝 RECENT CLASS HISTORY:\n"
    for i, cls in enumerate(class_history[:5], 1):
        try:
            if isinstance(cls['date'], str):
                date_obj = datetime.fromisoformat(cls['date'])
            else:
                date_obj = cls['date']
            date_str = date_obj.strftime('%d %b %Y')
        except:
            date_str = "Recent"
        class_emoji = "📚" if cls['class_type'] == 'theory' else "🔬"
        group_text = f" (Group {cls['group']})" if cls['group'] != 'BOTH' else ""
        text += f"\n{i}. {class_emoji} {date_str} - {cls['subject']} {cls['class_type'].upper()}{group_text}\n"
        if cls['topics']:
            text += f"   Topics: {cls['topics']}\n"
        elif cls['practical']:
            text += f"   Practical: {cls['practical']}\n"
    return text


# ==========================================
# HOMEWORK FUNCTIONS
# ==========================================

def save_homework(cr_student_id, subject_identifier, homework_description, submission_date,
                  submission_time=None, student_group='BOTH'):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    cr_info = is_cr(cr_student_id)
    if not cr_info['is_cr']:
        conn.close()
        return {'success': False, 'message': 'Not authorized as CR'}
    try:
        subject_name = detect_subject_name(subject_identifier)
        if not subject_name:
            search_term = f"%{subject_identifier}%"
            c.execute('''SELECT subject_name FROM syllabus
                         WHERE semester = ? AND (subject_name LIKE ? OR subject_code LIKE ?) LIMIT 1''',
                      (cr_info['semester'], search_term, search_term))
            result = c.fetchone()
            subject_name = result[0] if result else subject_identifier
        c.execute('''INSERT INTO homework 
                     (course, department, semester, student_group, subject_name, 
                      homework_description, submission_date, submission_time, posted_by, posted_at)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (cr_info['course'], cr_info['department'], cr_info['semester'],
                   student_group, subject_name, homework_description,
                   submission_date, submission_time, cr_student_id, datetime.now()))
        homework_id = c.lastrowid
        conn.commit()
        conn.close()
        return {'success': True, 'homework_id': homework_id, 'subject': subject_name,
                'submission_date': submission_date, 'course': cr_info['course'],
                'department': cr_info['department'], 'semester': cr_info['semester']}
    except Exception as e:
        conn.close()
        print(f"Error saving homework: {e}")
        return {'success': False, 'message': str(e)}


def get_pending_homework(course, department, semester, student_group):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        today = datetime.now().date()
        c.execute('''SELECT id, subject_name, homework_description, submission_date, submission_time
                     FROM homework
                     WHERE course = ? AND department = ? AND semester = ?
                     AND (student_group = ? OR student_group = 'BOTH')
                     AND is_completed = 0 AND submission_date >= ?
                     ORDER BY submission_date, submission_time''',
                  (course, department, semester, student_group, today))
        results = c.fetchall()
        conn.close()
        return [{'id': r[0], 'subject': r[1], 'description': r[2],
                 'submission_date': r[3], 'submission_time': r[4]} for r in results]
    except Exception as e:
        conn.close()
        print(f"Error getting homework: {e}")
        return []


def format_homework_for_ai(homework_list):
    if not homework_list:
        return "No pending homework"
    text = "\n📝 PENDING HOMEWORK:\n"
    for hw in homework_list:
        try:
            if isinstance(hw['submission_date'], str):
                date_obj = datetime.fromisoformat(hw['submission_date'])
            else:
                date_obj = hw['submission_date']
            date_str = date_obj.strftime('%d %B %Y (%A)')
        except:
            date_str = str(hw['submission_date'])
        time_str = f" at {hw['submission_time']}" if hw['submission_time'] else ""
        text += f"• {hw['subject']}: {hw['description']}\n"
        text += f"  Due: {date_str}{time_str}\n"
    return text


def get_homework_for_reminder():
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        today = datetime.now().date()
        c.execute('''SELECT h.id, h.course, h.department, h.semester, h.student_group,
                            h.subject_name, h.homework_description, h.submission_date, h.submission_time
                     FROM homework h
                     LEFT JOIN homework_reminders_sent hrs ON h.id = hrs.homework_id AND hrs.sent_date = ?
                     WHERE h.is_completed = 0 AND h.submission_date >= ? AND hrs.id IS NULL''',
                  (today, today))
        results = c.fetchall()
        conn.close()
        return [{'id': r[0], 'course': r[1], 'department': r[2], 'semester': r[3],
                 'student_group': r[4], 'subject': r[5], 'description': r[6],
                 'submission_date': r[7], 'submission_time': r[8]} for r in results]
    except Exception as e:
        conn.close()
        print(f"Error getting homework for reminder: {e}")
        return []


def mark_homework_reminder_sent(homework_id):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        today = datetime.now().date()
        c.execute('INSERT OR IGNORE INTO homework_reminders_sent (homework_id, sent_date) VALUES (?, ?)',
                  (homework_id, today))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        conn.close()
        print(f"Error marking reminder sent: {e}")
        return False


# ==========================================
# ATTENDANCE FUNCTIONS
# ==========================================

def get_student_attendance(student_id, month=None, year=None):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        if month and year:
            c.execute('''SELECT subject_name, classes_held, classes_attended, 
                                attendance_percentage, month, year
                         FROM attendance_records
                         WHERE student_id = ? AND month = ? AND year = ?
                         ORDER BY subject_name''', (student_id.upper(), month, year))
        else:
            c.execute('''SELECT subject_name, classes_held, classes_attended, 
                                attendance_percentage, month, year
                         FROM attendance_records WHERE student_id = ?
                         ORDER BY year DESC, CASE month
                             WHEN 'January' THEN 1 WHEN 'February' THEN 2 WHEN 'March' THEN 3
                             WHEN 'April' THEN 4 WHEN 'May' THEN 5 WHEN 'June' THEN 6
                             WHEN 'July' THEN 7 WHEN 'August' THEN 8 WHEN 'September' THEN 9
                             WHEN 'October' THEN 10 WHEN 'November' THEN 11 WHEN 'December' THEN 12
                         END DESC LIMIT 10''', (student_id.upper(),))
        results = c.fetchall()
        conn.close()
        if not results:
            return None
        latest_month = results[0][4]
        latest_year = results[0][5]
        return {
            'month': latest_month, 'year': latest_year,
            'subjects': [{'subject': r[0], 'classes_held': r[1], 'classes_attended': r[2],
                          'percentage': r[3]} for r in results
                         if r[4] == latest_month and r[5] == latest_year]
        }
    except Exception as e:
        conn.close()
        print(f"Error getting attendance: {e}")
        return None


def check_low_attendance(student_id, threshold=75):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT subject_name, attendance_percentage, month, year
                     FROM attendance_records WHERE student_id = ?
                     ORDER BY year DESC, CASE month
                         WHEN 'January' THEN 1 WHEN 'February' THEN 2 WHEN 'March' THEN 3
                         WHEN 'April' THEN 4 WHEN 'May' THEN 5 WHEN 'June' THEN 6
                         WHEN 'July' THEN 7 WHEN 'August' THEN 8 WHEN 'September' THEN 9
                         WHEN 'October' THEN 10 WHEN 'November' THEN 11 WHEN 'December' THEN 12
                     END DESC''', (student_id.upper(),))
        results = c.fetchall()
        conn.close()
        if not results:
            return []
        latest_month = results[0][2]
        latest_year = results[0][3]
        return [{'subject': r[0], 'percentage': r[1], 'month': r[2], 'year': r[3]}
                for r in results if r[2] == latest_month and r[3] == latest_year
                and r[1] < threshold]
    except Exception as e:
        conn.close()
        print(f"Error checking low attendance: {e}")
        return []


def format_attendance_for_ai(attendance_data, low_attendance_subjects):
    if not attendance_data:
        return "No attendance data available yet. We're working on it!"
    month = attendance_data['month']
    year = attendance_data['year']
    text = f"\n📊 ATTENDANCE DATA (As per {month} {year}):\n"
    for subject in attendance_data['subjects']:
        status_emoji = "✅" if subject['percentage'] >= 75 else "⚠️" if subject['percentage'] >= 60 else "❌"
        text += f"{status_emoji} {subject['subject']}: {subject['classes_attended']}/{subject['classes_held']} ({subject['percentage']:.1f}%)\n"
    if low_attendance_subjects:
        text += f"\n⚠️ LOW ATTENDANCE ALERT:\nYou need 75% attendance for Internal Exams!\n"
        for sub in low_attendance_subjects:
            text += f"• {sub['subject']}: {sub['percentage']:.1f}% (Need to improve!)\n"
    current_month = datetime.now().strftime('%B')
    if current_month != month:
        text += f"\n📌 NOTE: {current_month} {year} data will be updated soon.\n"
    return text


# ==========================================
# LEADERBOARD FUNCTIONS
# ==========================================

def get_attendance_leaderboard(month=None, year=None, limit=10):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        if not month or not year:
            c.execute('''SELECT month, year FROM attendance_records
                         ORDER BY year DESC, CASE month
                             WHEN 'January' THEN 1 WHEN 'February' THEN 2 WHEN 'March' THEN 3
                             WHEN 'April' THEN 4 WHEN 'May' THEN 5 WHEN 'June' THEN 6
                             WHEN 'July' THEN 7 WHEN 'August' THEN 8 WHEN 'September' THEN 9
                             WHEN 'October' THEN 10 WHEN 'November' THEN 11 WHEN 'December' THEN 12
                         END DESC LIMIT 1''')
            result = c.fetchone()
            if result:
                month, year = result
            else:
                conn.close()
                return None
        c.execute('''SELECT ar.student_id, ms.name,
                            SUM(ar.classes_attended) as total_attended,
                            SUM(ar.classes_held) as total_held,
                            (CAST(SUM(ar.classes_attended) AS REAL) / SUM(ar.classes_held) * 100) as overall_percentage
                     FROM attendance_records ar
                     JOIN master_students ms ON ar.student_id = ms.student_id
                     WHERE ar.month = ? AND ar.year = ?
                     GROUP BY ar.student_id, ms.name
                     HAVING total_held > 0
                     ORDER BY overall_percentage DESC LIMIT ?''', (month, year, limit))
        results = c.fetchall()
        conn.close()
        return {
            'month': month, 'year': year,
            'leaderboard': [{'rank': i+1, 'student_id': r[0], 'name': r[1],
                             'total_attended': r[2], 'total_held': r[3],
                             'overall_percentage': r[4]} for i, r in enumerate(results)]
        }
    except Exception as e:
        conn.close()
        print(f"Error getting leaderboard: {e}")
        return None


def get_student_rank(student_id, month=None, year=None):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        if not month or not year:
            c.execute('''SELECT month, year FROM attendance_records
                         ORDER BY year DESC, CASE month
                             WHEN 'January' THEN 1 WHEN 'February' THEN 2 WHEN 'March' THEN 3
                             WHEN 'April' THEN 4 WHEN 'May' THEN 5 WHEN 'June' THEN 6
                             WHEN 'July' THEN 7 WHEN 'August' THEN 8 WHEN 'September' THEN 9
                             WHEN 'October' THEN 10 WHEN 'November' THEN 11 WHEN 'December' THEN 12
                         END DESC LIMIT 1''')
            result = c.fetchone()
            if result:
                month, year = result
            else:
                conn.close()
                return None
        c.execute('''WITH ranked_students AS (
                         SELECT ar.student_id,
                                (CAST(SUM(ar.classes_attended) AS REAL) / SUM(ar.classes_held) * 100) as overall_percentage,
                                RANK() OVER (ORDER BY (CAST(SUM(ar.classes_attended) AS REAL) / SUM(ar.classes_held) * 100) DESC) as rank
                         FROM attendance_records ar
                         WHERE ar.month = ? AND ar.year = ?
                         GROUP BY ar.student_id HAVING SUM(ar.classes_held) > 0
                     )
                     SELECT rank, overall_percentage FROM ranked_students WHERE student_id = ?''',
                  (month, year, student_id.upper()))
        result = c.fetchone()
        conn.close()
        if result:
            return {'rank': result[0], 'overall_percentage': result[1], 'month': month, 'year': year}
        return None
    except Exception as e:
        conn.close()
        print(f"Error getting student rank: {e}")
        return None


def format_leaderboard_for_ai(leaderboard_data, student_rank=None):
    if not leaderboard_data:
        return "Leaderboard data not available yet."
    month = leaderboard_data['month']
    year = leaderboard_data['year']
    text = f"\n🏆 ATTENDANCE LEADERBOARD ({month} {year}):\nTOP 5 STUDENTS:\n"
    for student in leaderboard_data['leaderboard'][:5]:
        rank_emoji = "🥇" if student['rank'] == 1 else "🥈" if student['rank'] == 2 else "🥉" if student['rank'] == 3 else f"{student['rank']}."
        text += f"{rank_emoji} {student['name']}: {student['overall_percentage']:.1f}%\n"
    text += "\n⚠️⚠️⚠️ CRITICAL NOTE FOR AI ⚠️⚠️⚠️\n"
    text += "This is ONLY the leaderboard top 5.\n"
    text += "DO NOT use these positions to tell the student their rank!\n"
    text += "Use ONLY the 'STUDENT ACTUAL RANK' section for rank queries!\n"
    return text


# ==========================================
# EVENTS FUNCTIONS
# ==========================================

# ==========================================

def setup_events_table(conn):
    """Create events table if not exists"""
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS events
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  event_id TEXT UNIQUE,
                  title TEXT,
                  date DATE,
                  time TEXT,
                  venue TEXT,
                  description TEXT,
                  category TEXT,
                  poster_path TEXT,
                  poster_telegram_id TEXT,
                  is_active INTEGER DEFAULT 1,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()


def load_events_data(conn=None):
    """Load events from CSV into DB"""
    close_conn = False
    if conn is None:
        conn = sqlite3.connect('students.db', timeout=20)
        close_conn = True
    try:
        c = conn.cursor()
        setup_events_table(conn)

        with open('events_data.csv', 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                c.execute('''INSERT OR REPLACE INTO events
                             (event_id, title, date, time, venue, description,
                              category, poster_path, poster_telegram_id)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                          (row['event_id'].strip(),
                           row['title'].strip(),
                           row['date'].strip(),
                           row['time'].strip(),
                           row['venue'].strip(),
                           row['description'].strip(),
                           row['category'].strip().lower(),
                           row.get('poster_path', '').strip(),
                           row.get('poster_telegram_id', '').strip()))
        conn.commit()
        print("✅ Events data loaded!")
    except FileNotFoundError:
        print("⚠️ events_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading events: {e}")
    finally:
        if close_conn:
            conn.close()


def get_upcoming_events(category=None, days_ahead=30):
    """Get upcoming events, optionally filtered by category"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        setup_events_table(conn)
        today = datetime.now().date()
        future = today + timedelta(days=days_ahead)

        if category:
            c.execute('''SELECT event_id, title, date, time, venue, description,
                                category, poster_path, poster_telegram_id
                         FROM events
                         WHERE date >= ? AND date <= ? AND is_active = 1
                         AND category = ?
                         ORDER BY date, time''',
                      (today, future, category.lower()))
        else:
            c.execute('''SELECT event_id, title, date, time, venue, description,
                                category, poster_path, poster_telegram_id
                         FROM events
                         WHERE date >= ? AND date <= ? AND is_active = 1
                         ORDER BY date, time''',
                      (today, future))

        results = c.fetchall()
        conn.close()
        return [{'event_id': r[0], 'title': r[1], 'date': r[2], 'time': r[3],
                 'venue': r[4], 'description': r[5], 'category': r[6],
                 'poster_path': r[7], 'poster_telegram_id': r[8]} for r in results]
    except Exception as e:
        conn.close()
        print(f"Error getting events: {e}")
        return []


def get_event_by_id(event_id):
    """Get single event by ID"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        setup_events_table(conn)
        c.execute('''SELECT event_id, title, date, time, venue, description,
                            category, poster_path, poster_telegram_id
                     FROM events WHERE event_id = ? AND is_active = 1''',
                  (event_id.upper(),))
        r = c.fetchone()
        conn.close()
        if r:
            return {'event_id': r[0], 'title': r[1], 'date': r[2], 'time': r[3],
                    'venue': r[4], 'description': r[5], 'category': r[6],
                    'poster_path': r[7], 'poster_telegram_id': r[8]}
        return None
    except Exception as e:
        conn.close()
        print(f"Error getting event: {e}")
        return None


def get_events_for_reminder():
    """Get events happening tomorrow that need reminders"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        setup_events_table(conn)
        tomorrow = (datetime.now() + timedelta(days=1)).date()
        c.execute('''SELECT event_id, title, date, time, venue, description,
                            category, poster_path, poster_telegram_id
                     FROM events
                     WHERE date = ? AND is_active = 1
                     ORDER BY time''', (tomorrow,))
        results = c.fetchall()
        conn.close()
        return [{'event_id': r[0], 'title': r[1], 'date': r[2], 'time': r[3],
                 'venue': r[4], 'description': r[5], 'category': r[6],
                 'poster_path': r[7], 'poster_telegram_id': r[8]} for r in results]
    except Exception as e:
        conn.close()
        print(f"Error getting reminder events: {e}")
        return []


def save_poster_telegram_id(event_id, telegram_file_id):
    """Save Telegram file_id after first poster upload (for reuse)"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('UPDATE events SET poster_telegram_id = ? WHERE event_id = ?',
                  (telegram_file_id, event_id.upper()))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        conn.close()
        print(f"Error saving telegram ID: {e}")
        return False


def format_events_for_ai(events):
    """Format events list for AI context"""
    if not events:
        return "No upcoming events found."
    cat_emoji = {'academic': '📚', 'sports': '⚽', 'social': '🎉'}
    text = "UPCOMING EVENTS:\n"
    for e in events:
        try:
            d = datetime.strptime(str(e['date']), '%Y-%m-%d')
            date_str = d.strftime('%d %B %Y (%A)')
        except:
            date_str = str(e['date'])
        emoji = cat_emoji.get(e['category'], '📌')
        text += f"\n{emoji} [{e['event_id']}] {e['title']}\n"
        text += f"   Date: {date_str} at {e['time']}\n"
        text += f"   Venue: {e['venue']}\n"
        text += f"   {e['description']}\n"
    return text


# ==========================================
# NOTES / CSV STUDY MATERIAL FUNCTIONS
# ==========================================

def setup_notes_table(conn):
    """Create notes table if not exists"""
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS notes
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subject_code TEXT,
                  subject_name TEXT,
                  unit_number INTEGER,
                  unit_name TEXT,
                  topic TEXT,
                  content TEXT,
                  exam_tips TEXT,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    c.execute('''CREATE UNIQUE INDEX IF NOT EXISTS
                 idx_notes_unique ON notes(subject_code, unit_number, topic)''')
    conn.commit()


def load_notes_data(conn=None):
    """Load notes from notes_data.csv into DB"""
    close_conn = False
    if conn is None:
        conn = sqlite3.connect('students.db', timeout=20)
        close_conn = True
    try:
        c = conn.cursor()
        setup_notes_table(conn)
        c.execute('DELETE FROM notes')
        with open('notes_data.csv', 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            count = 0
            for row in reader:
                c.execute('''INSERT OR REPLACE INTO notes
                             (subject_code, subject_name, unit_number, unit_name,
                              topic, content, exam_tips)
                             VALUES (?, ?, ?, ?, ?, ?, ?)''',
                          (row['subject_code'].strip(),
                           row['subject_name'].strip(),
                           int(row['unit_number']),
                           row['unit_name'].strip(),
                           row['topic'].strip(),
                           row['content'].strip(),
                           row.get('exam_tips', '').strip()))
                count += 1
        conn.commit()
        print(f"✅ Notes data loaded! ({count} topics)")
    except FileNotFoundError:
        print("⚠️ notes_data.csv not found")
    except Exception as e:
        print(f"❌ Error loading notes: {e}")
    finally:
        if close_conn:
            conn.close()


def search_notes(query, subject_code=None, unit_number=None, limit=5):
    """Search notes by keyword — returns most relevant topics"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        query_lower = query.lower()
        words = [w for w in query_lower.split() if len(w) > 2]

        # Build base query
        params = []
        where_parts = []

        if subject_code:
            where_parts.append('subject_code = ?')
            params.append(subject_code.upper())
        if unit_number:
            where_parts.append('unit_number = ?')
            params.append(int(unit_number))

        where_sql = (' AND ' + ' AND '.join(where_parts)) if where_parts else ''

        c.execute(f'''SELECT subject_code, subject_name, unit_number, unit_name,
                              topic, content, exam_tips
                         FROM notes
                         WHERE 1=1 {where_sql}
                         ORDER BY unit_number''', params)
        all_rows = c.fetchall()
        conn.close()

        if not all_rows:
            return []

        # Score each row by keyword matches
        def score(row):
            text = (row[4] + ' ' + row[5] + ' ' + row[3]).lower()
            return sum(1 for w in words if w in text)

        scored = [(score(r), r) for r in all_rows]
        scored.sort(key=lambda x: -x[0])

        # Return top results — include all if unit/subject filter applied
        top = scored[:limit]
        return [r for s, r in top if s > 0] or [r for s, r in scored[:3]]

    except Exception as e:
        conn.close()
        print(f"Error searching notes: {e}")
        return []


def format_notes_for_ai(notes_rows, include_exam_tips=True):
    """Format notes rows into clean context string for AI"""
    if not notes_rows:
        return ""
    parts = []
    for row in notes_rows:
        subject_code, subject_name, unit_num, unit_name, topic, content, exam_tips = row
        text = f"[{subject_code} Unit {unit_num} — {topic}]\n{content}"
        if include_exam_tips and exam_tips:
            text += f"\nExam tip: {exam_tips}"
        parts.append(text)
    return "\n\n".join(parts)


def get_overall_attendance(student_id):
    """Get combined attendance across all months for exam eligibility"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT subject_name,
                            SUM(classes_held) as total_held,
                            SUM(classes_attended) as total_attended
                     FROM attendance_records
                     WHERE student_id = ?
                     GROUP BY subject_name
                     ORDER BY subject_name''', (student_id.upper(),))
        results = c.fetchall()
        conn.close()
        if not results:
            return None
        subjects = []
        for r in results:
            pct = (r[2] / r[1] * 100) if r[1] > 0 else 0
            classes_needed = max(0, int(0.75 * r[1]) - r[2] + 1)
            subjects.append({
                'subject': r[0],
                'total_held': r[1],
                'total_attended': r[2],
                'percentage': round(pct, 1),
                'eligible': pct >= 75,
                'classes_needed': classes_needed if pct < 75 else 0
            })
        return subjects
    except Exception as e:
        conn.close()
        print(f"Error getting overall attendance: {e}")
        return None


def format_overall_attendance_for_ai(overall_data):
    """Format overall/combined attendance for AI"""
    if not overall_data:
        return "No attendance data available."
    
    all_eligible = all(s['eligible'] for s in overall_data)
    total_held = sum(s['total_held'] for s in overall_data)
    total_attended = sum(s['total_attended'] for s in overall_data)
    overall_pct = round(total_attended / total_held * 100, 1) if total_held > 0 else 0

    text = f"OVERALL ATTENDANCE (Jan + Feb combined):\n"
    text += f"Overall: {total_attended}/{total_held} = {overall_pct}%\n\n"

    for s in overall_data:
        status = "✅" if s['eligible'] else "❌"
        text += f"{status} {s['subject']}: {s['total_attended']}/{s['total_held']} ({s['percentage']}%)"
        if not s['eligible']:
            text += f" — Need {s['classes_needed']} more classes"
        text += "\n"

    text += f"\nExam Eligibility: {'✅ ELIGIBLE' if all_eligible else '❌ NOT FULLY ELIGIBLE (75% required)'}"
    return text


def get_monthly_leaderboard(month=None, year=None, limit=10):
    """Leaderboard for a specific month"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        if not month or not year:
            # Get latest month
            c.execute('''SELECT month, year FROM attendance_records
                         ORDER BY year DESC, CASE month
                             WHEN 'January' THEN 1 WHEN 'February' THEN 2 WHEN 'March' THEN 3
                             WHEN 'April' THEN 4 WHEN 'May' THEN 5 WHEN 'June' THEN 6
                             WHEN 'July' THEN 7 WHEN 'August' THEN 8 WHEN 'September' THEN 9
                             WHEN 'October' THEN 10 WHEN 'November' THEN 11 WHEN 'December' THEN 12
                         END DESC LIMIT 1''')
            result = c.fetchone()
            if result:
                month, year = result
            else:
                conn.close()
                return None

        c.execute('''SELECT ar.student_id, ms.name,
                            SUM(ar.classes_attended) as attended,
                            SUM(ar.classes_held) as held,
                            ROUND(CAST(SUM(ar.classes_attended) AS REAL) / SUM(ar.classes_held) * 100, 1) as pct
                     FROM attendance_records ar
                     JOIN master_students ms ON ar.student_id = ms.student_id
                     WHERE ar.month = ? AND ar.year = ?
                     GROUP BY ar.student_id
                     HAVING held > 0
                     ORDER BY pct DESC LIMIT ?''', (month, year, limit))
        rows = c.fetchall()
        conn.close()
        return {
            'month': month, 'year': year,
            'data': [{'rank': i+1, 'student_id': r[0], 'name': r[1],
                      'attended': r[2], 'held': r[3], 'pct': r[4]}
                     for i, r in enumerate(rows)]
        }
    except Exception as e:
        conn.close()
        print(f"Error getting monthly leaderboard: {e}")
        return None


def get_overall_leaderboard(limit=10):
    """Leaderboard combining all months"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT ar.student_id, ms.name,
                            SUM(ar.classes_attended) as attended,
                            SUM(ar.classes_held) as held,
                            ROUND(CAST(SUM(ar.classes_attended) AS REAL) / SUM(ar.classes_held) * 100, 1) as pct
                     FROM attendance_records ar
                     JOIN master_students ms ON ar.student_id = ms.student_id
                     GROUP BY ar.student_id
                     HAVING held > 0
                     ORDER BY pct DESC LIMIT ?''', (limit,))
        rows = c.fetchall()
        conn.close()
        return {
            'month': 'Overall (Jan-Feb)', 'year': 2026,
            'data': [{'rank': i+1, 'student_id': r[0], 'name': r[1],
                      'attended': r[2], 'held': r[3], 'pct': r[4]}
                     for i, r in enumerate(rows)]
        }
    except Exception as e:
        conn.close()
        print(f"Error getting overall leaderboard: {e}")
        return None
