import os

print("BOT:", os.getenv("BOT_TOKEN"))
print("GROQ:", os.getenv("GROQ_API_KEY"))
print("TAVILY:", os.getenv("TAVILY_API_KEY"))
print("EMAIL:", os.getenv("SENDER_EMAIL"))
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, MessageHandler, CallbackQueryHandler, CommandHandler, filters, ContextTypes
from groq import Groq
import sqlite3
from datetime import datetime, timedelta
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import csv
import json
import asyncio
import re
import os
import openpyxl
from openpyxl import Workbook
import httpx
import base64
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# Import timetable functions
from timetable_functions import (
    get_today_schedule, get_next_class, get_subject_schedule,
    format_schedule_for_ai, format_updates_for_ai,
    is_cr, post_class_update, get_active_updates,
    setup_timetable,
    get_subject_syllabus, get_syllabus_progress,
    get_next_topic, format_syllabus_for_ai,
    mark_topic_complete, get_exam_tips,
    save_last_theory_class, save_last_practical_class,
    get_class_history, format_class_history_for_ai,
    detect_subject_name, get_subject_from_teacher,
    save_homework, get_pending_homework, format_homework_for_ai,
    get_homework_for_reminder, mark_homework_reminder_sent,
    get_student_attendance, check_low_attendance, format_attendance_for_ai,
    get_attendance_leaderboard, get_student_rank, format_leaderboard_for_ai,
    load_events_data, get_upcoming_events, get_event_by_id,
    get_events_for_reminder, save_poster_telegram_id, format_events_for_ai,
    get_overall_attendance, format_overall_attendance_for_ai,
    get_monthly_leaderboard, get_overall_leaderboard
)

# ✅ NOTES SYSTEM — PDF notes upload + ChromaDB search


# Configuration
BOT_TOKEN = os.getenv("BOT_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
ADMIN_ID = 6011716383
SARVAM_API_KEY = os.getenv("SARVAM_API_KEY")

# Initialize Groq client
groq_client = Groq(api_key=GROQ_API_KEY)

# Tavily web search
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY")
try:
    from tavily import TavilyClient
    tavily_client = TavilyClient(TAVILY_API_KEY)
    print("✅ Tavily web search ready!")
except ImportError:
    tavily_client = None
    print("⚠️ tavily-python not installed. Run: pip install tavily-python")

def is_admin(telegram_id):
    return telegram_id == ADMIN_ID

# ============================================================
# ✅ HELPER: Parse natural date from text
# ============================================================
def parse_date_from_text(text):
    """
    Extract a date from natural language text.
    Handles: '17th feb', 'kal', 'aaj', 'Monday', '17 february', etc.
    Returns a date object or None.
    """
    text_lower = text.lower().strip()
    today = datetime.now().date()

    # Relative keywords
    if any(w in text_lower for w in ['aaj', 'today']):
        return today
    if any(w in text_lower for w in ['kal', 'yesterday', 'kal ki', 'kal ka']):
        return today - timedelta(days=1)
    if any(w in text_lower for w in ['parso', 'day before yesterday']):
        return today - timedelta(days=2)

    month_map = {
        'jan': 1, 'january': 1, 'feb': 2, 'february': 2,
        'mar': 3, 'march': 3, 'apr': 4, 'april': 4,
        'may': 5, 'jun': 6, 'june': 6, 'jul': 7, 'july': 7,
        'aug': 8, 'august': 8, 'sep': 9, 'sept': 9, 'september': 9,
        'oct': 10, 'october': 10, 'nov': 11, 'november': 11,
        'dec': 12, 'december': 12
    }

    # Pattern: "17th feb", "17 february", "feb 17", "17-02", etc.
    patterns = [
        r'(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?([a-z]+)',
        r'([a-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?',
        r'(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?',
    ]

    for pattern in patterns:
        match = re.search(pattern, text_lower)
        if match:
            try:
                groups = match.groups()
                if groups[0] and groups[0].isdigit() and len(groups) >= 2:
                    day = int(groups[0])
                    month_str = groups[1]
                    if month_str.isdigit():
                        month = int(month_str)
                    else:
                        month = month_map.get(month_str[:3])
                    if month:
                        year = today.year
                        candidate = datetime(year, month, day).date()
                        if candidate > today:
                            candidate = datetime(year - 1, month, day).date()
                        return candidate
                elif groups[0] and not groups[0].isdigit() and len(groups) >= 2:
                    month_str = groups[0]
                    day = int(groups[1])
                    month = month_map.get(month_str[:3])
                    if month:
                        year = today.year
                        candidate = datetime(year, month, day).date()
                        if candidate > today:
                            candidate = datetime(year - 1, month, day).date()
                        return candidate
            except (ValueError, TypeError):
                continue

    # Day names → most recent occurrence
    days_map = {
        'monday': 0, 'tuesday': 1, 'wednesday': 2,
        'thursday': 3, 'friday': 4, 'saturday': 5, 'sunday': 6
    }
    for day_name, day_num in days_map.items():
        if day_name in text_lower:
            delta = (today.weekday() - day_num) % 7
            if delta == 0:
                delta = 7
            return today - timedelta(days=delta)

    return None


# ============================================================
# DATABASE SETUP
# ============================================================
def init_db():
    conn = sqlite3.connect('students.db')
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS master_students
                 (student_id TEXT PRIMARY KEY,
                  name TEXT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  email TEXT,
                  student_group TEXT DEFAULT 'A',
                  urdu_type TEXT DEFAULT 'regular')''')

    c.execute('''CREATE TABLE IF NOT EXISTS registered_users
                 (telegram_id INTEGER PRIMARY KEY,
                  student_id TEXT UNIQUE,
                  is_verified INTEGER DEFAULT 0,
                  otp TEXT,
                  registered_at TIMESTAMP,
                  last_active TIMESTAMP)''')

    c.execute('''CREATE TABLE IF NOT EXISTS conversations
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  telegram_id INTEGER,
                  role TEXT,
                  content TEXT,
                  timestamp TIMESTAMP)''')

    c.execute('''CREATE TABLE IF NOT EXISTS class_reminders_sent
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  class_key TEXT UNIQUE,
                  sent_at TIMESTAMP)''')

    c.execute('''CREATE TABLE IF NOT EXISTS pending_cr_actions
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  telegram_id INTEGER,
                  action_type TEXT,
                  action_data TEXT,
                  created_at TIMESTAMP)''')

    c.execute('''CREATE TABLE IF NOT EXISTS class_history_log
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  subject_name TEXT,
                  subject_code TEXT,
                  class_type TEXT,
                  student_group TEXT DEFAULT 'BOTH',
                  teacher_name TEXT,
                  unit_covered TEXT,
                  topics_covered TEXT,
                  key_points TEXT,
                  homework_assigned TEXT,
                  practical_work TEXT,
                  demonstrations TEXT,
                  class_date DATE,
                  recorded_by TEXT,
                  created_at TIMESTAMP,
                  photo_path TEXT DEFAULT NULL)''')
    try:
        c.execute("ALTER TABLE class_history_log ADD COLUMN photo_path TEXT DEFAULT NULL")
    except:
        pass

    c.execute('''CREATE TABLE IF NOT EXISTS homework_log
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  course TEXT,
                  department TEXT,
                  semester INTEGER,
                  student_group TEXT DEFAULT 'BOTH',
                  subject_name TEXT,
                  teacher_name TEXT,
                  description TEXT,
                  submission_date DATE,
                  submission_time TEXT,
                  posted_by TEXT,
                  reminder_sent INTEGER DEFAULT 0,
                  created_at TIMESTAMP)''')

    conn.commit()
    conn.close()

init_db()

# ============================================================
# LAST ACTIVE / USER STATS
# ============================================================
def update_last_active(telegram_id):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute("UPDATE registered_users SET last_active = ? WHERE telegram_id = ?",
              (datetime.now(), telegram_id))
    conn.commit()
    conn.close()

def get_monthly_active_users():
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    thirty_days_ago = datetime.now() - timedelta(days=30)
    c.execute('''SELECT COUNT(*) FROM registered_users
                 WHERE is_verified = 1 AND last_active >= ?''',
              (thirty_days_ago,))
    count = c.fetchone()[0]
    conn.close()
    return count

def export_users_to_excel():
    try:
        conn = sqlite3.connect('students.db')
        c = conn.cursor()
        c.execute('''SELECT ms.student_id, ms.name, ms.course, ms.department, ms.semester,
                            ms.email, ms.student_group, ms.urdu_type, ru.registered_at, ru.last_active
                     FROM registered_users ru
                     JOIN master_students ms ON ru.student_id = ms.student_id
                     WHERE ru.is_verified = 1
                     ORDER BY ru.registered_at DESC''')
        users = c.fetchall()
        conn.close()

        wb = Workbook()
        ws = wb.active
        ws.title = "Registered Users"
        headers = ['Student ID', 'Name', 'Course', 'Department', 'Semester', 'Email', 'Group', 'Urdu Type', 'Registered At', 'Last Active']
        ws.append(headers)
        for user in users:
            ws.append(list(user))
        filename = f'registered_users_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        wb.save(filename)
        return filename
    except Exception as e:
        print(f"Excel export error: {e}")
        return None

def get_user_count():
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM registered_users WHERE is_verified = 1")
    count = c.fetchone()[0]
    conn.close()
    return count

def reset_user_registration(telegram_id):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute("DELETE FROM registered_users WHERE telegram_id = ?", (telegram_id,))
    c.execute("DELETE FROM conversations WHERE telegram_id = ?", (telegram_id,))
    conn.commit()
    conn.close()

def load_student_data(csv_file='students_data.csv'):
    try:
        conn = sqlite3.connect('students.db')
        c = conn.cursor()
        with open(csv_file, 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                urdu_type = row.get('urdu_type', 'regular').strip().lower()
                if urdu_type not in ['regular', 'advanced']:
                    urdu_type = 'regular'
                c.execute('''INSERT OR REPLACE INTO master_students
                             (student_id, name, course, department, semester, email, student_group, urdu_type)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                          (row['student_id'].strip().upper(),
                           row['name'].strip(),
                           row['course'].strip(),
                           row['department'].strip(),
                           int(row['semester']),
                           row['email'].strip().lower(),
                           row.get('group', 'A').strip().upper(),
                           urdu_type))
        conn.commit()
        conn.close()
        print("✅ Student data loaded!")
    except FileNotFoundError:
        print("⚠️ students_data.csv not found.")
    except Exception as e:
        print(f"❌ Error loading data: {e}")

# ============================================================
# SAVE CLASS HISTORY TO LOCAL DB
# ============================================================
def save_class_history_to_db(course, department, semester, subject_name, subject_code,
                               class_type, student_group, teacher_name, unit_covered,
                               topics_covered, key_points, homework_assigned,
                               practical_work, demonstrations, class_date, recorded_by,
                               photo_path=None):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute('''INSERT INTO class_history_log
                 (course, department, semester, subject_name, subject_code, class_type,
                  student_group, teacher_name, unit_covered, topics_covered, key_points,
                  homework_assigned, practical_work, demonstrations, class_date, recorded_by, created_at,
                  photo_path)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (course, department, semester, subject_name, subject_code, class_type,
               student_group, teacher_name, unit_covered, topics_covered, key_points,
               homework_assigned, practical_work, demonstrations, str(class_date), recorded_by,
               datetime.now(), photo_path))
    conn.commit()
    conn.close()

def get_class_history_from_db(course, department, semester, subject_name=None,
                               class_date=None, class_type=None, student_group=None,
                               limit=20):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()

    query = '''SELECT subject_name, class_type, student_group, teacher_name, unit_covered,
                      topics_covered, key_points, homework_assigned, practical_work,
                      demonstrations, class_date, photo_path
               FROM class_history_log
               WHERE course = ? AND department = ? AND semester = ?'''
    params = [course, department, semester]

    if subject_name:
        query += ' AND (LOWER(subject_name) LIKE ? OR LOWER(subject_code) LIKE ?)'
        like_term = f'%{subject_name.lower()}%'
        params.extend([like_term, like_term])

    if class_date:
        query += ' AND class_date = ?'
        params.append(str(class_date))

    if class_type:
        query += ' AND class_type = ?'
        params.append(class_type)

    if student_group and student_group != 'BOTH':
        query += ' AND (student_group = ? OR student_group = "BOTH")'
        params.append(student_group)

    query += ' ORDER BY class_date DESC, created_at DESC LIMIT ?'
    params.append(limit)

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    results = []
    for row in rows:
        results.append({
            'subject': row[0], 'class_type': row[1], 'group': row[2],
            'teacher': row[3], 'unit': row[4], 'topics': row[5],
            'key_points': row[6], 'homework': row[7], 'practical_work': row[8],
            'demonstrations': row[9], 'date': row[10],
            'photo_path': row[11] if len(row) > 11 else None
        })
    return results

def format_class_history_for_context(history_records):
    if not history_records:
        return "No class history found."
    text = ""
    for rec in history_records:
        text += f"\n📅 Date: {rec['date']} | {rec['class_type'].upper()} | {rec['subject']}"
        if rec.get('group') and rec['group'] != 'BOTH':
            text += f" (Group {rec['group']})"
        text += "\n"
        if rec.get('teacher'):
            text += f"   👨‍🏫 Teacher: {rec['teacher']}\n"
        if rec.get('unit'):
            text += f"   📖 Unit: {rec['unit']}\n"
        if rec.get('topics'):
            text += f"   📝 Topics: {rec['topics']}\n"
        if rec.get('key_points'):
            text += f"   💡 Key Points: {rec['key_points']}\n"
        if rec.get('practical_work'):
            text += f"   🔬 Practical: {rec['practical_work']}\n"
        if rec.get('demonstrations'):
            text += f"   🎬 Demos: {rec['demonstrations']}\n"
        if rec.get('homework'):
            text += f"   📋 Homework: {rec['homework']}\n"
    return text

# ============================================================
# SAVE HOMEWORK TO LOCAL DB
# ============================================================
def save_homework_to_db(course, department, semester, student_group, subject_name,
                         teacher_name, description, submission_date, submission_time, posted_by):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute('''INSERT INTO homework_log
                 (course, department, semester, student_group, subject_name, teacher_name,
                  description, submission_date, submission_time, posted_by, created_at)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (course, department, semester, student_group, subject_name, teacher_name,
               description, str(submission_date), submission_time, posted_by, datetime.now()))
    conn.commit()
    conn.close()

def get_homework_from_db(course, department, semester, student_group=None, subject_name=None):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    today = datetime.now().date()

    query = '''SELECT subject_name, teacher_name, description, submission_date, submission_time
               FROM homework_log
               WHERE course = ? AND department = ? AND semester = ?
               AND submission_date >= ?'''
    params = [course, department, semester, str(today)]

    if student_group and student_group != 'BOTH':
        query += ' AND (student_group = ? OR student_group = "BOTH")'
        params.append(student_group)

    if subject_name:
        query += ' AND LOWER(subject_name) LIKE ?'
        params.append(f'%{subject_name.lower()}%')

    query += ' ORDER BY submission_date ASC'
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    return [{'subject': row[0], 'teacher': row[1], 'description': row[2],
             'due_date': row[3], 'due_time': row[4]} for row in rows]

# ============================================================
# CONVERSATION & STUDENT HELPERS
# ============================================================
def get_conversation_history(telegram_id, limit=20):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute('''SELECT role, content FROM conversations
                 WHERE telegram_id = ?
                 ORDER BY timestamp DESC LIMIT ?''',
              (telegram_id, limit))
    history = c.fetchall()
    conn.close()
    return [{"role": role, "content": content} for role, content in reversed(history)]

def save_message(telegram_id, role, content):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute('''INSERT INTO conversations (telegram_id, role, content, timestamp)
                 VALUES (?, ?, ?, ?)''',
              (telegram_id, role, content, datetime.now()))
    conn.commit()
    conn.close()

def get_student_details(student_id):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute("SELECT * FROM master_students WHERE student_id = ?", (student_id.upper(),))
    result = c.fetchone()
    conn.close()
    if result:
        return {
            'student_id': result[0], 'name': result[1], 'course': result[2],
            'department': result[3], 'semester': result[4], 'email': result[5],
            'group': result[6], 'urdu_type': result[7] if len(result) > 7 else 'regular'
        }
    return None

def get_registration_status(telegram_id):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute("SELECT student_id, is_verified FROM registered_users WHERE telegram_id = ?",
              (telegram_id,))
    result = c.fetchone()
    conn.close()
    if result:
        return {'student_id': result[0], 'is_verified': result[1]}
    return None

def save_otp(telegram_id, student_id, otp):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO registered_users
                 (telegram_id, student_id, otp, registered_at, is_verified, last_active)
                 VALUES (?, ?, ?, ?, 0, ?)''',
              (telegram_id, student_id, otp, datetime.now(), datetime.now()))
    conn.commit()
    conn.close()

def verify_otp(telegram_id, entered_otp):
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute("SELECT otp FROM registered_users WHERE telegram_id = ?", (telegram_id,))
    result = c.fetchone()
    if result and result[0] == entered_otp:
        c.execute("UPDATE registered_users SET is_verified = 1, last_active = ? WHERE telegram_id = ?",
                  (datetime.now(), telegram_id))
        conn.commit()
        conn.close()
        return True
    conn.close()
    return False

def send_otp_email(email, otp, name):
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = 'Your Verification Code - Zei'
        msg['From'] = f'Zei <{SENDER_EMAIL}>'
        msg['To'] = email
        msg['Reply-To'] = SENDER_EMAIL
        msg['X-Priority'] = '1'
        msg['X-MSMail-Priority'] = 'High'
        msg['Importance'] = 'High'

        text = f"""Hi {name},

Your verification code is: {otp}

This code will expire in 10 minutes.

Best regards,
Team Zei (A Zephy Company)
"""
        html = f"""
        <html>
          <body style="font-family: Arial, sans-serif; padding: 20px;">
            <div style="max-width: 600px; margin: 0 auto; background-color: #f9f9f9; padding: 30px; border-radius: 10px;">
              <h2 style="color: #333;">Hi {name}! 👋</h2>
              <p style="font-size: 16px; color: #555;">Your verification code is:</p>
              <div style="background-color: #4CAF50; color: white; padding: 15px; text-align: center; font-size: 24px; font-weight: bold; border-radius: 5px; margin: 20px 0;">
                {otp}
              </div>
              <p style="font-size: 14px; color: #777;">This code will expire in 10 minutes.</p>
              <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
              <p style="font-size: 12px; color: #999;">Best regards,<br>Team Zei (A Zephy Company)</p>
            </div>
          </body>
        </html>
        """
        msg.attach(MIMEText(text, 'plain'))
        msg.attach(MIMEText(html, 'html'))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Email error: {e}")
        return False

def get_greeting():
    hour = datetime.now().hour
    if 5 <= hour < 12:
        return "Good morning"
    elif 12 <= hour < 17:
        return "Good afternoon"
    elif 17 <= hour < 21:
        return "Good evening"
    else:
        return "Hello"

# ============================================================
# SCHEDULE HELPER
# ============================================================
def get_schedule_for_day(course, dept, sem, group, day_name):
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    c.execute('''SELECT subject_name, teacher_name, start_time, end_time, room_number, student_group
                 FROM class_schedule
                 WHERE course = ? AND department = ? AND semester = ?
                 AND day_of_week = ?
                 AND (student_group = ? OR student_group = 'BOTH')
                 AND is_active = 1
                 ORDER BY start_time''',
              (course, dept, sem, day_name, group))
    results = c.fetchall()
    conn.close()
    return [
        {'subject': row[0], 'teacher': row[1], 'start_time': row[2],
         'end_time': row[3], 'room': row[4], 'group': row[5]}
        for row in results
    ]

# ============================================================
# ✅ GROQ OCR - Extract text from image
# ============================================================
async def groq_ocr_from_image(image_bytes: bytes, mime_type: str = "image/jpeg") -> str:
    """Use Groq vision to extract text from image (OCR)"""
    try:
        b64_image = base64.b64encode(image_bytes).decode('utf-8')
        response = groq_client.chat.completions.create(
            model="meta-llama/llama-4-scout-17b-16e-instruct",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{b64_image}"
                            }
                        },
                        {
                            "type": "text",
                            "text": "Extract ALL text from this image exactly as written. Include everything: handwritten, printed, numbers, dates. Output only the extracted text, nothing else."
                        }
                    ]
                }
            ],
            temperature=0.1,
            max_tokens=1000
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"Groq OCR error: {e}")
        return ""

# ============================================================
# ✅ SARVAM BULBUL v3 TTS - speaker: shubh, 48khz
# ============================================================
async def sarvam_tts(text: str) -> bytes | None:
    """Convert text to speech using Sarvam Bulbul v3, speaker=shubh, 48khz
    NOTE: bulbul:v3 does NOT support pitch/loudness — only pace and temperature"""
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                "https://api.sarvam.ai/text-to-speech",
                headers={
                    "api-subscription-key": SARVAM_API_KEY,
                    "Content-Type": "application/json"
                },
                json={
                    "inputs": [text[:500]],
                    "target_language_code": "hi-IN",
                    "speaker": "shubh",
                    "model": "bulbul:v3",
                    "pace": 1.0,
                    "temperature": 0.7,
                    "sample_rate": 48000,
                    "enable_preprocessing": True
                }
            )
            if response.status_code == 200:
                data = response.json()
                audio_b64 = data.get("audios", [None])[0]
                if audio_b64:
                    return base64.b64decode(audio_b64)
            else:
                print(f"Sarvam TTS error: {response.status_code} {response.text}")
        return None
    except Exception as e:
        print(f"Sarvam TTS error: {e}")
        return None

# ============================================================
# ✅ HELPER: Clean AI text for TTS (remove markdown/emojis/JSON)
# ============================================================
def clean_text_for_tts(ai_response: str) -> str:
    text = ai_response
    if '{"action"' in text:
        json_start = text.find('{"action"')
        json_end = text.find('}', json_start) + 1
        text = text[:json_start] + text[json_end:]
    text = text.replace('*', '').replace('_', '').replace('`', '').replace('#', '')
    text = re.sub(
        r'[\U0001F300-\U0001F9FF'
        r'\U00002600-\U000027BF'
        r'\U0001FA00-\U0001FA6F'
        r'\U0001FA70-\U0001FAFF'
        r'\u2702-\u27B0'
        r'\u24C2-\U0001F251]+',
        '', text, flags=re.UNICODE
    )
    text = re.sub(r'\n+', '. ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    if len(text) > 490:
        trimmed = text[:490]
        last_stop = max(trimmed.rfind('.'), trimmed.rfind('!'), trimmed.rfind('?'))
        if last_stop > 200:
            text = trimmed[:last_stop + 1]
        else:
            text = trimmed + '...'
    return text

# ============================================================
# ✅ PDF GENERATOR — Convert AI response to clean PDF
# ============================================================
def fetch_image_for_pdf(query: str) -> str | None:
    """Search DuckDuckGo for a high-quality image. Returns local file path or None."""
    import requests, os, re, urllib.parse, tempfile
    try:
        # DuckDuckGo image search (no API key needed)
        q = urllib.parse.quote(query + " diagram high resolution")
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        # Step 1: get vqd token
        r = requests.get(f'https://duckduckgo.com/?q={q}&iax=images&ia=images',
                         headers=headers, timeout=8)
        vqd_match = re.search(r'vqd=([\d-]+)', r.text)
        if not vqd_match:
            return None
        vqd = vqd_match.group(1)

        # Step 2: fetch image results
        img_url = f'https://duckduckgo.com/i.js?q={q}&vqd={vqd}&f=,,,,,&p=1'
        r2 = requests.get(img_url, headers=headers, timeout=8)
        results = r2.json().get('results', [])

        # Pick first high-res image (prefer width > 600)
        chosen_url = None
        for res in results[:10]:
            w = res.get('width', 0)
            url = res.get('image', '')
            if w >= 600 and url.startswith('http'):
                chosen_url = url
                break
        if not chosen_url and results:
            chosen_url = results[0].get('image')
        if not chosen_url:
            return None

        # Step 3: download image
        img_data = requests.get(chosen_url, headers=headers, timeout=10, stream=True)
        ext = '.jpg'
        ct = img_data.headers.get('Content-Type', '')
        if 'png' in ct: ext = '.png'
        elif 'gif' in ct: ext = '.gif'
        elif 'webp' in ct: ext = '.webp'

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        for chunk in img_data.iter_content(8192):
            tmp.write(chunk)
        tmp.close()
        return tmp.name
    except Exception as e:
        print(f"[Image fetch error] {e}")
        return None


def generate_pdf_from_response(content: str, title: str = "Zei Response", image_path: str = None) -> str:
    """Single aesthetic minimalist PDF — clean typography, tables auto-detected, Zephy branding."""
    import re as _re
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     HRFlowable, Table, TableStyle, KeepTogether)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm, mm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

    # ── Fonts ──────────────────────────────────────────────
    FONT_DIR = '/usr/share/fonts/truetype/liberation/'
    try:
        pdfmetrics.registerFont(TTFont('LS',  FONT_DIR + 'LiberationSans-Regular.ttf'))
        pdfmetrics.registerFont(TTFont('LSB', FONT_DIR + 'LiberationSans-Bold.ttf'))
        pdfmetrics.registerFont(TTFont('LSI', FONT_DIR + 'LiberationSans-Italic.ttf'))
        R = 'LS'; B = 'LSB'; I = 'LSI'
    except:
        R = 'Helvetica'; B = 'Helvetica-Bold'; I = 'Helvetica-Oblique'

    # ── Palette ────────────────────────────────────────────
    BLACK   = colors.HexColor('#0f0f0f')
    DARK    = colors.HexColor('#1a1a1a')
    MID     = colors.HexColor('#444444')
    LIGHT   = colors.HexColor('#888888')
    XLIGHT  = colors.HexColor('#cccccc')
    ACCENT  = colors.HexColor('#1a1a1a')   # subtle black accent
    TBL_HDR = colors.HexColor('#1a1a1a')   # table header bg
    TBL_ALT = colors.HexColor('#f7f7f7')   # table alt row
    WHITE   = colors.white

    # ── Styles ─────────────────────────────────────────────
    styles_base = getSampleStyleSheet()
    def S(name, **kw):
        base_kw = dict(fontName=R, fontSize=10.5, textColor=DARK, leading=16)
        base_kw.update(kw)
        return ParagraphStyle(name, parent=styles_base['Normal'], **base_kw)

    title_s   = S('title',  fontName=B, fontSize=22, textColor=BLACK,
                  alignment=TA_CENTER, leading=28, spaceAfter=4)
    date_s    = S('date',   fontName=I, fontSize=9,  textColor=LIGHT,
                  alignment=TA_CENTER, spaceAfter=0)
    h1_s      = S('h1',     fontName=B, fontSize=13, textColor=BLACK,
                  spaceBefore=18, spaceAfter=5, leading=18)
    h2_s      = S('h2',     fontName=B, fontSize=11, textColor=MID,
                  spaceBefore=12, spaceAfter=3, leading=15)
    body_s    = S('body',   fontSize=10.5, textColor=DARK, leading=17, spaceAfter=4)
    bullet_s  = S('bullet', fontSize=10.5, textColor=DARK,
                  leftIndent=16, firstLineIndent=0, leading=16, spaceAfter=3)
    num_s     = S('num',    fontSize=10.5, textColor=DARK,
                  leftIndent=16, leading=16, spaceAfter=3)
    footer_s  = S('footer', fontName=I, fontSize=7.5, textColor=XLIGHT,
                  alignment=TA_CENTER)
    brand_s   = S('brand',  fontName=B, fontSize=8, textColor=LIGHT,
                  alignment=TA_CENTER, spaceAfter=0)

    # ── Helpers ────────────────────────────────────────────
    def hr(thick=0.6, col=None, before=4, after=8):
        return HRFlowable(width='100%', thickness=thick,
                          color=col or XLIGHT,
                          spaceBefore=before, spaceAfter=after)

    def md(text):
        text = _re.sub(r'[\U0001F300-\U0001FFFF\u2600-\u27BF]+', '', text)
        t = _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
        t = _re.sub(r'\*(.+?)\*',       r'<b>\1</b>', t)
        t = _re.sub(r'_(.+?)_',           r'<i>\1</i>', t)
        return t.strip()

    def make_table(rows):
        """Render a list-of-lists as a styled table."""
        col_count = max(len(r) for r in rows)
        page_w = A4[0] - 2*2.5*cm
        col_w = page_w / col_count

        tbl = Table(rows, colWidths=[col_w]*col_count, repeatRows=1)
        tbl.setStyle(TableStyle([
            # Header row
            ('BACKGROUND',  (0,0), (-1,0), TBL_HDR),
            ('TEXTCOLOR',   (0,0), (-1,0), WHITE),
            ('FONTNAME',    (0,0), (-1,0), B),
            ('FONTSIZE',    (0,0), (-1,0), 9.5),
            ('ALIGN',       (0,0), (-1,0), 'CENTER'),
            ('BOTTOMPADDING',(0,0),(-1,0), 7),
            ('TOPPADDING',  (0,0), (-1,0), 7),
            # Body rows
            ('FONTNAME',    (0,1), (-1,-1), R),
            ('FONTSIZE',    (0,1), (-1,-1), 9.5),
            ('TEXTCOLOR',   (0,1), (-1,-1), DARK),
            ('ROWBACKGROUNDS',(0,1),(-1,-1), [WHITE, TBL_ALT]),
            ('ALIGN',       (0,1), (-1,-1), 'LEFT'),
            ('TOPPADDING',  (0,1), (-1,-1), 5),
            ('BOTTOMPADDING',(0,1),(-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING',(0,0), (-1,-1), 8),
            # Border
            ('BOX',         (0,0), (-1,-1), 0.5, XLIGHT),
            ('INNERGRID',   (0,0), (-1,-1), 0.3, XLIGHT),
            ('LINEBELOW',   (0,0), (-1,0),  1.0, colors.HexColor('#333333')),
        ]))
        return tbl

    def try_parse_table(lines, i):
        """Try to parse markdown table starting at line i. Returns (table_flowable, next_i) or (None, i)."""
        if i >= len(lines): return None, i
        row0 = lines[i].strip()
        if not (row0.startswith('|') and row0.endswith('|')): return None, i
        # Collect table lines
        tbl_lines = []
        j = i
        while j < len(lines):
            l = lines[j].strip()
            if l.startswith('|') and l.endswith('|'):
                tbl_lines.append(l)
                j += 1
            else:
                break
        if len(tbl_lines) < 2: return None, i
        # Parse
        rows = []
        for idx, l in enumerate(tbl_lines):
            cells = [c.strip() for c in l.strip('|').split('|')]
            # Skip separator row (---|---)
            if all(_re.match(r'^[-: ]+$', c) for c in cells): continue
            rows.append(cells)
        if len(rows) < 1: return None, i
        return make_table(rows), j

    # ── Build story ────────────────────────────────────────
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    # Use title as filename — clean special chars
    import re as _re_fn
    safe_title = _re_fn.sub(r'[\\/*?:"<>|]', '', title).strip()
    safe_title = safe_title[:60] if len(safe_title) > 60 else safe_title
    pdf_path = f'{safe_title}.pdf' if safe_title else f'zei_{timestamp}.pdf' 

    doc = SimpleDocTemplate(
        pdf_path, pagesize=A4,
        leftMargin=2.5*cm, rightMargin=2.5*cm,
        topMargin=3*cm, bottomMargin=2.5*cm
    )

    story = []

    # Header: title + date + thick rule
    story.append(Paragraph(title, title_s))
    story.append(Paragraph(datetime.now().strftime('%d %B %Y'), date_s))
    story.append(Spacer(1, 6))
    story.append(hr(1.2, BLACK, before=2, after=18))

    # Image
    if image_path:
        try:
            from reportlab.platypus import Image as RLImage
            import os as _os
            if _os.path.exists(image_path):
                page_w = A4[0] - 2*2.5*cm
                img = RLImage(image_path, width=page_w, height=9*cm, kind='proportional')
                story.append(img)
                story.append(Spacer(1, 14))
        except Exception as e:
            print(f'[PDF image error] {e}')

    # Content parsing
    lines = content.split('\n')
    i = 0
    while i < len(lines):
        raw = lines[i].strip()

        # Try table
        tbl_flow, next_i = try_parse_table(lines, i)
        if tbl_flow:
            story.append(Spacer(1, 8))
            story.append(tbl_flow)
            story.append(Spacer(1, 10))
            i = next_i
            continue

        if not raw:
            story.append(Spacer(1, 5))
            i += 1
            continue

        # Skip dividers
        stripped = raw.replace(' ', '')
        if len(stripped) > 2 and len(set(stripped)) <= 2 and set(stripped) <= set('─-=_'):
            i += 1
            continue

        # Strip status emojis
        raw = _re.sub(r'[\u2705\u2b1c\u25fe\u2714\u274c\u2713\u25cb]', '', raw).strip()
        if not raw:
            i += 1
            continue

        # Heading detection
        # H1: ## text  or  *text* (bold alone on line)
        if raw.startswith('## '):
            story.append(Paragraph(md(raw[3:]), h1_s))
        elif raw.startswith('# '):
            story.append(Paragraph(md(raw[2:]), h1_s))
        elif raw.startswith('### '):
            story.append(Paragraph(md(raw[4:]), h2_s))
        elif raw.startswith('*') and raw.endswith('*') and raw.count('*') == 2 and len(raw) < 120:
            story.append(Paragraph(md(raw[1:-1]), h1_s))
        elif raw.startswith('**') and raw.endswith('**') and len(raw) < 120:
            story.append(Paragraph(md(raw), h2_s))
        elif raw.startswith('- ') or raw.startswith('• '):
            story.append(Paragraph(f'• {md(raw[2:].strip())}', bullet_s))
        elif _re.match(r'^\d+[.)]', raw):
            story.append(Paragraph(md(raw), num_s))
        else:
            story.append(Paragraph(md(raw), body_s))

        i += 1

    # ── Footer / Branding ──────────────────────────────────
    story.append(Spacer(1, 24))
    story.append(hr(0.5, XLIGHT, before=0, after=8))

    # Zephy branding block
    story.append(Paragraph('ZEI', brand_s))
    story.append(Paragraph(
        f'Powered by Zephy Intelligence  ·  {datetime.now().strftime("%d %b %Y")}',
        footer_s
    ))

    doc.build(story)
    return pdf_path


def is_pdf_request_text(message: str) -> bool:
    """Detect if user wants a PDF of the last response."""
    msg = message.lower()
    triggers = [
        # English
        'make pdf', 'create pdf', 'generate pdf', 'save as pdf',
        'convert to pdf', 'pdf please', 'send pdf', 'get pdf',
        'pdf format', 'download pdf', 'pdf download', 'as a pdf',
        'in pdf', 'to pdf', 'make a pdf', 'give pdf', 'pdf of this',
        # Hindi/Hinglish
        'pdf bana', 'pdf de', 'pdf mein', 'pdf kar', 'pdf dedo',
        'pdf chahiye', 'pdf bnao', 'pdf banao', 'pdf mein convert',
        'pdf save', 'pdf generate', 'isko pdf', 'yeh pdf', 'ye pdf',
    ]
    return any(t in msg for t in triggers)

def is_audio_request_text(message: str) -> bool:
    """Detect if user wants an audio/voice reply"""
    msg = message.lower().strip()
    if len(msg) > 100:
        return False
    triggers = [
        'audio bhej', 'audio do', 'audio send', 'audio chahiye',
        'voice bhej', 'voice do', 'voice send', 'voice chahiye',
        'awaaz mein bata', 'bol ke bata', 'sunao', 'suna do',
        'speak karo', 'bolo', 'bolke do', 'audio mein',
        'voice mein', 'audio reply', 'voice reply',
        'send audio', 'send voice', 'audio me batao',
    ]
    return any(t in msg for t in triggers)





# ============================================================
# ✅ REAL-TIME AI CONTEXT: Always reads fresh from DB
# ============================================================
def format_syllabus_direct(syllabus_data, progress_data, subject_filter=None) -> str:
    """Format syllabus directly — no AI, no truncation, 100% accurate."""
    lines = []
    for subject in syllabus_data:
        # Filter by subject if requested
        if subject_filter:
            sf = subject_filter.lower()
            sname = subject['subject_name'].lower()
            scode = subject['subject_code'].lower()
            if sf not in sname and sf not in scode:
                # Check common aliases
                aliases = {
                    'food production': 'bhm-201', 'f&b': 'bhm-202',
                    'food and beverage': 'bhm-202', 'fb': 'bhm-202',
                    'housekeeping': 'bhm-203', 'hk': 'bhm-203',
                    'front office': 'bhm-204', 'fo': 'bhm-204',
                }
                matched = False
                for alias, code in aliases.items():
                    if alias in sf and code in scode:
                        matched = True
                        break
                if not matched:
                    continue

        lines.append(f"*{subject['subject_name']} ({subject['subject_code']})*")
        lines.append("")
        seen_units = set()
        for unit in subject['units']:
            if unit['unit_number'] in seen_units:
                continue
            seen_units.add(unit['unit_number'])
            key = f"{subject['subject_code']}-{unit['unit_number']}"
            is_done = progress_data.get(key, False)
            marker = "\u2705" if is_done else "-"
            lines.append(f"{marker} Unit {unit['unit_number']}: {unit['unit_name']}")
            if unit.get('topics'):
                lines.append(unit['topics'])
            lines.append("")
        lines.append("")

    if not lines:
        return None
    lines.append("Kisi bhi unit ko samjhna ho toh bas bolo!")
    return "\n".join(lines)


def _duckduckgo_fallback(query: str) -> str:
    """DuckDuckGo fallback — free, no API key, less reliable"""
    try:
        import requests as _req, re as _re
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        # Get vqd token
        r = _req.get(f'https://duckduckgo.com/?q={_req.utils.quote(query)}&ia=web',
                     headers=headers, timeout=8)
        vqd = _re.search(r'vqd=([\d-]+)', r.text)
        if not vqd:
            return ""
        # Get results
        r2 = _req.get(
            f'https://links.duckduckgo.com/d.js?q={_req.utils.quote(query)}&vqd={vqd.group(1)}&p=1',
            headers=headers, timeout=8
        )
        snippets = _re.findall(r'"a":"([^"]{30,}?)"', r2.text)[:3]
        if not snippets:
            return ""
        import html as _html
        parts = [_html.unescape(s.replace('\\n', ' ').replace('\\t', ' ')) for s in snippets]
        return "\n\n".join(parts)
    except Exception as e:
        print(f"[DDG fallback error] {e}")
        return ""


def web_search(query: str, max_results: int = 4) -> dict:
    """Search web — Tavily primary, DuckDuckGo fallback.
    Returns dict: {'results': str, 'source': 'tavily'|'ddg'|'none', 'failed': bool}"""

    # ── Try Tavily ──────────────────────────────────────────
    if tavily_client:
        try:
            response = tavily_client.search(
                query=query,
                search_depth="advanced",
                max_results=max_results,
                include_answer=True,
                include_raw_content=False,
            )
            parts = []
            if response.get('answer'):
                parts.append(f"DIRECT ANSWER: {response['answer']}")
            for r in response.get('results', [])[:max_results]:
                title   = r.get('title', '')
                content = r.get('content', '')[:400]
                url     = r.get('url', '')
                domain  = url.split('/')[2] if url.startswith('http') else ''
                parts.append(f"[{domain}] {title}: {content}")
            if parts:
                print(f"[Tavily] ✅ Got results ({len(parts)} items)")
                return {'results': "\n\n".join(parts), 'source': 'tavily', 'failed': False}
            print("[Tavily] No results returned")
        except Exception as e:
            print(f"[Tavily] ❌ Failed: {e}")

    # ── Fallback: DuckDuckGo ────────────────────────────────
    print("[DDG] Trying DuckDuckGo fallback...")
    ddg_results = _duckduckgo_fallback(query)
    if ddg_results:
        print(f"[DDG] ✅ Got fallback results")
        return {'results': ddg_results, 'source': 'ddg', 'failed': False}

    # ── Both failed ─────────────────────────────────────────
    print("[Web] ❌ Both Tavily and DDG failed")
    return {'results': '', 'source': 'none', 'failed': True}


def needs_web_search(message: str) -> bool:
    """Search web for EVERYTHING except internal academic bot data."""
    msg = message.lower()

    # ❌ ONLY skip these — pure internal bot queries
    academic_skip = [
        'syllabus', 'timetable', 'homework', 'attendance',
        'mere classes', 'meri attendance', 'mera schedule',
        'aaj ki class', 'kal ki class', 'unit complete',
        'practical kab', 'theory kab',
        'akash sir', 'mohit sir', 'aarti maam', 'jaya maam', 'ummul maam',
        'pdf bana', 'make pdf', 'generate pdf',
        'audio bhej', 'voice bhej',
        'cr update', '/update', '/hw', '/complete',
    ]
    if any(w in msg for w in academic_skip):
        return False

    # Short greetings / chitchat — no search needed
    chitchat = ['hello', 'hi ', 'hey ', 'kya haal', 'kaise ho',
                'good morning', 'good night', 'shukriya', 'thanks', 'thank you',
                'ok', 'okay', 'haan', 'nahi', 'theek', 'accha']
    if msg.strip() in chitchat or any(msg.strip() == w for w in chitchat):
        return False
    if len(msg.strip()) < 6:
        return False

    # ✅ ALWAYS search — explicit requests
    explicit = ['search', 'google', 'internet pe', 'web pe', 'dhundho',
                'find out', 'look up', 'batao internet se']
    if any(w in msg for w in explicit):
        return True

    # ✅ ALWAYS search — life/death/status of people
    status_q = ['dead', 'alive', 'died', 'death', 'mar gaya', 'mar gayi',
                'zinda', 'murder', 'arrested', 'resign', 'fired', 'elected',
                'won election', 'lost election', 'president', 'prime minister']
    if any(w in msg for w in status_q):
        return True

    # ✅ ALWAYS search — current events & news
    news_q = ['news', 'latest', 'current', 'recent', 'today', 'abhi',
              'right now', 'just now', 'breaking', 'update', 'updates',
              'what happened', 'kya hua', 'kya ho raha', 'kya chal raha',
              'geopolit', 'war', 'attack', 'conflict', 'crisis', 'protest',
              'disaster', 'earthquake', 'flood', 'accident',
              'election', 'vote', 'result', 'winner']
    if any(w in msg for w in news_q):
        return True

    # ✅ ALWAYS search — people & public figures
    people_q = ['net worth', 'networth', 'salary', 'richest', 'billionaire',
                'ceo', 'founder', 'who is', 'who are', 'wife of', 'husband of',
                'age of', 'born', 'biography',
                'elon musk', 'bezos', 'zuckerberg', 'modi', 'trump', 'obama',
                'putin', 'gates', 'ambani', 'adani', 'kohli', 'sachin',
                'khamenei', 'biden', 'zelensky', 'xi jinping',
                'shah rukh', 'salman khan', 'taylor swift', 'celebrity']
    if any(w in msg for w in people_q):
        return True

    # ✅ ALWAYS search — finance & markets
    finance_q = ['stock', 'share price', 'sensex', 'nifty', 'nasdaq', 'dow jones',
                 'exchange rate', 'usd', 'inr', 'dollar', 'rupee', 'euro',
                 'bitcoin', 'crypto', 'ethereum', 'gold price', 'silver price',
                 'oil price', 'petrol', 'diesel price', 'inflation', 'gdp']
    if any(w in msg for w in finance_q):
        return True

    # ✅ ALWAYS search — sports
    sports_q = ['score', 'match', 'ipl', 'cricket', 'football', 'fifa',
                'olympics', 'tournament', 'championship', 'league', 'won',
                'lost', 'player', 'team']
    if any(w in msg for w in sports_q):
        return True

    # ✅ ALWAYS search — tech & products
    tech_q = ['released', 'launched', 'new iphone', 'new android', 'new model',
              'specs', 'review', 'price of', 'cost of', 'buy', 'available',
              'samsung', 'apple', 'google', 'microsoft', 'openai', 'chatgpt',
              'gemini', 'claude', 'grok', 'new ai']
    if any(w in msg for w in tech_q):
        return True

    # ✅ ALWAYS search — weather
    if any(w in msg for w in ['weather', 'mausam', 'temperature', 'rain', 'barish', 'garmi', 'sardi']):
        return True

    # ✅ Anything with a question word about the real world
    # "is X?", "was X?", "are X?", "did X?", "has X?"
    import re as _re
    world_patterns = [
        r'^is .{3,}',
        r'^was .{3,}',
        r'^are .{3,}',
        r'^did .{3,}',
        r'^has .{3,}',
        r'^what is .{3,}',
        r'^who is .{3,}',
        r'^when did .{3,}',
        r'^how much .{3,}',
        r'^how many .{3,}',
    ]
    for pat in world_patterns:
        if _re.match(pat, msg):
            return True

    return False


def get_ai_response(telegram_id, user_message):
    history = get_conversation_history(telegram_id)
    reg_status = get_registration_status(telegram_id)

    if reg_status and reg_status['is_verified'] == 1:
        update_last_active(telegram_id)

    if not reg_status:
        system_prompt = """You are Zei, an intelligent academic assistant created by Zephy Intelligence.

Current task: Get the student's Student ID for registration.

Instructions:
- Keep responses conversational (1-2 sentences)
- Match the user's language style (Hindi/English/Urdu)
- Ask for their Student ID naturally
- If they give you something that looks like a student ID (alphanumeric, 8-15 chars), extract it

When you identify a student ID, respond with ONLY this JSON:
{"action": "check_student_id", "student_id": "THE_ID_YOU_FOUND"}

Examples:
User: "hi" → "Hey! I need your Student ID to register you."
User: "202506413" → {"action": "check_student_id", "student_id": "202506413"}"""

    elif reg_status['is_verified'] == 0:
        system_prompt = """You are Zei, an intelligent academic assistant created by Zephy Intelligence.

Current task: Student is in OTP verification stage - CRITICAL PRIORITY.

ABSOLUTE RULES:
- If user sends ANY 6-digit number → respond with ONLY JSON: {"action": "verify_otp", "otp": "NUMBER"}
- DO NOT ask for OTP again
- DO NOT respond with any text
- ONLY JSON output allowed

Examples:
User: "202983" → {"action": "verify_otp", "otp": "202983"}
User: "123456" → {"action": "verify_otp", "otp": "123456"}

CRITICAL: NO TEXT RESPONSES - ONLY JSON!"""

    else:
        student_details = get_student_details(reg_status['student_id'])
        greeting = get_greeting()
        first_name = student_details['name'].split()[0]

        now = datetime.now()
        current_time = now.strftime('%H:%M')
        today_name = now.strftime('%A')
        today_date = now.strftime('%B %d, %Y')
        weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

        cr_info = is_cr(student_details['student_id'])
        cr_status_text = "🎖️ You are a Class Representative (CR)" if cr_info['is_cr'] else ""
        urdu_type = student_details.get('urdu_type', 'regular')

        today_schedule = get_today_schedule(
            student_details['course'], student_details['department'],
            student_details['semester'], student_details['group']
        )
        active_updates = get_active_updates(
            student_details['course'], student_details['department'], student_details['semester']
        )
        schedule_text = format_schedule_for_ai(today_schedule)
        updates_text = format_updates_for_ai(active_updates)

        today_date_obj = now.date()
        weekdays_num = {'Monday':0,'Tuesday':1,'Wednesday':2,'Thursday':3,'Friday':4,'Saturday':5,'Sunday':6}

        def get_cancelled_subjects_for_date(course, dept, sem, target_date):
            import sqlite3 as _sqlite3
            _conn = _sqlite3.connect('students.db', timeout=20)
            _c = _conn.cursor()
            _c.execute('''SELECT update_type, subject_name FROM class_updates
                         WHERE course = ? AND department = ? AND semester = ?
                         AND target_date = ? AND is_active = 1''',
                      (course, dept, sem, str(target_date)))
            updates = _c.fetchall()
            _conn.close()
            cancelled = set()
            all_cancelled = False
            for upd_type, subj in updates:
                if upd_type == 'cancelled':
                    subj_clean = (subj or '').strip()
                    if subj_clean == '':
                        all_cancelled = True
                    else:
                        cancelled.add(subj_clean.lower())
            return all_cancelled, cancelled

        week_schedule_text = "\n📅 COMPLETE WEEK SCHEDULE (already merged with updates — TRUST THIS):\n"
        week_schedule_text += "⚠️ If a class is missing below, it is CANCELLED. Never show cancelled classes as active.\n"
        for day in weekdays:
            day_num = weekdays_num.get(day, 0)
            days_ahead = (day_num - today_date_obj.weekday()) % 7
            day_date = today_date_obj if days_ahead == 0 else today_date_obj + timedelta(days=days_ahead)
            
            all_cancelled, cancelled_subjects = get_cancelled_subjects_for_date(
                student_details['course'], student_details['department'],
                student_details['semester'], day_date
            )
            
            day_schedule = get_schedule_for_day(
                student_details['course'], student_details['department'],
                student_details['semester'], student_details['group'], day
            )
            
            week_schedule_text += f"\n*{day} ({day_date.strftime('%d %b')}):*\n"
            
            if all_cancelled:
                week_schedule_text += "❌ SAARI CLASSES CANCEL\n"
                continue
                
            if day_schedule:
                any_shown = False
                for cls in day_schedule:
                    subj_lower = cls['subject'].lower()
                    is_cancelled = False
                    for csubj in cancelled_subjects:
                        if csubj == subj_lower or csubj in subj_lower or subj_lower in csubj:
                            is_cancelled = True
                            break
                    if not is_cancelled:
                        week_schedule_text += f"• {cls['start_time']}-{cls['end_time']}: {cls['subject']} ({cls['teacher']}) - {cls['room']}\n"
                        any_shown = True
                if not any_shown:
                    week_schedule_text += "❌ SAARI CLASSES CANCEL (updates ke baad)\n"
            else:
                week_schedule_text += "Koi class nahi\n"

        user_msg_lower = user_message.lower()

        asking_about_homework = any(w in user_msg_lower for w in [
            'homework', 'hw', 'assignment', 'submit', 'kya karna h', 'kya submit'])
        asking_about_attendance = any(w in user_msg_lower for w in [
            'attendance', 'present', 'absent', 'meri attendance', 'kitni attendance'])
        asking_about_syllabus = any(w in user_msg_lower for w in [
            'syllabus', 'pura syllabus', 'complete syllabus', 'silabus', 'unit',
            'kitne unit', 'units hue', 'unit hua', 'unit ho', 'unit complete',
            'units complete', 'progress', 'kitna hua', 'kya cover', 'cover hua',
            'pending unit', 'units pending', 'f&b unit', 'food production unit',
            'housekeeping unit', 'front office unit'])
        asking_about_history = any(w in user_msg_lower for w in [
            'kya padhaya', 'last class', 'pichli class', 'aakhri class', 'what was taught',
            'kya hua tha', 'class mein kya', 'practical mein kya', 'theory mein kya',
            'ko kya', 'date ko', 'january', 'february', 'march', 'april', 'may', 'june',
            'july', 'august', 'september', 'october', 'november', 'december',
            'kal kya', 'parso kya', 'last week', 'pichle hafte', 'class history',
            'padhaya', 'banaya', 'banayi', 'sikha', 'sikhaya'
        ])

        date_in_query = re.search(
            r'(\d{1,2})(?:st|nd|rd|th)?\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)',
            user_msg_lower)

        is_cr_update = cr_info['is_cr'] and any(w in user_msg_lower for w in [
            'cancel', 'cancelled', 'postpone', 'shift', 'change', 'room change',
            'nahi hai', 'nahi hogi', 'nahi ho rahi', 'class nahi'])
        is_unit_complete = cr_info['is_cr'] and 'unit' in user_msg_lower and any(
            w in user_msg_lower for w in ['complete', 'ho gaya', 'khatam'])

        context_parts = []
        today_date_full = now.strftime('%A, %d %B %Y')
        context_parts.append(f"""You are Zei by Zephy Intelligence - Smart Academic Assistant.

Student: {first_name} | {student_details['course']} Sem-{student_details['semester']} | Group {student_details['group']}
{cr_status_text}

⚠️ DATE CONTEXT (CRITICAL — USE THIS ONLY, NEVER GUESS):
- TODAY = {today_date_full}
- Current time = {current_time}
- Day number in week = {now.weekday()} (0=Monday, 6=Sunday)
- When user says "aaj" or "today" → {today_date_full}
- When user says "kal" or "tomorrow" → {(now + timedelta(days=1)).strftime('%A, %d %B %Y')}
- When user says "parso" → {(now + timedelta(days=2)).strftime('%A, %d %B %Y')}
- NEVER use any other date as today. NEVER confuse day name with date.

LANGUAGE RULE — MOST IMPORTANT:
- DEFAULT language is ENGLISH — always start in English
- If user writes in Hindi (Devanagari) → switch to Hindi
- If user writes in Hinglish (Roman Hindi like "samjhao", "batao") → switch to Hinglish
- If user writes in English → reply in English
- Mirror user's language from their message
- NEVER default to Hinglish unprompted""")

        if active_updates:
            context_parts.append(f"📢 ACTIVE UPDATES:\n{updates_text}")

        context_parts.append(f"📅 TODAY ({today_name}):\n{schedule_text}")
        context_parts.append(week_schedule_text)

        recent_history = get_class_history_from_db(
            student_details['course'], student_details['department'],
            student_details['semester'],
            student_group=student_details['group'],
            limit=5
        )
        if recent_history:
            context_parts.append(f"\n🔄 RECENTLY RECORDED CLASSES (Latest First):\n{format_class_history_for_context(recent_history)}")

        if asking_about_history or date_in_query:
            detected_subj_for_hist = detect_subject_name(user_message)
            teacher_subj = get_subject_from_teacher(user_message) if not detected_subj_for_hist else None
            subject_filter = detected_subj_for_hist or teacher_subj or None

            specific_date = parse_date_from_text(user_message)
            if date_in_query and not specific_date:
                month_map = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                             'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
                day_num = int(date_in_query.group(1))
                month_num = month_map.get(date_in_query.group(2))
                if month_num:
                    year = now.year
                    try:
                        specific_date = datetime(year, month_num, day_num).date()
                    except:
                        specific_date = None

            history_records = get_class_history_from_db(
                student_details['course'], student_details['department'],
                student_details['semester'],
                subject_name=subject_filter,
                class_date=specific_date,
                student_group=student_details['group'],
                limit=15
            )

            if history_records:
                hist_text = f"\n📚 CLASS HISTORY"
                if subject_filter:
                    hist_text += f" ({subject_filter})"
                if specific_date:
                    hist_text += f" - {specific_date}"
                hist_text += ":\n"
                hist_text += format_class_history_for_context(history_records)
                context_parts.append(hist_text)
            else:
                context_parts.append("\n📚 CLASS HISTORY: No records found for this query yet.")

        if asking_about_homework:
            local_hw = get_homework_from_db(
                student_details['course'], student_details['department'],
                student_details['semester'], student_details['group']
            )
            pending_homework = get_pending_homework(
                student_details['course'], student_details['department'],
                student_details['semester'], student_details['group']
            )
            hw_text = format_homework_for_ai(pending_homework)
            if local_hw:
                hw_text += "\n📝 HOMEWORK LOG:\n"
                for hw in local_hw:
                    hw_text += f"• {hw['subject']}"
                    if hw.get('teacher'):
                        hw_text += f" ({hw['teacher']})"
                    hw_text += f": {hw['description']} - Due: {hw['due_date']}"
                    if hw.get('due_time'):
                        hw_text += f" at {hw['due_time']}"
                    hw_text += "\n"
            context_parts.append(hw_text)

        if asking_about_attendance:
            attendance_data = get_student_attendance(student_details['student_id'])
            low_attendance = check_low_attendance(student_details['student_id'], threshold=75)
            context_parts.append(format_attendance_for_ai(attendance_data, low_attendance))

        if asking_about_syllabus:
            syllabus_data = get_subject_syllabus(student_details['semester'])
            progress_data = get_syllabus_progress(student_details['semester'])
            syllabus_text = "\nCOMPLETE SYLLABUS DATA (SIRF yahi topics batao jo listed hain — KABHI INVENT MAT KARO):\n"
            for subject in syllabus_data:
                syllabus_text += f"\n{subject['subject_name']} ({subject['subject_code']}):\n"
                seen = set()
                for unit in subject['units']:
                    if unit['unit_number'] in seen:
                        continue
                    seen.add(unit['unit_number'])
                    is_done = progress_data.get(f"{subject['subject_code']}-{unit['unit_number']}", False)
                    status = "COMPLETED" if is_done else "PENDING"
                    syllabus_text += f"  {status} | Unit {unit['unit_number']}: {unit['unit_name']}\n"
                    if unit.get('topics'):
                        syllabus_text += f"  Topics: {unit['topics']}\n"
            context_parts.append(syllabus_text)
            context_parts.append(
                "SYLLABUS DISPLAY FORMAT — clean and minimal:\n"
                "\n"
                "*BHM-201 — Food Production Foundation - II*\n"
                "\n"
                "\u2705 Unit 1: Soups\n"
                "Introduction, Classification and types - Broth, Bouillon, Puree, Cream, Veloute, Chowder, Bisque, Garnishes and accompaniments, International Soups\n"
                "\n"
                "- Unit 2: Egg, Poultry & Game\n"
                "Eggs: Introduction, Usage in Kitchen, Structure of an Egg, Classification, Grading of Eggs, Types, Selection, Storage and preparation of breakfast dishes with eggs. Poultry: Introduction, Classification, Selection Criterion, Cuts of Poultry, Yield and simple Indian preparations, Storage and Handling\n"
                "\n"
                "[repeat for ALL units exactly like this]\n"
                "\n"
                "FORMAT RULES:\n"
                "- Subject heading: *bold* — one line, nothing else\n"
                "- Completed unit: \u2705 Unit X: Name\n"
                "- Pending unit:   - Unit X: Name\n"
                "- Topics: COPY EXACTLY from DB data — do NOT summarize, do NOT shorten\n"
                "- Blank line between each unit\n"
                "- NO divider lines, NO [ ], NO [DONE], NO _italic_ topics\n"
                "- NO extra emojis except \u2705 for done units\n"
                "- End with: Kisi bhi unit ko samjhna ho toh bas bolo!"
            )

        # 🔥 UPDATED INTELLIGENT RULES - General Knowledge Fix
        context_parts.append(f"""
🤖 INTELLIGENT RULES:

1. **CLASS HISTORY QUERIES:**
   - User: "Akash sir ne kya padhaya?" → Search class_history by teacher
   - User: "Food Production last class?" → Filter by subject
   - User: "17 feb ko kya hua tha?" → Filter by date
   - User: "Kal ka practical kya tha?" → Yesterday's practical
   - User: "Group A ka last practical?" → Filter by group + class_type=practical
   - Always show: date, subject, teacher, topics, key points, homework assigned
   - ✅ REAL-TIME: Data is always fresh from DB - no lag

2. **SYLLABUS QUERIES — STRICT:**
   ✅ ONLY show syllabus topics that are present in the COMPLETE SYLLABUS data above.
   ✅ COMPLETED/PENDING status = EXACTLY what is in the DB data above — trust it 100%.
   ❌ NEVER use conversation history to determine completion status — DB is ground truth.
   ❌ NEVER add, invent, or expand topics that are not explicitly listed in the data.
   ❌ If a subject/unit is NOT in the syllabus data → say "Mujhe is subject ka syllabus nahi mila, admin se confirm karo"
   ❌ Do NOT use your own knowledge to fill gaps — only show what is in the DB data above.
   Example: User just marked Unit 2 complete → DB will show COMPLETED → show that. Don't second-guess.

3. **SCHEDULE QUERIES — CRITICAL:**
   The COMPLETE WEEK SCHEDULE above is GROUND TRUTH. Cancelled classes already removed.
   ✅ If a class NOT listed for a day → it IS CANCELLED. Say so confidently.
   ❌ NEVER show a class as active if it's missing from the merged schedule above.
   ❌ User poochtay rahe — answer wahi rehta hai jo schedule mein hai.
   Example: "Jaya ma'am ki class hogi?" → Check that day's list → not there = "Cancel hai".

4. **HOMEWORK QUERIES:**
   - User: "Kya homework h?" → Show all pending
   - User: "Mohit sir ka homework?" → Filter by teacher name
   - User: "Food Production assignment?" → Filter by subject

5. **CR ACTIONS (Detect naturally - NO SLASH NEEDED):**
   
   A. CLASS UPDATES (AI powered - no template):
   - User: "Kal English cancel hai" → {{"action": "cr_update", "text": "..."}}
   - User: "Food Production room 201 mein shift ho gayi" → {{"action": "cr_update", "text": "..."}}
   - User: "Monday ki F&B 3 baje hogi" → {{"action": "cr_update", "text": "..."}}
   
   B. UNIT COMPLETION:
   - User: "Food Production Unit 1 complete" → {{"action": "unit_complete", "text": "..."}}
   
   C. THEORY UPDATE:
   - User: "Akash sir ne aaj cuts padhaye" → {{"action": "theory_update", "text": "...", "teacher": "Akash"}}
   
   D. PRACTICAL UPDATE:
   - User: "Group A ka practical ho gaya, grilled fish banaya" → {{"action": "practical_update", "text": "...", "group": "A"}}

6. **DATES:**
   - Today = {today_name}, {today_date}
   - Kal/Tomorrow = Next day | Parso = Day after tomorrow
   - Monday/Tuesday/etc = Next occurrence

7. **GENERAL KNOWLEDGE vs ACADEMIC DATA — CRITICAL:**
   
   📚 ACADEMIC DATA (class, homework, attendance, schedule, syllabus):
   - ONLY answer from DB data provided above
   - If not in DB → "Mujhe nahi pata, CR ya Admin se confirm karo"
   - NEVER invent class details, homework dates, or schedules
   
   🌍 GENERAL KNOWLEDGE (recipes, facts, how-to, science, history, general questions):
   - Use your training knowledge freely and helpfully
   - Be detailed and informative
   - Examples that should be answered:
     * "Chicken tikka recipe?" → Give full recipe with ingredients & cooking steps
     * "How to make coffee?" → Explain brewing methods in detail
     * "What is AI?" → Explain artificial intelligence
     * "Recipe of biryani?" → Provide complete recipe
     * "How does photosynthesis work?" → Scientific explanation
     * "Tell me about Taj Mahal" → Historical facts

   📖 TEACHING MODE — when student says "samjha do", "explain karo", "batao", "padhao", "teach me", "what is [topic]":
   - Act like a TEACHER, not a syllabus printer
   - DO NOT just list syllabus points — actually EXPLAIN the topic with depth
   - Structure like this:
     1. Introduction — topic kya hai, kyu important hai (2-3 lines)
     2. Main content — proper explanation with examples, real-world context
     3. End with a follow-up question in the user's own language
   - Use the SAME language the user wrote in — mirror it exactly
   - Example: "Soups samjha do" →
     * First explain what soup is, why important in hospitality
     * Then explain broth vs bouillon vs puree with examples
     * End with: "Want to go deeper into classification?"
   - NEVER dump all subtopics as bullet points in one go
   - ONE concept at a time, conversational flow
   
   🎯 KEY PRINCIPLE:
   - Academic/institution data (classes, homework, attendance, schedule) = STRICT (DB only)
   - General world knowledge (recipes, facts, explanations) = HELPFUL (use training)
   - If user asks "recipe", "how to", "what is", "tell me about" → answer helpfully
   - If user asks "kya padhaya", "homework", "attendance", "class" → check DB strictly

8. **TONE & STYLE & LANGUAGE:**
   - MIRROR the user's language exactly:
     * User writes in Hindi (Devanagari) → reply in Hindi
     * User writes in English → reply in English
     * User writes in Hinglish (Roman Hindi) → reply in Hinglish
     * User writes in Urdu mix → reply with Urdu mix
   - DO NOT switch languages mid-reply unless the topic demands it (e.g. technical terms)
   - Friendly and conversational — like a smart classmate, not a textbook
   - For recipes/how-to: detailed, step-by-step
   - For academic: precise and data-driven

9. **FORMATTING — CRITICAL (Telegram):**
   Keep it simple, clean, minimal. Like a smart friend texting you — not a textbook.

   RULES:
   - Bullets: use - (hyphen) ONLY. Never * for bullets.
   - Bold: *text* (single star each side) — only for title/heading of a response, nothing else
   - NO emojis unless user uses them first
   - NO divider lines (────── or ═══)
   - NO italic _text_
   - NO heavy section headers like *Ingredients:* — just write "Ingredients:" plain
   - Short and conversational. If answer fits in 2 lines, write 2 lines.

   PERFECT EXAMPLE — Recipe:
   *Glazed Carrots*
   Serves: 2-3 | Time: 10-12 mins

   Ingredients:
   - 4-6 carrots, peeled and sliced
   - 2 tbsp butter, 1 tbsp honey, 1 tbsp brown sugar
   - Salt, pepper, pinch of cinnamon (optional)

   Instructions:
   1. Melt butter in pan on medium heat
   2. Add carrots, cook 5-7 mins till slightly soft
   3. Mix honey, sugar, salt, pepper — pour over carrots
   4. Cook 2-3 mins till glaze thickens
   5. Serve hot. Add lemon or parsley if you like.

   PERFECT EXAMPLE — Simple question:
   F&B mein 2 units ho gaye hain — Unit 1 aur Unit 2. Baaki 3 pending hain.

   NEVER DO THIS:
   ✅ *Unit 1: Soups* 📚
   ─────────────────
   _Introduction, Broth, Bouillon..._

CRITICAL: Distinguish academic queries (need DB) from general queries (use knowledge)!
Examples:
- "Chicken tikka recipe" → Answer with full recipe from your knowledge
- "Kya homework hai" → Check DB strictly
- "How to make pasta" → Answer with cooking steps from your knowledge
- "Akash sir ne kya padhaya" → Check DB strictly""")

        system_prompt = "\n\n".join(context_parts)

    # ✅ Web search — inject live data if query needs it
    if needs_web_search(user_message):
        print(f"[Web] Searching: {user_message}")
        search_result = web_search(user_message)

        if not search_result['failed'] and search_result['results']:
            src_label = 'Live Web' if search_result['source'] == 'tavily' else 'Web (DuckDuckGo fallback)'
            system_prompt += (
                f"\n\n━━━ LIVE WEB DATA ({src_label}) ━━━\n"
                f"{search_result['results']}\n"
                f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                f"RULES:\n"
                f"- Use ONLY this data for current facts, prices, news, net worth\n"
                f"- Ignore your training data for anything time-sensitive\n"
                f"- If DIRECT ANSWER is present, lead with that\n"
                f"- Cite source domain in brackets e.g. [forbes.com]\n"
                f"- End response with: _Source: {src_label}_"
            )
        else:
            # Both sources failed — tell AI to add disclaimer
            system_prompt += (
                f"\n\nWEB SEARCH FAILED: Could not fetch live data for this query.\n"
                f"IMPORTANT: You MUST add this disclaimer at the end of your response:\n"
                f"'⚠️ Live web search unavailable right now. This answer is based on my training data "
                f"and may be outdated. For latest info, please check Google or a news source.'"
            )
            print("[Web] ⚠️ Both sources failed — disclaimer will be added")

    messages = [{"role": "system", "content": system_prompt}]
    messages.extend(history[-10:])
    messages.append({"role": "user", "content": user_message})

    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            temperature=0.4,
            max_tokens=2500  # Increased for full syllabus display
        )
        ai_response = response.choices[0].message.content
        # Fix double-star bold (**text**) -> Telegram single-star (*text*)
        import re as _re
        ai_response = _re.sub('\\*\\*(.+?)\\*\\*', lambda m: '*' + m.group(1) + '*', ai_response)
        save_message(telegram_id, "user", user_message)
        save_message(telegram_id, "assistant", ai_response)
        return ai_response
    except Exception as e:
        print(f"Groq error: {e}")
        return "Sorry, there's an issue. Please try again."

# ============================================================
# ✅ GROQ OCR - IMAGE HANDLER
# ============================================================
async def handle_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    try:
        processing_msg = await update.message.reply_text("📷 Reading image... 🔍")

        # Get the highest resolution photo
        photo = update.message.photo[-1]
        file = await context.bot.get_file(photo.file_id)

        # Download image bytes
        img_path = f"ocr_{telegram_id}_{int(datetime.now().timestamp())}.jpg"
        await file.download_to_drive(img_path)

        with open(img_path, 'rb') as f:
            image_bytes = f.read()

        try:
            os.remove(img_path)
        except:
            pass

        # Run Groq OCR
        extracted_text = await groq_ocr_from_image(image_bytes, "image/jpeg")

        await processing_msg.delete()

        if not extracted_text:
            await update.message.reply_text("❌ Could not read text from image. Please try a clearer photo.")
            return

        # Show extracted text
        caption = update.message.caption or ""
        display_text = extracted_text[:300] + "..." if len(extracted_text) > 300 else extracted_text

        await update.message.reply_text(
            f"📷 *Text extracted from image:*\n\n`{display_text}`\n\n🤖 Processing...",
            parse_mode='Markdown'
        )

        # Combine caption + extracted text and pass to AI
        combined_message = f"{caption}\n\nImage text: {extracted_text}" if caption else f"Image text: {extracted_text}"
        ai_response = get_ai_response(telegram_id, combined_message)
        await process_ai_actions(update, context, ai_response, telegram_id, user_message=None)

    except Exception as e:
        print(f"Image OCR error: {e}")
        await update.message.reply_text("❌ Error reading image. Please try again.")

# ============================================================
# 🔥 VOICE HANDLER - transcribe → AI text reply → Sarvam voice reply
# ============================================================
async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    try:
        processing_msg = await update.message.reply_text("🎤 Processing...")
        voice = update.message.voice
        file = await context.bot.get_file(voice.file_id)
        voice_path = f"voice_{telegram_id}_{int(datetime.now().timestamp())}.ogg"
        await file.download_to_drive(voice_path)

        # Step 1: Transcribe audio with Groq Whisper
        with open(voice_path, "rb") as audio_file:
            transcript = groq_client.audio.transcriptions.create(
                model="whisper-large-v3-turbo",
                file=audio_file,
                response_format="text"
            )
        transcribed_text = transcript.strip()
        try:
            os.remove(voice_path)
        except:
            pass

        await processing_msg.delete()

        # Step 2: Show transcription to user
        await update.message.reply_text(
            f"🎤 *You said:*\n_{transcribed_text}_",
            parse_mode='Markdown'
        )

        # Step 3: Get AI text response
        ai_response = get_ai_response(telegram_id, transcribed_text)

        # Step 4: Send text reply first (handles CR actions too)
        await process_ai_actions(update, context, ai_response, telegram_id, user_message=transcribed_text)

        # ============================================================
        # ✅ Step 5: Sarvam Bulbul v3 voice reply — AFTER text reply
        # Skipped for CR actions (buttons/confirmations don't need TTS)
        # ============================================================
        is_cr_action = '{"action"' in ai_response
        if not is_cr_action:
            try:
                # Clean AI text for TTS — remove markdown, emojis, JSON
                tts_text = ai_response

                # Remove markdown symbols
                tts_text = tts_text.replace('*', '').replace('_', '').replace('`', '').replace('#', '')

                # Remove emojis
                tts_text = re.sub(
                    r'[\U0001F300-\U0001F9FF\U00002600-\U000027BF'
                    r'\U0001FA00-\U0001FAFF\u2702-\u27B0\u24C2-\U0001F251]+',
                    '', tts_text, flags=re.UNICODE
                )

                # Clean whitespace
                tts_text = re.sub(r'\n+', '. ', tts_text)
                tts_text = re.sub(r'\s+', ' ', tts_text).strip()

                # Trim to Sarvam 500 char limit at sentence boundary
                if len(tts_text) > 490:
                    trimmed = tts_text[:490]
                    last_stop = max(trimmed.rfind('.'), trimmed.rfind('!'), trimmed.rfind('?'))
                    tts_text = trimmed[:last_stop + 1] if last_stop > 200 else trimmed + '...'

                if tts_text:
                    audio_bytes = await sarvam_tts(tts_text)

                    if audio_bytes:
                        audio_path = f"tts_reply_{telegram_id}_{int(datetime.now().timestamp())}.wav"
                        with open(audio_path, 'wb') as f:
                            f.write(audio_bytes)

                        await update.message.reply_voice(
                            voice=open(audio_path, 'rb'),
                            caption="🔊 Voice reply"
                        )

                        try:
                            os.remove(audio_path)
                        except:
                            pass

                        print(f"✅ TTS voice reply sent to {telegram_id}")
                    else:
                        print(f"TTS returned no audio for {telegram_id}")

            except Exception as tts_error:
                # TTS optional — text reply already sent, user not affected
                print(f"TTS skipped for {telegram_id}: {tts_error}")

    except Exception as e:
        print(f"Voice error: {e}")
        await update.message.reply_text("❌ Couldn't process voice note!")

# ============================================================
# ✅ SARVAM TTS COMMAND - /speak
# ============================================================
async def handle_speak(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    text = update.message.text.replace('/speak', '').strip()
    if not text:
        await update.message.reply_text("Usage: `/speak [text to speak]`", parse_mode='Markdown')
        return

    processing_msg = await update.message.reply_text("🔊 Generating audio...")
    audio_bytes = await sarvam_tts(text)

    if audio_bytes:
        await processing_msg.delete()
        audio_path = f"tts_{telegram_id}_{int(datetime.now().timestamp())}.wav"
        with open(audio_path, 'wb') as f:
            f.write(audio_bytes)
        await update.message.reply_voice(voice=open(audio_path, 'rb'))
        try:
            os.remove(audio_path)
        except:
            pass
    else:
        await processing_msg.edit_text("❌ TTS failed. Please try again.")

# ============================================================
# ✅ IMPROVED PROCESS AI ACTIONS
# ============================================================
async def process_ai_actions(update, context, ai_response, telegram_id, user_message=None):
    try:
        if '{"action"' in ai_response:
            json_start = ai_response.find('{"action"')
            json_end = ai_response.find('}', json_start) + 1
            json_str = ai_response[json_start:json_end]
            action_data = json.loads(json_str)
            action = action_data.get('action')

            reg_status = get_registration_status(telegram_id)
            student_details = get_student_details(reg_status['student_id'])
            cr_info = is_cr(student_details['student_id'])

            if not cr_info['is_cr']:
                await _safe_send(update, ai_response)
                return

            if action == 'cr_update':
                conn = sqlite3.connect('students.db')
                c = conn.cursor()
                c.execute('INSERT INTO pending_cr_actions (telegram_id, action_type, action_data, created_at) VALUES (?, ?, ?, ?)',
                          (telegram_id, 'class_update', action_data.get('text', ''), datetime.now()))
                pending_id = c.lastrowid
                conn.commit()
                conn.close()

                # ✅ AI powered preview - show what AI understood, not a template
                update_text = action_data.get('text', '')
                keyboard = [[
                    InlineKeyboardButton("✅ Yes, Broadcast", callback_data=f'confirm_update_{pending_id}'),
                    InlineKeyboardButton("❌ Cancel", callback_data=f'cancel_update_{pending_id}')
                ]]
                await update.message.reply_text(
                    f"📢 *Broadcast this update?*\n\n{update_text}\n\n_Tap ✅ to send to all students_",
                    parse_mode='Markdown',
                    reply_markup=InlineKeyboardMarkup(keyboard)
                )
                return

            elif action == 'unit_complete':
                text = action_data.get('text', '')
                detected_subject = detect_subject_name(text)
                parse_prompt = f"""Extract from: "{text}"
Respond ONLY with JSON:
{{"subject_code": "...", "unit_number": "1/2/3/4/5/6"}}"""
                response = groq_client.chat.completions.create(
                    model="llama-3.3-70b-versatile",
                    messages=[{"role": "user", "content": parse_prompt}],
                    temperature=0.2, max_tokens=100
                )
                parsed = json.loads(response.choices[0].message.content)
                subject_id = detected_subject if detected_subject else parsed['subject_code']
                result = mark_topic_complete(
                    student_details['semester'], subject_id,
                    int(parsed['unit_number']), student_details['student_id']
                )
                if result['success']:
                    await update.message.reply_text(
                        f"✅ *Unit Marked Complete!*\n\n"
                        f"📚 {result['subject_name']}\n"
                        f"📖 Unit {result['unit_number']}: {result['unit_name']}\n\n"
                        f"Students can now see this as completed!",
                        parse_mode='Markdown'
                    )
                else:
                    await update.message.reply_text(f"❌ {result['message']}")
                return

            elif action == 'theory_update':
                text = action_data.get('text', '')
                teacher_hint = action_data.get('teacher', '')
                await _process_theory_save(update, context, student_details, text, teacher_hint)
                return

            elif action == 'practical_update':
                text = action_data.get('text', '')
                group_hint = action_data.get('group', student_details['group'])
                await _process_practical_save(update, context, student_details, text, group_hint)
                return

        # ✅ Check if student asking about last practical (Food Production) → send photo too
        await _send_reply_with_photo(update, context, ai_response, user_message, telegram_id)

    except Exception as e:
        print(f"Process error: {e}")
        await _safe_send(update, ai_response)

# ============================================================
# REMINDER SYSTEMS
# ============================================================
async def send_class_reminders(application):
    while True:
        try:
            now = datetime.now()
            current_day = now.strftime('%A')
            reminder_time = (now + timedelta(minutes=15)).strftime('%H:%M')

            conn = sqlite3.connect('students.db', timeout=20)
            c = conn.cursor()
            c.execute('''SELECT DISTINCT course, department, semester, student_group,
                                subject_name, teacher_name, start_time, end_time, room_number, urdu_type
                         FROM class_schedule
                         WHERE day_of_week = ? AND start_time = ? AND is_active = 1''',
                      (current_day, reminder_time))
            upcoming_classes = c.fetchall()

            for cls in upcoming_classes:
                course, dept, sem, group, subject, teacher, start, end, room = cls[:9]
                urdu_type = cls[9] if len(cls) > 9 else 'both'
                class_key = f"{current_day}_{start}_{subject}_{group}_{urdu_type}"

                c.execute('SELECT id FROM class_reminders_sent WHERE class_key = ? AND DATE(sent_at) = DATE(?)',
                          (class_key, now))
                if c.fetchone():
                    continue

                if group == 'BOTH' and urdu_type == 'both':
                    c.execute('''SELECT ru.telegram_id FROM registered_users ru
                                 JOIN master_students ms ON ru.student_id = ms.student_id
                                 WHERE ms.course = ? AND ms.department = ? AND ms.semester = ? AND ru.is_verified = 1''',
                              (course, dept, sem))
                elif group == 'BOTH':
                    c.execute('''SELECT ru.telegram_id FROM registered_users ru
                                 JOIN master_students ms ON ru.student_id = ms.student_id
                                 WHERE ms.course = ? AND ms.department = ? AND ms.semester = ?
                                 AND (ms.urdu_type = ? OR ms.urdu_type IS NULL) AND ru.is_verified = 1''',
                              (course, dept, sem, urdu_type))
                elif urdu_type == 'both':
                    c.execute('''SELECT ru.telegram_id FROM registered_users ru
                                 JOIN master_students ms ON ru.student_id = ms.student_id
                                 WHERE ms.course = ? AND ms.department = ? AND ms.semester = ?
                                 AND ms.student_group = ? AND ru.is_verified = 1''',
                              (course, dept, sem, group))
                else:
                    c.execute('''SELECT ru.telegram_id FROM registered_users ru
                                 JOIN master_students ms ON ru.student_id = ms.student_id
                                 WHERE ms.course = ? AND ms.department = ? AND ms.semester = ?
                                 AND ms.student_group = ? AND (ms.urdu_type = ? OR ms.urdu_type IS NULL)
                                 AND ru.is_verified = 1''',
                              (course, dept, sem, group, urdu_type))

                students = [row[0] for row in c.fetchall()]
                urdu_text = " [Advanced Urdu]" if urdu_type == 'advanced' else ""
                reminder_text = f"🔔 Reminder! {subject}{urdu_text} class 15 minutes mein - {start} se {teacher} ke saath {room} mein. Get ready! 📚"

                sent_count = 0
                for tid in students:
                    try:
                        await application.bot.send_message(chat_id=tid, text=reminder_text)
                        sent_count += 1
                    except Exception as e:
                        print(f"Failed to send to {tid}: {e}")

                c.execute('INSERT INTO class_reminders_sent (class_key, sent_at) VALUES (?, ?)', (class_key, now))
                conn.commit()
                print(f"✅ Sent class reminder to {sent_count} students")

            conn.close()
        except Exception as e:
            print(f"❌ Reminder error: {e}")

        await asyncio.sleep(60)

async def send_homework_reminders(application):
    while True:
        try:
            now = datetime.now()
            if now.hour == 20 and now.minute == 0:
                print("🔔 Checking homework reminders...")
                homework_list = get_homework_for_reminder()

                for hw in homework_list:
                    conn = sqlite3.connect('students.db', timeout=20)
                    c = conn.cursor()
                    if hw['student_group'] == 'BOTH':
                        c.execute('''SELECT ru.telegram_id FROM registered_users ru
                                     JOIN master_students ms ON ru.student_id = ms.student_id
                                     WHERE ms.course = ? AND ms.department = ? AND ms.semester = ? AND ru.is_verified = 1''',
                                  (hw['course'], hw['department'], hw['semester']))
                    else:
                        c.execute('''SELECT ru.telegram_id FROM registered_users ru
                                     JOIN master_students ms ON ru.student_id = ms.student_id
                                     WHERE ms.course = ? AND ms.department = ? AND ms.semester = ?
                                     AND ms.student_group = ? AND ru.is_verified = 1''',
                                  (hw['course'], hw['department'], hw['semester'], hw['student_group']))
                    students = [row[0] for row in c.fetchall()]
                    conn.close()

                    try:
                        if isinstance(hw['submission_date'], str):
                            date_obj = datetime.fromisoformat(hw['submission_date'])
                        else:
                            date_obj = hw['submission_date']
                        days_left = (date_obj.date() - now.date()).days
                        due_text = "TOMORROW" if days_left == 0 else (
                            f"in {days_left + 1} days" if days_left == 1 else f"in {days_left} days")
                        date_str = date_obj.strftime('%A, %B %d')
                    except:
                        due_text = "soon"
                        date_str = str(hw['submission_date'])

                    time_str = f" at {hw['submission_time']}" if hw['submission_time'] else ""
                    reminder_msg = f"""📝 *HOMEWORK REMINDER*

📚 Subject: {hw['subject']}
📋 Task: {hw['description']}
📅 Due: {date_str}{time_str}
⏰ {due_text.upper()}

Don't forget to complete it! 💪"""

                    sent_count = 0
                    for tid in students:
                        try:
                            await application.bot.send_message(chat_id=tid, text=reminder_msg, parse_mode='Markdown')
                            sent_count += 1
                        except Exception as e:
                            print(f"Failed homework reminder to {tid}: {e}")

                    mark_homework_reminder_sent(hw['id'])
                    print(f"✅ Sent homework reminder to {sent_count} students")

                await asyncio.sleep(60)
            else:
                await asyncio.sleep(60)
        except Exception as e:
            print(f"❌ Homework reminder error: {e}")
            await asyncio.sleep(60)

# ============================================================
# ✅ INTERNAL: Theory save with auto date detection
# ============================================================
async def _process_theory_save(update, context, student_details, text, teacher_hint=''):
    detected_subject = detect_subject_name(text)
    if not detected_subject:
        teacher_subject = get_subject_from_teacher(text)
        if teacher_subject:
            detected_subject = teacher_subject
    if not detected_subject and teacher_hint:
        teacher_subject = get_subject_from_teacher(teacher_hint)
        if teacher_subject:
            detected_subject = teacher_subject

    # ✅ Auto-detect date from text
    detected_date = parse_date_from_text(text)
    class_date = detected_date if detected_date else datetime.now().date()

    parse_prompt = f"""Extract theory class details from: "{text}"

{'Detected subject: ' + detected_subject if detected_subject else ''}
{'Teacher hint: ' + teacher_hint if teacher_hint else ''}
Class date detected: {class_date}

Respond ONLY with JSON (no extra text):
{{"subject_name": "...", "teacher_name": "...", "unit_covered": "...", "topics_covered": "...", "key_points": "...", "homework": "..."}}"""

    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": parse_prompt}],
            temperature=0.2, max_tokens=300
        )
        raw = response.choices[0].message.content.strip()
        # Extract JSON even if there's surrounding text
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        parsed = json.loads(json_match.group() if json_match else raw)
        if detected_subject:
            parsed['subject_name'] = detected_subject

        result = save_last_theory_class(
            cr_student_id=student_details['student_id'],
            subject_identifier=parsed['subject_name'],
            unit_covered=parsed.get('unit_covered', ''),
            topics_covered=parsed['topics_covered'],
            key_points=parsed.get('key_points', ''),
            homework_assigned=parsed.get('homework')
        )

        if result['success']:
            # ✅ Save to local DB immediately for real-time AI recall
            save_class_history_to_db(
                course=student_details['course'],
                department=student_details['department'],
                semester=student_details['semester'],
                subject_name=result.get('subject', parsed['subject_name']),
                subject_code=parsed.get('subject_code', ''),
                class_type='theory',
                student_group='BOTH',
                teacher_name=parsed.get('teacher_name', ''),
                unit_covered=parsed.get('unit_covered', ''),
                topics_covered=parsed['topics_covered'],
                key_points=parsed.get('key_points', ''),
                homework_assigned=parsed.get('homework', ''),
                practical_work='',
                demonstrations='',
                class_date=class_date,  # ✅ Uses detected date
                recorded_by=student_details['student_id'],
                photo_path=None
            )
            date_display = class_date.strftime('%A, %B %d') if class_date != datetime.now().date() else f"Today ({datetime.now().strftime('%B %d')})"
            await update.message.reply_text(
                f"✅ *Theory Class Recorded!*\n\n"
                f"📚 Subject: {result.get('subject', parsed['subject_name'])}\n"
                f"👨‍🏫 Teacher: {parsed.get('teacher_name', 'N/A')}\n"
                f"📖 Unit: {parsed.get('unit_covered', 'N/A')}\n"
                f"📝 Topics: {parsed['topics_covered']}\n"
                f"📅 Date: {date_display}\n\n"
                f"✅ AI memory updated instantly! Students can ask now.",
                parse_mode='Markdown'
            )
        else:
            await update.message.reply_text(f"❌ {result['message']}")
    except Exception as e:
        await update.message.reply_text(f"❌ Error saving theory: {e}")

# ============================================================
# ✅ INTERNAL: Practical save with auto date + subject detection
# ============================================================
async def _process_practical_save(update, context, student_details, text, group_hint='A', photo_path=None):
    # ✅ Auto-detect subject from teacher name or subject keywords
    detected_subject = detect_subject_name(text)
    if not detected_subject:
        teacher_subject = get_subject_from_teacher(text)
        if teacher_subject:
            detected_subject = teacher_subject

    # ✅ Auto-detect date from text
    detected_date = parse_date_from_text(text)
    class_date = detected_date if detected_date else datetime.now().date()

    parse_prompt = f"""Extract practical class details from this text: "{text}"

{'Detected subject: ' + detected_subject if detected_subject else 'No subject detected - extract from context (e.g. "shepherd pie" → Food Production)'}
Group hint: {group_hint}
Class date: {class_date}

Smart subject detection rules:
- Cooking/food items → Food Production / Culinary
- Service/F&B → Food & Beverage Service
- Rooms/housekeeping → Housekeeping
- Front desk/reception → Front Office

Respond ONLY with valid JSON:
{{"group": "A or B", "subject_name": "full subject name", "teacher_name": "teacher name or empty", "unit_covered": "unit info or empty", "practical_work": "what was made/done", "demonstrations": "demo details or empty", "key_points": "key learning points or empty"}}"""

    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": parse_prompt}],
            temperature=0.2, max_tokens=400
        )
        raw = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        parsed = json.loads(json_match.group() if json_match else raw)

        # Prefer detected subject over AI guessed one
        if detected_subject:
            parsed['subject_name'] = detected_subject

        result = save_last_practical_class(
            cr_student_id=student_details['student_id'],
            student_group=parsed.get('group', group_hint),
            subject_identifier=parsed['subject_name'],
            unit_covered=parsed.get('unit_covered', ''),
            practical_work=parsed['practical_work'],
            demonstrations=parsed.get('demonstrations', ''),
            key_points=parsed.get('key_points', '')
        )

        if result['success']:
            # ✅ Save to local DB immediately for real-time AI recall
            save_class_history_to_db(
                course=student_details['course'],
                department=student_details['department'],
                semester=student_details['semester'],
                subject_name=result.get('subject', parsed['subject_name']),
                subject_code=parsed.get('subject_code', ''),
                class_type='practical',
                student_group=parsed.get('group', group_hint),
                teacher_name=parsed.get('teacher_name', ''),
                unit_covered=parsed.get('unit_covered', ''),
                topics_covered=parsed['practical_work'],
                key_points=parsed.get('key_points', ''),
                homework_assigned='',
                practical_work=parsed['practical_work'],
                demonstrations=parsed.get('demonstrations', ''),
                class_date=class_date,  # ✅ Uses detected date
                recorded_by=student_details['student_id'],
                photo_path=photo_path
            )
            date_display = class_date.strftime('%A, %B %d') if class_date != datetime.now().date() else f"Today ({datetime.now().strftime('%B %d')})"
            await update.message.reply_text(
                f"✅ *Practical Recorded!*\n\n"
                f"🔬 Subject: {result.get('subject', parsed['subject_name'])}\n"
                f"👥 Group: {result.get('group', parsed.get('group', group_hint))}\n"
                f"👨‍🏫 Teacher: {parsed.get('teacher_name', 'N/A')}\n"
                f"🍳 Work: {parsed['practical_work']}\n"
                f"📅 Date: {date_display}\n"
                f"{'📸 Photo saved!' if photo_path else ''}\n\n"
                f"✅ AI memory updated! Students can ask now.",
                parse_mode='Markdown'
            )
        else:
            await update.message.reply_text(f"❌ {result['message']}")
    except Exception as e:
        await update.message.reply_text(f"❌ Error saving practical: {e}")

# ============================================================
# ✅ /update COMMAND - AI POWERED (no fixed template)
# ============================================================
async def handle_cr_update(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    student_details = get_student_details(reg_status['student_id'])
    cr_info = is_cr(student_details['student_id'])

    if not cr_info['is_cr']:
        await update.message.reply_text("❌ Not authorized as CR.")
        return

    update_text = update.message.text.replace('/update', '').strip()

    if not update_text:
        await update.message.reply_text(
            "📢 *Class Update*\n\n"
            "Bas likho kya update hai, main samajh lunga:\n\n"
            "• `/update Kal Food Production cancel hai`\n"
            "• `/update Monday English nahi hogi`\n"
            "• `/update Aaj F&B 3 baje hogi`\n"
            "• `/update Room change - ab Room 201 mein`\n\n"
            "🗣️ *Ya bina /update ke bhi bol sakte ho!*\n"
            "\"Kal English cancel hai\" → main khud detect kar lunga.",
            parse_mode='Markdown'
        )
        return

    # ✅ AI generates smart broadcast message (no fixed template)
    ai_msg_prompt = f"""You are Zei, a student bot assistant.

A Class Representative has sent this update: "{update_text}"

Generate a clear, friendly broadcast message for students in Hindi/English mix.
Keep it under 100 words. Include relevant emojis.
Do NOT use any template format. Write naturally based on what the CR said.

Examples:
Input: "Kal food production cancel hai"
Output: "📢 Class Update!\n\n❌ Kal (Thursday) Food Production ki class CANCEL hai. Koi class nahi hogi.\n\n- Posted by CR 🎖️"

Input: "Monday English 2 baje hogi"  
Output: "📢 Class Update!\n\n⏰ Monday ki English class ab 2:00 PM pe hogi (time change).\n\n- Posted by CR 🎖️"

Respond with ONLY the broadcast message text, nothing else."""

    try:
        ai_resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": ai_msg_prompt}],
            temperature=0.3, max_tokens=200
        )
        ai_broadcast_msg = ai_resp.choices[0].message.content.strip()
    except:
        ai_broadcast_msg = f"📢 Class Update!\n\n{update_text}\n\n- Posted by CR 🎖️"

    # ✅ Parse update to save to DB (so AI's merged schedule reflects cancel)
    parse_prompt = f"""Extract update from: "{update_text}"

Teacher to subject mapping:
- akash / chef akash / dr akash → "Food Production Foundation - II"
- mohit / mr mohit / mohit sir → "Food & Beverage Service Foundation - II"
- jaya / chef jaya / jaya maam → "Personality Development and Grooming"
- aarti / dr aarti → check context: housekeeping or front office

Detect:
- day: today/tomorrow/kal/Monday/Tuesday/Wednesday/Thursday/Friday/Saturday/Sunday
- update_type: cancelled/postponed/room_change
- subject_name: FULL subject name (use teacher mapping above). Leave EMPTY ONLY if explicitly ALL classes cancelled (e.g. "saari classes cancel", "all classes cancel")
- class_type: theory/practical/both

Respond ONLY with JSON:
{{"day": "Monday", "update_type": "cancelled", "subject_name": "Personality Development and Grooming", "class_type": "both", "new_time": null, "room_change": null, "reason": null}}"""

    try:
        parse_resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": parse_prompt}],
            temperature=0.2, max_tokens=200
        )
        raw = parse_resp.choices[0].message.content.strip()
        import re as _re
        json_match = _re.search(r'\{.*\}', raw, _re.DOTALL)
        parsed = json.loads(json_match.group() if json_match else raw)
    except:
        parsed = {"day": "today", "update_type": "cancelled", "subject_name": "", "class_type": "both"}

    # Resolve target date — FUTURE direction (nearest upcoming day)
    day_str = parsed.get('day', 'today').lower()
    today_d = datetime.now().date()
    days_map_cr = {'monday': 0, 'tuesday': 1, 'wednesday': 2, 'thursday': 3,
                   'friday': 4, 'saturday': 5, 'sunday': 6}
    if day_str in ['today', 'aaj']:
        target_date = today_d
    elif day_str in ['tomorrow', 'kal']:
        target_date = today_d + timedelta(days=1)
    elif day_str in days_map_cr:
        days_ahead = (days_map_cr[day_str] - today_d.weekday()) % 7
        if days_ahead == 0:
            days_ahead = 7  # Same day of week → next occurrence
        target_date = today_d + timedelta(days=days_ahead)
    else:
        target_date = today_d

    # Save to DB
    result = post_class_update(
        cr_student_id=student_details['student_id'],
        update_type=parsed.get('update_type', 'cancelled'),
        subject_name=parsed.get('subject_name', ''),
        original_time=None,
        new_time=parsed.get('new_time'),
        room_change=parsed.get('room_change'),
        reason=parsed.get('reason'),
        class_type=parsed.get('class_type', 'both'),
        target_date=target_date
    )

    if not result['success']:
        await update.message.reply_text(f"❌ Error: {result['message']}")
        return

    # ✅ Direct broadcast — no confirm button
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    c.execute('''SELECT ru.telegram_id FROM registered_users ru
                 JOIN master_students ms ON ru.student_id = ms.student_id
                 WHERE ms.course = ? AND ms.department = ? AND ms.semester = ? AND ru.is_verified = 1''',
              (result['course'], result['department'], result['semester']))
    students = [row[0] for row in c.fetchall()]
    conn.close()

    count = 0
    for tid in students:
        try:
            await context.bot.send_message(chat_id=tid, text=ai_broadcast_msg, parse_mode='Markdown')
            count += 1
        except:
            pass

    date_display = target_date.strftime('%A, %d %b')
    await update.message.reply_text(
        f"✅ *Broadcasted!*\n\n📅 {date_display}\n👥 Sent to {count} students",
        parse_mode='Markdown'
    )

# ============================================================
# ✅ /hw COMMAND - homework with teacher + date detection
# ============================================================
async def handle_homework(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    if not is_admin(telegram_id):
        await update.message.reply_text("❌ This command is for Admin only.")
        return

    student_details = get_student_details(reg_status['student_id'])

    hw_text = update.message.text.replace('/hw', '').strip()

    if not hw_text:
        await update.message.reply_text(
            "📝 *Post Homework:*\n\n"
            "Format: `/hw [teacher/subject] - [description]. Submit: [day/date] [time]`\n\n"
            "*With Teacher Name:*\n"
            "• `/hw Mohit sir - Journal complete karna hai. Submit: Monday 10:30`\n"
            "• `/hw Akash sir - Knife cuts practice. Submit: Friday`\n"
            "• `/hw Rajesh sir - F&B assignment. Submit: 20th February`\n\n"
            "*With Subject Name:*\n"
            "• `/hw Food Production - Assignment chapter 3. Submit: Friday`\n"
            "• `/hw English - Essay likhna h. Submit: Tomorrow`\n\n"
            "🎤 *Tip: Voice notes also work!*\n\n"
            "💡 Teacher name bolo toh main automatically subject detect kar lunga!",
            parse_mode='Markdown'
        )
        return

    detected_subject = detect_subject_name(hw_text)
    if not detected_subject:
        teacher_subject = get_subject_from_teacher(hw_text)
        if teacher_subject:
            detected_subject = teacher_subject

    parse_prompt = f"""Extract homework details from: "{hw_text}"

{'Detected subject: ' + detected_subject if detected_subject else 'No subject detected yet - extract from text'}

Extract teacher name if mentioned (like "Mohit sir", "Akash sir", "Rajesh ma'am", etc.)

Respond ONLY with JSON:
{{"subject": "{'...' if not detected_subject else detected_subject}", "teacher_name": "...", "description": "...", "submission_day": "Monday/Tuesday/etc or date like '20th February' or 'tomorrow'", "submission_time": "HH:MM or null"}}"""

    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": parse_prompt}],
            temperature=0.2, max_tokens=300
        )
        raw = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        parsed = json.loads(json_match.group() if json_match else raw)
        if detected_subject:
            parsed['subject'] = detected_subject

        # ✅ Use smart date parser
        submission_date = None

        # ✅ FUTURE date parser for homework — never go to past
        submission_day_str = parsed.get('submission_day', '').lower().strip()
        today = datetime.now().date()
        days_map_hw = {'monday': 0, 'tuesday': 1, 'wednesday': 2, 'thursday': 3,
                       'friday': 4, 'saturday': 5, 'sunday': 6}
        month_map_hw = {
            'jan': 1, 'january': 1, 'feb': 2, 'february': 2,
            'mar': 3, 'march': 3, 'apr': 4, 'april': 4,
            'may': 5, 'jun': 6, 'june': 6, 'jul': 7, 'july': 7,
            'aug': 8, 'august': 8, 'sep': 9, 'sept': 9, 'september': 9,
            'oct': 10, 'october': 10, 'nov': 11, 'november': 11,
            'dec': 12, 'december': 12
        }
        if submission_day_str in ['today', 'aaj']:
            submission_date = today
        elif submission_day_str in ['tomorrow', 'kal']:
            submission_date = today + timedelta(days=1)
        elif submission_day_str in days_map_hw:
            days_ahead = (days_map_hw[submission_day_str] - today.weekday()) % 7
            if days_ahead == 0:
                days_ahead = 7  # Same weekday this week → next occurrence
            submission_date = today + timedelta(days=days_ahead)
        else:
            # Specific date like "25th feb", "20 march"
            date_match = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s*([a-z]+)', submission_day_str)
            if date_match:
                try:
                    d = int(date_match.group(1))
                    m = month_map_hw.get(date_match.group(2)[:3])
                    if m:
                        candidate = datetime(today.year, m, d).date()
                        if candidate < today:
                            candidate = datetime(today.year + 1, m, d).date()
                        submission_date = candidate
                except:
                    pass
        if submission_date is None:
            submission_date = today + timedelta(days=7)  # Default: next week

        result = save_homework(
            cr_student_id=student_details['student_id'],
            subject_identifier=parsed['subject'],
            homework_description=parsed['description'],
            submission_date=submission_date,
            submission_time=parsed.get('submission_time'),
            student_group='BOTH'
        )

        if result['success']:
            save_homework_to_db(
                course=result['course'],
                department=result['department'],
                semester=result['semester'],
                student_group='BOTH',
                subject_name=result['subject'],
                teacher_name=parsed.get('teacher_name', ''),
                description=parsed['description'],
                submission_date=submission_date,
                submission_time=parsed.get('submission_time'),
                posted_by=student_details['student_id']
            )

            date_str = submission_date.strftime('%A, %B %d')
            time_str = f" at {parsed.get('submission_time')}" if parsed.get('submission_time') else ""
            teacher_str = f"\n👨‍🏫 Teacher: {parsed.get('teacher_name')}" if parsed.get('teacher_name') else ""

            await update.message.reply_text(
                f"✅ *Homework Posted!*\n\n"
                f"📚 Subject: {result['subject']}{teacher_str}\n"
                f"📋 Task: {parsed['description']}\n"
                f"📅 Due: {date_str}{time_str}\n\n"
                f"Students will get daily reminders at 8 PM! 🔔",
                parse_mode='Markdown'
            )

            conn = sqlite3.connect('students.db')
            c = conn.cursor()
            c.execute('''SELECT ru.telegram_id FROM registered_users ru
                         JOIN master_students ms ON ru.student_id = ms.student_id
                         WHERE ms.course = ? AND ms.department = ? AND ms.semester = ? AND ru.is_verified = 1''',
                      (result['course'], result['department'], result['semester']))
            students = [row[0] for row in c.fetchall()]
            conn.close()

            teacher_line = f"\n👨‍🏫 Teacher: {parsed.get('teacher_name')}" if parsed.get('teacher_name') else ""
            broadcast_msg = f"""📝 *NEW HOMEWORK*

📚 Subject: {result['subject']}{teacher_line}
📋 Task: {parsed['description']}
📅 Submit by: {date_str}{time_str}

You'll get daily reminders at 8 PM! 🔔"""

            count = 0
            for tid in students:
                try:
                    await context.bot.send_message(chat_id=tid, text=broadcast_msg, parse_mode='Markdown')
                    count += 1
                except:
                    pass
            print(f"✅ Homework broadcast to {count} students")
        else:
            await update.message.reply_text(f"❌ {result['message']}")

    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

# ============================================================
# ✅ /lasttheory COMMAND - auto date + teacher detection
# ============================================================
async def handle_last_theory(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    student_details = get_student_details(reg_status['student_id'])
    cr_info = is_cr(student_details['student_id'])

    if not cr_info['is_cr']:
        await update.message.reply_text("❌ Only CR can update class history.")
        return

    theory_text = update.message.text.replace('/lasttheory', '').strip()

    if not theory_text:
        await update.message.reply_text(
            "📚 *Update Theory Class:*\n\n"
            "Bas likhdo kya hua, date bhi bol sakte ho:\n\n"
            "• `/lasttheory Akash sir - Types of cuts (julienne, brunoise)`\n"
            "• `/lasttheory 17th feb Akash sir - Stocks and sauces`\n"
            "• `/lasttheory Kal Mohit sir - F&B service types`\n"
            "• `/lasttheory Food Production - Knife skills`\n\n"
            "🎤 *Voice notes also work!*\n\n"
            "💡 Date mention karo toh woh date pe save hoga,\n"
            "warna aaj ki date lega.",
            parse_mode='Markdown'
        )
        return

    await _process_theory_save(update, context, student_details, theory_text)

# ============================================================
# ✅ /lastpractical COMMAND - auto date + subject detection
# ============================================================
async def handle_last_practical(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    if not is_admin(telegram_id):
        await update.message.reply_text("❌ This command is for Admin only.")
        return
    reg_status = get_registration_status(telegram_id)
    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return
    student_details = get_student_details(reg_status['student_id'])

    # ✅ Text from caption if photo sent, else from message
    if update.message.photo:
        practical_text = (update.message.caption or '').replace('/lastpractical', '').strip()
    else:
        practical_text = update.message.text.replace('/lastpractical', '').strip()

    if not practical_text:
        await update.message.reply_text(
            "🔬 *Update Practical Class (Admin Only):*\n\n"
            "📸 *Photo ke saath bhi bhej sakte ho!*\n\n"
            "• `/lastpractical Group A shepherd pie` + photo attach karo\n"
            "• `/lastpractical 13th feb Akash sir ne Bruschetta banawaya` + photo\n"
            "• `/lastpractical Group B Mohit sir - Table setting aaj`\n"
            "💡 Food Production mein photo attach karo — students query pe photo milega!",
            parse_mode='Markdown')
        return

    # ✅ Save photo if attached
    photo_path = None
    if update.message.photo:
        try:
            photo = update.message.photo[-1]
            file = await context.bot.get_file(photo.file_id)
            os.makedirs('practical_photos', exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            photo_filename = f"practical_photos/practical_{timestamp}.jpg"
            await file.download_to_drive(photo_filename)
            photo_path = photo_filename
            print(f"✅ Practical photo saved: {photo_filename}")
        except Exception as e:
            print(f"Photo save error: {e}")

    group_hint = student_details['group']
    group_match = re.search(r'group\s*([ab])', practical_text.lower())
    if group_match:
        group_hint = group_match.group(1).upper()
    await _process_practical_save(update, context, student_details, practical_text, group_hint, photo_path=photo_path)
# ============================================================
# /complete COMMAND
# ============================================================
async def handle_complete_topic(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    student_details = get_student_details(reg_status['student_id'])
    cr_info = is_cr(student_details['student_id'])

    if not cr_info['is_cr']:
        await update.message.reply_text("❌ Only CR can mark topics complete.")
        return

    complete_text = update.message.text.replace('/complete', '').strip()

    if not complete_text:
        await update.message.reply_text(
            "✅ *Mark Unit Complete:*\n\n"
            "Format: `/complete [subject] Unit [number]`\n\n"
            "Examples:\n"
            "• `/complete Food Production Unit 1`\n"
            "• `/complete F&B Unit 2`\n"
            "• `/complete English Unit 3`\n"
            "• `/complete Akash sir Unit 1` ← Teacher name bhi chalega!\n\n"
            "🗣️ *Natural language bhi kaam karta hai:*\n"
            "  'Food Production Unit 1 complete ho gaya'\n"
            "  'F&B Unit 2 khatam'",
            parse_mode='Markdown'
        )
        return

    detected_subject = detect_subject_name(complete_text)
    if not detected_subject:
        teacher_subject = get_subject_from_teacher(complete_text)
        if teacher_subject:
            detected_subject = teacher_subject

    parse_prompt = f"""Extract from: "{complete_text}"

{'Detected: ' + detected_subject if detected_subject else ''}

Respond ONLY with JSON:
{{"subject_code": "{'...' if not detected_subject else detected_subject}", "unit_number": "1/2/3/4/5/6"}}"""

    try:
        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": parse_prompt}],
            temperature=0.2, max_tokens=100
        )
        raw = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        parsed = json.loads(json_match.group() if json_match else raw)
        subject_id = detected_subject if detected_subject else parsed['subject_code']

        result = mark_topic_complete(
            student_details['semester'], subject_id,
            int(parsed['unit_number']), student_details['student_id']
        )

        if result['success']:
            await update.message.reply_text(
                f"✅ *Unit Marked Complete!*\n\n"
                f"📚 {result['subject_name']}\n"
                f"📖 Unit {result['unit_number']}: {result['unit_name']}\n\n"
                f"Students can now see this as completed! 🎉",
                parse_mode='Markdown'
            )
        else:
            await update.message.reply_text(f"❌ {result['message']}")

    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")

# ============================================================
# ATTENDANCE COMMAND
# ============================================================
async def handle_attendance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = str(update.message.from_user.id)
    reg_status = get_registration_status(telegram_id)
    if not reg_status or reg_status.get('is_verified') != 1:
        await update.message.reply_text("Please register first using /start!")
        return

    student_id = reg_status['student_id']
    args = context.args
    arg = args[0].lower() if args else ''

    if arg in ['overall', 'semester', 'exam', 'eligibility', 'total']:
        # Overall combined attendance
        overall = get_overall_attendance(student_id)
        if not overall:
            await update.message.reply_text("No attendance data found.")
            return
        text = format_overall_attendance_for_ai(overall)
        await _safe_send(update, text, parse_mode='Markdown')

    elif arg in ['january', 'jan']:
        att = get_student_attendance(student_id, month='January', year=2026)
        low = check_low_attendance(student_id) if att else []
        text = format_attendance_for_ai(att, low) + "\n_Showing: January 2026_"
        await _safe_send(update, text, parse_mode='Markdown')

    elif arg in ['february', 'feb']:
        att = get_student_attendance(student_id, month='February', year=2026)
        low = check_low_attendance(student_id) if att else []
        text = format_attendance_for_ai(att, low) + "\n_Showing: February 2026_"
        await _safe_send(update, text, parse_mode='Markdown')

    else:
        # Default — latest month (February) + overall summary
        att = get_student_attendance(student_id)
        low = check_low_attendance(student_id) if att else []
        text = format_attendance_for_ai(att, low)

        # Add overall at end
        overall = get_overall_attendance(student_id)
        if overall:
            total_held = sum(s['total_held'] for s in overall)
            total_att  = sum(s['total_attended'] for s in overall)
            overall_pct = round(total_att / total_held * 100, 1) if total_held > 0 else 0
            all_elig = all(s['eligible'] for s in overall)
            text += f"\n*Overall (Jan + Feb):* {total_att}/{total_held} = *{overall_pct}%*"
            text += f"\nExam eligibility: {'✅ Eligible' if all_elig else '❌ Not fully eligible'}"
            text += "\n\n_For details: /attendance overall_"
        await _safe_send(update, text, parse_mode='Markdown')


async def handle_leaderboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = str(update.message.from_user.id)
    reg_status = get_registration_status(telegram_id)
    if not reg_status or reg_status.get('is_verified') != 1:
        await update.message.reply_text("Please register first using /start!")
        return

    student_id = reg_status['student_id']
    args = context.args
    arg = args[0].lower() if args else ''

    if arg in ['overall', 'semester', 'total']:
        lb = get_overall_leaderboard(limit=1000)
        title = "Overall Leaderboard (Jan + Feb)"
    elif arg in ['january', 'jan']:
        lb = get_monthly_leaderboard(month='January', year=2026, limit=1000)
        title = "January 2026 Leaderboard"
    elif arg in ['february', 'feb']:
        lb = get_monthly_leaderboard(month='February', year=2026, limit=1000)
        title = "February 2026 Leaderboard"
    else:
        # Default: overall
        lb = get_overall_leaderboard(limit=1000)
        title = "Overall Leaderboard (Jan + Feb)"

    if not lb or not lb['data']:
        await update.message.reply_text("No leaderboard data available yet.")
        return

    # Find student rank
    student_rank = None
    for entry in lb['data']:
        if entry['student_id'] == student_id:
            student_rank = entry['rank']
            break

    rank_emojis = {1: "🥇", 2: "🥈", 3: "🥉"}
    lines = [f"*{title}*\n"]
    for entry in lb['data']:
        em = rank_emojis.get(entry['rank'], f"{entry['rank']}.")
        lines.append(f"{em} {entry['name']}: *{entry['pct']}%* ({entry['attended']}/{entry['held']})")

    if student_rank:
        lines.append(f"\n_Your rank: #{student_rank}_")
    lines.append("\n`/leaderboard january` · `/leaderboard february` · `/leaderboard overall`")

    await _safe_send(update, "\n".join(lines), parse_mode='Markdown')

    # Generate personal attendance PDF with graph
    try:
        processing = await update.message.reply_text("📊 Generating your attendance report...")
        # Get student name
        conn_pdf = sqlite3.connect('students.db', timeout=20)
        c_pdf = conn_pdf.cursor()
        c_pdf.execute("SELECT name FROM master_students WHERE student_id = ?", (student_id,))
        row_pdf = c_pdf.fetchone()
        conn_pdf.close()
        student_name = row_pdf[0] if row_pdf else student_id

        # Get both months data
        jan_raw = get_student_attendance(student_id, month='January', year=2026)
        feb_raw = get_student_attendance(student_id, month='February', year=2026)
        jan_subjects = jan_raw['subjects'] if jan_raw else []
        feb_subjects = feb_raw['subjects'] if feb_raw else []

        # Full leaderboard for PDF
        full_lb = get_overall_leaderboard(limit=1000)

        pdf_path = generate_attendance_pdf(
            student_id, student_name,
            jan_subjects, feb_subjects, full_lb
        )
        await processing.delete()
        with open(pdf_path, 'rb') as f_pdf:
            await update.message.reply_document(
                document=f_pdf,
                filename=f"Zei_Attendance_{student_name.replace(' ', '_')}.pdf",
                caption="📈 Your attendance report with class leaderboard"
            )
        try: os.remove(pdf_path)
        except: pass
    except Exception as e:
        print(f"[Attendance PDF error] {e}")
        try: await processing.delete()
        except: pass


async def handle_syllabus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    student_details = get_student_details(reg_status['student_id'])
    syllabus_data = get_subject_syllabus(student_details['semester'])
    progress_data = get_syllabus_progress(student_details['semester'])

    if not syllabus_data:
        await update.message.reply_text(
            "📚 *Syllabus*\n\nNo syllabus data available yet.\nPlease contact admin.",
            parse_mode='Markdown'
        )
        return

    response_text = f"📚 *Semester {student_details['semester']} Syllabus:*\n\n"
    for subject in syllabus_data:
        subject_code = subject['subject_code']
        units = subject['units']
        completed = sum(1 for u in units if progress_data.get(f"{subject_code}-{u['unit_number']}", False))
        total = len(units)
        response_text += f"*{subject['subject_name']}*\nProgress: {completed}/{total}\n\n"
        for unit in units:
            status = "✅" if progress_data.get(f"{subject_code}-{u['unit_number']}", False) else "⬜"
            response_text += f"{status} Unit {unit['unit_number']}: {unit['unit_name']}\n"
        response_text += "\n"

    await _safe_send(update, response_text, parse_mode='Markdown')

# ============================================================
# PROGRESS COMMAND
# ============================================================
async def handle_progress(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status['is_verified'] == 0:
        await update.message.reply_text("❌ Please register first!")
        return

    student_details = get_student_details(reg_status['student_id'])
    syllabus_data = get_subject_syllabus(student_details['semester'])
    progress_data = get_syllabus_progress(student_details['semester'])

    response_text = "📊 *Progress:*\n\n"
    total_completed = 0
    total_units = 0

    for subject in syllabus_data:
        subject_code = subject['subject_code']
        units = subject['units']
        completed = sum(1 for u in units if progress_data.get(f"{subject_code}-{u['unit_number']}", False))
        total = len(units)
        total_completed += completed
        total_units += total
        percentage = int((completed / total) * 100) if total > 0 else 0
        bar = "█" * (percentage // 10) + "░" * (10 - percentage // 10)
        response_text += f"*{subject['subject_name']}*\n{bar} {percentage}%\n\n"

    overall = int((total_completed / total_units) * 100) if total_units > 0 else 0
    response_text += f"*Overall:* {overall}%"
    await _safe_send(update, response_text, parse_mode='Markdown')

# ============================================================
# LOGOUT COMMAND
# ============================================================
async def handle_logout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    reg_status = get_registration_status(telegram_id)

    if not reg_status:
        await update.message.reply_text("You're not logged in!")
        return

    student_details = get_student_details(reg_status['student_id'])
    name = student_details['name'] if student_details else "User"
    reset_user_registration(telegram_id)

    await update.message.reply_text(
        f"👋 *Goodbye {name}!*\n\n"
        f"You've been logged out successfully.\n\n"
        f"To login again, send me your Student ID.\n\n"
        f"See you soon! 🎓",
        parse_mode='Markdown'
    )

# ============================================================
# ADMIN COMMANDS
# ============================================================
def get_retention_stats():
    """Calculate all retention metrics from message_history"""
    import sqlite3 as _sq
    from datetime import datetime as _dt, timedelta as _td
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    stats = {}
    try:
        today = _dt.now().date()
        week_ago = today - _td(days=7)
        month_ago = today - _td(days=30)
        prev_week_start = today - _td(days=14)

        # DAU — unique users today
        c.execute("""SELECT COUNT(DISTINCT telegram_id) FROM message_history
                     WHERE DATE(timestamp) = ?""", (today,))
        stats['dau'] = c.fetchone()[0]

        # MAU — unique users last 30 days
        c.execute("""SELECT COUNT(DISTINCT telegram_id) FROM message_history
                     WHERE DATE(timestamp) >= ?""", (month_ago,))
        stats['mau'] = c.fetchone()[0]

        # Session frequency — avg days active per user (last 30 days)
        c.execute("""SELECT telegram_id, COUNT(DISTINCT DATE(timestamp)) as active_days
                     FROM message_history WHERE DATE(timestamp) >= ?
                     GROUP BY telegram_id""", (month_ago,))
        rows = c.fetchall()
        if rows:
            stats['avg_active_days'] = round(sum(r[1] for r in rows) / len(rows), 1)
            # Distribution
            stats['casual'] = sum(1 for r in rows if r[1] <= 3)
            stats['regular'] = sum(1 for r in rows if 4 <= r[1] <= 10)
            stats['power'] = sum(1 for r in rows if r[1] > 10)
        else:
            stats['avg_active_days'] = 0
            stats['casual'] = stats['regular'] = stats['power'] = 0

        # Power users — 10+ messages total
        c.execute("""SELECT telegram_id, COUNT(*) as msg_count
                     FROM message_history WHERE role = 'user'
                     GROUP BY telegram_id HAVING msg_count >= 10
                     ORDER BY msg_count DESC LIMIT 5""")
        stats['top_users'] = c.fetchall()

        # Inactive users — registered but silent 7+ days
        c.execute("""SELECT COUNT(*) FROM master_students ms
                     WHERE ms.is_verified = 1
                     AND ms.telegram_id NOT IN (
                         SELECT DISTINCT telegram_id FROM message_history
                         WHERE DATE(timestamp) >= ?
                     )""", (week_ago,))
        stats['inactive_7d'] = c.fetchone()[0]

        # Week-over-week retention
        c.execute("""SELECT COUNT(DISTINCT telegram_id) FROM message_history
                     WHERE DATE(timestamp) >= ? AND DATE(timestamp) < ?""",
                  (prev_week_start, week_ago))
        users_prev_week = set(r[0] for r in c.fetchall()) if False else None
        # Prev week users
        c.execute("""SELECT DISTINCT telegram_id FROM message_history
                     WHERE DATE(timestamp) >= ? AND DATE(timestamp) < ?""",
                  (prev_week_start, week_ago))
        prev_users = set(r[0] for r in c.fetchall())
        # This week users
        c.execute("""SELECT DISTINCT telegram_id FROM message_history
                     WHERE DATE(timestamp) >= ?""", (week_ago,))
        curr_users = set(r[0] for r in c.fetchall())

        if prev_users:
            retained = len(prev_users & curr_users)
            stats['wow_retention'] = round(retained / len(prev_users) * 100, 1)
            stats['prev_week_users'] = len(prev_users)
            stats['retained_users'] = retained
        else:
            stats['wow_retention'] = 0
            stats['prev_week_users'] = 0
            stats['retained_users'] = 0

        # Total messages
        c.execute("SELECT COUNT(*) FROM message_history WHERE role = 'user'")
        stats['total_messages'] = c.fetchone()[0]

    except Exception as e:
        print(f"[Stats error] {e}")
    finally:
        conn.close()
    return stats


async def send_daily_retention_report(context: ContextTypes.DEFAULT_TYPE):
    """Daily job: send retention report to admin"""
    try:
        s = get_retention_stats()
        total_users = get_user_count()
        msg = (
            f"*Zei Daily Report*\n"
            f"_{datetime.now().strftime('%d %B %Y')}_\n\n"
            f"*Users*\n"
            f"DAU: *{s.get('dau', 0)}* | MAU: *{s.get('mau', 0)}* | Total: *{total_users}*\n\n"
            f"*Retention*\n"
            f"Week-over-week: *{s.get('wow_retention', 0)}%* "
            f"({s.get('retained_users', 0)}/{s.get('prev_week_users', 0)} users returned)\n"
            f"Inactive 7d+: *{s.get('inactive_7d', 0)}* users\n\n"
            f"*Engagement (last 30 days)*\n"
            f"Avg active days/user: *{s.get('avg_active_days', 0)}*\n"
            f"Casual (1-3 days): {s.get('casual', 0)} | "
            f"Regular (4-10): {s.get('regular', 0)} | "
            f"Power (10+): {s.get('power', 0)}\n\n"
            f"*Total messages sent:* {s.get('total_messages', 0)}"
        )
        await context.bot.send_message(chat_id=ADMIN_ID, text=msg, parse_mode='Markdown')
    except Exception as e:
        print(f"[Daily report error] {e}")


async def handle_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    if telegram_id != ADMIN_ID:
        await update.message.reply_text("❌ Admin access only!")
        return

    processing = await update.message.reply_text("Calculating...")
    s = get_retention_stats()
    total_users = get_user_count()

    # Top users with names
    top_user_lines = ""
    if s.get('top_users'):
        conn = sqlite3.connect('students.db', timeout=20)
        c = conn.cursor()
        for uid, count in s['top_users']:
            c.execute("SELECT name FROM master_students WHERE telegram_id = ?", (uid,))
            row = c.fetchone()
            name = row[0] if row else str(uid)
            top_user_lines += f"  {name}: {count} msgs\n"
        conn.close()

    msg = (
        f"*Zei Analytics*\n"
        f"_{datetime.now().strftime('%d %B %Y, %H:%M')}_\n\n"
        f"*Overview*\n"
        f"Total users: *{total_users}*\n"
        f"DAU: *{s.get('dau', 0)}* | MAU: *{s.get('mau', 0)}*\n\n"
        f"*Retention*\n"
        f"Week-over-week: *{s.get('wow_retention', 0)}%*\n"
        f"({s.get('retained_users', 0)} of {s.get('prev_week_users', 0)} prev-week users returned)\n"
        f"Inactive 7d+: *{s.get('inactive_7d', 0)}* users\n\n"
        f"*Engagement (30 days)*\n"
        f"Avg active days/user: *{s.get('avg_active_days', 0)}*\n"
        f"Casual (1-3d): {s.get('casual', 0)} | "
        f"Regular (4-10d): {s.get('regular', 0)} | "
        f"Power (10d+): {s.get('power', 0)}\n\n"
        f"*Top Power Users*\n"
        f"{top_user_lines if top_user_lines else 'Not enough data yet'}\n"
        f"*Total messages:* {s.get('total_messages', 0)}\n\n"
        f"`/export` - Download full user data"
    )

    await processing.delete()
    await update.message.reply_text(msg, parse_mode='Markdown')

async def handle_export_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    if telegram_id != ADMIN_ID:
        await update.message.reply_text("❌ Admin access only!")
        return

    processing_msg = await update.message.reply_text("📊 Generating Excel file...")
    filename = export_users_to_excel()

    if filename:
        await update.message.reply_document(document=open(filename, 'rb'))
        await processing_msg.delete()
        try:
            os.remove(filename)
        except:
            pass
    else:
        await processing_msg.edit_text("❌ Failed to generate Excel file!")

# ============================================================
# MAIN MESSAGE HANDLER
# ============================================================

# ============================================================
# ✅ PHOTO-AWARE REPLY — sends practical photo if query matches
# ============================================================
def get_practical_photo(course, department, semester, student_group, target_date=None, offset=0):
    """Get photo_path of Food Production practical.
    offset=0 → latest, offset=1 → second latest (last to last), etc.
    target_date → specific date's photo"""
    import sqlite3 as _sq
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        if target_date:
            c.execute("""SELECT photo_path, class_date FROM class_history_log
                         WHERE course = ? AND department = ? AND semester = ?
                         AND (student_group = ? OR student_group = 'BOTH')
                         AND class_type = 'practical'
                         AND subject_name LIKE '%Food Production%'
                         AND photo_path IS NOT NULL AND photo_path != ''
                         AND class_date = ?
                         ORDER BY created_at DESC LIMIT 1""",
                      (course, department, semester, student_group, str(target_date)))
            row = c.fetchone()
        else:
            c.execute("""SELECT photo_path, class_date FROM class_history_log
                         WHERE course = ? AND department = ? AND semester = ?
                         AND (student_group = ? OR student_group = 'BOTH')
                         AND class_type = 'practical'
                         AND subject_name LIKE '%Food Production%'
                         AND photo_path IS NOT NULL AND photo_path != ''
                         ORDER BY class_date DESC, created_at DESC""",
                      (course, department, semester, student_group))
            rows = c.fetchall()
            row = rows[offset] if rows and offset < len(rows) else None
        conn.close()
        return (row[0], row[1]) if row else (None, None)
    except Exception as e:
        conn.close()
        print(f"Photo fetch error: {e}")
        return (None, None)

# Keep old name for backward compat
def get_latest_practical_photo(course, department, semester, student_group):
    path, _ = get_practical_photo(course, department, semester, student_group)
    return path


async def _safe_send(update, text, parse_mode=None):
    """Send message, auto-splitting if > 4096 chars (Telegram limit)."""
    MAX = 4000  # safe margin below 4096
    if len(text) <= MAX:
        await update.message.reply_text(text, parse_mode=parse_mode)
        return
    # Split on newlines to avoid cutting mid-word
    chunks = []
    current = ""
    for line in text.split('\n'):
        if len(current) + len(line) + 1 > MAX:
            if current:
                chunks.append(current.strip())
            current = line
        else:
            current += ('\n' if current else '') + line
    if current:
        chunks.append(current.strip())
    for chunk in chunks:
        if chunk:
            await update.message.reply_text(chunk, parse_mode=parse_mode)


async def _send_reply_with_photo(update, context, ai_response, user_message, telegram_id):
    """Send AI reply + practical photo — date-aware"""
    import os as _os
    import re as _re
    from datetime import datetime as _dt
    
    msg_lower = (user_message or '').lower()
    is_practical_query = 'practical' in msg_lower

    photo_sent = False
    if is_practical_query:
        reg_status = get_registration_status(telegram_id)
        if reg_status and reg_status['is_verified'] == 1:
            student_details = get_student_details(reg_status['student_id'])

            # ✅ Detect specific date from query (e.g. "13th feb", "13 february")
            target_date = None
            date_match = _re.search(
                r'(\d{1,2})(?:st|nd|rd|th)?\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)',
                msg_lower
            )
            if date_match:
                month_map = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
                             'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
                try:
                    d = int(date_match.group(1))
                    m = month_map.get(date_match.group(2)[:3])
                    if m:
                        target_date = _dt(_dt.now().year, m, d).date()
                except:
                    pass

            # Count "last to last" occurrences → offset dynamically
            # "last practical" = 0, "last to last" = 1, "last to last to last" = 2, etc.
            if not target_date:
                # Count how many times "last" appears beyond the first
                last_count = msg_lower.count('last')
                # "last practical" = 1 last → offset 0 (latest)
                # "last to last" = 2 lasts → offset 1
                # "last to last to last" = 3 lasts → offset 2
                photo_offset = max(0, last_count - 1)
                # Also handle Hindi: "pichla pichla pichla" 
                pichla_count = msg_lower.count('pichla') + msg_lower.count('pichli')
                if pichla_count > 1:
                    photo_offset = max(photo_offset, pichla_count - 1)
            else:
                photo_offset = 0

            photo_path, photo_date = get_practical_photo(
                student_details['course'],
                student_details['department'],
                student_details['semester'],
                student_details['group'],
                target_date=target_date,
                offset=photo_offset
            )
            print(f"[PHOTO] date={target_date} offset={photo_offset} path={photo_path} photo_date={photo_date}")

            if photo_path:
                abs_path = _os.path.abspath(photo_path)
                if _os.path.exists(abs_path):
                    try:
                        await _safe_send(update, ai_response)
                        # Caption with date
                        try:
                            cap_date = _dt.strptime(str(photo_date), '%Y-%m-%d').strftime('%d %b %Y')
                        except:
                            cap_date = str(photo_date or '')
                        with open(abs_path, 'rb') as photo_file:
                            await update.message.reply_photo(
                                photo=photo_file,
                                caption=f"📸 Practical Dish — {cap_date}"
                            )
                        photo_sent = True
                        print("[PHOTO] ✅ Sent!")
                    except Exception as e:
                        print(f"[PHOTO] Error: {e}")
                else:
                    print(f"[PHOTO] File not found: {abs_path}")

    if not photo_sent:
        await _safe_send(update, ai_response)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    telegram_id = update.effective_user.id
    user_message = update.message.text

    reg_status = get_registration_status(telegram_id)

    if reg_status and reg_status['is_verified'] == 1:
        if len(user_message) >= 8 and len(user_message) <= 15 and user_message.isalnum():
            await update.message.reply_text(
                "⚠️ You're already logged in!\n\n"
                "To logout and login with a different ID, use:\n"
                "`/logout`",
                parse_mode='Markdown'
            )
            return

    # ✅ Image request? Fetch and send relevant image

    # ✅ Audio request — user text mein "audio bhej do" type kiya
    if is_audio_request_text(user_message):
        last_response = context.user_data.get('last_ai_response', '')
        if not last_response:
            await _safe_send(update, "Pehle kuch poochho, phir audio bhej deta hun!")
            return
        processing_msg = await update.message.reply_text("🎤 Generating audio... ⏳")
        try:
            import re as _re_tts
            tts_text = last_response
            tts_text = tts_text.replace('*', '').replace('_', '').replace('`', '').replace('#', '')
            tts_text = _re_tts.sub(r'[\U00010000-\U0010ffff]', '', tts_text)
            tts_text = _re_tts.sub(r'\n+', '. ', tts_text)
            tts_text = _re_tts.sub(r'\s+', ' ', tts_text).strip()
            if len(tts_text) > 490:
                trimmed = tts_text[:490]
                last_stop = max(trimmed.rfind('.'), trimmed.rfind('!'), trimmed.rfind('?'))
                tts_text = trimmed[:last_stop + 1] if last_stop > 200 else trimmed + '...'
            audio_bytes = await sarvam_tts(tts_text)
            await processing_msg.delete()
            if audio_bytes:
                import io as _io
                await update.message.reply_voice(
                    voice=_io.BytesIO(audio_bytes),
                    caption="🔊 Here you go!"
                )
            else:
                await _safe_send(update, "❌ Could not generate audio. Please try again.")
        except Exception as e:
            print(f"[Audio request error] {e}")
            try: await processing_msg.delete()
            except: pass
            await _safe_send(update, "❌ Could not generate audio.")
        return

    if is_pdf_request_text(user_message):
        last_response = context.user_data.get('last_ai_response')
        if last_response:
            try:
                processing_msg = await update.message.reply_text("📄 Generating PDF... ⏳")
                # Smart title via Groq
                last_user_msg = context.user_data.get('last_user_message', '')
                try:
                    title_resp = groq_client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=[{"role": "user", "content": f"Generate a short, clean document title (max 6 words, title case, no quotes) for this content. User asked: '{last_user_msg}'. Content preview: '{last_response[:200]}'. Respond with ONLY the title, nothing else."}],
                        temperature=0.3, max_tokens=20
                    )
                    title = title_resp.choices[0].message.content.strip().strip('"').strip("'")
                except:
                    title = last_user_msg[:50] if last_user_msg else 'Zei Document'
                # Fetch image for PDF
                img_path = None
                try:
                    img_query_resp = groq_client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=[{"role": "user", "content": f"Give me a short Google image search query (3-5 words, no quotes) to find a high-quality educational diagram or photo for this topic: '{last_user_msg}'. Only respond with the search query, nothing else."}],
                        temperature=0.3, max_tokens=15
                    )
                    img_query = img_query_resp.choices[0].message.content.strip().strip('"').strip("'")
                    img_path = fetch_image_for_pdf(img_query)
                except Exception as ie:
                    print(f"[Image query error] {ie}")

                pdf_path = generate_pdf_from_response(last_response, title=title, image_path=img_path)

                # Cleanup temp image
                if img_path:
                    try: os.remove(img_path)
                    except: pass

                await processing_msg.delete()
                import re as _re_fn2
                safe_fn = _re_fn2.sub(r'[\\/*?:"<>|]', '', title).strip()[:60]
                safe_fn = safe_fn if safe_fn else 'Zei Document'
                with open(pdf_path, 'rb') as pdf_file:
                    await update.message.reply_document(
                        document=pdf_file,
                        filename=f"{safe_fn}.pdf",
                        caption=f"📄 *{safe_fn}*\n\n_Zei by Zephy Intelligence_",
                        parse_mode='Markdown'
                    )
                try:
                    os.remove(pdf_path)
                except:
                    pass
                return
            except Exception as e:
                print(f"PDF generation error: {e}")
                await update.message.reply_text("❌ Could not generate PDF. Please try again.")
                return
        else:
            await update.message.reply_text("⚠️ First ask me something, then I can make a PDF!")
            return


    # ✅ Events query — AI powered with conversation context
    event_keywords = [
        'event', 'events', 'function', 'functions', 'fest', 'festival',
        'upcoming', 'programme', 'program', 'activity', 'activities',
        'celebration', 'party', 'competition', 'seminar', 'lecture',
        'sports day', 'ethnic day', 'fresher', 'farewell', 'convocation',
        'iftaar', 'dawat', 'cultural', 'annual',
        'koi event', 'kya event', 'events hai', 'event hai', 'koi function',
        'aane wala', 'aane wale', 'next event', 'coming event'
    ]
    # Also trigger if last response was about events (followup like "recent mein konsa?")
    is_event_query = any(w in user_message.lower() for w in event_keywords)

    if is_event_query:
        reg_status_ev = get_registration_status(telegram_id)
        if reg_status_ev and reg_status_ev.get('is_verified') == 1:
            try:
                all_events = get_upcoming_events()
                if not all_events:
                    await _safe_send(update, "No upcoming events right now! Check back later. 😊")
                    return

                # Build FULL events context — every field, exact data
                cat_emoji = {'academic': '📚', 'sports': '⚽', 'social': '🎉'}
                events_ctx = "UPCOMING EVENTS (use ONLY this data, never guess):\n\n"
                for i, e in enumerate(all_events, 1):
                    try:
                        d = datetime.strptime(str(e['date']), '%Y-%m-%d')
                        date_str = d.strftime('%d %B %Y (%A)')
                    except:
                        date_str = str(e['date'])
                    em = cat_emoji.get(e['category'], '📌')
                    events_ctx += (f"{i}. {em} Title: {e['title']}\n"
                                   f"   Date: {date_str}\n"
                                   f"   Time: {e['time']}\n"
                                   f"   Venue: {e['venue']}\n"
                                   f"   About: {e['description']}\n"
                                   f"   Category: {e['category']}\n\n")

                # Include conversation history for context
                chat_history = context.user_data.get('chat_history', [])
                recent_history = chat_history[-4:] if len(chat_history) > 4 else chat_history

                events_system = f"""You are Zei by Zephy Intelligence — a smart academic bot.

{events_ctx}

STRICT RULES:
1. Answer ONLY from the event data above — NEVER guess, NEVER hallucinate dates or details
2. If asked about a specific event → give its exact date, time, venue from the data
3. If asked "recent mein konsa?" or "next one?" → find the nearest upcoming date from today
4. If asked "any events?" → give a 1-line teaser of each event (name + date only)
5. Followup questions refer to the same events topic — use conversation history
6. LANGUAGE: Mirror user exactly — Hinglish reply to Hinglish, English to English
7. Keep replies SHORT and conversational
8. For single event details → give full info (date, time, venue, description)
9. NEVER say "mujhe pata nahi" if the event exists in the data above"""

                history_msgs = [{"role": m["role"], "content": m["content"]} for m in recent_history]

                ev_response = groq_client.chat.completions.create(
                    model="llama-3.3-70b-versatile",
                    messages=[
                        {"role": "system", "content": events_system},
                        *history_msgs,
                        {"role": "user", "content": user_message}
                    ],
                    temperature=0.2,
                    max_tokens=400
                )
                import re as _re2
                ev_reply = ev_response.choices[0].message.content
                ev_reply = _re2.sub(r'\*\*(.+?)\*\*', lambda m: '*' + m.group(1) + '*', ev_reply)
                context.user_data['last_ai_response'] = ev_reply
                context.user_data['last_user_message'] = user_message
                # Save to chat history for followup context
                if 'chat_history' not in context.user_data:
                    context.user_data['chat_history'] = []
                context.user_data['chat_history'].append({"role": "user", "content": user_message})
                context.user_data['chat_history'].append({"role": "assistant", "content": ev_reply})
                if len(context.user_data['chat_history']) > 20:
                    context.user_data['chat_history'] = context.user_data['chat_history'][-20:]

                # ✅ Smart poster — find best matching event and send poster
                ev_reply_lower = ev_reply.lower()

                def _score(title):
                    tl = title.lower()
                    if tl in ev_reply_lower: return 3
                    words = [w for w in tl.split() if len(w) > 3]
                    if words and all(w in ev_reply_lower for w in words): return 2
                    matched = sum(1 for w in words if w in ev_reply_lower)
                    return 1 if words and matched/len(words) >= 0.6 else 0

                scored = sorted([(e, _score(e['title'])) for e in all_events], key=lambda x: -x[1])
                scored = [(e, s) for e, s in scored if s > 0]

                # Send poster only if one event clearly matched
                if scored and (len(scored) == 1 or scored[0][1] > (scored[1][1] if len(scored) > 1 else 0)):
                    best = scored[0][0]
                    await _send_event_with_poster(update, best, ev_reply)
                else:
                    await _safe_send(update, ev_reply, parse_mode='Markdown')
                return
            except Exception as e:
                print(f"[Events AI error] {e}")
                # Fall through to main AI

    # ✅ Direct syllabus display — bypass AI for accuracy
    syllabus_keywords = ['syllabus', 'silabus', 'units', 'unit list', 'pura syllabus',
                         'complete syllabus', 'sabhi unit', 'all units', 'sare unit']
    is_direct_syllabus = any(w in user_message.lower() for w in syllabus_keywords)

    if is_direct_syllabus:
        reg_status = get_registration_status(telegram_id)
        if reg_status and reg_status.get('is_verified') == 1:
            student_details = get_student_details(reg_status['student_id'])
            if student_details:
                try:
                    syllabus_data = get_subject_syllabus(student_details['semester'])
                    progress_data = get_syllabus_progress(student_details['semester'])
                    # Detect subject filter from message
                    msg_lower = user_message.lower()
                    subject_filter = None
                    for kw in ['food production', 'f&b', 'food and beverage', 'housekeeping',
                               'front office', 'bhm-201', 'bhm-202', 'bhm-203', 'bhm-204']:
                        if kw in msg_lower:
                            subject_filter = kw
                            break
                    direct_response = format_syllabus_direct(syllabus_data, progress_data, subject_filter)
                    if direct_response:
                        context.user_data['last_ai_response'] = direct_response
                        context.user_data['last_user_message'] = user_message
                        await _safe_send(update, direct_response, parse_mode='Markdown')
                        return
                except Exception as e:
                    print(f"[Direct syllabus error] {e}")
                    # Fall through to AI

    # ✅ Web search indicator — show before AI call if search needed
    search_msg = None
    if needs_web_search(user_message):
        indicators = [
            "🌐 Searching the web...",
            "🔍 Fetching live data...",
            "📡 Connecting to internet...",
            "🌍 Looking it up online...",
        ]
        import hashlib as _hs
        idx = int(_hs.md5(user_message.encode()).hexdigest(), 16) % len(indicators)
        search_msg = await update.message.reply_text(indicators[idx])

    ai_response = get_ai_response(telegram_id, user_message)

    if search_msg:
        try:
            await search_msg.delete()
        except:
            pass

    # ✅ Faculty notes disclaimer — append if notes exist for this unit
    try:
        import re as _re_disc
        msg_d = user_message.lower()
        subj_map = {
            'f&b': 'BHM-202', 'food and beverage': 'BHM-202', 'food & beverage': 'BHM-202',
            'beverage service': 'BHM-202', 'fnb': 'BHM-202', 'bhm-202': 'BHM-202', 'bhm202': 'BHM-202',
            'food production': 'BHM-201', 'bhm-201': 'BHM-201', 'bhm201': 'BHM-201',
            'housekeeping': 'BHM-203', 'bhm-203': 'BHM-203',
            'front office': 'BHM-204', 'bhm-204': 'BHM-204',
            'personality': 'BHM-205', 'bhm-205': 'BHM-205',
        }
        subj_short = {
            'BHM-201': 'Food Production', 'BHM-202': 'F&B',
            'BHM-203': 'Housekeeping', 'BHM-204': 'Front Office',
            'BHM-205': 'Personality Development',
        }
        disc_subject = None
        for kw, sc in subj_map.items():
            if kw in msg_d:
                disc_subject = sc
                break
        unit_m = _re_disc.search(r'unit[\s]*([\d]+)', msg_d)
        if disc_subject and unit_m:
            disc_unit = int(unit_m.group(1))
            note = get_note_file(disc_subject, disc_unit)
            if note:
                short = subj_short.get(disc_subject, disc_subject.lower().replace('-',''))
                ai_response += "\n\n📎 *Faculty notes available:* `/classnotes " + short + " " + str(disc_unit) + "`"
    except Exception as _e:
        pass

    # ✅ Save last response for PDF generation
    context.user_data['last_ai_response'] = ai_response
    context.user_data['last_user_message'] = user_message

    try:
        if ai_response.strip().startswith('{') and '"action"' in ai_response:
            action_data = json.loads(ai_response)

            if action_data.get('action') == 'check_student_id':
                student_id = action_data['student_id']
                student_details = get_student_details(student_id)

                if student_details:
                    context.user_data['pending_student_details'] = student_details
                    cr_info = is_cr(student_details['student_id'])
                    cr_badge = "\n🎖️ Class Representative\n" if cr_info['is_cr'] else ""
                    urdu_type = student_details.get('urdu_type', 'regular')
                    urdu_display = f"\n📖 Urdu: {urdu_type.title()}\n" if urdu_type == 'advanced' else ""

                    keyboard = [[
                        InlineKeyboardButton("✅ Yes", callback_data='confirm_yes'),
                        InlineKeyboardButton("❌ No", callback_data='confirm_no')
                    ]]

                    await update.message.reply_text(
                        f"📋 *Found:*\n\n"
                        f"🆔 ID: `{student_details['student_id']}`\n"
                        f"👤 Name: *{student_details['name']}*\n"
                        f"📖 Course: {student_details['course']}\n"
                        f"🏢 Department: {student_details['department']}\n"
                        f"📚 Semester: {student_details['semester']}\n"
                        f"👥 Group: {student_details['group']}\n"
                        f"📧 Email: {student_details['email']}"
                        f"{urdu_display}"
                        f"{cr_badge}\n"
                        f"*Correct?*",
                        parse_mode='Markdown',
                        reply_markup=InlineKeyboardMarkup(keyboard)
                    )
                    return
                else:
                    await update.message.reply_text("❌ Student ID not found.")
                    return

            elif action_data.get('action') == 'verify_otp':
                otp = action_data['otp']
                if verify_otp(telegram_id, otp):
                    reg_status = get_registration_status(telegram_id)
                    student_details = get_student_details(reg_status['student_id'])
                    greeting = get_greeting()
                    first_name = student_details['name'].split()[0]
                    cr_info = is_cr(student_details['student_id'])
                    urdu_type = student_details.get('urdu_type', 'regular')
                    urdu_display = f" (Advanced Urdu)" if urdu_type == 'advanced' else ""

                    is_admin = str(telegram_id) == str(ADMIN_ID)

                    if is_admin:
                        success_message = (
                            f"{greeting}, *{first_name}*\n\n"
                            f"*Admin Commands*\n"
                            f"`/addnotes` - Upload faculty notes PDF\n"
                            f"`/addposter` - Upload event poster\n"
                            f"`/events` - View events\n"
                            f"`/listnotes` - View uploaded notes\n"
                            f"`/export` - Export student data\n"
                            f"`/stats` - Bot statistics\n\n"
                            f"_Welcome to Zei · Powered by Zephy Intelligence_"
                        )
                    elif cr_info['is_cr']:
                        success_message = (
                            f"{greeting}, *{first_name}* (CR){urdu_display}\n\n"
                            f"*CR Commands*\n"
                            f"`/update` - Cancel / shift / room change\n"
                            f"`/hw` - Post homework\n"
                            f"`/lasttheory` - Log theory class\n"
                            f"`/lastpractical` - Log practical class\n"
                            f"`/complete` - Mark unit complete\n\n"
                            f"*Or just type:*\n"
                            f"Kal F&B cancel hai\n"
                            f"Akash sir ne cuts padhaye\n\n"
                            f"`/events` - `/classnotes F&B 1` - `/listnotes`\n\n"
                            f"_Welcome to Zei · Powered by Zephy Intelligence_"
                        )
                    else:
                        success_message = (
                            f"{greeting}, *{first_name}*{urdu_display}\n\n"
                            f"*Ask me anything:*\n"
                            f"Monday ka schedule?\n"
                            f"F&B unit 1 explain karo\n"
                            f"Meri attendance?\n"
                            f"Any upcoming events?\n\n"
                            f"`/events` - `/classnotes F&B 1` - `/listnotes`\n\n"
                            f"_Welcome to Zei · Powered by Zephy Intelligence_"
                        )
                    
                    save_message(telegram_id, "assistant", success_message)
                    await update.message.reply_text(success_message, parse_mode='Markdown')
                    return
                else:
                    await update.message.reply_text("❌ Wrong OTP.")
                    return

        await process_ai_actions(update, context, ai_response, telegram_id, user_message=user_message)

    except json.JSONDecodeError:
        await _safe_send(update, ai_response)

# ============================================================
# CALLBACK HANDLER
# ============================================================
async def handle_button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    telegram_id = update.effective_user.id

    if query.data == 'confirm_yes':
        student_details = context.user_data.get('pending_student_details')
        if not student_details:
            await query.edit_message_text("❌ Session expired.")
            return

        otp = str(random.randint(100000, 999999))
        save_otp(telegram_id, student_details['student_id'], otp)
        email_sent = send_otp_email(student_details['email'], otp, student_details['name'])

        if email_sent:
            await query.edit_message_text(
                f"✅ *Confirmed!*\n\n📧 OTP sent to *{student_details['email']}*\n\nEnter the 6-digit OTP.",
                parse_mode='Markdown'
            )
        else:
            await query.edit_message_text("❌ Failed to send OTP.")

    elif query.data == 'confirm_no':
        await query.edit_message_text("❌ Details incorrect.\n\nContact admin.")

    elif query.data.startswith('confirm_update_'):
        pending_id = int(query.data.split('_')[2])

        conn = sqlite3.connect('students.db')
        c = conn.cursor()
        c.execute('SELECT action_data FROM pending_cr_actions WHERE id = ?', (pending_id,))
        result = c.fetchone()

        if not result:
            await query.edit_message_text("❌ Update expired.")
            return

        # ✅ Support both new (JSON with ai_msg) and old (plain text) format
        try:
            stored = json.loads(result[0])
            update_text = stored.get('original', result[0])
            ai_broadcast_msg = stored.get('ai_msg', None)
        except:
            update_text = result[0]
            ai_broadcast_msg = None

        # ✅ AI powered parsing - extract day, subject, type
        parse_prompt = f"""Extract update from: "{update_text}"

Detect day (today/tomorrow/kal/Monday/Tuesday/etc), subject, and type.

Respond ONLY with JSON:
{{"day": "today/tomorrow/kal/Monday/etc", "update_type": "cancelled/postponed/room_change", "subject_name": "...", "class_type": "theory/practical/both", "new_time": "...", "room_change": "...", "reason": "..."}}"""

        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": parse_prompt}],
            temperature=0.3, max_tokens=300
        )
        raw = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        parsed = json.loads(json_match.group() if json_match else raw)

        day_str = parsed.get('day', 'today').lower()
        today = datetime.now().date()
        days_map = {'monday': 0, 'tuesday': 1, 'wednesday': 2, 'thursday': 3,
                    'friday': 4, 'saturday': 5, 'sunday': 6}

        if day_str in ['today', 'aaj']:
            target_date = today
        elif day_str in ['tomorrow', 'kal']:
            target_date = today + timedelta(days=1)
        elif day_str in days_map:
            days_ahead = days_map[day_str] - today.weekday()
            if days_ahead <= 0:
                days_ahead += 7
            target_date = today + timedelta(days=days_ahead)
        else:
            target_date = today

        reg_status = get_registration_status(telegram_id)
        student_details = get_student_details(reg_status['student_id'])

        result = post_class_update(
            cr_student_id=student_details['student_id'],
            update_type=parsed.get('update_type'),
            subject_name=parsed.get('subject_name', ''),
            original_time=None,
            new_time=parsed.get('new_time'),
            room_change=parsed.get('room_change'),
            reason=parsed.get('reason'),
            class_type=parsed.get('class_type', 'both'),
            target_date=target_date
        )

        if result['success']:
            date_display = target_date.strftime('%A, %B %d')

            # ✅ Use AI generated message if available, else build smart one
            if ai_broadcast_msg:
                broadcast_msg = ai_broadcast_msg
            else:
                if parsed['update_type'] == 'cancelled':
                    broadcast_msg = f"📢 *CLASS UPDATE*\n\n❌ CANCELLED\n📅 {date_display}\n🔔 {parsed.get('class_type', 'both').upper()}: {parsed.get('subject_name', 'ALL CLASSES')}\n\n- Posted by CR 🎖️"
                elif parsed['update_type'] == 'postponed':
                    broadcast_msg = f"📢 *CLASS UPDATE*\n\n⏰ TIME CHANGE\n📅 {date_display}\n🕐 New time: {parsed.get('new_time', 'TBD')}\n\n- Posted by CR 🎖️"
                else:
                    broadcast_msg = f"📢 *CLASS UPDATE*\n\n🔄 ROOM CHANGE\n📅 {date_display}\n📍 New room: {parsed.get('room_change', 'TBD')}\n\n- Posted by CR 🎖️"

            conn = sqlite3.connect('students.db')
            c = conn.cursor()
            c.execute('''SELECT ru.telegram_id FROM registered_users ru
                         JOIN master_students ms ON ru.student_id = ms.student_id
                         WHERE ms.course = ? AND ms.department = ? AND ms.semester = ? AND ru.is_verified = 1''',
                      (result['course'], result['department'], result['semester']))
            students = [row[0] for row in c.fetchall()]
            c.execute('DELETE FROM pending_cr_actions WHERE id = ?', (pending_id,))
            conn.commit()
            conn.close()

            count = 0
            for tid in students:
                try:
                    await context.bot.send_message(chat_id=tid, text=broadcast_msg, parse_mode='Markdown')
                    count += 1
                except:
                    pass

            await query.edit_message_text(
                f"✅ *Update Broadcasted!*\n\n📅 {date_display}\n👥 Sent to {count} students",
                parse_mode='Markdown'
            )
        else:
            await query.edit_message_text(f"❌ {result['message']}")

    elif query.data.startswith('edit_update_'):
        pending_id = int(query.data.split('_')[2])
        await query.edit_message_text(
            "✏️ *Edit Mode*\n\nUse `/update [your new message]` to send a custom update.\n\nOr just broadcast the current one by pressing ✅.",
            parse_mode='Markdown'
        )

    elif query.data.startswith('cancel_update_'):
        pending_id = int(query.data.split('_')[2])
        conn = sqlite3.connect('students.db')
        c = conn.cursor()
        c.execute('DELETE FROM pending_cr_actions WHERE id = ?', (pending_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("❌ Update cancelled.")

# ============================================================
# MAIN
# ============================================================
async def handle_pdf_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle PDF uploads — extract text and let AI explain it."""
    # Skip if this is an /addnotes upload — handled separately
    caption = (update.message.caption or "").strip()
    if caption.lower().startswith('/addnotes'):
        return

    telegram_id = str(update.message.from_user.id)
    reg_status = get_registration_status(telegram_id)

    if not reg_status or reg_status.get('is_verified') != 1:
        await update.message.reply_text("Please register first using /start!")
        return

    doc = update.message.document
    caption = (update.message.caption or "").strip()

    processing_msg = await update.message.reply_text("📄 Reading PDF... ⏳")

    try:
        import pdfplumber, tempfile, os as _os

        # Download PDF
        file = await context.bot.get_file(doc.file_id)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
            tmp_path = tmp.name
        await file.download_to_drive(tmp_path)

        # Extract text
        extracted = []
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted.append(text.strip())
        _os.remove(tmp_path)

        full_text = "\n\n".join(extracted)

        if not full_text.strip():
            await processing_msg.delete()
            await update.message.reply_text("❌ No readable text found in PDF (might be a scanned image).")
            return

        # Trim to avoid token overflow — keep first ~6000 chars
        MAX_CHARS = 6000
        truncated = False
        if len(full_text) > MAX_CHARS:
            full_text = full_text[:MAX_CHARS]
            truncated = True

        # Build prompt
        user_prompt = caption if caption else "Please explain this document clearly and helpfully."

        system_prompt = f"""You are Zei by Zephy Intelligence — a smart academic assistant.
The student has uploaded a PDF document. Here is the extracted content:

--- PDF CONTENT START ---
{full_text}
--- PDF CONTENT END ---
{"(Note: PDF was long, showing first portion only)" if truncated else ""}

LANGUAGE RULE: Detect the language of the user's request and reply in the same language.
- English request → English reply
- Hinglish request → Hinglish reply  
- Hindi request → Hindi reply

TASK: Answer the student's request based on the PDF content above.
If they say "explain like I'm 10" → use very simple language, analogies, examples.
If they ask a specific question → answer from the PDF.
If no specific instruction → give a clear summary of what the PDF is about.
Keep it conversational and helpful."""

        await processing_msg.delete()
        processing_msg2 = await update.message.reply_text("🤔 Analyzing... ⏳")

        response = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.5,
            max_tokens=2000
        )
        ai_reply = response.choices[0].message.content
        import re as _re
        ai_reply = _re.sub(r'\*\*(.+?)\*\*', lambda m: '*' + m.group(1) + '*', ai_reply)

        # Save for PDF generation
        context.user_data['last_ai_response'] = ai_reply
        context.user_data['last_user_message'] = user_prompt

        await processing_msg2.delete()
        await _safe_send(update, ai_reply)

    except ImportError:
        await processing_msg.delete()
        await update.message.reply_text("❌ pdfplumber not installed. Run `pip install pdfplumber` on the server.")
    except Exception as e:
        print(f"[PDF Upload Error] {e}")
        try:
            await processing_msg.delete()
        except:
            pass
        await update.message.reply_text("❌ Could not process PDF. Please try again.")


# ============================================================
# EVENTS SYSTEM
# ============================================================

def setup_events_table(conn=None):
    """Create events table with correct schema, migrate if needed"""
    close_conn = False
    if conn is None:
        conn = sqlite3.connect('students.db', timeout=20)
        close_conn = True
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS events
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  event_id TEXT UNIQUE,
                  title TEXT,
                  date DATE,
                  time TEXT,
                  venue TEXT,
                  description TEXT,
                  category TEXT DEFAULT 'general',
                  poster_path TEXT,
                  poster_telegram_id TEXT,
                  is_active INTEGER DEFAULT 1,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    # Migrate old schema — add poster_telegram_id if missing
    existing_cols = [row[1] for row in c.execute("PRAGMA table_info(events)").fetchall()]
    if 'poster_telegram_id' not in existing_cols:
        try: c.execute("ALTER TABLE events ADD COLUMN poster_telegram_id TEXT")
        except: pass
    conn.commit()
    if close_conn:
        conn.close()


def load_events_from_csv():
    """Load events from events_data.csv into DB"""
    import csv as _csv
    try:
        conn = sqlite3.connect('students.db', timeout=20)
        setup_events_table(conn)
        c = conn.cursor()
        c.execute('DELETE FROM events')  # fresh reload
        with open('events_data.csv', 'r', encoding='utf-8') as f:
            reader = _csv.DictReader(f)
            count = 0
            for row in reader:
                c.execute('''INSERT OR REPLACE INTO events
                             (event_id, title, date, time, venue, description,
                              category, poster_path, poster_telegram_id)
                             VALUES (?,?,?,?,?,?,?,?,?)''',
                          (row['event_id'].strip(),
                           row['title'].strip(),
                           row['date'].strip(),
                           row.get('time','').strip(),
                           row.get('venue','').strip(),
                           row.get('description','').strip(),
                           row.get('category','general').strip().lower(),
                           row.get('poster_path','').strip(),
                           row.get('poster_telegram_id','').strip()))
                count += 1
        conn.commit()
        conn.close()
        print(f"✅ Events loaded from CSV ({count} records)")
    except FileNotFoundError:
        print("⚠️ events_data.csv not found — skipping")
    except Exception as e:
        print(f"❌ Error loading events: {e}")


def get_upcoming_events(category=None, days_ahead=30):
    """Get upcoming active events"""
    conn = sqlite3.connect('students.db', timeout=20)
    setup_events_table(conn)
    c = conn.cursor()
    today = datetime.now().date().isoformat()
    future = (datetime.now() + timedelta(days=days_ahead)).date().isoformat()
    if category:
        c.execute('''SELECT event_id, title, date, time, venue, description,
                              category, poster_path, poster_telegram_id
                     FROM events
                     WHERE date >= ? AND date <= ? AND is_active=1 AND category=?
                     ORDER BY date, time''', (today, future, category))
    else:
        c.execute('''SELECT event_id, title, date, time, venue, description,
                              category, poster_path, poster_telegram_id
                     FROM events
                     WHERE date >= ? AND date <= ? AND is_active=1
                     ORDER BY date, time''', (today, future))
    rows = c.fetchall()
    conn.close()
    return [{'event_id': r[0], 'title': r[1], 'date': r[2], 'time': r[3],
             'venue': r[4], 'description': r[5], 'category': r[6],
             'poster_path': r[7] or '', 'poster_telegram_id': r[8] or ''} for r in rows]


def get_event_by_id(event_id):
    """Get single event by event_id"""
    conn = sqlite3.connect('students.db', timeout=20)
    setup_events_table(conn)
    c = conn.cursor()
    c.execute('''SELECT event_id, title, date, time, venue, description,
                          category, poster_path, poster_telegram_id
                 FROM events WHERE event_id=? AND is_active=1''', (str(event_id).upper(),))
    row = c.fetchone()
    conn.close()
    if row:
        return {'event_id': row[0], 'title': row[1], 'date': row[2], 'time': row[3],
                'venue': row[4], 'description': row[5], 'category': row[6],
                'poster_path': row[7] or '', 'poster_telegram_id': row[8] or ''}
    return None


def get_events_for_reminder():
    """Get events happening tomorrow"""
    conn = sqlite3.connect('students.db', timeout=20)
    setup_events_table(conn)
    c = conn.cursor()
    tomorrow = (datetime.now() + timedelta(days=1)).date().isoformat()
    c.execute('''SELECT event_id, title, date, time, venue, description,
                          category, poster_path, poster_telegram_id
                 FROM events WHERE date=? AND is_active=1
                 ORDER BY time''', (tomorrow,))
    rows = c.fetchall()
    conn.close()
    return [{'event_id': r[0], 'title': r[1], 'date': r[2], 'time': r[3],
             'venue': r[4], 'description': r[5], 'category': r[6],
             'poster_path': r[7] or '', 'poster_telegram_id': r[8] or ''} for r in rows]


def save_poster_telegram_id(event_id, telegram_file_id):
    """Cache Telegram file_id for faster future sends"""
    conn = sqlite3.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('UPDATE events SET poster_telegram_id=? WHERE event_id=?',
                  (telegram_file_id, event_id))
        conn.commit()
    except Exception as e:
        print(f"Error saving telegram ID: {e}")
    finally:
        conn.close()


def format_event_message(event):
    """Format a single event for display"""
    try:
        d = datetime.strptime(event['date'], '%Y-%m-%d')
        date_str = d.strftime('%A, %d %B %Y')
    except:
        date_str = event['date']
    cat_emoji = {'academic': '📚', 'sports': '⚽', 'social': '🎉', 'cultural': '🎭', 'general': '📌'}
    em = cat_emoji.get(event.get('category','general'), '📌')
    msg = f"{em} *{event['title']}*\n📅 {date_str}"
    if event.get('time'): msg += f" at {event['time']}"
    if event.get('venue'): msg += f"\n📍 {event['venue']}"
    if event.get('description'): msg += f"\n\n{event['description']}"
    return msg


async def _send_event_with_poster(update_or_bot, event, caption, chat_id=None):
    """Helper — send event with poster if available, else text only"""
    poster_sent = False
    if event.get('poster_telegram_id'):
        try:
            if chat_id:
                await update_or_bot.send_photo(chat_id=chat_id, photo=event['poster_telegram_id'],
                                               caption=caption, parse_mode='Markdown')
            else:
                await update_or_bot.message.reply_photo(photo=event['poster_telegram_id'],
                                                        caption=caption, parse_mode='Markdown')
            poster_sent = True
        except Exception as e:
            print(f"[Poster send error] {e}")
    if not poster_sent:
        if chat_id:
            await update_or_bot.send_message(chat_id=chat_id, text=caption, parse_mode='Markdown')
        else:
            await _safe_send(update_or_bot, caption, parse_mode='Markdown')


async def handle_addposter(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin: /addposter E001 — attach photo with this caption to add poster"""
    telegram_id = str(update.message.from_user.id)
    if telegram_id != str(ADMIN_ID):
        await update.message.reply_text("❌ Admin only.")
        return

    if not update.message.photo:
        await update.message.reply_text(
            "*Usage:* Send a photo with caption `/addposter <event_id>`\n\n"
            "Example: `/addposter E001`\n\n"
            "To get event IDs, type: `any event`",
            parse_mode='Markdown'
        )
        return

    caption = (update.message.caption or "").strip()
    import re as _re
    match = _re.search(r'/addposter\s+(\S+)', caption, _re.IGNORECASE)
    if not match:
        await update.message.reply_text("❌ Caption must be: `/addposter E001`", parse_mode='Markdown')
        return

    event_id = match.group(1).upper()
    event = get_event_by_id(event_id)
    if not event:
        await update.message.reply_text(f"❌ Event `{event_id}` not found.", parse_mode='Markdown')
        return

    file_id = update.message.photo[-1].file_id
    save_poster_telegram_id(event_id, file_id)

    await update.message.reply_text(
        f"✅ *Poster saved for {event['title']}!*\n\n"
        f"Event ID: `{event_id}`\n"
        f"Students will now see this poster when they ask about events.",
        parse_mode='Markdown'
    )


async def handle_events(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show upcoming events — next 2-3 with posters, rest as text list"""
    telegram_id = str(update.message.from_user.id)
    reg_status = get_registration_status(telegram_id)
    if not reg_status or reg_status.get('is_verified') != 1:
        await update.message.reply_text("Please register first using /start!")
        return

    events = get_upcoming_events()
    if not events:
        await update.message.reply_text("No upcoming events right now!")
        return

    cat_emoji = {'academic': '📚', 'sports': '⚽', 'social': '🎉', 'general': '📌'}

    # Send next 2-3 events with posters
    poster_events = events[:3]
    remaining = events[3:]

    for event in poster_events:
        try:
            d = datetime.strptime(str(event['date']), '%Y-%m-%d')
            date_str = d.strftime('%d %B %Y (%A)')
        except:
            date_str = str(event['date'])
        em = cat_emoji.get(event['category'], '📌')
        caption = (
            f"{em} *{event['title']}*\n"
            f"📅 {date_str} at {event['time']}\n"
            f"📍 {event['venue']}\n\n"
            f"{event['description']}"
        )
        await _send_event_with_poster(update, event, caption)

    # Remaining events as compact text list
    if remaining:
        lines = ["\n*More upcoming events:*"]
        for e in remaining:
            try:
                d = datetime.strptime(str(e['date']), '%Y-%m-%d')
                date_str = d.strftime('%d %B (%A)')
            except:
                date_str = str(e['date'])
            em = cat_emoji.get(e['category'], '📌')
            lines.append(f"{em} *{e['title']}* — {date_str} at {e['time']}, {e['venue']}")
        await _safe_send(update, "\n".join(lines), parse_mode='Markdown')


async def send_event_reminders(context: ContextTypes.DEFAULT_TYPE):
    """Daily job: send event reminders for tomorrow's events"""
    events = get_events_for_reminder()
    if not events:
        return

    # Get all registered students
    conn = sqlite3.connect('students.db', timeout=20)
    cur = conn.cursor()
    cur.execute("SELECT telegram_id FROM master_students WHERE is_verified = 1")
    students = [row[0] for row in cur.fetchall()]
    conn.close()

    cat_emoji = {'academic': '📚', 'sports': '⚽', 'social': '🎉'}

    for event in events:
        try:
            d = datetime.strptime(str(event['date']), '%Y-%m-%d')
            date_str = d.strftime('%d %B %Y (%A)')
        except:
            date_str = str(event['date'])

        emoji = cat_emoji.get(event['category'], '📌')
        reminder_text = (
            f"🔔 *Event Tomorrow!*\n\n"
            f"{emoji} *{event['title']}*\n"
            f"📅 {date_str} at {event['time']}\n"
            f"📍 {event['venue']}\n\n"
            f"{event['description']}"
        )

        for student_id in students:
            try:
                if event.get('poster_telegram_id'):
                    await context.bot.send_photo(
                        chat_id=student_id,
                        photo=event['poster_telegram_id'],
                        caption=reminder_text, parse_mode='Markdown'
                    )
                elif event.get('poster_path'):
                    import os as _os
                    if _os.path.exists(event['poster_path']):
                        with open(event['poster_path'], 'rb') as img:
                            msg = await context.bot.send_photo(
                                chat_id=student_id, photo=img,
                                caption=reminder_text, parse_mode='Markdown'
                            )
                        save_poster_telegram_id(event['event_id'], msg.photo[-1].file_id)
                    else:
                        await context.bot.send_message(
                            chat_id=student_id, text=reminder_text, parse_mode='Markdown'
                        )
                else:
                    await context.bot.send_message(
                        chat_id=student_id, text=reminder_text, parse_mode='Markdown'
                    )
            except Exception as e:
                print(f"[Event reminder error] student {student_id}: {e}")


# ============================================================
# ✅ PDF NOTES SYSTEM — Admin uploads, students request
# ============================================================

def setup_notes_table_db():
    """Create notes_files table in DB"""
    import sqlite3 as _sq
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS notes_files
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  subject_code TEXT,
                  subject_name TEXT,
                  unit_number INTEGER,
                  unit_name TEXT,
                  file_id TEXT,
                  file_name TEXT,
                  uploaded_by TEXT,
                  uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  UNIQUE(subject_code, unit_number))'''
    )
    conn.commit()
    conn.close()

setup_notes_table_db()


def save_note_file(subject_code, subject_name, unit_number, unit_name, file_id, file_name, admin_id):
    import sqlite3 as _sq
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''INSERT OR REPLACE INTO notes_files
                     (subject_code, subject_name, unit_number, unit_name, file_id, file_name, uploaded_by)
                     VALUES (?, ?, ?, ?, ?, ?, ?)''',
                  (subject_code.upper(), subject_name, int(unit_number),
                   unit_name, file_id, file_name, admin_id))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        conn.close()
        print(f"[Notes save error] {e}")
        return False


def get_note_file(subject_code, unit_number):
    import sqlite3 as _sq
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT subject_code, subject_name, unit_number, unit_name, file_id, file_name
                     FROM notes_files WHERE subject_code = ? AND unit_number = ?''',
                  (subject_code.upper(), int(unit_number)))
        r = c.fetchone()
        conn.close()
        if r:
            return {'subject_code': r[0], 'subject_name': r[1], 'unit_number': r[2],
                    'unit_name': r[3], 'file_id': r[4], 'file_name': r[5]}
        return None
    except Exception as e:
        conn.close()
        return None


def list_notes():
    import sqlite3 as _sq
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    try:
        c.execute('''SELECT subject_code, subject_name, unit_number, unit_name, file_name
                     FROM notes_files ORDER BY subject_code, unit_number''')
        rows = c.fetchall()
        conn.close()
        return [{'subject_code': r[0], 'subject_name': r[1], 'unit_number': r[2],
                 'unit_name': r[3], 'file_name': r[4]} for r in rows]
    except Exception as e:
        conn.close()
        return []


def detect_notes_request(message: str):
    """Detect if user is requesting notes. Returns (subject_code, unit_number) or None."""
    import re as _re
    msg = message.lower().strip()

    # Patterns: "f&b unit 1", "bhm-202 unit 2", "food production unit 3", "unit 1 f&b"
    subject_map = {
        'f&b': 'BHM-202', 'food and beverage': 'BHM-202', 'food & beverage': 'BHM-202',
        'beverage service': 'BHM-202', 'fnb': 'BHM-202', 'bhm-202': 'BHM-202', 'bhm 202': 'BHM-202',
        'food production': 'BHM-201', 'bhm-201': 'BHM-201', 'bhm 201': 'BHM-201',
        'housekeeping': 'BHM-203', 'bhm-203': 'BHM-203', 'bhm 203': 'BHM-203',
        'front office': 'BHM-204', 'bhm-204': 'BHM-204', 'bhm 204': 'BHM-204',
        'personality': 'BHM-205', 'bhm-205': 'BHM-205',
    }

    # Must have "notes", "pdf", "bhej", "send", "de do", "chahiye" etc.
    note_triggers = ['notes', 'pdf', 'bhej', 'send', 'de do', 'chahiye', 'do bhai',
                     'bhejdo', 'share', 'upload', 'dedo', 'note']
    if not any(t in msg for t in note_triggers):
        return None

    # Detect unit number
    unit_match = _re.search(r'unit\\s*(\\d+)', msg)
    if not unit_match:
        return None
    unit_number = int(unit_match.group(1))

    # Detect subject
    subject_code = None
    for kw, code in subject_map.items():
        if kw in msg:
            subject_code = code
            break

    return (subject_code, unit_number) if subject_code else None


async def handle_addnotes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin: Upload PDF notes. Caption: /addnotes BHM-202 1  OR  /addnotes Unit 1 F&B"""
    telegram_id = str(update.message.from_user.id)
    if telegram_id != str(ADMIN_ID):
        await update.message.reply_text("❌ Admin only.")
        return

    doc = update.message.document
    if not doc or not doc.file_name.lower().endswith('.pdf'):
        await update.message.reply_text(
            "*Usage:* Send a PDF with caption:\n"
            "`/addnotes BHM-202 1` — F&B Unit 1\n"
            "`/addnotes BHM-201 3` — Food Production Unit 3\n\n"
            "*Subject codes:* BHM-201 BHM-202 BHM-203 BHM-204",
            parse_mode='Markdown'
        )
        return

    caption = (update.message.caption or "").strip()

    # ── Parse caption ──────────────────────────────────────
    import re as _re
    subject_names = {
        'BHM-201': 'Food Production Foundation - II',
        'BHM-202': 'Food & Beverage Service Foundation - II',
        'BHM-203': 'Housekeeping Skills - II',
        'BHM-204': 'Front Office Foundation - II',
        'BHM-205': 'Personality Development and Grooming',
    }
    subject_keywords = {
        'f&b': 'BHM-202', 'food and beverage': 'BHM-202', 'food & beverage': 'BHM-202',
        'fnb': 'BHM-202', 'beverage': 'BHM-202',
        'food production': 'BHM-201', 'food prod': 'BHM-201',
        'housekeeping': 'BHM-203', 'hk': 'BHM-203',
        'front office': 'BHM-204', 'fo': 'BHM-204',
        'personality': 'BHM-205', 'grooming': 'BHM-205',
    }

    cap_lower = caption.lower()

    # Detect subject code — either BHM-XXX directly or keyword
    subject_code = None
    bhm_match = _re.search(r'bhm[-\s]?(\d{3})', cap_lower)
    if bhm_match:
        subject_code = f"BHM-{bhm_match.group(1)}"
    else:
        for kw, code in subject_keywords.items():
            if kw in cap_lower:
                subject_code = code
                break

    # Detect unit number
    unit_match = _re.search(r'unit\s*(\d+)|(\d+)', cap_lower)
    unit_number = None
    if unit_match:
        unit_number = int(unit_match.group(1) or unit_match.group(2))

    if not subject_code or not unit_number:
        await update.message.reply_text(
            "❌ Could not detect subject or unit.\n\n"
            "Try: `/addnotes BHM-202 1` or `/addnotes F&B Unit 1`",
            parse_mode='Markdown'
        )
        return

    # ── Auto-fetch unit name from syllabus DB ──────────────
    unit_name = None
    try:
        import sqlite3 as _sq
        conn = _sq.connect('students.db', timeout=20)
        c = conn.cursor()
        c.execute(
            'SELECT unit_name FROM syllabus WHERE subject_code = ? AND unit_number = ? LIMIT 1',
            (subject_code, unit_number)
        )
        row = c.fetchone()
        conn.close()
        if row:
            unit_name = row[0]
    except Exception as e:
        print(f"[Syllabus lookup error] {e}")

    if not unit_name:
        unit_name = f"Unit {unit_number}"  # fallback

    subject_name = subject_names.get(subject_code, subject_code)

    # ── Save ───────────────────────────────────────────────
    success = save_note_file(subject_code, subject_name, unit_number, unit_name,
                             doc.file_id, doc.file_name, telegram_id)
    if success:
        await update.message.reply_text(
            f"✅ *Notes saved!*\n\n"
            f"📚 {subject_name}\n"
            f"Unit {unit_number}: {unit_name}\n\n"
            f"Students can request: _\"F&B unit {unit_number} notes bhej do\"_",
            parse_mode='Markdown'
        )
    else:
        await update.message.reply_text("❌ Could not save notes. Try again.")


async def handle_listnotes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/listnotes — Show all uploaded notes"""
    telegram_id = str(update.message.from_user.id)
    reg_status = get_registration_status(telegram_id)
    if not reg_status or reg_status.get('is_verified') != 1:
        await update.message.reply_text("Please register first using /start!")
        return

    notes = list_notes()
    if not notes:
        await update.message.reply_text("No notes uploaded yet.")
        return

    lines = ["*📚 Available Notes:*\n"]
    current_subject = None
    for n in notes:
        if n['subject_code'] != current_subject:
            current_subject = n['subject_code']
            lines.append(f"\n*{n['subject_name']} ({n['subject_code']})*")
        lines.append(f"  Unit {n['unit_number']}: {n['unit_name']}")

    lines.append("\n_Request: \"F&B unit 1 notes bhej do\"_")
    await _safe_send(update, "\n".join(lines), parse_mode='Markdown')


async def handle_classnotes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/classnotes F&B 1 — Send faculty uploaded PDF notes"""
    telegram_id = str(update.message.from_user.id)
    reg_status = get_registration_status(telegram_id)
    if not reg_status or reg_status.get('is_verified') != 1:
        await update.message.reply_text("Please register first using /start!")
        return

    args = context.args
    if not args:
        # Show available notes list
        notes = list_notes()
        if not notes:
            await update.message.reply_text(
                "No faculty notes uploaded yet.\n\nUsage: `/classnotes F&B 1`",
                parse_mode='Markdown'
            )
        else:
            lines = ["*📚 Available Faculty Notes:*\n"]
            cur = None
            for n in notes:
                if n['subject_code'] != cur:
                    cur = n['subject_code']
                    lines.append(f"\n*{n['subject_name']}*")
                lines.append(f"  `/classnotes {n['subject_name'].split()[0]} {n['unit_number']}` — Unit {n['unit_number']}: {n['unit_name']}")
            await _safe_send(update, "\n".join(lines), parse_mode='Markdown')
        return

    # Parse: last arg = unit number, everything before = subject
    import re as _re
    full_input = " ".join(args).strip()

    # Extract unit number — last number in input
    unit_match = _re.search(r'(\d+)\s*$', full_input)
    if not unit_match:
        await update.message.reply_text(
            "❌ Please include unit number.\n"
            "Example: `/classnotes F&B 1` or `/classnotes Food Production 2`",
            parse_mode='Markdown'
        )
        return

    unit_number = int(unit_match.group(1))
    subj_input = full_input[:unit_match.start()].strip().lower()

    # Subject mapping — natural names supported
    subj_map = {
        'f&b': 'BHM-202', 'food and beverage': 'BHM-202', 'food & beverage': 'BHM-202',
        'food and beverages': 'BHM-202', 'beverage service': 'BHM-202',
        'fnb': 'BHM-202', 'fb': 'BHM-202', 'bhm-202': 'BHM-202', 'bhm202': 'BHM-202',
        'food production': 'BHM-201', 'food prod': 'BHM-201', 'fp': 'BHM-201',
        'bhm-201': 'BHM-201', 'bhm201': 'BHM-201',
        'housekeeping': 'BHM-203', 'house keeping': 'BHM-203', 'hk': 'BHM-203',
        'bhm-203': 'BHM-203', 'bhm203': 'BHM-203',
        'front office': 'BHM-204', 'fo': 'BHM-204',
        'bhm-204': 'BHM-204', 'bhm204': 'BHM-204',
        'personality': 'BHM-205', 'personality development': 'BHM-205',
        'grooming': 'BHM-205', 'pd': 'BHM-205',
        'bhm-205': 'BHM-205', 'bhm205': 'BHM-205',
    }

    subject_code = subj_map.get(subj_input)

    # Fuzzy fallback — partial match
    if not subject_code:
        for kw, sc in subj_map.items():
            if kw in subj_input or subj_input in kw:
                subject_code = sc
                break

    if not subject_code:
        await update.message.reply_text(
            f"❌ Subject not recognized: *{subj_input}*\n\n"
            "Try: `/classnotes F&B 1` or `/classnotes Food Production 2`\n"
            "Or just `/classnotes` to see all available notes.",
            parse_mode='Markdown'
        )
        return

    note = get_note_file(subject_code, unit_number)
    if not note:
        await update.message.reply_text(
            f"❌ No faculty notes uploaded for *{subject_code} Unit {unit_number}* yet.\n"
            "Ask your CR or Admin to upload.",
            parse_mode='Markdown'
        )
        return

    await update.message.reply_document(
        document=note['file_id'],
        caption=(
            f"📚 *{note['subject_name']}*\n"
            f"Unit {note['unit_number']}: {note['unit_name']}\n\n"
            f"_Uploaded by faculty via Zei_"
        ),
        parse_mode='Markdown'
    )


async def send_morning_briefing(context: ContextTypes.DEFAULT_TYPE):
    """Daily 6 AM: Send personalized morning briefing + voice to all students"""
    import io as _io
    import sqlite3 as _sq

    today_name = datetime.now().strftime('%A')
    today_date = datetime.now().strftime('%d %B %Y')

    # Get all verified students
    conn = _sq.connect('students.db', timeout=20)
    c = conn.cursor()
    c.execute("""SELECT ru.telegram_id, ms.name, ms.student_id,
                        ms.course, ms.department, ms.semester, ms.student_group
                 FROM master_students ms
                 JOIN registered_users ru ON ms.student_id = ru.student_id
                 WHERE ru.is_verified = 1""")
    students = c.fetchall()
    conn.close()

    if not students:
        return

    # Fetch top 3 news once for all students
    news_text = ""
    news_voice_text = ""
    try:
        search_result = web_search("India top news today", max_results=3)
        if not search_result['failed'] and search_result['results']:
            # Ask Groq to summarize into 3 bullet points
            news_resp = groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{
                    "role": "user",
                    "content": (
                        f"Based on this news data, give me exactly 3 top current news headlines "
                        f"in 1 line each. Format: just the 3 lines, no numbering, no extra text.\n\n"
                        f"{search_result['results'][:1500]}"
                    )
                }],
                temperature=0.3, max_tokens=150
            )
            raw_news = news_resp.choices[0].message.content.strip()
            news_lines = [l.strip() for l in raw_news.split('\n') if l.strip()][:3]
            news_text = "\n".join(f"• {l}" for l in news_lines)
            news_voice_text = ". ".join(news_lines)
    except Exception as e:
        print(f"[Morning brief news error] {e}")
        news_text = ""
        news_voice_text = ""

    for row in students:
        tg_id, name, student_id, course, dept, semester, group = row
        first_name = name.split()[0] if name else "Student"

        try:
            # Get today's schedule
            schedule = get_today_schedule(course, dept, semester, group)

            # Build schedule text
            if schedule:
                sched_lines = []
                for cls in schedule:
                    sched_lines.append(
                        f"• {cls['start_time']}-{cls['end_time']}: "
                        f"{cls['subject']} ({cls['teacher']})"
                    )
                sched_text = "\n".join(sched_lines)
                sched_voice = (
                    f"You have {len(schedule)} classes today. "
                    + ", ".join(f"{c['subject']} at {c['start_time']}" for c in schedule[:3])
                )
            else:
                sched_text = "_No classes today — enjoy your day!_ 🌟"
                sched_voice = "You have no classes today. Enjoy your day!"

            # Build text message
            text_msg = (
                f"Good morning, *{first_name}*! ☀️\n"
                f"_{today_date}_\n\n"
                f"*Today's Schedule ({today_name})*\n"
                f"{sched_text}\n\n"
            )
            if news_text:
                text_msg += f"*Top News*\n{news_text}\n\n"
            text_msg += "_Powered by Zei · Zephy Intelligence_"

            # Send text
            await context.bot.send_message(
                chat_id=tg_id,
                text=text_msg,
                parse_mode='Markdown'
            )

            # Motivational line for no-class days
            motivational = ""
            if not schedule:
                motivational = " Make the most of your free day!"

            # Voice 1: Greeting + schedule (max 490 chars)
            voice1 = (
                f"Good morning {first_name}! "
                f"Today is {today_name}, {today_date}. "
                f"{sched_voice}.{motivational}"
            )[:490]

            audio1 = await sarvam_tts(voice1)
            if audio1:
                await context.bot.send_voice(
                    chat_id=tg_id,
                    voice=_io.BytesIO(audio1),
                    caption="🔊 Good morning!"
                )

            # Voice 2: News (only if available, max 490 chars)
            if news_voice_text:
                voice2 = f"Here are today's top 3 news. {news_voice_text}"[:490]
                audio2 = await sarvam_tts(voice2)
                if audio2:
                    await context.bot.send_voice(
                        chat_id=tg_id,
                        voice=_io.BytesIO(audio2),
                        caption="📰 Top news"
                    )

        except Exception as e:
            print(f"[Morning brief error] {tg_id}: {e}")


async def handle_test_morning(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/testmorning — Admin: test morning briefing right now"""
    telegram_id = str(update.message.from_user.id)
    if telegram_id != str(ADMIN_ID):
        await update.message.reply_text("❌ Admin only.")
        return
    await send_morning_briefing(context)


def generate_attendance_pdf(student_id, student_name, jan_data, feb_data, leaderboard_data):
    """Generate SpaceX-themed B&W attendance PDF with line chart + leaderboard"""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    from matplotlib.gridspec import GridSpec
    import numpy as np
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Image as RLImage, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import tempfile, os as _os

    # ── Fonts ──────────────────────────────────────────────
    FONT_DIR = '/usr/share/fonts/truetype/liberation/'
    try:
        pdfmetrics.registerFont(TTFont('LS',  FONT_DIR + 'LiberationSans-Regular.ttf'))
        pdfmetrics.registerFont(TTFont('LSB', FONT_DIR + 'LiberationSans-Bold.ttf'))
        pdfmetrics.registerFont(TTFont('LSI', FONT_DIR + 'LiberationSans-Italic.ttf'))
        R = 'LS'; B = 'LSB'; I = 'LSI'
    except:
        R = 'Helvetica'; B = 'Helvetica-Bold'; I = 'Helvetica-Oblique'

    # ── Colors (SpaceX B&W) ────────────────────────────────
    BLACK  = colors.HexColor('#0a0a0a')
    WHITE  = colors.white
    GRAY   = colors.HexColor('#888888')
    LGRAY  = colors.HexColor('#cccccc')
    XGRAY  = colors.HexColor('#f0f0f0')

    # ── Generate matplotlib chart ──────────────────────────
    subjects = [s['subject'].replace('Foundation - II', '').replace('Skills - II', '').replace('Development and Grooming', '').strip() for s in (jan_data or feb_data or [])]

    # Get percentages per subject per month
    jan_pcts = []
    feb_pcts = []

    all_subjects = list({s['subject'] for s in (jan_data or []) + (feb_data or [])})
    all_subjects.sort()
    short_subjects = [s.replace('Food Production Foundation - II', 'Food Prod')
                       .replace('Food & Beverage Service Foundation - II', 'F&B')
                       .replace('Front Office Foundation - II', 'Front Office')
                       .replace('Housekeeping Skills - II', 'Housekeeping')
                       .replace('Personality Development and Grooming', 'Personality') for s in all_subjects]

    for subj in all_subjects:
        jan_pct = next((s['percentage'] for s in (jan_data or []) if s['subject'] == subj), 0)
        feb_pct = next((s['percentage'] for s in (feb_data or []) if s['subject'] == subj), 0)
        jan_pcts.append(jan_pct)
        feb_pcts.append(feb_pct)

    # Create figure — SpaceX style
    fig = plt.figure(figsize=(10, 4), facecolor='#0a0a0a')
    ax = fig.add_subplot(111)
    ax.set_facecolor('#0a0a0a')

    x = np.arange(len(short_subjects))
    w = 0.35

    # Lines with markers
    if jan_pcts:
        ax.plot(x, jan_pcts, 'o-', color='#ffffff', linewidth=2, markersize=6,
                markerfacecolor='#ffffff', label='January', zorder=3)
    if feb_pcts:
        ax.plot(x, feb_pcts, 's--', color='#888888', linewidth=2, markersize=6,
                markerfacecolor='#888888', label='February', zorder=3)

    # 75% threshold line
    ax.axhline(y=75, color='#444444', linestyle=':', linewidth=1.5, alpha=0.8, label='75% threshold')

    # Styling
    ax.set_xticks(x)
    ax.set_xticklabels(short_subjects, color='#cccccc', fontsize=9, rotation=15, ha='right')
    ax.set_yticks([0, 25, 50, 75, 100])
    ax.set_yticklabels(['0%', '25%', '50%', '75%', '100%'], color='#cccccc', fontsize=9)
    ax.set_ylim(0, 110)
    ax.grid(axis='y', color='#222222', linewidth=0.8, zorder=0)
    ax.spines['bottom'].set_color('#333333')
    ax.spines['left'].set_color('#333333')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.legend(facecolor='#111111', edgecolor='#333333', labelcolor='#cccccc', fontsize=8)

    # Value labels on points
    for i, (j, f) in enumerate(zip(jan_pcts, feb_pcts)):
        if j: ax.annotate(f'{j:.0f}%', (i, j), textcoords="offset points",
                          xytext=(0, 10), ha='center', color='#ffffff', fontsize=7)
        if f: ax.annotate(f'{f:.0f}%', (i, f), textcoords="offset points",
                          xytext=(0, -16), ha='center', color='#aaaaaa', fontsize=7)

    plt.tight_layout(pad=0.5)

    # Save chart to temp file
    chart_tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    plt.savefig(chart_tmp.name, dpi=150, bbox_inches='tight',
                facecolor='#0a0a0a', edgecolor='none')
    plt.close()
    chart_path = chart_tmp.name

    # ── Build PDF ──────────────────────────────────────────
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_path = f'zei_attendance_{timestamp}.pdf'

    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm)

    styles = getSampleStyleSheet()
    def S(name, **kw):
        base = dict(fontName=R, fontSize=10, textColor=BLACK, leading=15)
        base.update(kw)
        return ParagraphStyle(name, parent=styles['Normal'], **base)

    title_s   = S('t', fontName=B, fontSize=18, textColor=BLACK, alignment=TA_LEFT, spaceAfter=2)
    sub_s     = S('s', fontSize=9, textColor=GRAY, alignment=TA_LEFT)
    h2_s      = S('h2', fontName=B, fontSize=11, textColor=BLACK, spaceBefore=14, spaceAfter=4)
    body_s    = S('b', fontSize=9.5, textColor=BLACK, leading=15)
    footer_s  = S('f', fontName=I, fontSize=7.5, textColor=GRAY, alignment=TA_CENTER)

    def hr(col='#cccccc', thick=0.5, before=4, after=8):
        return HRFlowable(width='100%', thickness=thick,
                          color=colors.HexColor(col), spaceBefore=before, spaceAfter=after)

    story = []

    # Header
    # Header — centered ZEI + subtitle
    story.append(Spacer(1, 10))
    story.append(Paragraph('ZEI', S('logo', fontName=B, fontSize=32, textColor=BLACK,
                                     alignment=TA_CENTER, spaceAfter=0, leading=36)))
    story.append(Paragraph('ATTENDANCE REPORT', S('subt', fontSize=8, textColor=GRAY,
                                                    letterSpacing=5, alignment=TA_CENTER, spaceAfter=0)))
    story.append(Spacer(1, 12))
    story.append(hr('#0a0a0a', 1.5, before=0, after=12))
    story.append(Paragraph(f'Student: <b>{student_name}</b>', body_s))
    story.append(Paragraph(f'ID: {student_id}  ·  Period: January – February 2026', sub_s))
    story.append(Spacer(1, 14))

    # Chart
    story.append(Paragraph('ATTENDANCE TREND', S('ch', fontName=B, fontSize=9, textColor=GRAY, letterSpacing=3)))
    story.append(hr('#cccccc', 0.4, before=2, after=6))
    page_w = A4[0] - 4*cm
    chart_img = RLImage(chart_path, width=page_w, height=page_w*0.4)
    story.append(chart_img)
    story.append(Spacer(1, 14))

    # Overall summary table
    story.append(Paragraph('SUBJECT BREAKDOWN', S('ch2', fontName=B, fontSize=9, textColor=GRAY, letterSpacing=3)))
    story.append(hr('#cccccc', 0.4, before=2, after=6))

    tbl_data = [['Subject', 'Jan', 'Feb', 'Overall', 'Status']]
    for subj in all_subjects:
        short = subj.replace('Food Production Foundation - II', 'Food Production') \
                    .replace('Food & Beverage Service Foundation - II', 'F&B Service') \
                    .replace('Front Office Foundation - II', 'Front Office') \
                    .replace('Housekeeping Skills - II', 'Housekeeping') \
                    .replace('Personality Development and Grooming', 'Personality Dev')
        j = next((s for s in (jan_data or []) if s['subject'] == subj), None)
        f = next((s for s in (feb_data or []) if s['subject'] == subj), None)
        j_str = f"{j['percentage']:.1f}%" if j else '—'
        f_str = f"{f['percentage']:.1f}%" if f else '—' 
        # Overall
        j_held = j['classes_held'] if j else 0
        j_att  = j['classes_attended'] if j else 0
        f_held = f['classes_held'] if f else 0
        f_att  = f['classes_attended'] if f else 0
        total_held = j_held + f_held
        total_att  = j_att + f_att
        overall_pct = round(total_att / total_held * 100, 1) if total_held > 0 else 0
        status = '✓' if overall_pct >= 75 else '✗'
        tbl_data.append([short, j_str, f_str, f'{overall_pct:.1f}%', status])

    tbl = Table(tbl_data, colWidths=[5.5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 1.5*cm])
    tbl.setStyle(TableStyle([
        ('FONTNAME',    (0,0), (-1,0), B),
        ('FONTSIZE',    (0,0), (-1,-1), 9),
        ('FONTNAME',    (0,1), (-1,-1), R),
        ('BACKGROUND',  (0,0), (-1,0), BLACK),
        ('TEXTCOLOR',   (0,0), (-1,0), WHITE),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [WHITE, XGRAY]),
        ('ALIGN',       (1,0), (-1,-1), 'CENTER'),
        ('ALIGN',       (0,0), (0,-1), 'LEFT'),
        ('TOPPADDING',  (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('BOX',         (0,0), (-1,-1), 0.5, LGRAY),
        ('INNERGRID',   (0,0), (-1,-1), 0.3, LGRAY),
        ('LINEBELOW',   (0,0), (-1,0), 1, BLACK),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 16))

    # Leaderboard
    if leaderboard_data:
        story.append(Paragraph('CLASS LEADERBOARD', S('lb', fontName=B, fontSize=9, textColor=GRAY, letterSpacing=3)))
        story.append(hr('#cccccc', 0.4, before=2, after=6))
        lb_tbl_data = [['Rank', 'Student', 'Attended', 'Overall %']]
        student_rank_row = None
        for entry in leaderboard_data['data']:
            is_me = entry['student_id'] == student_id
            row = [f"#{entry['rank']}", entry['name'],
                   f"{entry['attended']}/{entry['held']}", f"{entry['pct']}%"]
            lb_tbl_data.append(row)
            if is_me:
                student_rank_row = len(lb_tbl_data) - 1

        lb_tbl = Table(lb_tbl_data, colWidths=[1.5*cm, 6.5*cm, 3*cm, 2.5*cm])
        style_cmds = [
            ('FONTNAME',    (0,0), (-1,0), B),
            ('FONTSIZE',    (0,0), (-1,-1), 9),
            ('FONTNAME',    (0,1), (-1,-1), R),
            ('BACKGROUND',  (0,0), (-1,0), BLACK),
            ('TEXTCOLOR',   (0,0), (-1,0), WHITE),
            ('ROWBACKGROUNDS',(0,1),(-1,-1), [WHITE, XGRAY]),
            ('ALIGN',       (0,0), (-1,-1), 'CENTER'),
            ('ALIGN',       (1,0), (1,-1), 'LEFT'),
            ('TOPPADDING',  (0,0), (-1,-1), 5),
            ('BOTTOMPADDING',(0,0),(-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('BOX',         (0,0), (-1,-1), 0.5, LGRAY),
            ('INNERGRID',   (0,0), (-1,-1), 0.3, LGRAY),
            ('LINEBELOW',   (0,0), (-1,0), 1, BLACK),
        ]
        # Highlight student row
        if student_rank_row:
            style_cmds += [
                ('BACKGROUND', (0, student_rank_row), (-1, student_rank_row), colors.HexColor('#e8e8e8')),
                ('FONTNAME',   (0, student_rank_row), (-1, student_rank_row), B),
            ]
        lb_tbl.setStyle(TableStyle(style_cmds))
        story.append(lb_tbl)

    # Footer
    story.append(Spacer(1, 20))
    story.append(hr('#0a0a0a', 1, before=0, after=6))
    story.append(Paragraph(
        f'ZEI  ·  Zephy Intelligence  ·  {datetime.now().strftime("%d %B %Y")}',
        footer_s
    ))

    doc.build(story)

    # Cleanup chart
    try: _os.remove(chart_path)
    except: pass

    return pdf_path

def main():
    load_student_data()
    setup_timetable()
    # Load events
    try:
        load_events_from_csv()
    except Exception as e:
        print(f"⚠️ Events load error: {e}")
    setup_events_table()
    load_events_from_csv()

    application = Application.builder().token(BOT_TOKEN).build()

    application.job_queue.run_repeating(
        lambda context: asyncio.create_task(send_class_reminders(application)),
        interval=60, first=10
    )
    application.job_queue.run_repeating(
        lambda context: asyncio.create_task(send_homework_reminders(application)),
        interval=60, first=10
    )
    application.job_queue.run_repeating(
        lambda context: asyncio.create_task(send_event_reminders(application)),
        interval=3600, first=30  # Check every hour
    )

    # All handlers
    application.add_handler(CommandHandler('logout', handle_logout))
    application.add_handler(CommandHandler('update', handle_cr_update))
    application.add_handler(CommandHandler('hw', handle_homework))
    application.add_handler(CommandHandler('complete', handle_complete_topic))
    application.add_handler(CommandHandler('lasttheory', handle_last_theory))
    application.add_handler(CommandHandler('lastpractical', handle_last_practical))
    application.add_handler(CommandHandler('syllabus', handle_syllabus))
    application.add_handler(CommandHandler('progress', handle_progress))
    application.add_handler(CommandHandler('attendance', handle_attendance))
    application.add_handler(CommandHandler('leaderboard', handle_leaderboard))
    application.add_handler(CommandHandler('stats', handle_stats))
    application.add_handler(CommandHandler('export', handle_export_users))
    application.add_handler(CommandHandler('speak', handle_speak))
    application.add_handler(CommandHandler('events', handle_events))




    application.add_handler(MessageHandler(filters.VOICE, handle_voice))
    # ✅ Photo with /lastpractical caption → practical handler, else → OCR
    application.add_handler(MessageHandler(
        filters.PHOTO & filters.CaptionRegex(r'^/lastpractical'),
        handle_last_practical
    ))
    application.add_handler(MessageHandler(
        filters.PHOTO & filters.CaptionRegex(r'^/addposter'),
        handle_addposter
    ))
    application.add_handler(MessageHandler(filters.PHOTO, handle_image))  # ✅ OCR
    application.add_handler(MessageHandler(
        filters.Document.PDF & filters.CaptionRegex(r'^/addnotes'),
        handle_addnotes
    ))
    application.add_handler(MessageHandler(filters.Document.PDF, handle_pdf_upload))
    application.add_handler(CommandHandler("events", handle_events))
    application.add_handler(CommandHandler("listnotes", handle_listnotes))
    application.add_handler(CommandHandler("testmorning", handle_test_morning))
    application.add_handler(CommandHandler("classnotes", handle_classnotes))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_handler(CallbackQueryHandler(handle_button_callback))

    total_users = get_user_count()
    monthly_active = get_monthly_active_users()

    print("🤖 Zei by Zephy Intelligence - FULLY UPGRADED!")
    print(f"  👥 Total Users: {total_users}")
    print(f"  📈 Monthly Active: {monthly_active}")
    print(f"  🔑 Admin ID: {ADMIN_ID}")
    print("  ✅ /update - AI powered (no template, smart preview)")
    print("  ✅ /hw - Homework with teacher + auto date detection")
    print("  ✅ /complete - Unit complete (teacher name bhi)")
    print("  ✅ /lasttheory - Theory class + auto date + DB save")
    print("  ✅ /lastpractical - Auto subject/date detect + DB save")
    print("  ✅ Real-time AI memory (saves → instantly queryable)")
    print("  ✅ Date-wise history: '17th feb ko kya tha?'")
    print("  ✅ Teacher-wise: 'Akash sir ne practical mein kya banaya?'")
    print("  ✅ Subject auto-detect: 'shepherd pie' → Food Production")
    print("  ✅ 🎤 Voice input (Hindi/Urdu/English) via Groq Whisper")
    print("  ✅ 🔊 Voice reply (Sarvam Bulbul v3, speaker: shubh, 48khz)")
    print("  ✅ 📷 Groq OCR (image text extraction)")

    # Schedule daily jobs
    job_queue = application.job_queue
    job_queue.run_daily(send_morning_briefing, time=datetime.strptime("06:00", "%H:%M").time())
    job_queue.run_daily(send_event_reminders, time=datetime.strptime("08:00", "%H:%M").time())
    job_queue.run_daily(send_daily_retention_report, time=datetime.strptime("21:00", "%H:%M").time())

    application.run_polling()

if __name__ == '__main__':
    main()
