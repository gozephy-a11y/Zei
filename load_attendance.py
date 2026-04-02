import sqlite3
import csv
from datetime import datetime

def load_attendance_from_csv():
    """Load attendance from the CSV you provided"""
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    
    print("🔄 Loading attendance data...")
    
    # Create attendance tables if not exist
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
    
    # Your CSV data
    csv_file = 'attendance_january_2026.csv'
    month = 'January'
    year = 2026
    
    try:
        with open(csv_file, 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            count = 0
            
            for row in csv_reader:
                student_id = row['student_id'].strip().upper()
                subject_name = row['subject_name'].strip()
                classes_held = int(row['classes_held'])
                classes_attended = int(row['classes_attended'])
                
                # Calculate percentage
                percentage = (classes_attended / classes_held * 100) if classes_held > 0 else 0
                
                c.execute('''INSERT OR REPLACE INTO attendance_records
                             (student_id, subject_name, month, year, 
                              classes_held, classes_attended, attendance_percentage, updated_at)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                          (student_id, subject_name, month, year,
                           classes_held, classes_attended, percentage, datetime.now()))
                count += 1
        
        conn.commit()
        print(f"✅ Loaded {count} attendance records for {month} {year}!")
        
        # Verify data
        c.execute("SELECT COUNT(DISTINCT student_id) FROM attendance_records")
        student_count = c.fetchone()[0]
        print(f"📊 Total students with attendance: {student_count}")
        
        # Show sample
        c.execute('''SELECT ms.name, ar.student_id, 
                            SUM(ar.classes_attended) as total_attended,
                            SUM(ar.classes_held) as total_held,
                            (CAST(SUM(ar.classes_attended) AS REAL) / SUM(ar.classes_held) * 100) as percentage
                     FROM attendance_records ar
                     JOIN master_students ms ON ar.student_id = ms.student_id
                     WHERE ar.month = ? AND ar.year = ?
                     GROUP BY ar.student_id, ms.name
                     ORDER BY percentage DESC
                     LIMIT 5''',
                  (month, year))
        
        print(f"\n🏆 Top 5 Students:")
        for i, row in enumerate(c.fetchall(), 1):
            print(f"{i}. {row[0]}: {row[4]:.1f}%")
        
    except FileNotFoundError:
        print(f"❌ {csv_file} not found!")
        print("Create the file with your attendance data first.")
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        conn.close()

if __name__ == '__main__':
    load_attendance_from_csv()