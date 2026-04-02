# migration.py - PEHLE YEH CHALAO!
import sqlite3

def migrate_database():
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    
    print("🔄 Starting migration...")
    
    try:
        # Add urdu_type to master_students
        try:
            c.execute("ALTER TABLE master_students ADD COLUMN urdu_type TEXT DEFAULT 'regular'")
            print("✅ Added urdu_type to master_students")
        except sqlite3.OperationalError:
            print("⚠️ urdu_type already exists in master_students")
        
        # Add urdu_type to class_schedule
        try:
            c.execute("ALTER TABLE class_schedule ADD COLUMN urdu_type TEXT DEFAULT 'both'")
            print("✅ Added urdu_type to class_schedule")
        except sqlite3.OperationalError:
            print("⚠️ urdu_type already exists in class_schedule")
        
        # 🔥 FIX class_updates table structure
        print("🔥 Fixing class_updates table...")
        
        # Check if columns exist
        c.execute("PRAGMA table_info(class_updates)")
        columns = [row[1] for row in c.fetchall()]
        
        # Drop and recreate if structure is wrong
        if 'class_type' not in columns or 'target_date' not in columns:
            print("   Recreating class_updates table with correct structure...")
            
            # Backup existing data
            c.execute("SELECT * FROM class_updates")
            backup_data = c.fetchall()
            
            # Drop old table
            c.execute("DROP TABLE IF EXISTS class_updates")
            
            # Create new structure
            c.execute('''CREATE TABLE class_updates
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
            
            print("   ✅ class_updates table recreated")
        else:
            print("   ✅ class_updates table already correct")
        
        conn.commit()
        print("\n✅ Migration completed!")
        
    except Exception as e:
        print(f"\n❌ Migration failed: {e}")
        conn.rollback()
    finally:
        conn.close()

if __name__ == '__main__':
    migrate_database()