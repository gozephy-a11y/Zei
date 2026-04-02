# fix_security.py
import sqlite3
from datetime import datetime

def fix_database():
    conn = sqlite3.connect('students.db')
    c = conn.cursor()
    
    print("🔄 Fixing database for security updates...")
    
    try:
        # 1. Add is_locked column to registered_users
        print("1️⃣ Adding is_locked column...")
        try:
            c.execute("ALTER TABLE registered_users ADD COLUMN is_locked INTEGER DEFAULT 1")
            print("   ✅ Added is_locked column")
        except sqlite3.OperationalError as e:
            if "duplicate column" in str(e).lower():
                print("   ⚠️ Column already exists")
            else:
                raise
        
        # 2. Set all existing users to locked
        print("2️⃣ Locking existing accounts...")
        c.execute("UPDATE registered_users SET is_locked = 1 WHERE is_locked IS NULL")
        updated = c.rowcount
        print(f"   ✅ Locked {updated} existing accounts")
        
        # 3. Create unlock_requests table
        print("3️⃣ Creating unlock_requests table...")
        c.execute('''CREATE TABLE IF NOT EXISTS unlock_requests
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      telegram_id INTEGER,
                      student_id TEXT,
                      reason TEXT,
                      requested_at TIMESTAMP,
                      status TEXT DEFAULT 'pending',
                      admin_response TEXT,
                      responded_at TIMESTAMP)''')
        print("   ✅ unlock_requests table created")
        
        conn.commit()
        print("\n✅ Database fixed successfully!")
        print("\n📌 All accounts are now LOCKED for security")
        print("📌 Users will need admin approval to unlock")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        conn.rollback()
    finally:
        conn.close()

if __name__ == '__main__':
    fix_database()