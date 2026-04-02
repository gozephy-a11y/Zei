import sqlite3

# Tumhara Telegram ID (bot se /start karne wale ka)
TELEGRAM_ID = 6011716383  # Ye tumhara telegram ID hai (Ali ka)

conn = sqlite3.connect('students.db')
c = conn.cursor()

# Delete user from registered_users
c.execute("DELETE FROM registered_users WHERE telegram_id = ?", (TELEGRAM_ID,))

# Delete conversation history
c.execute("DELETE FROM conversations WHERE telegram_id = ?", (TELEGRAM_ID,))

conn.commit()
conn.close()

print(f"✅ User {TELEGRAM_ID} reset done! you can now register again.")