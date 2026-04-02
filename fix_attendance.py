import os
import shutil

# Check if attendance file exists
files = os.listdir('.')

attendance_files = [f for f in files if 'attendance' in f.lower() and f.endswith('.csv')]

if attendance_files:
    current_file = attendance_files[0]
    target_file = 'attendance_january_2026.csv'
    
    if current_file != target_file:
        shutil.copy(current_file, target_file)
        print(f"✅ Copied {current_file} → {target_file}")
    else:
        print(f"✅ {target_file} already exists!")
else:
    print("❌ No attendance CSV file found!")
    print("Please upload your attendance CSV file")