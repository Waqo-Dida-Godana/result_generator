import sqlite3
conn = sqlite3.connect('school_report.db')
cursor = conn.cursor()
cursor.execute('PRAGMA table_info(marks)')
print('Marks table columns:')
for row in cursor.fetchall():
    print(row)
conn.close()
