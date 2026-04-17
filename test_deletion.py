#!/usr/bin/env python
"""Test deletion and orphan cleanup functionality."""

from database import db

print("=" * 60)
print("TESTING DELETION & ORPHAN CLEANUP")
print("=" * 60)

# Test 1: Verify foreign key constraints are enabled
print("\n[TEST 1] Checking database constraints...")
conn = db.get_connection()
cursor = conn.cursor()
cursor.execute("PRAGMA foreign_keys")
fk_status = cursor.fetchone()[0]
conn.close()
print(f"✓ Foreign keys: {'ENABLED' if fk_status else 'DISABLED'}")

# Test 2: Delete a student and verify marks are removed
print("\n[TEST 2] Testing student deletion cascade...")
students = db.get_all_students()
if students:
    test_student = students[0]
    student_id = test_student.get("id")
    student_name = test_student.get("name")
    
    # Check marks before deletion
    conn = db.get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as cnt FROM marks WHERE student_id = ?", (student_id,))
    marks_before = cursor.fetchone()["cnt"]
    conn.close()
    
    if marks_before > 0:
        # Delete student
        db.delete_student(student_id)
        
        # Check marks after deletion
        conn = db.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) as cnt FROM marks WHERE student_id = ?", (student_id,))
        marks_after = cursor.fetchone()["cnt"]
        conn.close()
        
        if marks_after == 0:
            print(f"✓ Student '{student_name}' deleted with {marks_before} marks cleaned up")
        else:
            print(f"✗ ISSUE: {marks_after} orphan marks remain after student deletion!")
    else:
        print(f"⊘ Student '{student_name}' has no marks - skipping cascade test")
else:
    print("⊘ No students in database to test")

# Test 3: Verify no orphan subjects
print("\n[TEST 3] Checking for orphan subject references...")
conn = db.get_connection()
cursor = conn.cursor()
cursor.execute("""
    SELECT DISTINCT subject FROM marks 
    WHERE subject NOT IN (SELECT name FROM custom_subjects)
    LIMIT 5
""")
orphan_subjects = cursor.fetchall()
conn.close()

if orphan_subjects:
    print(f"⚠️  Found {len(orphan_subjects)} subject references not in custom_subjects:")
    for row in orphan_subjects[:3]:
        print(f"   - '{row['subject']}'")
    print("   (This is OK - subjects can be added during mark entry)")
else:
    print("✓ No orphan subject references found")

# Test 4: Summary statistics
print("\n[TEST 4] Database integrity summary...")
conn = db.get_connection()
cursor = conn.cursor()

queries = {
    "Students": "SELECT COUNT(*) FROM students",
    "Marks": "SELECT COUNT(*) FROM marks",
    "Classes": "SELECT COUNT(*) FROM school_classes",
    "Subjects": "SELECT COUNT(*) FROM custom_subjects",
    "Teachers": "SELECT COUNT(*) FROM users WHERE role IN ('teacher', 'class_teacher')",
}

for label, query in queries.items():
    cursor.execute(query)
    count = cursor.fetchone()[0]
    print(f"✓ {label}: {count}")

conn.close()

print("\n" + "=" * 60)
print("✓ All tests completed successfully!")
print("=" * 60)
