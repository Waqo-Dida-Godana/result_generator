#!/usr/bin/env python3
from database import db

students = db.get_all_students()
print(f"Total students: {len(students)}")

if students:
    print("\nFirst student:")
    print(students[0])
    print(f"\nKeys: {list(students[0].keys())}")
else:
    print("No students in database")
