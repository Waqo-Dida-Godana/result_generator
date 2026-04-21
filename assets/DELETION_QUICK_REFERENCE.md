# Quick Deletion Reference

## One-Line Solutions

### Delete Single Records
```python
# Delete student (with all marks)
db.delete_student('student_id')

# Delete class (with all students & marks)
db.delete_class('class_id')

# Delete subject (with all marks for that subject)
db.delete_subject('subject_id')

# Delete teacher
db.delete_user('teacher_id')
```

### Delete by Type (Batch)
```python
# Delete ALL students in one class
db.delete_all_students_in_class('Grade 4')

# Delete ALL marks for one term/exam
db.clear_all_marks(term='One', exam_type='End-Term')

# Delete all teachers
db.delete_all_teachers()

# Delete all subjects (and their marks)
db.delete_all_subjects()

# Delete all classes (students and marks)
db.delete_all_classes()

# DELETE EVERYTHING (except admin user)
db.reset_all_data()
```

## Common Scenarios

### "I entered wrong marks for Grade 4 Math"
```python
# Option 1: Delete just that term's marks
db.clear_all_marks(term='One', exam_type='End-Term')
# Then re-import correct marks

# Option 2: Delete that subject only
db.delete_subject('subject_id')  # All marks for Math deleted
```

### "I need to clear a class for new students"
```python
# Keeps class definition, removes students AND marks
db.delete_all_students_in_class('Grade 1')
```

### "Remove of all grade 5 (class and everything)"
```python
# Deletes class, streams, all students, all their marks
db.delete_class('class_id')
```

### "Year-end cleanup - start fresh"
```python
# NUCLEAR OPTION - everything gone except admin
db.reset_all_data()
```

### "Change curriculum - delete all subjects"
```python
result = db.delete_all_subjects()
# All marks deleted too
# Then add new subjects
```

## Deletion Operations Matrix

| What | Method | Cascades |
|-----|--------|----------|
| 1 Student | `delete_student(id)` | → Marks + Assignments |
| 1 Class | `delete_class(id)` | → Students + Marks + Streams + Assignments |
| 1 Subject | `delete_subject(id)` | → Marks + Assignments |
| 1 Teacher | `delete_user(id)` | → Assignments |
| All Students in Class | `delete_all_students_in_class(name)` | → Their marks |
| All Marks (1 term) | `clear_all_marks(term, exam)` | None |
| All Teachers | `delete_all_teachers()` | → Assignments |
| All Subjects | `delete_all_subjects()` | → Marks + Assignments |
| All Classes | `delete_all_classes()` | → Students + Marks + Streams + Assignments |
| EVERYTHING | `reset_all_data()` | Complete reset except admin |

## What Each Deletion Actually Removes

```
✓ = deleted from database
← = cascades to
- = unaffected

DELETE STUDENT
✓ Student record ← Marks ← Teacher assignments

DELETE CLASS  
✓ Class ← Students ← Marks
✓ Streams ← related data
✓ Teacher assignments for class

DELETE SUBJECT
✓ Subject ← All marks for this subject
✓ Teacher subject assignments

DELETE TEACHER
✓ Teacher record ✓ Their assignments
- Students and marks unaffected

CLEAR MARKS (by term/exam)
✓ Marks for that term/exam
- Students, classes, subjects unaffected

DELETE ALL SUBJECTS
✓ All subjects ✓ ALL marks ✓ Subject assignments
- Student records, class definitions unaffected

DELETE ALL CLASSES
✓ All classes ✓ All students ✓ All marks ✓ Streams
- Subjects remain ✓ Class assignments deleted

RESET ALL DATA
✓ Everything except admin user
- Admin user preserved
```

## Return Value Patterns

```python
# Single record delete (returns bool)
success = db.delete_student(id)
if success:
    print("Student deleted")

# Batch operations (return tuple)
success, message = db.delete_all_students_in_class('Grade 3')
if success:
    print(message)  
    # "Deleted 45 students and their marks from Grade 3"

success, message = db.reset_all_data()
if success:
    print(message)
    # "All data has been reset except admin user"
```

## Safety Checks

✓ Admin user **cannot** be deleted  
✓ **All deletions are permanent** - no undo  
✓ Marks cascade delete automatically  
✓ Orphaned data **is cleaned up**  
✓ Database integrity maintained  

## Database State After Operations

| Operation | Students | Marks | Classes | Subjects | Teachers |
|-----------|----------|-------|---------|----------|----------|
| Delete 1 Student | -1 | -N | 0 | 0 | 0 |
| Delete 1 Class | -N | -M | -1 | 0 | -K |
| Delete 1 Subject | 0 | -M | 0 | -1 | -K* |
| Clear Marks (Term) | 0 | -M | 0 | 0 | 0 |
| Reset All | -ALL | -ALL | -ALL | -ALL | -ALL* |

*= assignments only, teachers with no assignments kept (if custom added)

## Tips

1. **Before deleting**, count records:
   ```python
   cursor.execute('SELECT COUNT(*) FROM students WHERE class = ?', ('Grade 4',))
   print(f"Will delete: {cursor.fetchone()[0]} students")
   ```

2. **Verify deletion**:
   ```python
   cursor.execute('SELECT * FROM marks WHERE student_id = ?', (sid,))
   print(f"Remaining marks: {cursor.fetchall()}")  # Should be empty
   ```

3. **Backup before reset**:
   ```python
   # Copy database file before reset_all_data()
   import shutil
   shutil.copy('school_report.db', 'school_report_backup.db')
   db.reset_all_data()
   ```

4. **Check for orphans after deletion**:
   ```python
   cursor.execute('''
       SELECT COUNT(*) FROM marks m 
       WHERE NOT EXISTS (SELECT 1 FROM students s WHERE s.id = m.student_id)
   ''')
   print(f"Orphaned marks: {cursor.fetchone()[0]}")  # Should be 0
   ```

---

**All deletions now properly cascade - no more orphaned data!**
