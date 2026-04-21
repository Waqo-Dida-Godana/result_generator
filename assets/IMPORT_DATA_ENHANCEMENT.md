# Data Import Enhancement - Marks, Classes, and Subjects

## Overview
This enhancement enables automatic detection and import of referenced data (classes and subjects) when importing marks or students from Excel files. The system now automatically registers new classes and subjects found in the imported data without creating redundancy.

## Changes Made

### 1. Enhanced Student Import (`import_excel` method)

#### Added Columns Used:
- **subject** (alongside existing student data):
  - Accepts variations: "subject", "subject_name", "subj", "course", "course_name"
  - Optional: Only processes if column contains data

#### New Features:
```python
# Automatically extracts and imports unique subjects
- Detects all unique subjects in the Excel data
- Checks if each subject already exists in database (by name)
- If not found, adds it with auto-determined level
- Prevents redundancy: does not re-add existing subjects

# Automatically extracts and imports unique classes
- Detects all unique classes in the Excel data
- Checks if each class already exists in database (by name)
- If not found, adds it with intelligently determined level
- Prevents redundancy: does not re-add existing classes
```

#### Import Preview Shows:
- Sheets ready for import
- Total rows found
- Fallback class (if no class column or default used)

#### Summary after import includes:
- New subjects added
- New classes added
- Traditional metrics (students added/updated, admission numbers generated)

### 2. Enhanced Marks Import (`_import_marks_workbook` method)

#### New Features:
```python
# When importing whole school marks:
- Scans assessment sheets for subject names in mark columns
- Extracts class names from sheet structure
- Adds any new subjects not in database
- Adds any new classes not in database
- All additions happen with no redundancy check avoiding duplicates
```

#### Summary Report Enhancement:
```
Import Complete
===============
Term: [Term]
Exam: [Exam Type]
Sheets imported: N
Sheets skipped: N
New subjects added: N
New classes added: N
Student records updated: N
...
```

### 3. New Helper Method: `_determine_class_level(class_name)`

Intelligently infers the appropriate level for a class during import:

```python
Logic:
1. Check against known CLASSES_BY_LEVEL mappings
2. Detect patterns in class name:
   - PP1, PP2, Baby, Nursery, Pre → "Pre-Primary"
   - Grade/Class/Std + number:
     - 1-3 → "Lower Primary"
     - 4-6 → "Upper Primary"
     - 7+ → "Junior Secondary"
3. Default: "Primary"
```

### 4. Column Alias Updates

Added new column detection patterns for better flexibility:

#### Class Column:
- "class" ✓ (existing)
- "grade" ✓ (existing)
- "class_name" ✓ (existing)
- "student_class" ✓ *(new)*

#### Subject Column (new):
- "subject"
- "subject_name"
- "subj"
- "course"
- "course_name"

## Use Cases

### Scenario 1: Import New Students with Subjects
```
Excel File (students_new.xlsx):
┌─────────┬──────────┬────────┬─────────────┐
│ Name    │ Class    │ Stream │ Subject     │
├─────────┼──────────┼────────┼─────────────┤
│ Alice   │ Grade 4  │ Blue   │ English     │
│ Bob     │ Grade 5  │ Red    │ Mathematics │
└─────────┴──────────┴────────┴─────────────┘

Result:
✓ Grade 4 class added (if not exists)
✓ Grade 5 class added (if not exists)
✓ English subject added (if not exists)
✓ Mathematics subject added (if not exists)
✓ Students imported with their data
✓ No duplicate classes/subjects created
```

### Scenario 2: Import Whole School Marks
```
Excel File (marks_term1.xlsx):
- Sheet: "Grade 3 Stream A"
- Headers: Name | English | Math | Science | ...
- Multiple sheets with different classes

Result:
✓ Classes Grade 3, Grade 4, etc. added (if not exist)
✓ Subjects English, Math, Science added (if not exist)
✓ All marks imported correctly
✓ No redundancy in master subject/class lists
```

### Scenario 3: Bulk Import from Multiple Sources
```
Can safely import from multiple Excel files:
- Same subjects appear in different files
- Same classes appear in different files
- System checks existence before adding
- Result: Clean, non-redundant database
```

## Database Operations

### Classes Table (school_classes)
```sql
- name (unique, so duplicate names are skipped)
- level (auto-determined or specified)
- stream (optional)
- abbreviation
```

### Subjects Table (custom_subjects)
```sql
- name (checked for existing entries)
- level (auto-determined)
- category (default: "Core")
- is_optional
- code/abbreviation
```

### Redundancy Prevention
- **Subjects**: Checked via `db.get_subject_by_name(subject_name)`
- **Classes**: Checked via `db.get_class_by_name(class_name)`
- Only added if NOT found (prevents duplicates)

## Benefits

1. **Simplified Workflow**: No need to manually add classes/subjects first
2. **Data Integrity**: Automatic level assignment based on class name
3. **Redundancy Prevention**: Smart duplicate detection
4. **Better Tracking**: Import summary shows what was added
5. **Flexible Input**: Accepts multiple column name variations
6. **Backward Compatible**: Existing import functionality unchanged

## Error Handling

- If subject/class addition fails, import continues with marks/students
- Import summary shows success counts (0 if all already existed)
- All exceptions properly caught and reported to user

## Example Excel Templates

### For Student Import with Subjects:
```
admission_no | student_name | class      | stream | gender | subject        | photo
-------------|--------------|------------|--------|--------|----------------|------
A001         | John Doe     | Grade 3    | Blue   | Male   | English        | john.jpg
A002         | Jane Smith   | Grade 3    | Blue   | Female | Mathematics    | jane.jpg
A003         | Mark Wilson  | Grade 4    | Red    | Male   | Science        | mark.jpg
```

### For Marks Import:
Each sheet name becomes the class identifier. Column headers become subject names.

## Testing Recommendations

1. Test with new subjects not in database
2. Test with existing subjects (should not duplicate)
3. Test with new classes not in database
4. Test with mixed old and new data
5. Verify import summary shows accurate counts
6. Check database after import for duplicates
7. Test with various class naming conventions (Grade X, Class X, PP1, etc.)

## Future Enhancements

- Allow custom level assignment during import
- Batch import with conflict resolution options
- Category assignment for subjects during import
- Subject abbreviation auto-generation
- Class abbreviation auto-generation
