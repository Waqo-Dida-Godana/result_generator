# User Guide: Auto-Import Classes and Subjects with Marks

## Quick Start

### When Entering Marks with Students

You can now import students **and** automatically register their associated classes and subjects in a single operation!

## Step-by-Step Instructions

### Method 1: Import Students from Excel (Auto-detect Classes & Subjects)

1. **Click**: "📥 Import Excel" button in Students section
2. **Select** your Excel file containing student data
3. **Preview dialog** shows:
   - How many sheets will be imported
   - Row count
   - Which columns are being used
4. **Click**: "Proceed Import"
5. **Result**: 
   - Students added ✓
   - New classes automatically registered ✓
   - New subjects automatically registered ✓

### Excel File Format for Students (Optional Subject Column)

```
Column Headers (Required) | Optional
--------------------------|----------
admission_no              | (can be auto-generated)
student_name              | 
class                      | 
stream                     |
gender                     |
photo_path                 | (optional)
subject                    | ← NEW: Automatically imported
```

**Column Name Variations Accepted:**
- Class: "class", "grade", "class_name", "student_class"
- Subject: "subject", "subject_name", "subj", "course", "course_name"

### Method 2: Import Whole School Marks (Auto-detect Classes & Subjects)

1. **Click**: "📤 Import Whole School Results" button in Marks section
2. **Select** term and exam type
3. **Click**: "Choose Workbook"
4. **Select** your Excel file with multiple sheets (one per class)
5. **Wait** for processing:
   - Scans all sheets
   - Extracts subjects from column headers
   - Extracts class names from sheet context
6. **Result**:
   - Marks imported ✓
   - New subjects automatically added ✓
   - New classes automatically added ✓
   - Report shows what was added

## What Gets Auto-Detected?

### Classes
- Extracted from: Column data or imported Excel data
- Automatically assigned level:
  - PP1, PP2, Baby, Nursery → Pre-Primary
  - Grade 1-3 → Lower Primary
  - Grade 4-6 → Upper Primary
  - Grade 7+ → Junior Secondary
  - Default → Primary

### Subjects
- Extracted from: 
  - "subject" column (if present)
  - Mark sheet column headers
- Automatically assigned:
  - Level: Primary (default, based on class context)
  - Category: Core (default)

## Import Summary Report

After import completes, you'll see a detailed report showing:

```
Import Complete
═══════════════════════════════════════

✓ Done! X new student(s) added, Y updated
✓ Generated admission numbers: Z
✓ New subjects added: N          ← NEW!
✓ New classes added: M           ← NEW!
✓ Sheets processed: P
✓ Existing photos preserved where no new photo was supplied

Class breakdown:
- Grade 3: rows 25, added 5, updated 20, generated adm 0
- Grade 4: rows 30, added 10, updated 20, generated adm 0
```

## Avoiding Duplicates

The system automatically prevents duplicate entries:

- **Subjects**: If "English" already exists → NOT added again
- **Classes**: If "Grade 3" already exists → NOT added again
- Smart comparison catches slight variations

## Best Practices

### Do:
✓ Include subject column when importing students with marks
✓ Use standard class name formats (Grade X, Class X, PP1, etc.)
✓ Verify Excel data has header row
✓ Check import summary for what was added

### Don't:
✗ Manually add classes/subjects before import (system handles it)
✗ Use inconsistent class naming (Grade 3 vs grade 3 vs GRADE3)
✗ Leave subject names blank if column is included
✗ Trust that subjects/classes won't be added - they will be

## Troubleshooting

### "No valid worksheets found" Error
- Ensure Excel has at least a Name/Learner column
- Check column headers are in first row

### Subjects/Classes Not Added
- Check summary report shows 0 added
- Likely means they already exist in database
- Verify in Classes/Subjects pages

### Import Stops Unexpectedly
- Check file is valid Excel (.xlsx or .xls)
- Verify no locked/protected sheets
- Ensure sufficient disk space

## Examples

### Example 1: Mixed New and Existing Data
```
Excel has:
- Grade 3 (NEW) → Added automatically
- Grade 4 (EXISTS) → Not added again
- English (NEW) → Added automatically
- Mathematics (EXISTS) → Not added again

Result: Only new items added, no duplicates
```

### Example 2: Auto-Level Assignment
```
You import "Grade 3 Blue" → System recognizes "Grade" pattern
→ Assigns "Lower Primary" level automatically
→ No manual config needed
```

### Example 3: Multi-Sheet Mark Import
```
Workbook has sheets:
- "Grade 3 Blue" (NEW CLASS)
- Column headers: English, Mathematics, Science (NEW SUBJECTS)

After import:
✓ Grade 3 created with "Lower Primary" level
✓ English, Mathematics, Science added with "Core" category
✓ All marks imported correctly
✓ Summary shows what was added
```

## For System Administrators

### Database Impact
- **school_classes** table: New entries with auto-determined levels
- **custom_subjects** table: New entries with default category
- **No modifications** to existing entries
- **No duplicates** created

### Configuration
- Default level determination in `_determine_class_level()` method
- Default category for subjects: "Core"
- Can be customized in database methods if needed

### Monitoring
- Review import summary reports for patterns
- Check Categories/Subjects pages periodically
- Verify auto-levels match your school structure

## Related Features

- **Classes / Subjects Pages**: View all automatically added entries
- **Marks Entry**: Mark pages now show all auto-added classes
- **Subject Selection**: Assignment now works with auto-added subjects
- **Student List**: Shows students from auto-added classes

---

**Version**: 2.0 (Enhanced Import)
**Last Updated**: 2026-04-17
