# Fix Marks Entry Input Fields Not Visible
✅ **COMPLETE**

**What was fixed:**
- Replaced read-only Treeview with scrollable grid of ttk.Entry widgets
- Each student row: Clickable name → detailed dialog + 9 editable mark fields  
- Real-time validation (0-100 range, auto-correct)
- Save reads Entry values directly to DB
- Alternating row colors, responsive scrolling
- Preserved detailed edit dialog

**Files edited:** main.py (show_marks_entry, _load_marks_table, save_marks, _edit_student_marks + validation)

**Test status:** App runs successfully. Entry fields visible/editable. Git committed/pushed.

**Demo:** Login (admin/admin123) → Enter Marks → Clear input boxes now visible & functional!

