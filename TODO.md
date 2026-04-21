Student Lists Print Functionality - Implementation Plan
Generated: [Current Date/Time]

## Status: 🔄 In Progress (1/7 complete)

✅ **1. Add helper queries to database.py** ✓

✅ **Plan Approved** - Class/stream-specific lists with recent averages, button in Students UI.

## Breakdown of Steps (Approved Plan)



### 2. [ ] Create student_list_pdf() function in main.py
   - Mirror report card structure (header/footer/tables)
   - Params: class_name, stream_name, include_averages=True
   - A4 portrait, subjects summary table

### 3. [ ] Add Students tab/page in main.py sidebar/UI
   - Similar to Classes/Reports: filters (class/stream), search
   - Treeview table of students
   - "Print Student List" button → call student_list_pdf()

### 4. [ ] Integrate print button + save dialog
   - tk.filedialog.asksaveasfilename("Student_List_[Class]_[Stream]_[Date].pdf")
   - Progress dialog for PDF generation

### 5. [ ] Test PDF generation (small class)
   - Verify letterhead/header/footer renders
   - Check table columns: Adm#, Name, Gender, Guardian, Email, Stream, Avg

### 6. [ ] Test large lists (100+ students)
   - Multi-page support
   - Performance + memory

### 7. [ ] Final validation + completion
   - A4 print preview
   - Match report card quality

## Notes
- Reuse existing ReportLab setup (styles/fonts/tables)
- No new dependencies needed
- Handle empty classes gracefully

**Next**: Step 2 - PDF generator function in main.py

---
*Updated automatically after each completed step.*

