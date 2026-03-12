# School Management System - Python Tkinter

## Requirements

- Python 3.8 or higher
- matplotlib>=3.5.0

## Installation

1. Make sure you have Python installed
2. Install the required packages:
   ```
   pip install matplotlib
   ```

## Running the Application

Navigate to the `python_tkinter` folder and run:
```
python main.py
```

## Default Login

- Username: admin
- Password: admin123

## Features

1. **Dashboard** - Overview of school statistics including total students, class average, top student, and grading scale

2. **Students Management** - Add, edit, delete, and search students with their admission number, name, class, and gender

3. **Marks Entry** - Enter marks for 9 subjects (Math, Eng, Kis, Int Sci, Agri, SST, CRE, CIA, Pre-Tech) across 3 terms

4. **Reports & Rankings** - View student rankings, subject performance, and export data to CSV

5. **Charts** - Visual analysis including:
   - Subject averages bar chart
   - Grade distribution pie chart
   - Top 5 students bar chart
   - Class performance comparison

6. **Report Cards** - Generate and print individual or bulk report cards

## Data Storage

The application uses SQLite for local data storage. A database file (`school_report.db`) will be created automatically in the application folder.

## Grading Scale

| Grade | Range | Description |
|-------|-------|-------------|
| EE | 80-100 | Exceeding Expectations |
| ME | 70-79 | Meeting Expectations |
| AE | 60-69 | Approaching Expectations |
| BE | 50-59 | Below Expectations |
| IE | 0-49 | Inadequate |
