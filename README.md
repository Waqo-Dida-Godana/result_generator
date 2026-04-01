# MOAS Result Generator / School MIS

Desktop school result management system built with Python, Tkinter, SQLite, matplotlib, pandas, and report/export tooling. The application is centered around learner records, marks capture, ranking, charts, report cards, exam analytics, and end-of-year promotion management.

## System Overview

This project is a single-user or small-school desktop MIS for CBC-style result processing. It combines:

- learner registration and class placement
- class, stream, subject, grading-scale, and teacher setup
- marks entry by term and exam type
- ranking, averages, charts, and result exports
- printable and emailable report cards
- exam deviation analytics across exam types over time
- year-end promotion workflows with audit logging and scheduled execution

The application runs as a Tkinter desktop app from `main.py` and persists all operational data in a local SQLite database managed by `database.py`.

## Architecture At a Glance

```text
User
  -> Tkinter UI in main.py
      -> Database service in database.py
      -> Exam analytics engine in exam_analytics.py
      -> Promotion orchestration in promotion.py
      -> Letterhead extraction helper in extract_letterhead.py
      -> File export layers: CSV, Excel, PDF, charts, email attachments
      -> Scheduled promotion CLI in run_promotion_task.py
```

## Core Module Map

| File | Role | Key Responsibilities |
| --- | --- | --- |
| `main.py` | Application shell and UI | Login, navigation, dashboard, settings, student/class/teacher screens, marks entry, ranking views, charts, report cards, email, promotions UI, analytics UI |
| `database.py` | Persistence layer | SQLite schema creation, migrations, CRUD operations, marks storage, ranking queries, settings, email logs, promotion history/audit support |
| `exam_analytics.py` | Analytical engine | Builds exam sessions from marks data, computes deviation metrics, ANOVA/time-series style comparisons, report text, and chart exports |
| `promotion.py` | Promotion service | Promotion settings, academic-year resolution, eligibility preview, repeat/no-data handling, auto-promotion checks, batch execution |
| `run_promotion_task.py` | CLI scheduler entry point | Manual or scheduled promotion execution, dry-runs, force/manual scope support, logging to console and `promotion_task.log` |
| `extract_letterhead.py` | Asset pre-processor | Extracts header/footer text and images from `assets/letterhead.docx` into PNG/JSON assets used by report outputs |
| `check_libs.py` | Environment sanity check | Confirms key reporting/export libraries import correctly |
| `run_promotion.ps1` and `run_promotion.bat` | Scheduler wrappers | Windows-friendly launch scripts for the promotion task |
| `install_and_run.ps1` | Environment helper | Example PowerShell helper for dependency installation checks |
| `test_exam_analytics_feature.py` | Analytics tests | Verifies session grouping, exam-type summaries, deviation matrix, date filtering, report generation |
| `test_promotion_feature.py` | Promotion tests | Verifies preview grouping, promotion execution, rollback, audit/history, trigger-date behavior |
| `test_all_imports.py`, `test_import.py`, `test_db.py`, `test_ver.py` | Smoke checks | Validate imports, database availability, and selected dependency versions |
| `build_exe.bat` and `*.spec` | Packaging helpers | PyInstaller-based desktop executable build path |

## Major Subsystems and How They Interact

### 1. Presentation Layer

`main.py` contains the `SchoolReportApp` class, which owns:

- login/auth screens
- the responsive topbar and collapsible sidebar
- page routing for all major modules
- table-heavy management screens
- export, print, and email actions

This layer never talks directly to raw SQL. It delegates persistence to the global `db` object from `database.py`, analytics to `exam_analytics`, and promotions to `promotion_manager`.

### 2. Data Layer

`database.py` exposes a `Database` class and a global `db = Database()`.

It handles:

- initial schema creation
- backwards-compatible schema migration
- CRUD for students, classes, subjects, streams, teachers, assignments, grading scales
- marks storage per `student + term + exam_type + subject`
- application settings and SMTP settings
- email delivery logging
- promotion history and audit logging

This module is the backbone of the system. Every UI action that persists or reads business data flows through it.

### 3. Analytics Layer

`exam_analytics.py` reads marks through the database layer and transforms them into `ExamSession` objects, then into an `ExamComparison` containing:

- deviation metrics
- anomalies and patterns
- recommendations
- ANOVA-style statistical output
- time-series trend output
- exam-type summaries
- subject deviation rows such as `Opener -> Mid-Term`

The analytics UI in `main.py` uses this service for on-screen analysis, text report generation, and chart exports.

### 4. Promotion Layer

`promotion.py` sits between the UI/CLI and the database promotion tables.

It is responsible for:

- validating promotion settings
- deciding the current academic year
- checking whether the annual promotion date has been reached
- previewing promoted, repeating, no-data, terminal, and already-processed students
- converting the preview into a validated batch request
- delegating the final transaction to `database.py`

### 5. Export and Document Layer

The application supports multiple output paths:

- CSV result exports
- Excel exports, including “Western Spotlight” formatting
- PDF report cards via ReportLab
- chart PNG exports via matplotlib
- result emails with PDF attachments via SMTP
- report card visual preview inside the Tkinter app

Letterhead assets are extracted once from `assets/letterhead.docx` and reused across PDF/Excel/report-card rendering.

## Database Design

The schema is created and migrated in `database.py`. The most important tables are:

| Table | Purpose |
| --- | --- |
| `users` | Login records plus role, name, email, abbreviation |
| `students` | Learner biographical and placement data |
| `marks` | Subject scores per student, term, and exam type |
| `school_classes` | Class catalog with level and abbreviation |
| `streams` | Stream definitions attached to classes |
| `custom_subjects` | Subject catalog by level and category |
| `teacher_assignments` | Subject-teacher and class-teacher assignments |
| `student_comments` | Per-student term comments |
| `grading_scales` | Configurable grade bands per class |
| `app_settings` | Generic app configuration, including SMTP settings |
| `email_logs` | Sent/failed email audit trail |
| `promotion_settings` | Promotion date and promotion rule configuration |
| `promotion_history` | Student-level promotion outcomes |
| `promotion_audit_log` | Batch and per-student promotion audit entries |

### Data Integrity Rules

- `marks` enforces uniqueness on `student_id + term + exam_type + subject`
- promotion batches are validated before commit
- invalid promotion batches are rejected without partial writes
- promotion history prevents duplicate yearly processing
- foreign keys are used across core relational tables

## Key Features

### Authentication and User Roles

- Default admin user is created automatically on first run.
- Login uses the `users` table.
- Teacher records can be created and assigned subject/class responsibilities.

Default credentials:

- Username: `admin`
- Password: `admin123`

### Academic Structure Management

The system supports CBC-oriented setup for:

- school levels
- classes
- streams
- subjects
- grading scales
- subject and class-teacher assignments

`main.py` also seeds default class and subject catalog records so a fresh database starts with usable baseline data.

### Student Management

Students can be:

- added manually
- edited
- deleted
- searched
- exported to Excel
- imported in batch from Excel templates

Student records include class, stream, gender, admission number, guardian name, parent email, and optional photo path.

### Marks Entry and Import

Marks are entered by:

- class
- stream
- term
- exam type
- subject

The system supports:

- direct grid-based entry
- single-student edit dialogs
- Excel template download
- Excel import for structured mark sheets
- workbook parsing for assessment-style layouts

Example exam types used in the app include:

- `Opener`
- `Mid-Term`
- `End-Term`
- `Quiz`
- `Assignment`
- `CAT`

### Results and Ranking

The Results page calculates:

- total score
- average score
- grade
- performance level
- class position
- subject averages

Exports include:

- standard CSV result sheets
- formatted “Western Spotlight” Excel workbooks

### Dashboard and Charts

The Dashboard summarizes:

- total students
- new students
- active classes
- average score
- exam-session performance trend
- enrollment movement

The Charts page visualizes:

- subject averages
- grade distribution
- top students
- class performance comparison

### Report Cards and Email Delivery

The Report Cards module provides:

- in-app report card preview
- single-student PDF export
- bulk report-card PDF export
- SMTP-based result emailing
- bulk email sending
- failed email review and retry

Email settings are stored in `app_settings`, and delivery attempts are tracked in `email_logs`.

### Exam Analytics

The analytics subsystem compares exam sessions over time and across exam types.

Key metrics include:

- score variance
- pass-rate fluctuation
- difficulty index shift
- mean score deviation
- standard deviation
- grade distribution shift
- performance consistency

It also generates subject-level comparisons such as:

| Subject | Baseline | Compared | Deviation |
| --- | --- | --- | --- |
| Mathematics | Opener 58.5 | Mid-Term 59.9 | `+1.4` |
| English | Opener 74.3 | Mid-Term 63.1 | `-11.2` |

This framework is described in more detail in `EXAM_ANALYTICS_README.md`, but the functionality is fully integrated into the main application.

### Promotions

The Promotions module supports:

- configurable annual promotion date
- optional academic-year override
- minimum passing average
- preview before execution
- repeat-year handling
- no-data detection
- terminal-class detection
- already-processed detection
- manual execution from the UI
- scheduled or manual CLI execution
- transaction-safe batch updates
- audit logging

Additional details are preserved in `PROMOTION_FEATURE_README.md`.

## Primary Operational Workflows

### Startup and Login Flow

1. Run `python main.py`.
2. `SchoolReportApp` initializes UI state, default catalog records, and grading scales.
3. `database.py` ensures schema and migrations are up to date.
4. The login screen appears.
5. After successful authentication, the main shell loads dashboard, sidebar, and topbar.

### Marks Entry Workflow

1. Open `Enter Marks`.
2. Choose class and optional stream.
3. Select term and exam type.
4. The app loads students and the applicable subject set.
5. Enter or import marks.
6. `save_marks()` collects values and writes them via `db.save_student_marks(...)`.
7. Marks are stored per subject and can later be reused by results, charts, analytics, promotions, and report cards.

### Exam Analysis Workflow

1. Open `Exam Analytics`.
2. Filter by class, term, or exam type.
3. Click `Run Analysis`.
4. `main.py` calls `exam_analytics.get_exam_sessions(...)`.
5. The analytics engine groups marks into distinct subject-level exam sessions.
6. `compare_exam_sessions(...)` computes deviation metrics, patterns, anomalies, summaries, ANOVA-style results, trend output, and subject deviation rows.
7. The UI displays:
   - summary text
   - deviation metrics table
   - subject deviation view
8. `Generate Report` creates a detailed text report.
9. `Export Charts` writes PNG visualizations for downstream reporting.

### Result Generation Workflow

1. Open `Results`.
2. Select class, term, and exam type.
3. `load_reports()` builds ranked result rows from the stored marks.
4. The screen shows subject-average tiles and a searchable ranking table.
5. Users can export:
   - CSV result tables
   - Spotlight-style Excel workbooks

### Report Card Workflow

1. Open `Report Cards`.
2. Select class, stream, term, exam type, and student.
3. The system builds a result object for that student.
4. The report card renderer creates:
   - learner info block
   - subject performance table
   - total/average/position summary
   - performance chart
   - comments area
   - optional header/footer letterhead
5. The user can preview, print/export PDF, or email the result.

### Promotion Workflow

1. Configure promotion settings in `Promotions`.
2. The app resolves the academic year and promotion due date.
3. `promotion_manager.get_promotion_preview()` classifies learners into:
   - eligible
   - repeating
   - no data
   - terminal
   - already processed
4. The user reviews the preview.
5. `Execute Promotions` converts preview data to a batch request.
6. `database.batch_promote_students(...)` validates the batch.
7. If validation passes, all updates, history rows, and audit rows commit in one transaction.
8. If validation fails, the batch is rejected and logged without partial promotion.

## Technical Specifications

### Runtime Stack

- Python desktop application
- Tkinter / ttk UI
- SQLite local database
- matplotlib for charts
- pandas and openpyxl for Excel import/export
- ReportLab and FPDF for PDF/report generation
- Pillow for image handling
- SMTP via Python standard library for email sending
- optional SciPy support for richer analytics statistics

### Data and Output Formats

- Database: `.db` SQLite files
- Charts: `.png`
- Reports: `.csv`, `.xlsx`, `.pdf`, `.txt`
- Letterhead metadata: `.json`

### UI Characteristics

- single-window desktop shell
- responsive topbar/sidebar refinements
- searchable paginated data tables
- integrated chart canvases
- modal dialogs for editing and assignment flows

## Setup Instructions

### Requirements

- Python 3 with Tkinter support
- pip
- Windows desktop environment is the primary target based on helper scripts and packaging files

### Install Dependencies

```powershell
python -m pip install -r requirements.txt
```

Optional analytics dependency:

```powershell
python -m pip install scipy
```

### Verify Libraries

```powershell
python check_libs.py
python test_all_imports.py
```

### Run the Application

```powershell
python main.py
```

## Usage Guide

### Recommended First-Time Setup

1. Log in with the default admin account.
2. Review `Settings` for classes, streams, subjects, teachers, and grading scales.
3. Add students or import them.
4. Enter marks by class, term, and exam type.
5. Review results and charts.
6. Configure email settings if report emailing is needed.
7. Configure promotion settings before the year-end cycle.

### Useful Commands

Run the desktop app:

```powershell
python main.py
```

Run promotion task manually:

```powershell
python run_promotion_task.py
```

Dry-run promotions:

```powershell
python run_promotion_task.py --dry-run
```

Process a single class:

```powershell
python run_promotion_task.py --class "Grade 7"
```

Force a manual promotion run before the trigger date:

```powershell
python run_promotion_task.py --force
```

Verbose promotion logging:

```powershell
python run_promotion_task.py --verbose
```

## Testing and Validation

Focused feature tests currently include:

```powershell
python -m unittest test_promotion_feature.py test_exam_analytics_feature.py
```

Additional smoke checks:

```powershell
python test_db.py
python test_import.py
python test_ver.py
```

## Build and Packaging

The repository includes Windows packaging helpers:

- `build_exe.bat`
- `MOAS_CBC_Report.spec`
- `MOAS_MIS.spec`

Typical build flow:

1. create or activate `.venv`
2. install `requirements.txt`
3. install `pyinstaller`
4. run `build_exe.bat`

The batch script bundles:

- `main.py`
- icons
- database files
- required Python dependencies

## Project Structure

```text
result_generator/
  assets/
    letterhead.docx
    letterhead.png
    letterhead_footer.png
    letterhead.json
  database.py
  main.py
  exam_analytics.py
  promotion.py
  run_promotion_task.py
  run_promotion.ps1
  run_promotion.bat
  extract_letterhead.py
  check_libs.py
  install_and_run.ps1
  requirements.txt
  EXAM_ANALYTICS_README.md
  PROMOTION_FEATURE_README.md
  test_*.py
  build_exe.bat
  *.spec
```

## Notes and Practical Considerations

- The default storage database used by the code is `school_report.db`.
- The app contains migration logic to keep older databases usable as new features are added.
- Analytics degrade gracefully when optional SciPy support is absent.
- Promotion execution is intentionally conservative and records audit history for review.
- Email delivery depends on valid SMTP settings stored inside the application.

## Summary

This codebase is a full desktop academic operations stack rather than a single report generator. `main.py` provides the working UI, `database.py` centralizes persistent business logic, `exam_analytics.py` adds comparative performance intelligence, and `promotion.py` manages year-end progression safely. Together they support the end-to-end cycle from learner setup to marks entry, analysis, reporting, communication, and promotion.
