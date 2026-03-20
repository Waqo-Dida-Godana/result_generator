"""
Database module for School Report Management System
Uses SQLite for local data storage
"""

import sqlite3
import uuid
from datetime import datetime
from typing import List, Dict, Optional, Tuple

DEFAULT_EXAM_TYPE = 'End-Term'


class Database:
    def __init__(self, db_name: str = "school_report.db"):
        self.db_name = db_name
        self.init_database()
    
    def get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_name)
        conn.row_factory = sqlite3.Row
        return conn
    
    def init_database(self):
        """Initialize database tables"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Users table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id TEXT PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        ''')
        
        # Students table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS students (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                class TEXT NOT NULL,
                gender TEXT NOT NULL CHECK (gender IN ('Male', 'Female')),
                admission_no TEXT UNIQUE NOT NULL,
                photo_path TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        ''')
        
        # Marks table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS marks (
                id TEXT PRIMARY KEY,
                student_id TEXT NOT NULL,
                term TEXT NOT NULL,
                exam_type TEXT NOT NULL DEFAULT 'End-Term',
                subject TEXT NOT NULL,
                marks INTEGER NOT NULL CHECK (marks >= 0 AND marks <= 100),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                UNIQUE(student_id, term, exam_type, subject)
            )
        ''')
        
        # Create indexes
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_students_class ON students(class)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_student_id ON marks(student_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_term ON marks(term)')
        
        # Create default admin user if not exists
        cursor.execute('SELECT * FROM users WHERE username = ?', ('admin',))
        if not cursor.fetchone():
            cursor.execute(
                'INSERT INTO users (id, username, password, created_at) VALUES (?, ?, ?, ?)',
                (str(uuid.uuid4()), 'admin', 'admin123', datetime.now().isoformat())
            )
        
        conn.commit()
        conn.close()
        self._migrate_schema()

    def _migrate_schema(self):
        """Add new columns/tables without breaking existing data."""
        conn = self.get_connection()
        cursor = conn.cursor()

        # Extend students table for legacy databases
        for col, defn in [
            ('photo_path', "TEXT DEFAULT ''"),
            ('stream', "TEXT NOT NULL DEFAULT ''"),
            ('guardian_name', "TEXT NOT NULL DEFAULT ''"),
            ('parent_email', "TEXT NOT NULL DEFAULT ''"),
            ('created_at', "TEXT NOT NULL DEFAULT ''"),
            ('updated_at', "TEXT NOT NULL DEFAULT ''"),
        ]:
            try:
                cursor.execute(f'ALTER TABLE students ADD COLUMN {col} {defn}')
            except Exception:
                pass  # column already exists

        # Extend users table with role, full_name, email
        for col, defn in [
            ('role',      "TEXT NOT NULL DEFAULT 'admin'"),
            ('full_name', "TEXT NOT NULL DEFAULT ''"),
            ('email',     "TEXT NOT NULL DEFAULT ''"),
            ('abbreviation', "TEXT NOT NULL DEFAULT ''"),
        ]:
            try:
                cursor.execute(f'ALTER TABLE users ADD COLUMN {col} {defn}')
            except Exception:
                pass  # column already exists

        # Teacher assignments (subject or class-teacher role per class)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS teacher_assignments (
                id               TEXT PRIMARY KEY,
                teacher_id       TEXT NOT NULL,
                class_name       TEXT NOT NULL,
                subject          TEXT,
                is_class_teacher INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (teacher_id) REFERENCES users(id) ON DELETE CASCADE,
                UNIQUE(teacher_id, class_name, subject)
            )
        ''')
        
        # School Classes table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS school_classes (
                id TEXT PRIMARY KEY,
                name TEXT UNIQUE NOT NULL,
                level TEXT NOT NULL,
                stream TEXT,
                abbreviation TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            )
        ''')
        
        # Custom Subjects table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS custom_subjects (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                level TEXT NOT NULL,
                category TEXT NOT NULL,
                is_optional INTEGER DEFAULT 0,
                code TEXT NOT NULL DEFAULT '',
                abbreviation TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                UNIQUE(name, level)
            )
        ''')

        for col, defn in [
            ('abbreviation', "TEXT NOT NULL DEFAULT ''"),
        ]:
            try:
                cursor.execute(f'ALTER TABLE school_classes ADD COLUMN {col} {defn}')
            except Exception:
                pass

        for col, defn in [
            ('code', "TEXT NOT NULL DEFAULT ''"),
            ('abbreviation', "TEXT NOT NULL DEFAULT ''"),
        ]:
            try:
                cursor.execute(f'ALTER TABLE custom_subjects ADD COLUMN {col} {defn}')
            except Exception:
                pass
        
        # Streams table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS streams (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                class_id TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (class_id) REFERENCES school_classes(id) ON DELETE CASCADE,
                UNIQUE(name, class_id)
            )
        ''')

        # Class-teacher comments per student per term
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS student_comments (
                id           TEXT PRIMARY KEY,
                student_id   TEXT NOT NULL,
                teacher_id   TEXT NOT NULL,
                term         TEXT NOT NULL,
                comment_text TEXT NOT NULL,
                created_at   TEXT NOT NULL,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                UNIQUE(student_id, term)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS grading_scales (
                id         TEXT PRIMARY KEY,
                class_name TEXT NOT NULL,
                min_mark   REAL NOT NULL,
                max_mark   REAL NOT NULL,
                grade_code TEXT NOT NULL,
                grade_name TEXT NOT NULL DEFAULT '',
                sort_order INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS email_logs (
                id TEXT PRIMARY KEY,
                student_id TEXT NOT NULL,
                term TEXT NOT NULL,
                exam_type TEXT NOT NULL DEFAULT 'End-Term',
                recipient_email TEXT NOT NULL,
                status TEXT NOT NULL,
                error_message TEXT NOT NULL DEFAULT '',
                sent_at TEXT NOT NULL,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
            )
        ''')

        self._migrate_marks_exam_type(cursor)

        conn.commit()
        conn.close()

    def _migrate_marks_exam_type(self, cursor):
        """Expand marks storage to support multiple exam types per term."""
        cursor.execute("PRAGMA table_info(marks)")
        columns = [row['name'] for row in cursor.fetchall()]
        if not columns:
            return

        if 'exam_type' in columns:
            cursor.execute(
                "UPDATE marks SET exam_type = ? WHERE exam_type IS NULL OR TRIM(exam_type) = ''",
                (DEFAULT_EXAM_TYPE,)
            )
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_term_exam ON marks(term, exam_type)')
            return

        cursor.execute('ALTER TABLE marks RENAME TO marks_legacy')
        cursor.execute('''
            CREATE TABLE marks (
                id TEXT PRIMARY KEY,
                student_id TEXT NOT NULL,
                term TEXT NOT NULL,
                exam_type TEXT NOT NULL DEFAULT 'End-Term',
                subject TEXT NOT NULL,
                marks INTEGER NOT NULL CHECK (marks >= 0 AND marks <= 100),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                UNIQUE(student_id, term, exam_type, subject)
            )
        ''')
        cursor.execute('''
            INSERT INTO marks (id, student_id, term, exam_type, subject, marks, created_at, updated_at)
            SELECT id, student_id, term, ?, subject, marks, created_at, updated_at
            FROM marks_legacy
        ''', (DEFAULT_EXAM_TYPE,))
        cursor.execute('DROP TABLE marks_legacy')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_student_id ON marks(student_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_term ON marks(term)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_term_exam ON marks(term, exam_type)')
    
    # User operations
    def authenticate(self, username: str, password: str) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ? AND password = ?', (username, password))
        row = cursor.fetchone()
        conn.close()
        if row:
            return dict(row)
        return None
    
    def register_user(self, name: str, email: str, password: str) -> bool:
        """Register a new user; email is stored as username."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM users WHERE username = ?', (email,))
        if cursor.fetchone():
            conn.close()
            return False
        cursor.execute(
            'INSERT INTO users (id, username, password, created_at) VALUES (?, ?, ?, ?)',
            (str(uuid.uuid4()), email, password, datetime.now().isoformat())
        )
        conn.commit()
        conn.close()
        return True

    # ── Teacher / user management ────────────────────────────────────────────
    def get_all_teachers(self) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM users WHERE role IN ('teacher','class_teacher') ORDER BY full_name, username"
        )
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def add_teacher(self, full_name: str, username: str, password: str, role: str, abbreviation: str = '') -> Tuple[bool, str]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM users WHERE username = ?', (username,))
        if cursor.fetchone():
            conn.close()
            return False, 'Username already exists'
        try:
            cursor.execute(
                'INSERT INTO users (id, username, password, role, full_name, email, abbreviation, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                (str(uuid.uuid4()), username, password, role, full_name, '', abbreviation.strip(), datetime.now().isoformat())
            )
            conn.commit()
            conn.close()
            return True, 'Teacher added successfully'
        except Exception as e:
            conn.close()
            return False, str(e)

    def update_teacher(self, teacher_id: str, full_name: str, username: str, role: str,
                       abbreviation: str = '', password: str = '') -> Tuple[bool, str]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM users WHERE username = ? AND id != ?', (username, teacher_id))
        if cursor.fetchone():
            conn.close()
            return False, 'Username already exists'
        try:
            if password.strip():
                cursor.execute(
                    'UPDATE users SET full_name = ?, username = ?, role = ?, abbreviation = ?, password = ? WHERE id = ? AND role != ?',
                    (full_name, username, role, abbreviation.strip(), password.strip(), teacher_id, 'admin')
                )
            else:
                cursor.execute(
                    'UPDATE users SET full_name = ?, username = ?, role = ?, abbreviation = ? WHERE id = ? AND role != ?',
                    (full_name, username, role, abbreviation.strip(), teacher_id, 'admin')
                )
            success = cursor.rowcount > 0
            conn.commit()
            conn.close()
            return (True, 'Teacher updated successfully') if success else (False, 'Teacher not found')
        except Exception as e:
            conn.close()
            return False, str(e)

    def delete_user(self, user_id: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE id = ? AND role != 'admin'", (user_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success

    # ── Assignment management ────────────────────────────────────────────────
    def assign_subject_teacher(self, teacher_id: str, class_name: str, subject: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                'INSERT OR REPLACE INTO teacher_assignments (id, teacher_id, class_name, subject, is_class_teacher) VALUES (?, ?, ?, ?, 0)',
                (str(uuid.uuid4()), teacher_id, class_name, subject)
            )
            conn.commit()
            conn.close()
            return True
        except Exception:
            conn.close()
            return False

    def assign_class_teacher(self, teacher_id: str, class_name: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                'DELETE FROM teacher_assignments WHERE class_name = ? AND is_class_teacher = 1',
                (class_name,)
            )
            cursor.execute(
                'INSERT INTO teacher_assignments (id, teacher_id, class_name, subject, is_class_teacher) VALUES (?, ?, ?, NULL, 1)',
                (str(uuid.uuid4()), teacher_id, class_name)
            )
            conn.commit()
            conn.close()
            return True
        except Exception:
            conn.close()
            return False

    def remove_assignment(self, assignment_id: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM teacher_assignments WHERE id = ?', (assignment_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success

    def get_all_subject_assignments(self) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT ta.id, ta.teacher_id, ta.class_name, ta.subject,
                   u.full_name, u.username, u.role
            FROM teacher_assignments ta
            JOIN users u ON ta.teacher_id = u.id
            WHERE ta.is_class_teacher = 0 AND ta.subject IS NOT NULL
            ORDER BY u.full_name, ta.class_name
        ''')
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_all_class_assignments(self) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT ta.id, ta.teacher_id, ta.class_name,
                   u.full_name, u.username, u.role
            FROM teacher_assignments ta
            JOIN users u ON ta.teacher_id = u.id
            WHERE ta.is_class_teacher = 1
            ORDER BY ta.class_name
        ''')
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_subject_teacher_assignments(self) -> List[Dict]:
        """Alias for get_all_subject_assignments (used by main.py)."""
        return self.get_all_subject_assignments()

    def get_class_teacher_assignments(self) -> List[Dict]:
        """Alias for get_all_class_assignments (used by main.py)."""
        return self.get_all_class_assignments()

    def get_teacher_subjects(self, teacher_id: str) -> List[Dict]:
        """Subjects+classes a subject teacher is assigned to."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT class_name, subject FROM teacher_assignments WHERE teacher_id = ? AND is_class_teacher = 0 AND subject IS NOT NULL',
            (teacher_id,)
        )
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_teacher_classes(self, teacher_id: str) -> List[str]:
        """Classes where user is assigned as class teacher."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT class_name FROM teacher_assignments WHERE teacher_id = ? AND is_class_teacher = 1',
            (teacher_id,)
        )
        rows = cursor.fetchall()
        conn.close()
        return [r['class_name'] for r in rows]
    
    # ── Class Management ─────────────────────────────────────────────────
    def add_class(self, name: str, level: str, stream: str = None, abbreviation: str = '') -> Tuple[bool, str]:
        """Add a new class"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                'INSERT INTO school_classes (id, name, level, stream, abbreviation, created_at) VALUES (?, ?, ?, ?, ?, ?)',
                (str(uuid.uuid4()), name, level, stream, abbreviation.strip(), datetime.now().isoformat())
            )
            conn.commit()
            conn.close()
            return True, 'Class added successfully'
        except Exception as e:
            conn.close()
            return False, str(e)

    def update_class(self, class_id: str, name: str, level: str, stream: str = None,
                     abbreviation: str = '') -> Tuple[bool, str]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('SELECT name FROM school_classes WHERE id = ?', (class_id,))
            row = cursor.fetchone()
            if not row:
                conn.close()
                return False, 'Class not found'
            old_name = row['name']
            cursor.execute(
                'UPDATE school_classes SET name = ?, level = ?, stream = ?, abbreviation = ? WHERE id = ?',
                (name, level, stream, abbreviation.strip(), class_id)
            )
            if old_name != name:
                cursor.execute('UPDATE students SET class = ? WHERE class = ?', (name, old_name))
                cursor.execute('UPDATE teacher_assignments SET class_name = ? WHERE class_name = ?', (name, old_name))
            success = cursor.rowcount > 0 or old_name != name
            conn.commit()
            conn.close()
            return (True, 'Class updated successfully') if success else (False, 'Class not found')
        except Exception as e:
            conn.close()
            return False, str(e)
    
    def get_all_classes(self) -> List[Dict]:
        """Get all classes"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM school_classes ORDER BY name')
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_class_by_name(self, name: str) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM school_classes WHERE name = ?', (name,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None
    
    def delete_class(self, class_id: str) -> bool:
        """Delete a class"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM school_classes WHERE id = ?', (class_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success
    
    # ── Stream Management ────────────────────────────────────────────────
    def add_stream(self, name: str, class_id: str) -> Tuple[bool, str]:
        """Add a new stream"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                'INSERT INTO streams (id, name, class_id, created_at) VALUES (?, ?, ?, ?)',
                (str(uuid.uuid4()), name, class_id, datetime.now().isoformat())
            )
            conn.commit()
            conn.close()
            return True, 'Stream added successfully'
        except Exception as e:
            conn.close()
            return False, str(e)
    
    def get_streams_for_class(self, class_id: str) -> List[Dict]:
        """Get all streams for a class"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM streams WHERE class_id = ? ORDER BY name', (class_id,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    
    def delete_stream(self, stream_id: str) -> bool:
        """Delete a stream"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM streams WHERE id = ?', (stream_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success
    
    # ── Subject Management ────────────────────────────────────────────────
    def add_subject(self, name: str, level: str, category: str, is_optional: bool = False,
                    abbreviation: str = '', code: str = '') -> Tuple[bool, str]:
        """Add a new subject"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            subject_code = (code or abbreviation).strip().upper()
            cursor.execute(
                'INSERT INTO custom_subjects (id, name, level, category, is_optional, code, abbreviation, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                (str(uuid.uuid4()), name, level, category, 1 if is_optional else 0, subject_code, subject_code or abbreviation.strip(), datetime.now().isoformat())
            )
            conn.commit()
            conn.close()
            return True, 'Subject added successfully'
        except Exception as e:
            conn.close()
            return False, str(e)

    def update_subject(self, subject_id: str, name: str, level: str, category: str,
                       is_optional: bool = False, abbreviation: str = '', code: str = '') -> Tuple[bool, str]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('SELECT name FROM custom_subjects WHERE id = ?', (subject_id,))
            row = cursor.fetchone()
            if not row:
                conn.close()
                return False, 'Subject not found'
            old_name = row['name']
            subject_code = (code or abbreviation).strip().upper()
            cursor.execute(
                'UPDATE custom_subjects SET name = ?, level = ?, category = ?, is_optional = ?, code = ?, abbreviation = ? WHERE id = ?',
                (name, level, category, 1 if is_optional else 0, subject_code, subject_code or abbreviation.strip(), subject_id)
            )
            if old_name != name:
                cursor.execute('UPDATE marks SET subject = ? WHERE subject = ?', (name, old_name))
                cursor.execute('UPDATE teacher_assignments SET subject = ? WHERE subject = ?', (name, old_name))
            success = cursor.rowcount > 0 or old_name != name
            conn.commit()
            conn.close()
            return (True, 'Subject updated successfully') if success else (False, 'Subject not found')
        except Exception as e:
            conn.close()
            return False, str(e)
    
    def get_subjects_by_level(self, level: str = None) -> List[Dict]:
        """Get subjects by level or all subjects"""
        conn = self.get_connection()
        cursor = conn.cursor()
        if level:
            cursor.execute('SELECT * FROM custom_subjects WHERE level = ? ORDER BY category, code, name', (level,))
        else:
            cursor.execute('SELECT * FROM custom_subjects ORDER BY level, category, code, name')
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_subject_by_name(self, name: str, level: str = None) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if level:
            cursor.execute('SELECT * FROM custom_subjects WHERE name = ? AND level = ?', (name, level))
        else:
            cursor.execute('SELECT * FROM custom_subjects WHERE name = ? ORDER BY level LIMIT 1', (name,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None
    
    def delete_subject(self, subject_id: str) -> bool:
        """Delete a subject"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM custom_subjects WHERE id = ?', (subject_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success

    def replace_subject_catalog(self, subjects: List[Dict]) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('DELETE FROM custom_subjects')
        for subject in subjects:
            code = str(subject.get('code', '') or '').strip().upper()
            cursor.execute(
                '''INSERT INTO custom_subjects
                   (id, name, level, category, is_optional, code, abbreviation, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (
                    str(uuid.uuid4()),
                    str(subject.get('name', '')).strip(),
                    str(subject.get('level', '')).strip(),
                    str(subject.get('category', 'Core')).strip(),
                    1 if subject.get('is_optional') else 0,
                    code,
                    code,
                    now
                )
            )
        conn.commit()
        conn.close()
        return True

    # Grading scale management
    def add_grading_scale(self, class_name: str, min_mark: float, max_mark: float,
                          grade_code: str, grade_name: str = '', sort_order: int = 0) -> Tuple[bool, str]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                '''INSERT INTO grading_scales
                   (id, class_name, min_mark, max_mark, grade_code, grade_name, sort_order, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (str(uuid.uuid4()), class_name, float(min_mark), float(max_mark),
                 grade_code.strip(), grade_name.strip(), int(sort_order), datetime.now().isoformat())
            )
            conn.commit()
            conn.close()
            return True, 'Grading scale added successfully'
        except Exception as e:
            conn.close()
            return False, str(e)

    def update_grading_scale(self, scale_id: str, class_name: str, min_mark: float, max_mark: float,
                             grade_code: str, grade_name: str = '', sort_order: int = 0) -> Tuple[bool, str]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                '''UPDATE grading_scales
                   SET class_name = ?, min_mark = ?, max_mark = ?, grade_code = ?, grade_name = ?, sort_order = ?
                   WHERE id = ?''',
                (class_name, float(min_mark), float(max_mark), grade_code.strip(),
                 grade_name.strip(), int(sort_order), scale_id)
            )
            success = cursor.rowcount > 0
            conn.commit()
            conn.close()
            return (True, 'Grading scale updated successfully') if success else (False, 'Grading scale not found')
        except Exception as e:
            conn.close()
            return False, str(e)

    def get_grading_scales(self, class_name: str = None) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if class_name:
            cursor.execute(
                'SELECT * FROM grading_scales WHERE class_name = ? ORDER BY sort_order, max_mark DESC, min_mark DESC',
                (class_name,)
            )
        else:
            cursor.execute(
                'SELECT * FROM grading_scales ORDER BY class_name, sort_order, max_mark DESC, min_mark DESC'
            )
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_grading_scale(self, scale_id: str) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM grading_scales WHERE id = ?', (scale_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def delete_grading_scale(self, scale_id: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM grading_scales WHERE id = ?', (scale_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success
    
    # ── Comments ─────────────────────────────────────────────────────────────
    def save_comment(self, student_id: str, teacher_id: str, term: str, comment_text: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute(
            'INSERT OR REPLACE INTO student_comments (id, student_id, teacher_id, term, comment_text, created_at) VALUES (?, ?, ?, ?, ?, ?)',
            (str(uuid.uuid4()), student_id, teacher_id, term, comment_text, now)
        )
        conn.commit()
        conn.close()
        return True

    def get_student_comment(self, student_id: str, term: str) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT sc.*, u.full_name FROM student_comments sc JOIN users u ON sc.teacher_id = u.id WHERE sc.student_id = ? AND sc.term = ?',
            (student_id, term)
        )
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_class_comments(self, class_name: str, term: str) -> Dict[str, str]:
        """Returns {student_id: comment_text} for a class/term."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT sc.student_id, sc.comment_text
            FROM student_comments sc
            JOIN students st ON sc.student_id = st.id
            WHERE st.class = ? AND sc.term = ?
        ''', (class_name, term))
        rows = cursor.fetchall()
        conn.close()
        return {r['student_id']: r['comment_text'] for r in rows}

    def change_password(self, username: str, old_password: str, new_password: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('UPDATE users SET password = ? WHERE username = ? AND password = ?', 
                       (new_password, username, old_password))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success
    
    # Student operations
    def get_all_students(self) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM students ORDER BY name')
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    def get_students_by_class(self, class_name: str) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM students WHERE class = ? ORDER BY name', (class_name,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_students_by_class_and_stream(self, class_name: str, stream: str = '') -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if stream:
            cursor.execute(
                'SELECT * FROM students WHERE class = ? AND stream = ? ORDER BY name',
                (class_name, stream)
            )
        else:
            cursor.execute('SELECT * FROM students WHERE class = ? ORDER BY name', (class_name,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    def add_student(self, name: str, class_name: str, gender: str, admission_no: str, photo_path: str = "",
                    guardian_name: str = "", parent_email: str = "", stream: str = "") -> Dict:
        conn = self.get_connection()
        cursor = conn.cursor()
        student_id = str(uuid.uuid4())
        now = datetime.now().isoformat()
        cursor.execute(
            '''INSERT INTO students
               (id, name, class, gender, admission_no, photo_path, stream, guardian_name, parent_email, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (student_id, name, class_name, gender, admission_no, photo_path, stream.strip(), guardian_name.strip(), parent_email.strip(), now, now)
        )
        conn.commit()
        conn.close()
        return {
            'id': student_id,
            'name': name,
            'class': class_name,
            'gender': gender,
            'admission_no': admission_no,
            'photo_path': photo_path,
            'stream': stream.strip(),
            'guardian_name': guardian_name.strip(),
            'parent_email': parent_email.strip(),
        }
    
    def update_student(self, student_id: str, name: str, class_name: str, gender: str, admission_no: str,
                       photo_path: str = "", guardian_name: str = "", parent_email: str = "",
                       stream: str = "") -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        if photo_path:
            # Update including new photo
            cursor.execute(
                '''UPDATE students
                   SET name = ?, class = ?, gender = ?, admission_no = ?, photo_path = ?,
                       stream = ?, guardian_name = ?, parent_email = ?, updated_at = ?
                   WHERE id = ?''',
                (name, class_name, gender, admission_no, photo_path, stream.strip(), guardian_name.strip(), parent_email.strip(), now, student_id)
            )
        else:
            # Preserve existing photo when no new photo is provided
            cursor.execute(
                '''UPDATE students
                   SET name = ?, class = ?, gender = ?, admission_no = ?,
                       stream = ?, guardian_name = ?, parent_email = ?, updated_at = ?
                   WHERE id = ?''',
                (name, class_name, gender, admission_no, stream.strip(), guardian_name.strip(), parent_email.strip(), now, student_id)
            )
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success

    def get_student_by_admission_no(self, admission_no: str) -> Optional[Dict]:
        """Find a student by their admission number."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM students WHERE admission_no = ?', (admission_no,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None
    
    def delete_student(self, student_id: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM students WHERE id = ?', (student_id,))
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success
    
    def search_students(self, query: str) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        search = f'%{query}%'
        cursor.execute(
            'SELECT * FROM students WHERE name LIKE ? OR admission_no LIKE ? OR class LIKE ? OR stream LIKE ? ORDER BY name',
            (search, search, search, search)
        )
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    # Marks operations
    def get_marks(self, term: str = 'One', exam_type: str = DEFAULT_EXAM_TYPE) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM marks WHERE term = ? AND exam_type = ?', (term, exam_type))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def set_setting(self, key: str, value: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)',
            (str(key).strip(), str(value or ''))
        )
        conn.commit()
        conn.close()
        return True

    def get_setting(self, key: str, default: str = '') -> str:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT value FROM app_settings WHERE key = ?', (str(key).strip(),))
        row = cursor.fetchone()
        conn.close()
        return str(row['value']) if row else default

    def get_settings(self, keys: List[str]) -> Dict[str, str]:
        return {key: self.get_setting(key, '') for key in keys}

    def log_email_delivery(self, student_id: str, term: str, exam_type: str,
                           recipient_email: str, status: str, error_message: str = '') -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            '''INSERT INTO email_logs
               (id, student_id, term, exam_type, recipient_email, status, error_message, sent_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
            (
                str(uuid.uuid4()),
                student_id,
                term,
                exam_type or DEFAULT_EXAM_TYPE,
                recipient_email.strip(),
                status.strip(),
                error_message.strip(),
                datetime.now().isoformat(),
            )
        )
        conn.commit()
        conn.close()
        return True

    def get_email_logs(self, class_name: str = '', term: str = '', exam_type: str = '',
                       stream: str = '',
                       status: str = '') -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT el.*, st.name AS student_name, st.class AS class_name, st.stream AS student_stream, st.admission_no
            FROM email_logs el
            JOIN students st ON el.student_id = st.id
            WHERE 1 = 1
        '''
        params = []
        if class_name:
            query += ' AND st.class = ?'
            params.append(class_name)
        if term:
            query += ' AND el.term = ?'
            params.append(term)
        if exam_type:
            query += ' AND el.exam_type = ?'
            params.append(exam_type)
        if stream:
            query += ' AND st.stream = ?'
            params.append(stream)
        if status:
            query += ' AND el.status = ?'
            params.append(status)
        query += ' ORDER BY el.sent_at DESC'
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_student(self, student_id: str) -> Optional[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM students WHERE id = ?', (student_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None
    
    def get_student_marks(self, student_id: str, term: str = 'One', exam_type: str = DEFAULT_EXAM_TYPE) -> Dict[str, int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT subject, marks FROM marks WHERE student_id = ? AND term = ? AND exam_type = ?',
            (student_id, term, exam_type)
        )
        rows = cursor.fetchall()
        conn.close()
        return {row['subject']: row['marks'] for row in rows}
    
    def save_student_marks(self, student_id: str, marks: Dict[str, int], term: str = 'One',
                           exam_type: str = DEFAULT_EXAM_TYPE) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        
        # Delete existing marks for this student, term, and exam type only.
        cursor.execute(
            'DELETE FROM marks WHERE student_id = ? AND term = ? AND exam_type = ?',
            (student_id, term, exam_type)
        )
        
        # Insert new marks
        for subject, value in marks.items():
            if value is not None and value != '':
                cursor.execute(
                    'INSERT INTO marks (id, student_id, term, exam_type, subject, marks, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                    (str(uuid.uuid4()), student_id, term, exam_type, subject, int(value), now, now)
                )
        
        conn.commit()
        conn.close()
        return True
    
    def save_all_marks(self, student_marks: Dict[str, Dict[str, int]], term: str = 'One',
                       exam_type: str = DEFAULT_EXAM_TYPE) -> bool:
        for student_id, marks in student_marks.items():
            self.save_student_marks(student_id, marks, term, exam_type)
        return True
    
    # Statistics
    def get_statistics(self, term: str = 'One', exam_type: str = DEFAULT_EXAM_TYPE) -> Dict:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Total students
        cursor.execute('SELECT COUNT(*) as count FROM students')
        total_students = cursor.fetchone()['count']
        
        # Get all marks for the term
        cursor.execute('''
            SELECT s.id, s.name, s.class, m.subject, m.marks 
            FROM students s 
            LEFT JOIN marks m ON s.id = m.student_id AND m.term = ? AND m.exam_type = ?
        ''', (term, exam_type))
        rows = cursor.fetchall()
        conn.close()
        
        # Calculate statistics
        marks_by_student = {}
        for row in rows:
            sid = row['id']
            if sid not in marks_by_student:
                marks_by_student[sid] = {'name': row['name'], 'class': row['class'], 'marks': []}
            if row['marks'] is not None:
                marks_by_student[sid]['marks'].append(row['marks'])
        
        if marks_by_student:
            all_averages = [sum(m['marks']) / len(m['marks']) if m['marks'] else 0 for m in marks_by_student.values()]
            avg_score = round(sum(all_averages) / len(all_averages), 1) if all_averages else 0
            top_student = max(marks_by_student.values(), key=lambda x: sum(x['marks']) if x['marks'] else 0)
            top_name = top_student['name'] if top_student['marks'] else '—'
        else:
            avg_score = 0
            top_name = '—'
        
        return {
            'students': total_students,
            'avg_score': avg_score,
            'top_student': top_name,
            'subjects': 9
        }
    
    # Results calculation
    def calculate_results(self, class_filter: str = 'All', term: str = 'One',
                          exam_type: str = DEFAULT_EXAM_TYPE) -> List[Dict]:
        subjects = ['Math', 'Eng', 'Kis', 'Int Sci', 'Agri', 'SST', 'CRE', 'CIA', 'Pre-Tech']
        
        conn = self.get_connection()
        cursor = conn.cursor()
        
        if class_filter == 'All':
            cursor.execute('SELECT * FROM students ORDER BY name')
        else:
            cursor.execute('SELECT * FROM students WHERE class = ? ORDER BY name', (class_filter,))
        
        students = [dict(row) for row in cursor.fetchall()]
        
        results = []
        for student in students:
            cursor.execute(
                'SELECT subject, marks FROM marks WHERE student_id = ? AND term = ? AND exam_type = ?',
                (student['id'], term, exam_type)
            )
            marks_rows = cursor.fetchall()
            marks = {row['subject']: row['marks'] for row in marks_rows}
            
            marks_values = [marks.get(s, 0) for s in subjects]
            total = sum(marks_values)
            average = round(total / len(subjects), 1) if subjects else 0
            
            # Get grade
            if average >= 80:
                grade = 'EE'
            elif average >= 70:
                grade = 'ME'
            elif average >= 60:
                grade = 'AE'
            elif average >= 50:
                grade = 'BE'
            else:
                grade = 'IE'
            
            results.append({
                'student': student,
                'marks': marks,
                'total': total,
                'average': average,
                'grade': grade
            })
        
        conn.close()
        
        # Sort by total descending and assign positions
        results.sort(key=lambda x: x['total'], reverse=True)
        for i, r in enumerate(results):
            r['position'] = i + 1
        
        return results

    # ── Class Exam History ─────────────────────────────────────────────────
    def get_class_exam_history(self, class_name: str) -> List[Dict]:
        """Get all exam sessions (term/exam_type combinations) for a specific class."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Get distinct term/exam_type combinations where this class has marks
        cursor.execute('''
            SELECT DISTINCT m.term, m.exam_type, m.created_at
            FROM marks m
            JOIN students s ON m.student_id = s.id
            WHERE s.class = ?
            ORDER BY m.created_at DESC, m.term, m.exam_type
        ''', (class_name,))
        
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    
    def get_all_classes_exam_history(self) -> List[Dict]:
        """Get exam history summary for all classes."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Get all classes
        cursor.execute('SELECT * FROM school_classes ORDER BY name')
        classes = [dict(r) for r in cursor.fetchall()]
        
        result = []
        for cls in classes:
            class_name = cls.get('name', '')
            
            # Get distinct exam sessions for this class
            cursor.execute('''
                SELECT DISTINCT m.term, m.exam_type
                FROM marks m
                JOIN students s ON m.student_id = s.id
                WHERE s.class = ?
                ORDER BY m.term DESC, m.exam_type
            ''', (class_name,))
            
            exams = [dict(r) for r in cursor.fetchall()]
            
            # Get student count
            cursor.execute('SELECT COUNT(*) as count FROM students WHERE class = ?', (class_name,))
            student_count = cursor.fetchone()['count']
            
            # Get average score for latest exam
            avg_score = None
            if exams:
                latest = exams[0]
                cursor.execute('''
                    SELECT AVG(m.marks) as avg
                    FROM marks m
                    JOIN students s ON m.student_id = s.id
                    WHERE s.class = ? AND m.term = ? AND m.exam_type = ?
                ''', (class_name, latest.get('term'), latest.get('exam_type')))
                avg_row = cursor.fetchone()
                avg_score = round(avg_row['avg'], 1) if avg_row and avg_row['avg'] else None
            
            result.append({
                'class': cls,
                'class_name': class_name,
                'level': cls.get('level', ''),
                'stream': cls.get('stream', ''),
                'student_count': student_count,
                'exam_count': len(exams),
                'exams': exams,
                'latest_avg': avg_score
            })
        
        conn.close()
        return result
    
    def get_class_exam_details(self, class_name: str, term: str, exam_type: str) -> List[Dict]:
        """Get detailed exam results for a specific class, term, and exam type."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Get distinct subjects for this exam (excluding 'TOTAL' which is a summary row)
        cursor.execute('''
            SELECT DISTINCT m.subject FROM marks m
            JOIN students s ON m.student_id = s.id
            WHERE s.class = ? AND m.term = ? AND m.exam_type = ? AND s.name != 'TOTAL'
        ''', (class_name, term, exam_type))
        subjects = [r['subject'] for r in cursor.fetchall()]
        
        # Get all students in the class
        cursor.execute('SELECT * FROM students WHERE class = ? ORDER BY name', (class_name,))
        students = [dict(r) for r in cursor.fetchall()]
        
        results = []
        for student in students:
            # Get marks for this student
            cursor.execute('''
                SELECT subject, marks FROM marks 
                WHERE student_id = ? AND term = ? AND exam_type = ?
            ''', (student['id'], term, exam_type))
            
            marks_rows = cursor.fetchall()
            marks = {row['subject']: row['marks'] for row in marks_rows}
            
            # Calculate total and average using actual subjects stored
            marks_values = [v for v in marks.values() if v is not None]
            total = sum(marks_values) if marks_values else 0
            average = round(total / len(marks_values), 1) if marks_values else 0
            
            # Determine grade
            if average >= 80:
                grade = 'EE'
            elif average >= 70:
                grade = 'ME'
            elif average >= 60:
                grade = 'AE'
            elif average >= 50:
                grade = 'BE'
            else:
                grade = 'IE'
            
            results.append({
                'student_id': student['id'],
                'student_name': student['name'],
                'admission_no': student.get('admission_no', ''),
                'marks': marks,
                'total': total,
                'average': average,
                'grade': grade
            })
        
        # Sort by total descending
        results.sort(key=lambda x: x['total'], reverse=True)
        
        # Assign positions
        for i, r in enumerate(results):
            r['position'] = i + 1
        
        conn.close()
        return results
    
    def get_available_exam_sessions(self) -> List[Dict]:
        """Get all available exam sessions across all classes."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT DISTINCT term, exam_type
            FROM marks
            ORDER BY term, exam_type
        ''')
        
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]


# Global database instance
db = Database()
