"""
Database module for School Report Management System
Uses SQLite for local data storage
"""

import re
import sqlite3
import uuid
import os
import sys
from datetime import datetime
from typing import List, Dict, Optional, Tuple

DEFAULT_EXAM_TYPE = 'End-Term'


class Database:
    def __init__(self, db_name: str = "school_report.db"):
        self.db_name = self._resolve_db_path(db_name)
        self.init_database()

    def _resolve_db_path(self, db_name: str) -> str:
        if os.path.isabs(db_name):
            return db_name
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, db_name)
    
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

        self._ensure_class_scoped_student_admission_numbers(conn, cursor)

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
                stream_name      TEXT NOT NULL DEFAULT '',
                subject          TEXT,
                is_class_teacher INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (teacher_id) REFERENCES users(id) ON DELETE CASCADE,
                UNIQUE(teacher_id, class_name, stream_name, subject, is_class_teacher)
            )
        ''')
        self._ensure_stream_scoped_teacher_assignments(conn, cursor)
        
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

        # Promotion history table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS promotion_history (
                id TEXT PRIMARY KEY,
                student_id TEXT NOT NULL,
                from_class TEXT NOT NULL,
                to_class TEXT NOT NULL,
                promotion_date TEXT NOT NULL,
                academic_year TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'promoted',
                reason TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_promotion_history_student_year ON promotion_history(student_id, academic_year)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_promotion_history_year ON promotion_history(academic_year)')

        # Promotion audit log table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS promotion_audit_log (
                id TEXT PRIMARY KEY,
                promotion_batch_id TEXT NOT NULL,
                action TEXT NOT NULL,
                details TEXT NOT NULL DEFAULT '',
                performed_by TEXT,
                performed_at TEXT NOT NULL,
                FOREIGN KEY (performed_by) REFERENCES users(id) ON DELETE SET NULL
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_promotion_audit_batch ON promotion_audit_log(promotion_batch_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_promotion_audit_time ON promotion_audit_log(performed_at)')

        # Promotion settings table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS promotion_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL DEFAULT ''
            )
        ''')

        self._migrate_marks_exam_type(cursor)

        conn.commit()
        conn.close()

    def _ensure_stream_scoped_teacher_assignments(self, conn, cursor):
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='teacher_assignments'")
        if not cursor.fetchone():
            return

        cursor.execute('PRAGMA table_info(teacher_assignments)')
        columns = [row[1] for row in cursor.fetchall()]

        cursor.execute('PRAGMA index_list(teacher_assignments)')
        unique_indexes = []
        for row in cursor.fetchall():
            if not bool(row[2]):
                continue
            index_name = row[1]
            cursor.execute(f'PRAGMA index_info("{index_name}")')
            unique_indexes.append([info[2] for info in cursor.fetchall()])

        expected_unique = ['teacher_id', 'class_name', 'stream_name', 'subject', 'is_class_teacher']
        if 'stream_name' in columns and expected_unique in unique_indexes:
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_teacher_assignments_class_stream ON teacher_assignments(class_name, stream_name)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_teacher_assignments_teacher ON teacher_assignments(teacher_id)')
            return

        stream_select = 'COALESCE(stream_name, \'\')' if 'stream_name' in columns else "''"
        cursor.execute('PRAGMA foreign_keys = OFF')
        cursor.execute('DROP TABLE IF EXISTS teacher_assignments_new')
        cursor.execute('''
            CREATE TABLE teacher_assignments_new (
                id               TEXT PRIMARY KEY,
                teacher_id       TEXT NOT NULL,
                class_name       TEXT NOT NULL,
                stream_name      TEXT NOT NULL DEFAULT '',
                subject          TEXT,
                is_class_teacher INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (teacher_id) REFERENCES users(id) ON DELETE CASCADE,
                UNIQUE(teacher_id, class_name, stream_name, subject, is_class_teacher)
            )
        ''')
        cursor.execute(f'''
            INSERT OR IGNORE INTO teacher_assignments_new (
                id, teacher_id, class_name, stream_name, subject, is_class_teacher
            )
            SELECT
                id, teacher_id, class_name, {stream_select}, subject, is_class_teacher
            FROM teacher_assignments
        ''')
        cursor.execute('DROP TABLE teacher_assignments')
        cursor.execute('ALTER TABLE teacher_assignments_new RENAME TO teacher_assignments')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_teacher_assignments_class_stream ON teacher_assignments(class_name, stream_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_teacher_assignments_teacher ON teacher_assignments(teacher_id)')
        cursor.execute('PRAGMA foreign_keys = ON')

    def _ensure_class_scoped_student_admission_numbers(self, conn, cursor):
        """Allow admission numbers like 1,2,3 to repeat across different classes."""
        cursor.execute('PRAGMA index_list(students)')
        index_rows = cursor.fetchall()
        has_class_scoped_unique = False
        has_global_admission_unique = False

        for row in index_rows:
            index_name = row[1]
            is_unique = bool(row[2])
            if not is_unique:
                continue
            cursor.execute(f'PRAGMA index_info("{index_name}")')
            columns = [info[2] for info in cursor.fetchall()]
            if columns == ['class', 'admission_no']:
                has_class_scoped_unique = True
            if columns == ['admission_no']:
                has_global_admission_unique = True

        if has_class_scoped_unique and not has_global_admission_unique:
            return

        cursor.execute('PRAGMA foreign_keys = OFF')
        cursor.execute('DROP TABLE IF EXISTS students_new')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS students_new (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                class TEXT NOT NULL,
                gender TEXT NOT NULL CHECK (gender IN ('Male', 'Female')),
                admission_no TEXT NOT NULL,
                photo_path TEXT DEFAULT '',
                stream TEXT NOT NULL DEFAULT '',
                guardian_name TEXT NOT NULL DEFAULT '',
                parent_email TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT ''
            )
        ''')
        cursor.execute('''
            INSERT INTO students_new (
                id, name, class, gender, admission_no, photo_path,
                stream, guardian_name, parent_email, created_at, updated_at
            )
            SELECT
                id, name, class, gender, admission_no, photo_path,
                stream, guardian_name, parent_email, created_at, updated_at
            FROM students
        ''')
        cursor.execute('DROP TABLE students')
        cursor.execute('ALTER TABLE students_new RENAME TO students')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_students_class ON students(class)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_students_admission_no ON students(admission_no)')
        cursor.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_students_class_admission_no ON students(class, admission_no)')
        cursor.execute('PRAGMA foreign_keys = ON')

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
    def assign_subject_teacher(self, teacher_id: str, class_name: str, subject: str, stream_name: str = '') -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            stream_name = (stream_name or '').strip()
            cursor.execute(
                '''DELETE FROM teacher_assignments
                   WHERE teacher_id = ? AND class_name = ? AND stream_name = ? AND subject = ? AND is_class_teacher = 0''',
                (teacher_id, class_name, stream_name, subject)
            )
            cursor.execute(
                '''INSERT INTO teacher_assignments
                   (id, teacher_id, class_name, stream_name, subject, is_class_teacher)
                   VALUES (?, ?, ?, ?, ?, 0)''',
                (str(uuid.uuid4()), teacher_id, class_name, stream_name, subject)
            )
            conn.commit()
            conn.close()
            return True
        except Exception:
            conn.close()
            return False

    def assign_class_teacher(self, teacher_id: str, class_name: str, stream_name: str = '') -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            stream_name = (stream_name or '').strip()
            cursor.execute(
                'DELETE FROM teacher_assignments WHERE class_name = ? AND stream_name = ? AND is_class_teacher = 1',
                (class_name, stream_name)
            )
            cursor.execute(
                '''INSERT INTO teacher_assignments
                   (id, teacher_id, class_name, stream_name, subject, is_class_teacher)
                   VALUES (?, ?, ?, ?, NULL, 1)''',
                (str(uuid.uuid4()), teacher_id, class_name, stream_name)
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
            SELECT ta.id, ta.teacher_id, ta.class_name, ta.stream_name, ta.subject,
                   u.full_name, u.username, u.role
            FROM teacher_assignments ta
            JOIN users u ON ta.teacher_id = u.id
            WHERE ta.is_class_teacher = 0 AND ta.subject IS NOT NULL
            ORDER BY u.full_name, ta.class_name, ta.stream_name, ta.subject
        ''')
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_all_class_assignments(self) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT ta.id, ta.teacher_id, ta.class_name, ta.stream_name,
                   u.full_name, u.username, u.role
            FROM teacher_assignments ta
            JOIN users u ON ta.teacher_id = u.id
            WHERE ta.is_class_teacher = 1
            ORDER BY ta.class_name, ta.stream_name
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
            '''SELECT class_name, stream_name, subject
               FROM teacher_assignments
               WHERE teacher_id = ? AND is_class_teacher = 0 AND subject IS NOT NULL''',
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
            '''SELECT DISTINCT class_name
               FROM teacher_assignments
               WHERE teacher_id = ? AND is_class_teacher = 1''',
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

    def update_stream(self, stream_id: str, name: str, class_id: str) -> Tuple[bool, str]:
        """Update stream name and/or parent class."""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                'UPDATE streams SET name = ?, class_id = ? WHERE id = ?',
                (name, class_id, stream_id)
            )
            success = cursor.rowcount > 0
            conn.commit()
            conn.close()
            return (True, 'Stream updated successfully') if success else (False, 'Stream not found')
        except Exception as e:
            conn.close()
            return False, str(e)
    
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

    def get_student_by_admission_no(self, admission_no: str, class_name: str = None) -> Optional[Dict]:
        """Find a student by their admission number."""
        conn = self.get_connection()
        cursor = conn.cursor()
        if class_name:
            cursor.execute(
                'SELECT * FROM students WHERE admission_no = ? AND class = ? ORDER BY updated_at DESC, created_at DESC LIMIT 1',
                (admission_no, class_name)
            )
        else:
            cursor.execute('SELECT * FROM students WHERE admission_no = ? ORDER BY updated_at DESC, created_at DESC LIMIT 1', (admission_no,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_next_class_admission_no(self, class_name: str, exclude_student_id: str = '') -> str:
        conn = self.get_connection()
        cursor = conn.cursor()
        if exclude_student_id:
            cursor.execute(
                'SELECT admission_no FROM students WHERE class = ? AND id != ?',
                (class_name, exclude_student_id)
            )
        else:
            cursor.execute('SELECT admission_no FROM students WHERE class = ?', (class_name,))
        rows = cursor.fetchall()
        conn.close()

        used_numbers = set()
        for row in rows:
            admission_no = str(row['admission_no'] or '').strip()
            if re.fullmatch(r'\d+', admission_no):
                value = int(admission_no)
                if value > 0:
                    used_numbers.add(value)

        next_number = 1
        while next_number in used_numbers:
            next_number += 1
        return str(next_number)
    
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

    # ── Promotion Management ────────────────────────────────────────────────
    def get_promotion_setting(self, key: str, default: str = '') -> str:
        """Get a promotion setting value."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT value FROM promotion_settings WHERE key = ?', (str(key).strip(),))
        row = cursor.fetchone()
        conn.close()
        return str(row['value']) if row else default

    def set_promotion_setting(self, key: str, value: str) -> bool:
        """Set a promotion setting value."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'INSERT OR REPLACE INTO promotion_settings (key, value) VALUES (?, ?)',
            (str(key).strip(), str(value or ''))
        )
        conn.commit()
        conn.close()
        return True

    def get_promotion_settings(self, keys: List[str]) -> Dict[str, str]:
        """Get multiple promotion settings."""
        return {key: self.get_promotion_setting(key, '') for key in keys}

    def _get_level_sort_order(self, level: str) -> int:
        level_name = str(level or '').strip().lower()
        level_order = {
            'pre-primary': 0,
            'lower primary (grade 1-3)': 1,
            'upper primary (grade 4-6)': 2,
            'junior school (grade 7-9)': 3,
        }
        return level_order.get(level_name, 99)

    def _get_class_sort_number(self, class_name: str) -> int:
        class_label = str(class_name or '').strip().lower()
        aliases = {
            'baby class': -4,
            'play group': -3,
            'pp1': -2,
            'pre-primary 1': -2,
            'pp2': -1,
            'pre-primary 2': -1,
        }
        if class_label in aliases:
            return aliases[class_label]

        match = re.search(r'(\d+)', class_label)
        if match:
            return int(match.group(1))
        return 10_000

    def get_class_progression_order(self) -> List[str]:
        """Get the ordered list of classes for promotion progression."""
        classes = self.get_all_classes()
        sorted_classes = sorted(
            classes,
            key=lambda x: (
                self._get_level_sort_order(x.get('level', '')),
                self._get_class_sort_number(x.get('name', '')),
                str(x.get('name', '')).strip().lower(),
            )
        )
        return [c['name'] for c in sorted_classes]

    def get_next_class(self, current_class: str) -> Optional[str]:
        """Get the next class in the progression order."""
        progression = self.get_class_progression_order()
        try:
            current_index = progression.index(current_class)
            if current_index < len(progression) - 1:
                return progression[current_index + 1]
        except ValueError:
            pass
        return None

    def get_latest_exam_session_for_class(self, class_name: str) -> Optional[Dict]:
        """Return the most recent exam session recorded for a class."""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            '''
            SELECT
                m.term,
                m.exam_type,
                MAX(COALESCE(NULLIF(m.updated_at, ''), m.created_at)) AS recorded_at
            FROM marks m
            JOIN students s ON m.student_id = s.id
            WHERE s.class = ?
            GROUP BY m.term, m.exam_type
            ORDER BY recorded_at DESC, m.term DESC, m.exam_type DESC
            LIMIT 1
            ''',
            (class_name,)
        )
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def _get_students_with_promotion_averages(self, class_name: str) -> Tuple[Optional[Dict], List[Dict]]:
        latest_exam = self.get_latest_exam_session_for_class(class_name)
        if not latest_exam:
            return None, []

        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            '''
            SELECT
                s.id,
                s.name,
                s.admission_no,
                s.class,
                s.stream,
                AVG(m.marks) AS average_marks,
                COUNT(m.id) AS marks_count
            FROM students s
            LEFT JOIN marks m
                ON s.id = m.student_id
               AND m.term = ?
               AND m.exam_type = ?
            WHERE s.class = ?
            GROUP BY s.id, s.name, s.admission_no, s.class, s.stream
            ORDER BY s.name
            ''',
            (latest_exam['term'], latest_exam['exam_type'], class_name)
        )
        rows = []
        for row in cursor.fetchall():
            item = dict(row)
            item['term'] = latest_exam['term']
            item['exam_type'] = latest_exam['exam_type']
            item['recorded_at'] = latest_exam.get('recorded_at')
            rows.append(item)
        conn.close()
        return latest_exam, rows

    def get_students_eligible_for_promotion(self, class_name: str, min_average: float = 50.0) -> List[Dict]:
        """Get students eligible for promotion from a class based on their performance."""
        _, students = self._get_students_with_promotion_averages(class_name)
        eligible = [
            student for student in students
            if student.get('marks_count', 0) > 0 and student.get('average_marks') is not None
            and float(student['average_marks']) >= float(min_average)
        ]
        return sorted(eligible, key=lambda item: (-float(item.get('average_marks') or 0), item.get('name', '')))

    def get_students_repeating(self, class_name: str, min_average: float = 50.0) -> List[Dict]:
        """Get students who should repeat the class based on performance."""
        _, students = self._get_students_with_promotion_averages(class_name)
        repeating = [
            student for student in students
            if student.get('marks_count', 0) > 0 and student.get('average_marks') is not None
            and float(student['average_marks']) < float(min_average)
        ]
        return sorted(repeating, key=lambda item: (float(item.get('average_marks') or 0), item.get('name', '')))

    def get_students_without_promotion_data(self, class_name: str) -> List[Dict]:
        """Get students in a class who have no marks for the latest exam session."""
        _, students = self._get_students_with_promotion_averages(class_name)
        no_data = [student for student in students if int(student.get('marks_count') or 0) == 0]
        return sorted(no_data, key=lambda item: item.get('name', ''))

    def has_promotion_history(self, academic_year: str, student_id: str = None) -> bool:
        """Return True when promotion decisions already exist for the academic year."""
        conn = self.get_connection()
        cursor = conn.cursor()
        if student_id:
            cursor.execute(
                'SELECT 1 FROM promotion_history WHERE academic_year = ? AND student_id = ? LIMIT 1',
                (academic_year, student_id)
            )
        else:
            cursor.execute(
                'SELECT 1 FROM promotion_history WHERE academic_year = ? LIMIT 1',
                (academic_year,)
            )
        row = cursor.fetchone()
        conn.close()
        return row is not None

    def _resolve_promotion_actor(self, cursor, performed_by: str = None) -> Optional[str]:
        actor_id = str(performed_by or '').strip()
        if not actor_id:
            return None
        cursor.execute('SELECT id FROM users WHERE id = ?', (actor_id,))
        row = cursor.fetchone()
        return row['id'] if row else None

    def _insert_promotion_audit_entry(self, cursor, batch_id: str, action: str, details: str,
                                      performed_by: str = None, performed_at: str = None) -> None:
        cursor.execute(
            '''INSERT INTO promotion_audit_log
               (id, promotion_batch_id, action, details, performed_by, performed_at)
               VALUES (?, ?, ?, ?, ?, ?)''',
            (
                str(uuid.uuid4()),
                batch_id,
                action,
                str(details or ''),
                performed_by,
                performed_at or datetime.now().isoformat(),
            )
        )

    def _write_promotion_audit_entry(self, batch_id: str, action: str, details: str,
                                     performed_by: str = None, performed_at: str = None) -> None:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            actor_id = self._resolve_promotion_actor(cursor, performed_by)
            self._insert_promotion_audit_entry(cursor, batch_id, action, details, actor_id, performed_at)
            conn.commit()
        finally:
            conn.close()

    def promote_student(self, student_id: str, from_class: str, to_class: str,
                       academic_year: str, status: str = 'promoted',
                       reason: str = '', performed_by: str = None) -> Tuple[bool, str]:
        """Promote a single student to the next class."""
        success, message, _ = self.batch_promote_students(
            [
                {
                    'student_id': student_id,
                    'from_class': from_class,
                    'to_class': to_class,
                    'status': status,
                    'reason': reason,
                }
            ],
            academic_year,
            performed_by,
        )
        return success, message

    def batch_promote_students(self, promotions: List[Dict], academic_year: str,
                              performed_by: str = None) -> Tuple[bool, str, Dict]:
        """Promote multiple students in a batch with transaction support."""
        conn = self.get_connection()
        cursor = conn.cursor()
        batch_id = str(uuid.uuid4())
        now = datetime.now().isoformat()
        results = {
            'batch_id': batch_id,
            'academic_year': academic_year,
            'promoted': 0,
            'repeating': 0,
            'failed': 0,
            'total': len(promotions),
            'errors': []
        }

        try:
            actor_id = self._resolve_promotion_actor(cursor, performed_by)
            student_ids = [str(promo.get('student_id', '')).strip() for promo in promotions]
            duplicate_student_ids = sorted({sid for sid in student_ids if sid and student_ids.count(sid) > 1})

            if duplicate_student_ids:
                results['failed'] = len(duplicate_student_ids)
                results['errors'] = [f'Duplicate promotion entries for student {sid}' for sid in duplicate_student_ids]
                self._write_promotion_audit_entry(
                    batch_id,
                    'BATCH_REJECTED',
                    '; '.join(results['errors']),
                    actor_id,
                    now,
                )
                conn.close()
                return False, 'Batch promotion rejected: duplicate students in request', results

            students_by_id = {}
            if student_ids:
                placeholders = ','.join('?' for _ in student_ids)
                cursor.execute(
                    f'SELECT id, name, class, admission_no FROM students WHERE id IN ({placeholders})',
                    student_ids
                )
                students_by_id = {row['id']: dict(row) for row in cursor.fetchall()}
                cursor.execute(
                    f'''
                    SELECT student_id
                    FROM promotion_history
                    WHERE academic_year = ?
                      AND student_id IN ({placeholders})
                    ''',
                    [academic_year, *student_ids]
                )
                already_processed = {row['student_id'] for row in cursor.fetchall()}
            else:
                already_processed = set()

            valid_classes = {row['name'] for row in self.get_all_classes()}
            for promo in promotions:
                student_id = str(promo.get('student_id', '')).strip()
                from_class = str(promo.get('from_class', '')).strip()
                to_class = str(promo.get('to_class', '')).strip()
                status = str(promo.get('status', 'promoted')).strip().lower() or 'promoted'

                if not student_id:
                    results['errors'].append('Promotion entry missing student_id')
                    continue
                if status not in ('promoted', 'repeating'):
                    results['errors'].append(f'Unsupported promotion status "{status}" for student {student_id}')
                    continue

                student = students_by_id.get(student_id)
                if not student:
                    results['errors'].append(f'Student {student_id} not found')
                    continue
                if student_id in already_processed:
                    results['errors'].append(f'{student["name"]} already has a promotion decision for {academic_year}')
                    continue
                if from_class != student['class']:
                    results['errors'].append(
                        f'{student["name"]} is currently in {student["class"]}, expected {from_class}'
                    )
                    continue
                if not to_class:
                    results['errors'].append(f'Promotion target missing for {student["name"]}')
                    continue
                if to_class != from_class and to_class not in valid_classes:
                    results['errors'].append(f'Target class {to_class} does not exist for {student["name"]}')
                    continue

            if results['errors']:
                results['failed'] = len(results['errors'])
                self._write_promotion_audit_entry(
                    batch_id,
                    'BATCH_REJECTED',
                    '; '.join(results['errors']),
                    actor_id,
                    now,
                )
                conn.close()
                return False, 'Batch promotion rejected during validation', results

            cursor.execute('BEGIN IMMEDIATE')
            self._insert_promotion_audit_entry(
                cursor,
                batch_id,
                'BATCH_START',
                f'Starting batch promotion for {len(promotions)} students in {academic_year}',
                actor_id,
                now,
            )

            for promo in promotions:
                student_id = str(promo.get('student_id', '')).strip()
                from_class = str(promo.get('from_class', '')).strip()
                to_class = str(promo.get('to_class', '')).strip()
                status = str(promo.get('status', 'promoted')).strip().lower() or 'promoted'
                reason = str(promo.get('reason', '')).strip()
                student = students_by_id[student_id]

                cursor.execute(
                    'UPDATE students SET class = ?, updated_at = ? WHERE id = ?',
                    (to_class, now, student_id)
                )
                cursor.execute(
                    '''INSERT INTO promotion_history
                       (id, student_id, from_class, to_class, promotion_date, academic_year, status, reason, created_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                    (str(uuid.uuid4()), student_id, from_class, to_class, now, academic_year, status, reason, now)
                )

                if status == 'promoted':
                    results['promoted'] += 1
                else:
                    results['repeating'] += 1

                detail = (
                    f'{student["name"]} ({student.get("admission_no", "")}): '
                    f'{from_class} -> {to_class} [{status}] {reason}'
                ).strip()
                self._insert_promotion_audit_entry(
                    cursor,
                    batch_id,
                    'STUDENT_PROCESSED',
                    detail,
                    actor_id,
                    now,
                )

            self._insert_promotion_audit_entry(
                cursor,
                batch_id,
                'BATCH_COMPLETE',
                f'Completed: {results["promoted"]} promoted, {results["repeating"]} repeating, 0 failed',
                actor_id,
                now,
            )
            conn.commit()
            conn.close()
            message = (
                f'Batch promotion completed: {results["promoted"]} promoted, '
                f'{results["repeating"]} repeating'
            )
            return True, message, results
        except Exception as e:
            conn.rollback()
            conn.close()
            results['failed'] = len(promotions)
            results['errors'].append(str(e))
            self._write_promotion_audit_entry(
                batch_id,
                'BATCH_FAILED',
                str(e),
                performed_by,
                now,
            )
            return False, f'Batch promotion failed: {str(e)}', results

    def get_promotion_history(self, student_id: str = None, academic_year: str = None,
                             class_name: str = None) -> List[Dict]:
        """Get promotion history with optional filters."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = '''
            SELECT ph.*, s.name as student_name, s.admission_no
            FROM promotion_history ph
            JOIN students s ON ph.student_id = s.id
            WHERE 1 = 1
        '''
        params = []
        
        if student_id:
            query += ' AND ph.student_id = ?'
            params.append(student_id)
        if academic_year:
            query += ' AND ph.academic_year = ?'
            params.append(academic_year)
        if class_name:
            query += ' AND (ph.from_class = ? OR ph.to_class = ?)'
            params.extend([class_name, class_name])
        
        query += ' ORDER BY ph.promotion_date DESC, s.name'
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_promotion_audit_log(self, batch_id: str = None, limit: int = 100) -> List[Dict]:
        """Get promotion audit log entries."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        if batch_id:
            cursor.execute(
                '''SELECT pal.*, u.full_name as performed_by_name
                   FROM promotion_audit_log pal
                   LEFT JOIN users u ON pal.performed_by = u.id
                   WHERE pal.promotion_batch_id = ?
                   ORDER BY pal.performed_at DESC''',
                (batch_id,)
            )
        else:
            cursor.execute(
                '''SELECT pal.*, u.full_name as performed_by_name
                   FROM promotion_audit_log pal
                   LEFT JOIN users u ON pal.performed_by = u.id
                   ORDER BY pal.performed_at DESC
                   LIMIT ?''',
                (limit,)
            )
        
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_promotion_statistics(self, academic_year: str = None) -> Dict:
        """Get promotion statistics for an academic year."""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        if academic_year:
            cursor.execute(
                '''SELECT 
                       COUNT(*) as total_promotions,
                       COALESCE(SUM(CASE WHEN status = 'promoted' THEN 1 ELSE 0 END), 0) as promoted_count,
                       COALESCE(SUM(CASE WHEN status = 'repeating' THEN 1 ELSE 0 END), 0) as repeating_count
                   FROM promotion_history
                   WHERE academic_year = ?''',
                (academic_year,)
            )
        else:
            cursor.execute(
                '''SELECT 
                       COUNT(*) as total_promotions,
                       COALESCE(SUM(CASE WHEN status = 'promoted' THEN 1 ELSE 0 END), 0) as promoted_count,
                       COALESCE(SUM(CASE WHEN status = 'repeating' THEN 1 ELSE 0 END), 0) as repeating_count
                   FROM promotion_history'''
            )
        
        row = cursor.fetchone()
        conn.close()
        
        return dict(row) if row else {
            'total_promotions': 0,
            'promoted_count': 0,
            'repeating_count': 0
        }


# Global database instance
db = Database()
