"""
Database module for School Report Management System
Uses SQLite for local data storage
"""

import sqlite3
import uuid
from datetime import datetime
from typing import List, Dict, Optional, Tuple


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
                subject TEXT NOT NULL,
                marks INTEGER NOT NULL CHECK (marks >= 0 AND marks <= 100),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                UNIQUE(student_id, term, subject)
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
    
    def add_student(self, name: str, class_name: str, gender: str, admission_no: str, photo_path: str = "") -> Dict:
        conn = self.get_connection()
        cursor = conn.cursor()
        student_id = str(uuid.uuid4())
        now = datetime.now().isoformat()
        cursor.execute(
            'INSERT INTO students (id, name, class, gender, admission_no, photo_path, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            (student_id, name, class_name, gender, admission_no, photo_path, now, now)
        )
        conn.commit()
        conn.close()
        return {'id': student_id, 'name': name, 'class': class_name, 'gender': gender, 'admission_no': admission_no, 'photo_path': photo_path}
    
    def update_student(self, student_id: str, name: str, class_name: str, gender: str, admission_no: str, photo_path: str = "") -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute(
            'UPDATE students SET name = ?, class = ?, gender = ?, admission_no = ?, photo_path = ?, updated_at = ? WHERE id = ?',
            (name, class_name, gender, admission_no, photo_path, now, student_id)
        )
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        return success
    
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
            'SELECT * FROM students WHERE name LIKE ? OR admission_no LIKE ? OR class LIKE ? ORDER BY name',
            (search, search, search)
        )
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    # Marks operations
    def get_marks(self, term: str = 'One') -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM marks WHERE term = ?', (term,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    def get_student_marks(self, student_id: str, term: str = 'One') -> Dict[str, int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT subject, marks FROM marks WHERE student_id = ? AND term = ?', (student_id, term))
        rows = cursor.fetchall()
        conn.close()
        return {row['subject']: row['marks'] for row in rows}
    
    def save_student_marks(self, student_id: str, marks: Dict[str, int], term: str = 'One') -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        
        # Delete existing marks for this student and term
        cursor.execute('DELETE FROM marks WHERE student_id = ? AND term = ?', (student_id, term))
        
        # Insert new marks
        for subject, value in marks.items():
            if value is not None and value != '':
                cursor.execute(
                    'INSERT INTO marks (id, student_id, term, subject, marks, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)',
                    (str(uuid.uuid4()), student_id, term, subject, int(value), now, now)
                )
        
        conn.commit()
        conn.close()
        return True
    
    def save_all_marks(self, student_marks: Dict[str, Dict[str, int]], term: str = 'One') -> bool:
        for student_id, marks in student_marks.items():
            self.save_student_marks(student_id, marks, term)
        return True
    
    # Statistics
    def get_statistics(self, term: str = 'One') -> Dict:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Total students
        cursor.execute('SELECT COUNT(*) as count FROM students')
        total_students = cursor.fetchone()['count']
        
        # Get all marks for the term
        cursor.execute('''
            SELECT s.id, s.name, s.class, m.subject, m.marks 
            FROM students s 
            LEFT JOIN marks m ON s.id = m.student_id AND m.term = ?
        ''', (term,))
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
    def calculate_results(self, class_filter: str = 'All', term: str = 'One') -> List[Dict]:
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
            cursor.execute('SELECT subject, marks FROM marks WHERE student_id = ? AND term = ?', 
                          (student['id'], term))
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


# Global database instance
db = Database()
