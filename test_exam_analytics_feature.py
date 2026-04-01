import os
import unittest
import uuid
from datetime import datetime

from database import Database
from exam_analytics import exam_analytics, ExamAnalytics


class ExamAnalyticsFeatureTests(unittest.TestCase):
    def setUp(self):
        base_tmp_dir = os.path.join(os.getcwd(), '.test_tmp')
        os.makedirs(base_tmp_dir, exist_ok=True)
        self.db_path = os.path.join(base_tmp_dir, f'exam_analytics_{uuid.uuid4().hex}.db')
        self.db = Database(self.db_path)
        self.analytics = ExamAnalytics()
        self.analytics.db = self.db

        self.db.add_class('Grade 7', 'Junior School (Grade 7-9)')
        self.student_one = self.db.add_student('Amina', 'Grade 7', 'Female', 'ANA001')
        self.student_two = self.db.add_student('Brian', 'Grade 7', 'Male', 'BRI001')

        self.db.save_student_marks(
            self.student_one['id'],
            {'Mathematics': 50, 'English': 80},
            term='One',
            exam_type='Opener',
        )
        self.db.save_student_marks(
            self.student_two['id'],
            {'Mathematics': 70, 'English': 60},
            term='One',
            exam_type='Opener',
        )

        self.db.save_student_marks(
            self.student_one['id'],
            {'Mathematics': 80, 'English': 60},
            term='One',
            exam_type='Mid-Term',
        )
        self.db.save_student_marks(
            self.student_two['id'],
            {'Mathematics': 90, 'English': 60},
            term='One',
            exam_type='Mid-Term',
        )

    def tearDown(self):
        if os.path.exists(self.db_path):
            os.remove(self.db_path)

    def test_get_exam_sessions_groups_rows_into_unique_sessions(self):
        sessions = self.analytics.get_exam_sessions(class_name='Grade 7', term='One')

        self.assertEqual(len(sessions), 4)
        session_keys = {(session.exam_type, session.subject) for session in sessions}
        self.assertEqual(
            session_keys,
            {
                ('Opener', 'Mathematics'),
                ('Opener', 'English'),
                ('Mid-Term', 'Mathematics'),
                ('Mid-Term', 'English'),
            },
        )

    def test_exam_type_summaries_and_subject_deviation_rows_match_expected_shifts(self):
        sessions = self.analytics.get_exam_sessions(class_name='Grade 7', term='One')
        comparison = self.analytics.compare_exam_sessions(sessions)

        self.assertEqual([row['exam_type'] for row in comparison.exam_type_summaries], ['Opener', 'Mid-Term'])
        opener_summary = comparison.exam_type_summaries[0]
        mid_summary = comparison.exam_type_summaries[1]
        self.assertAlmostEqual(opener_summary['mean_score'], 65.0)
        self.assertAlmostEqual(mid_summary['mean_score'], 72.5)

        math_row = next(row for row in comparison.subject_deviation_rows if row['subject'] == 'Mathematics')
        english_row = next(row for row in comparison.subject_deviation_rows if row['subject'] == 'English')
        self.assertEqual(math_row['baseline_exam_type'], 'Opener')
        self.assertEqual(math_row['comparison_exam_type'], 'Mid-Term')
        self.assertAlmostEqual(math_row['score_deviation'], 25.0)
        self.assertAlmostEqual(english_row['score_deviation'], -10.0)

    def test_date_filters_can_exclude_sessions(self):
        sessions = self.analytics.get_exam_sessions(
            class_name='Grade 7',
            start_date=datetime(2000, 1, 1),
            end_date=datetime(2000, 12, 31),
        )
        self.assertEqual(sessions, [])

    def test_report_includes_statistical_and_subject_deviation_sections(self):
        sessions = self.analytics.get_exam_sessions(class_name='Grade 7', term='One')
        comparison = self.analytics.compare_exam_sessions(sessions)
        report = self.analytics.generate_comparison_report(comparison)

        self.assertIn('STATISTICAL TESTS', report)
        self.assertIn('EXAM TYPE SUMMARY', report)
        self.assertIn('SUBJECT DEVIATION MATRIX', report)
        self.assertIn('Mathematics', report)


if __name__ == '__main__':
    unittest.main()
