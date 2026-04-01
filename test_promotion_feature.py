import os
import shutil
import unittest
import uuid
from datetime import date

from database import Database
from promotion import PromotionManager


class PromotionFeatureTests(unittest.TestCase):
    def setUp(self):
        base_tmp_dir = os.path.join(os.getcwd(), '.test_tmp')
        os.makedirs(base_tmp_dir, exist_ok=True)
        self.db_path = os.path.join(base_tmp_dir, f'promotion_test_{uuid.uuid4().hex}.db')
        self.db = Database(self.db_path)
        self.manager = PromotionManager(
            db_instance=self.db,
            today_provider=lambda: date(2026, 12, 5),
        )

        self.manager.update_settings({
            'promotion_date': '12-01',
            'min_passing_average': '50',
            'promotion_academic_year': '2026/2027',
            'auto_promote_enabled': 'true',
        })

        self.db.add_class('Grade 1', 'Lower Primary (Grade 1-3)')
        self.db.add_class('Grade 2', 'Lower Primary (Grade 1-3)')
        self.db.add_class('Grade 3', 'Lower Primary (Grade 1-3)')

        self.alice = self.db.add_student('Alice', 'Grade 1', 'Female', 'ADM001')
        self.bob = self.db.add_student('Bob', 'Grade 1', 'Male', 'ADM002')
        self.cara = self.db.add_student('Cara', 'Grade 1', 'Female', 'ADM003')

        self.db.save_student_marks(
            self.alice['id'],
            {'Mathematics': 80, 'English': 70},
            term='Three',
            exam_type='End-Term',
        )
        self.db.save_student_marks(
            self.bob['id'],
            {'Mathematics': 40, 'English': 30},
            term='Three',
            exam_type='End-Term',
        )

    def tearDown(self):
        if os.path.exists(self.db_path):
            os.remove(self.db_path)
        base_tmp_dir = os.path.dirname(self.db_path)
        if os.path.isdir(base_tmp_dir):
            shutil.rmtree(base_tmp_dir, ignore_errors=True)

    def test_preview_separates_promoted_repeating_and_no_data_students(self):
        preview = self.manager.get_promotion_preview('Grade 1')

        self.assertEqual(len(preview['eligible']), 1)
        self.assertEqual(len(preview['repeating']), 1)
        self.assertEqual(len(preview['no_data']), 1)
        self.assertEqual(preview['eligible'][0]['name'], 'Alice')
        self.assertEqual(preview['eligible'][0]['next_class'], 'Grade 2')
        self.assertEqual(preview['repeating'][0]['name'], 'Bob')
        self.assertEqual(preview['no_data'][0]['name'], 'Cara')

    def test_preview_marks_classes_without_any_exam_session_as_no_data(self):
        self.db.add_student('Dan', 'Grade 3', 'Male', 'ADM004')

        preview = self.manager.get_promotion_preview('Grade 3')

        self.assertEqual(len(preview['eligible']), 0)
        self.assertEqual(len(preview['repeating']), 0)
        self.assertEqual(len(preview['no_data']), 1)
        self.assertEqual(preview['no_data'][0]['name'], 'Dan')

    def test_execute_promotion_updates_students_history_and_audit_log(self):
        success, message, results = self.manager.execute_promotion('Grade 1')

        self.assertTrue(success, message)
        self.assertEqual(results['promoted'], 1)
        self.assertEqual(results['repeating'], 1)
        self.assertEqual(results['no_data'], 1)
        self.assertTrue(results.get('batch_id'))

        self.assertEqual(self.db.get_student(self.alice['id'])['class'], 'Grade 2')
        self.assertEqual(self.db.get_student(self.bob['id'])['class'], 'Grade 1')
        self.assertEqual(self.db.get_student(self.cara['id'])['class'], 'Grade 1')

        history = self.manager.get_promotion_history(academic_year='2026/2027')
        self.assertEqual(len(history), 2)

        audit_log = self.manager.get_promotion_audit_log(batch_id=results['batch_id'])
        actions = [row['action'] for row in audit_log]
        self.assertIn('BATCH_COMPLETE', actions)
        self.assertEqual(actions.count('STUDENT_PROCESSED'), 2)

    def test_promote_single_student_blocks_duplicate_yearly_decision(self):
        success, _, _ = self.manager.execute_promotion('Grade 1')
        self.assertTrue(success)

        second_success, second_message = self.manager.promote_single_student(
            self.bob['id'],
            status='repeating',
            reason='Second attempt should fail',
        )

        self.assertFalse(second_success)
        self.assertIn('already has a promotion decision', second_message)

    def test_batch_promotions_roll_back_when_validation_fails(self):
        eva = self.db.add_student('Eva', 'Grade 1', 'Female', 'ADM005')

        success, message, results = self.db.batch_promote_students(
            [
                {
                    'student_id': self.alice['id'],
                    'from_class': 'Grade 1',
                    'to_class': 'Grade 2',
                    'status': 'promoted',
                    'reason': 'Valid promotion',
                },
                {
                    'student_id': eva['id'],
                    'from_class': 'Grade 1',
                    'to_class': 'Grade 99',
                    'status': 'promoted',
                    'reason': 'Invalid target class',
                },
            ],
            academic_year='2026/2027',
        )

        self.assertFalse(success)
        self.assertIn('rejected', message.lower())
        self.assertEqual(self.db.get_student(self.alice['id'])['class'], 'Grade 1')
        self.assertEqual(self.db.get_student(eva['id'])['class'], 'Grade 1')
        self.assertEqual(len(self.manager.get_promotion_history(academic_year='2026/2027')), 0)
        self.assertGreaterEqual(results['failed'], 1)

    def test_auto_promotion_respects_configured_trigger_date(self):
        gated_manager = PromotionManager(
            db_instance=self.db,
            today_provider=lambda: date(2026, 11, 30),
        )
        gated_manager.update_settings({
            'promotion_date': '12-01',
            'auto_promote_enabled': 'true',
            'promotion_academic_year': '2026/2027',
        })

        success, message, _ = gated_manager.check_and_execute_auto_promotion()

        self.assertFalse(success)
        self.assertIn('not been reached', message)


if __name__ == '__main__':
    unittest.main()
