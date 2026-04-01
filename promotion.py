"""
Student promotion orchestration for the School Report Management System.
"""

from datetime import date, datetime
from typing import Callable, Dict, List, Optional, Tuple

from database import db


class PromotionManager:
    """Manages student promotions with validation, previewing, and audit support."""

    DEFAULT_SETTINGS = {
        'promotion_date': '12-01',
        'min_passing_average': '50.0',
        'auto_promote_enabled': 'false',
        'promotion_academic_year': '',
        'require_manual_approval': 'true',
        'notify_on_promotion': 'false',
    }

    def __init__(self, db_instance=None, today_provider: Optional[Callable[[], date]] = None):
        self.db = db_instance or db
        self.today_provider = today_provider or date.today
        self._initialize_settings()

    def _initialize_settings(self) -> None:
        for key, default_value in self.DEFAULT_SETTINGS.items():
            current_value = self.db.get_promotion_setting(key, '')
            if current_value == '':
                self.db.set_promotion_setting(key, default_value)

    def _today(self) -> date:
        return self.today_provider()

    def _parse_bool(self, value, default: bool = False) -> bool:
        if value is None:
            return default
        return str(value).strip().lower() in ('1', 'true', 'yes', 'on')

    def _validate_promotion_date(self, value: str) -> Tuple[bool, str, str]:
        raw_value = str(value or '').strip()
        try:
            month, day = map(int, raw_value.split('-'))
            normalized = f'{month:02d}-{day:02d}'
            date(2024, month, day)
            return True, '', normalized
        except (TypeError, ValueError):
            return False, 'Promotion date must be in MM-DD format.', raw_value

    def _validate_min_average(self, value: str) -> Tuple[bool, str, str]:
        raw_value = str(value or '').strip()
        try:
            numeric = float(raw_value)
        except (TypeError, ValueError):
            return False, 'Minimum passing average must be a number.', raw_value

        if numeric < 0 or numeric > 100:
            return False, 'Minimum passing average must be between 0 and 100.', raw_value
        return True, '', f'{numeric:.1f}'

    def _validate_academic_year(self, value: str) -> Tuple[bool, str, str]:
        raw_value = str(value or '').strip()
        if not raw_value:
            return True, '', ''
        parts = raw_value.split('/')
        if len(parts) != 2 or not all(part.isdigit() and len(part) == 4 for part in parts):
            return False, 'Academic year must look like YYYY/YYYY.', raw_value
        start_year, end_year = map(int, parts)
        if end_year != start_year + 1:
            return False, 'Academic year must span consecutive years.', raw_value
        return True, '', f'{start_year}/{end_year}'

    def get_settings(self) -> Dict[str, str]:
        settings = {}
        for key, default_value in self.DEFAULT_SETTINGS.items():
            current_value = self.db.get_promotion_setting(key, default_value)
            settings[key] = current_value if current_value != '' else default_value
        return settings

    def update_settings(self, settings: Dict[str, str]) -> Tuple[bool, str]:
        normalized = {}

        if 'promotion_date' in settings:
            valid, message, value = self._validate_promotion_date(settings.get('promotion_date', ''))
            if not valid:
                return False, message
            normalized['promotion_date'] = value

        if 'min_passing_average' in settings:
            valid, message, value = self._validate_min_average(settings.get('min_passing_average', ''))
            if not valid:
                return False, message
            normalized['min_passing_average'] = value

        if 'promotion_academic_year' in settings:
            valid, message, value = self._validate_academic_year(settings.get('promotion_academic_year', ''))
            if not valid:
                return False, message
            normalized['promotion_academic_year'] = value

        for key in ('auto_promote_enabled', 'require_manual_approval', 'notify_on_promotion'):
            if key in settings:
                normalized[key] = 'true' if self._parse_bool(settings.get(key)) else 'false'

        try:
            for key, value in normalized.items():
                if key in self.DEFAULT_SETTINGS:
                    self.db.set_promotion_setting(key, value)
            return True, 'Promotion settings saved successfully.'
        except Exception as exc:
            return False, f'Unable to save promotion settings: {exc}'

    def is_promotion_due(self) -> bool:
        settings = self.get_settings()
        valid, _, promotion_date_str = self._validate_promotion_date(settings.get('promotion_date', '12-01'))
        if not valid:
            return False

        month, day = map(int, promotion_date_str.split('-'))
        today = self._today()
        return today >= date(today.year, month, day)

    def get_current_academic_year(self) -> str:
        settings = self.get_settings()
        valid, _, override = self._validate_academic_year(settings.get('promotion_academic_year', ''))
        if valid and override:
            return override

        valid_date, _, promotion_date_str = self._validate_promotion_date(settings.get('promotion_date', '12-01'))
        today = self._today()
        if not valid_date:
            return f'{today.year}/{today.year + 1}'

        month, day = map(int, promotion_date_str.split('-'))
        promotion_date = date(today.year, month, day)
        if today < promotion_date:
            return f'{today.year - 1}/{today.year}'
        return f'{today.year}/{today.year + 1}'

    def _get_processed_lookup(self, academic_year: str) -> Dict[str, Dict]:
        history = self.db.get_promotion_history(academic_year=academic_year)
        return {row['student_id']: row for row in history}

    def _attach_student_context(self, student: Dict, current_class: str, latest_exam: Optional[Dict]) -> Dict:
        record = dict(student)
        record['current_class'] = current_class
        record['exam_term'] = latest_exam.get('term') if latest_exam else ''
        record['exam_type'] = latest_exam.get('exam_type') if latest_exam else ''
        record['exam_recorded_at'] = latest_exam.get('recorded_at') if latest_exam else ''
        return record

    def get_promotion_preview(self, class_name: str = None) -> Dict:
        settings = self.get_settings()
        min_average = float(settings.get('min_passing_average', '50.0'))
        academic_year = self.get_current_academic_year()
        processed_lookup = self._get_processed_lookup(academic_year)

        preview = {
            'academic_year': academic_year,
            'promotion_due': self.is_promotion_due(),
            'promotion_date': settings.get('promotion_date', '12-01'),
            'min_passing_average': min_average,
            'eligible': [],
            'repeating': [],
            'no_data': [],
            'terminal': [],
            'already_processed': [],
            'class_summaries': {},
        }

        if class_name:
            selected = self.db.get_class_by_name(class_name)
            classes = [selected] if selected else []
        else:
            classes = self.db.get_all_classes()

        for cls in classes:
            cls_name = cls['name']
            next_class = self.db.get_next_class(cls_name)
            latest_exam = self.db.get_latest_exam_session_for_class(cls_name)
            all_students = self.db.get_students_by_class(cls_name)

            if latest_exam:
                passing_students = self.db.get_students_eligible_for_promotion(cls_name, min_average)
                repeating_students = self.db.get_students_repeating(cls_name, min_average)
                no_data_students = self.db.get_students_without_promotion_data(cls_name)
            else:
                passing_students = []
                repeating_students = []
                no_data_students = [dict(student) for student in all_students]

            class_processed = []
            class_eligible = []
            class_terminal = []
            class_repeating = []
            class_no_data = []

            for student in passing_students:
                enriched = self._attach_student_context(student, cls_name, latest_exam)
                history = processed_lookup.get(student['id'])
                if history:
                    enriched.update({
                        'status': history.get('status', ''),
                        'to_class': history.get('to_class', ''),
                        'promotion_date': history.get('promotion_date', ''),
                    })
                    class_processed.append(enriched)
                    continue

                if next_class:
                    enriched['next_class'] = next_class
                    class_eligible.append(enriched)
                else:
                    enriched['status'] = 'terminal'
                    enriched['reason'] = 'No next class is configured after this class.'
                    class_terminal.append(enriched)

            for student in repeating_students:
                enriched = self._attach_student_context(student, cls_name, latest_exam)
                history = processed_lookup.get(student['id'])
                if history:
                    enriched.update({
                        'status': history.get('status', ''),
                        'to_class': history.get('to_class', ''),
                        'promotion_date': history.get('promotion_date', ''),
                    })
                    class_processed.append(enriched)
                    continue
                class_repeating.append(enriched)

            for student in no_data_students:
                enriched = self._attach_student_context(student, cls_name, latest_exam)
                history = processed_lookup.get(student['id'])
                if history:
                    enriched.update({
                        'status': history.get('status', ''),
                        'to_class': history.get('to_class', ''),
                        'promotion_date': history.get('promotion_date', ''),
                    })
                    class_processed.append(enriched)
                    continue
                class_no_data.append(enriched)

            preview['eligible'].extend(class_eligible)
            preview['repeating'].extend(class_repeating)
            preview['no_data'].extend(class_no_data)
            preview['terminal'].extend(class_terminal)
            preview['already_processed'].extend(class_processed)
            preview['class_summaries'][cls_name] = {
                'total_students': len(all_students),
                'eligible_count': len(class_eligible),
                'repeating_count': len(class_repeating),
                'no_data_count': len(class_no_data),
                'terminal_count': len(class_terminal),
                'already_processed_count': len(class_processed),
                'next_class': next_class,
                'latest_exam_term': latest_exam.get('term') if latest_exam else '',
                'latest_exam_type': latest_exam.get('exam_type') if latest_exam else '',
            }

        return preview

    def _build_promotions_from_preview(self, preview: Dict) -> List[Dict]:
        min_average = float(preview.get('min_passing_average', 50.0))
        promotions = []

        for student in preview['eligible']:
            promotions.append({
                'student_id': student['id'],
                'from_class': student['current_class'],
                'to_class': student['next_class'],
                'status': 'promoted',
                'reason': (
                    f'Average: {float(student.get("average_marks") or 0):.1f}% '
                    f'in {student.get("exam_term", "")} {student.get("exam_type", "")}'
                ).strip(),
            })

        for student in preview['repeating']:
            promotions.append({
                'student_id': student['id'],
                'from_class': student['current_class'],
                'to_class': student['current_class'],
                'status': 'repeating',
                'reason': (
                    f'Average: {float(student.get("average_marks") or 0):.1f}% '
                    f'(below {min_average:.1f}%)'
                ),
            })

        return promotions

    def execute_promotion(self, class_name: str = None, performed_by: str = None,
                         dry_run: bool = False) -> Tuple[bool, str, Dict]:
        preview = self.get_promotion_preview(class_name)
        promotions = self._build_promotions_from_preview(preview)

        summary = {
            'academic_year': preview['academic_year'],
            'promoted': len(preview['eligible']),
            'repeating': len(preview['repeating']),
            'failed': 0,
            'total': len(promotions),
            'no_data': len(preview['no_data']),
            'terminal': len(preview['terminal']),
            'already_processed': len(preview['already_processed']),
            'class_summaries': preview['class_summaries'],
        }

        if not promotions:
            return True, 'No students are ready for batch promotion.', summary

        if dry_run:
            summary['dry_run'] = True
            return True, f'Dry run: {len(promotions)} students would be processed.', summary

        success, message, results = self.db.batch_promote_students(
            promotions,
            preview['academic_year'],
            performed_by,
        )
        results.update({
            'academic_year': preview['academic_year'],
            'no_data': len(preview['no_data']),
            'terminal': len(preview['terminal']),
            'already_processed': len(preview['already_processed']),
            'class_summaries': preview['class_summaries'],
        })
        return success, message, results

    def promote_single_student(self, student_id: str, to_class: str = None,
                              status: str = 'promoted', reason: str = '',
                              performed_by: str = None) -> Tuple[bool, str]:
        student = self.db.get_student(student_id)
        if not student:
            return False, 'Student not found.'

        academic_year = self.get_current_academic_year()
        if self.db.has_promotion_history(academic_year, student_id):
            return False, f'This student already has a promotion decision for {academic_year}.'

        from_class = student['class']
        normalized_status = str(status or 'promoted').strip().lower()

        if to_class is None:
            if normalized_status == 'promoted':
                to_class = self.db.get_next_class(from_class)
                if not to_class:
                    return False, f'No next class is configured after {from_class}.'
            else:
                to_class = from_class

        return self.db.promote_student(
            student_id,
            from_class,
            to_class,
            academic_year,
            normalized_status,
            reason,
            performed_by,
        )

    def get_promotion_history(self, student_id: str = None, academic_year: str = None,
                             class_name: str = None) -> List[Dict]:
        return self.db.get_promotion_history(student_id, academic_year, class_name)

    def get_promotion_audit_log(self, batch_id: str = None, limit: int = 100) -> List[Dict]:
        return self.db.get_promotion_audit_log(batch_id, limit)

    def get_promotion_statistics(self, academic_year: str = None) -> Dict:
        return self.db.get_promotion_statistics(academic_year)

    def check_and_execute_auto_promotion(self, performed_by: str = None) -> Tuple[bool, str, Dict]:
        settings = self.get_settings()
        if not self._parse_bool(settings.get('auto_promote_enabled', 'false')):
            return False, 'Auto-promotion is disabled.', {}

        if not self.is_promotion_due():
            return False, 'Promotion date has not been reached yet.', {}

        success, message, results = self.execute_promotion(
            class_name=None,
            performed_by=performed_by,
            dry_run=False,
        )

        if success and results.get('total', 0) == 0 and results.get('already_processed', 0) > 0:
            return False, f'Promotion decisions already exist for {results.get("academic_year", self.get_current_academic_year())}.', results
        return success, message, results


promotion_manager = PromotionManager()


def run_scheduled_promotion():
    print(f'[{datetime.now().isoformat()}] Starting scheduled promotion check...')

    success, message, results = promotion_manager.check_and_execute_auto_promotion(
        performed_by=None
    )

    if success:
        print(f'[{datetime.now().isoformat()}] Promotion completed successfully:')
        print(f'  - Promoted: {results.get("promoted", 0)}')
        print(f'  - Repeating: {results.get("repeating", 0)}')
        print(f'  - No data: {results.get("no_data", 0)}')
        print(f'  - Terminal classes: {results.get("terminal", 0)}')
        print(f'  - Already processed: {results.get("already_processed", 0)}')
    else:
        print(f'[{datetime.now().isoformat()}] Promotion not executed: {message}')

    return success, message, results


if __name__ == '__main__':
    run_scheduled_promotion()
