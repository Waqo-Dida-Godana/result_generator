"""
Exam Analytics Framework for School Report Management System
Comprehensive analytical framework to compare different types of exams
conducted over a year by measuring and analyzing their deviations.
"""

import numpy as np
from datetime import datetime, timedelta
from importlib import import_module
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass, field
from enum import Enum
import statistics
from database import db


class DeviationMetric(Enum):
    """Enumeration of deviation metrics for exam comparison."""
    SCORE_VARIANCE = "score_variance"
    PASS_RATE_FLUCTUATION = "pass_rate_fluctuation"
    DIFFICULTY_INDEX_SHIFT = "difficulty_index_shift"
    MEAN_SCORE_DEVIATION = "mean_score_deviation"
    STANDARD_DEVIATION = "standard_deviation"
    GRADE_DISTRIBUTION_SHIFT = "grade_distribution_shift"
    PERFORMANCE_CONSISTENCY = "performance_consistency"


@dataclass
class ExamSession:
    """Represents a single exam session with metadata."""
    term: str
    exam_type: str
    class_name: str
    subject: str
    date: Optional[datetime] = None
    student_count: int = 0
    mean_score: float = 0.0
    median_score: float = 0.0
    std_deviation: float = 0.0
    pass_rate: float = 0.0
    difficulty_index: float = 0.0
    grade_distribution: Dict[str, int] = field(default_factory=dict)
    scores: List[float] = field(default_factory=list)


@dataclass
class DeviationAnalysis:
    """Results of deviation analysis between exam sessions."""
    metric: DeviationMetric
    value: float
    interpretation: str
    severity: str  # 'low', 'medium', 'high', 'critical'
    details: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ExamComparison:
    """Comparison results between two or more exam sessions."""
    exam_sessions: List[ExamSession]
    deviations: List[DeviationAnalysis]
    overall_similarity_score: float  # 0-100, higher = more similar
    anomalies: List[str]
    patterns: List[str]
    recommendations: List[str]
    anova_results: Dict[str, Any] = field(default_factory=dict)
    time_series_results: Dict[str, Any] = field(default_factory=dict)
    exam_type_summaries: List[Dict[str, Any]] = field(default_factory=list)
    subject_deviation_rows: List[Dict[str, Any]] = field(default_factory=list)


class ExamAnalytics:
    """
    Comprehensive analytical framework for comparing exam types and measuring deviations.
    
    This class provides methods to:
    1. Calculate various deviation metrics
    2. Perform statistical comparisons between exam sessions
    3. Identify patterns and anomalies
    4. Generate visualizations and reports
    """
    
    # Thresholds for deviation severity
    SEVERITY_THRESHOLDS = {
        DeviationMetric.SCORE_VARIANCE: {'low': 100, 'medium': 250, 'high': 500},
        DeviationMetric.PASS_RATE_FLUCTUATION: {'low': 5, 'medium': 15, 'high': 30},
        DeviationMetric.DIFFICULTY_INDEX_SHIFT: {'low': 0.1, 'medium': 0.25, 'high': 0.5},
        DeviationMetric.MEAN_SCORE_DEVIATION: {'low': 5, 'medium': 10, 'high': 20},
        DeviationMetric.STANDARD_DEVIATION: {'low': 5, 'medium': 10, 'high': 15},
        DeviationMetric.GRADE_DISTRIBUTION_SHIFT: {'low': 10, 'medium': 25, 'high': 50},
        DeviationMetric.PERFORMANCE_CONSISTENCY: {'low': 0.8, 'medium': 0.6, 'high': 0.4},
    }
    EXAM_TYPE_ORDER = {
        'Opener': 1,
        'Quiz': 2,
        'Assignment': 3,
        'Mid-Term': 4,
        'CAT': 5,
        'End-Term': 6,
        'Final': 7,
    }
    
    def __init__(self):
        self.db = db
    
    # ── Data Retrieval Methods ──────────────────────────────────────────────
    
    def get_exam_sessions(self, class_name: str = None, term: str = None,
                         exam_type: str = None, subject: str = None,
                         start_date: datetime = None, end_date: datetime = None) -> List[ExamSession]:
        """
        Retrieve exam sessions from the database with optional filters.
        
        Args:
            class_name: Filter by class name
            term: Filter by term
            exam_type: Filter by exam type
            subject: Filter by subject
            start_date: Filter by start date
            end_date: Filter by end date
        
        Returns:
            List of ExamSession objects
        """
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        query = '''
            SELECT
                m.term,
                m.exam_type,
                s.class,
                m.subject,
                MAX(COALESCE(NULLIF(m.updated_at, ''), m.created_at)) AS recorded_at
            FROM marks m
            JOIN students s ON m.student_id = s.id
            WHERE 1 = 1
        '''
        params = []
        
        if class_name:
            query += ' AND s.class = ?'
            params.append(class_name)
        if term:
            query += ' AND m.term = ?'
            params.append(term)
        if exam_type:
            query += ' AND m.exam_type = ?'
            params.append(exam_type)
        if subject:
            query += ' AND m.subject = ?'
            params.append(subject)
        if start_date:
            query += " AND COALESCE(NULLIF(m.updated_at, ''), m.created_at) >= ?"
            params.append(start_date.isoformat())
        if end_date:
            query += " AND COALESCE(NULLIF(m.updated_at, ''), m.created_at) <= ?"
            params.append(end_date.isoformat())
        
        query += '''
            GROUP BY m.term, m.exam_type, s.class, m.subject
            ORDER BY recorded_at DESC, m.term, m.exam_type, s.class, m.subject
        '''
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        
        sessions = []
        for row in rows:
            session = self._build_exam_session(
                row['term'],
                row['exam_type'],
                row['class'],
                row['subject'],
                row['recorded_at'],
            )
            if session:
                sessions.append(session)
        
        conn.close()
        return sessions
    
    def _sort_exam_types(self, exam_types: List[str]) -> List[str]:
        return sorted(
            exam_types,
            key=lambda exam_name: (
                self.EXAM_TYPE_ORDER.get(str(exam_name or '').strip(), 999),
                str(exam_name or '').strip().lower(),
            )
        )

    def _parse_datetime(self, value: Any) -> Optional[datetime]:
        raw_value = str(value or '').strip()
        if not raw_value:
            return None
        try:
            return datetime.fromisoformat(raw_value.replace('Z', '+00:00'))
        except ValueError:
            return None

    def _aggregate_session_scores(self, sessions: List[ExamSession]) -> Dict[str, float]:
        all_scores = [score for session in sessions for score in session.scores]
        mean_scores = [session.mean_score for session in sessions if session.mean_score > 0]

        if not all_scores:
            return {
                'student_count': 0,
                'mean_score': 0.0,
                'median_score': 0.0,
                'std_deviation': 0.0,
                'pass_rate': 0.0,
                'difficulty_index': 0.0,
                'score_variance': 0.0,
            }

        return {
            'student_count': len(all_scores),
            'mean_score': statistics.mean(all_scores),
            'median_score': statistics.median(all_scores),
            'std_deviation': statistics.stdev(all_scores) if len(all_scores) > 1 else 0.0,
            'pass_rate': (sum(1 for score in all_scores if score >= 50) / len(all_scores)) * 100,
            'difficulty_index': statistics.mean(all_scores) / 100.0,
            'score_variance': statistics.variance(mean_scores) if len(mean_scores) > 1 else 0.0,
        }

    def _grade_distribution_from_scores(self, scores: List[float]) -> Dict[str, int]:
        distribution = {'EE': 0, 'ME': 0, 'AE': 0, 'BE': 0, 'IE': 0}
        for score in scores:
            if score >= 80:
                distribution['EE'] += 1
            elif score >= 70:
                distribution['ME'] += 1
            elif score >= 60:
                distribution['AE'] += 1
            elif score >= 50:
                distribution['BE'] += 1
            else:
                distribution['IE'] += 1
        return distribution

    def _build_exam_session(self, term: str, exam_type: str,
                           class_name: str, subject: str,
                           recorded_at: Any = None) -> Optional[ExamSession]:
        """Build an ExamSession object with calculated statistics."""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Get all marks for this exam session
        cursor.execute('''
            SELECT m.marks, m.created_at
            FROM marks m
            JOIN students s ON m.student_id = s.id
            WHERE s.class = ? AND m.term = ? AND m.exam_type = ? AND m.subject = ?
        ''', (class_name, term, exam_type, subject))
        
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            return None
        
        scores = [row['marks'] for row in rows]
        row_dates = [self._parse_datetime(row['created_at']) for row in rows]
        row_dates = [parsed for parsed in row_dates if parsed]
        
        # Parse date
        exam_date = self._parse_datetime(recorded_at)
        if exam_date is None and row_dates:
            exam_date = max(row_dates)
        
        # Calculate statistics
        session = ExamSession(
            term=term,
            exam_type=exam_type,
            class_name=class_name,
            subject=subject,
            date=exam_date,
            student_count=len(scores),
            scores=scores
        )
        
        if scores:
            session.mean_score = statistics.mean(scores)
            session.median_score = statistics.median(scores)
            session.std_deviation = statistics.stdev(scores) if len(scores) > 1 else 0.0
            session.pass_rate = (sum(1 for s in scores if s >= 50) / len(scores)) * 100
            session.difficulty_index = session.mean_score / 100.0  # Normalized to 0-1
            session.grade_distribution = self._grade_distribution_from_scores(scores)
        
        return session

    def summarize_exam_types(self, sessions: List[ExamSession]) -> List[Dict[str, Any]]:
        """Aggregate performance statistics by exam type."""
        grouped = {}
        for session in sessions:
            grouped.setdefault(session.exam_type, []).append(session)

        summaries = []
        for exam_type, grouped_sessions in grouped.items():
            aggregated = self._aggregate_session_scores(grouped_sessions)
            summaries.append({
                'exam_type': exam_type,
                'session_count': len(grouped_sessions),
                'subject_count': len({s.subject for s in grouped_sessions}),
                'class_count': len({s.class_name for s in grouped_sessions}),
                **aggregated,
            })

        return sorted(
            summaries,
            key=lambda row: (
                self.EXAM_TYPE_ORDER.get(row['exam_type'], 999),
                str(row['exam_type']).lower(),
            )
        )

    def calculate_subject_deviation_rows(self, sessions: List[ExamSession]) -> List[Dict[str, Any]]:
        """
        Create a subject-by-subject deviation matrix similar to a manual opener vs mid-term table.

        The first available exam type in the configured order is treated as the baseline, and every
        later exam type is compared back to that baseline for each subject.
        """
        exam_types = self._sort_exam_types(sorted({session.exam_type for session in sessions}))
        if len(exam_types) < 2:
            return []

        baseline_exam_type = exam_types[0]
        grouped = {}
        for session in sessions:
            grouped.setdefault((session.subject, session.exam_type), []).append(session)

        rows = []
        subjects = sorted({session.subject for session in sessions})
        for subject in subjects:
            baseline_sessions = grouped.get((subject, baseline_exam_type), [])
            if not baseline_sessions:
                continue

            baseline_stats = self._aggregate_session_scores(baseline_sessions)
            for comparison_exam_type in exam_types[1:]:
                comparison_sessions = grouped.get((subject, comparison_exam_type), [])
                if not comparison_sessions:
                    continue

                comparison_stats = self._aggregate_session_scores(comparison_sessions)
                score_deviation = comparison_stats['mean_score'] - baseline_stats['mean_score']
                pass_rate_deviation = comparison_stats['pass_rate'] - baseline_stats['pass_rate']
                difficulty_shift = comparison_stats['difficulty_index'] - baseline_stats['difficulty_index']
                deviation_percent = (
                    (score_deviation / baseline_stats['mean_score']) * 100
                    if baseline_stats['mean_score'] else 0.0
                )

                rows.append({
                    'subject': subject,
                    'baseline_exam_type': baseline_exam_type,
                    'comparison_exam_type': comparison_exam_type,
                    'baseline_mean': baseline_stats['mean_score'],
                    'comparison_mean': comparison_stats['mean_score'],
                    'score_deviation': score_deviation,
                    'deviation_percent': deviation_percent,
                    'baseline_pass_rate': baseline_stats['pass_rate'],
                    'comparison_pass_rate': comparison_stats['pass_rate'],
                    'pass_rate_deviation': pass_rate_deviation,
                    'difficulty_shift': difficulty_shift,
                    'baseline_student_count': baseline_stats['student_count'],
                    'comparison_student_count': comparison_stats['student_count'],
                    'severity': self._classify_severity(
                        DeviationMetric.MEAN_SCORE_DEVIATION,
                        abs(score_deviation),
                    ),
                })

        return sorted(
            rows,
            key=lambda row: (
                -abs(row['score_deviation']),
                row['subject'].lower(),
                self.EXAM_TYPE_ORDER.get(row['comparison_exam_type'], 999),
            )
        )
    
    # ── Deviation Metrics Calculation ───────────────────────────────────────
    
    def calculate_score_variance(self, sessions: List[ExamSession]) -> DeviationAnalysis:
        """
        Calculate score variance across exam sessions.
        
        High variance indicates inconsistent difficulty or grading standards.
        """
        if len(sessions) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.SCORE_VARIANCE,
                value=0.0,
                interpretation="Insufficient data for variance calculation",
                severity="low"
            )
        
        # Calculate variance of mean scores across sessions
        mean_scores = [s.mean_score for s in sessions if s.mean_score > 0]
        if not mean_scores:
            return DeviationAnalysis(
                metric=DeviationMetric.SCORE_VARIANCE,
                value=0.0,
                interpretation="No valid mean scores available",
                severity="low"
            )
        
        variance = statistics.variance(mean_scores) if len(mean_scores) > 1 else 0.0
        severity = self._classify_severity(DeviationMetric.SCORE_VARIANCE, variance)
        
        interpretation = f"Score variance of {variance:.2f} across {len(sessions)} exam sessions. "
        if severity == 'low':
            interpretation += "Scores are consistent across sessions."
        elif severity == 'medium':
            interpretation += "Moderate variation in scores detected."
        elif severity == 'high':
            interpretation += "Significant variation in scores - review exam difficulty."
        else:
            interpretation += "Critical variation - immediate review required."
        
        return DeviationAnalysis(
            metric=DeviationMetric.SCORE_VARIANCE,
            value=variance,
            interpretation=interpretation,
            severity=severity,
            details={'mean_scores': mean_scores, 'session_count': len(sessions)}
        )
    
    def calculate_pass_rate_fluctuation(self, sessions: List[ExamSession]) -> DeviationAnalysis:
        """
        Calculate pass rate fluctuations across exam sessions.
        
        Large fluctuations may indicate inconsistent teaching quality or exam difficulty.
        """
        if len(sessions) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.PASS_RATE_FLUCTUATION,
                value=0.0,
                interpretation="Insufficient data for pass rate analysis",
                severity="low"
            )
        
        pass_rates = [s.pass_rate for s in sessions]
        max_rate = max(pass_rates)
        min_rate = min(pass_rates)
        fluctuation = max_rate - min_rate
        
        severity = self._classify_severity(DeviationMetric.PASS_RATE_FLUCTUATION, fluctuation)
        
        interpretation = f"Pass rate fluctuation of {fluctuation:.1f}% (range: {min_rate:.1f}% - {max_rate:.1f}%). "
        if severity == 'low':
            interpretation += "Pass rates are stable."
        elif severity == 'medium':
            interpretation += "Moderate fluctuation in pass rates."
        elif severity == 'high':
            interpretation += "Significant pass rate variation detected."
        else:
            interpretation += "Critical pass rate instability - investigate causes."
        
        return DeviationAnalysis(
            metric=DeviationMetric.PASS_RATE_FLUCTUATION,
            value=fluctuation,
            interpretation=interpretation,
            severity=severity,
            details={'pass_rates': pass_rates, 'max': max_rate, 'min': min_rate}
        )
    
    def calculate_difficulty_index_shift(self, sessions: List[ExamSession]) -> DeviationAnalysis:
        """
        Calculate difficulty index shifts across exam sessions.
        
        Difficulty index = mean_score / 100. Values closer to 1 indicate easier exams.
        """
        if len(sessions) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.DIFFICULTY_INDEX_SHIFT,
                value=0.0,
                interpretation="Insufficient data for difficulty analysis",
                severity="low"
            )
        
        difficulty_indices = [s.difficulty_index for s in sessions]
        max_diff = max(difficulty_indices)
        min_diff = min(difficulty_indices)
        shift = max_diff - min_diff
        
        severity = self._classify_severity(DeviationMetric.DIFFICULTY_INDEX_SHIFT, shift)
        
        interpretation = f"Difficulty index shift of {shift:.3f} (range: {min_diff:.3f} - {max_diff:.3f}). "
        if severity == 'low':
            interpretation += "Exam difficulty is consistent."
        elif severity == 'medium':
            interpretation += "Moderate difficulty variation detected."
        elif severity == 'high':
            interpretation += "Significant difficulty variation - review exam standards."
        else:
            interpretation += "Critical difficulty inconsistency - standardization needed."
        
        return DeviationAnalysis(
            metric=DeviationMetric.DIFFICULTY_INDEX_SHIFT,
            value=shift,
            interpretation=interpretation,
            severity=severity,
            details={'difficulty_indices': difficulty_indices, 'max': max_diff, 'min': min_diff}
        )
    
    def calculate_mean_score_deviation(self, sessions: List[ExamSession]) -> DeviationAnalysis:
        """
        Calculate mean score deviation from overall average.
        
        Identifies sessions that deviate significantly from the norm.
        """
        if len(sessions) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.MEAN_SCORE_DEVIATION,
                value=0.0,
                interpretation="Insufficient data for mean score analysis",
                severity="low"
            )
        
        mean_scores = [s.mean_score for s in sessions if s.mean_score > 0]
        if not mean_scores:
            return DeviationAnalysis(
                metric=DeviationMetric.MEAN_SCORE_DEVIATION,
                value=0.0,
                interpretation="No valid mean scores available",
                severity="low"
            )
        
        overall_mean = statistics.mean(mean_scores)
        deviations = [abs(score - overall_mean) for score in mean_scores]
        max_deviation = max(deviations)
        
        severity = self._classify_severity(DeviationMetric.MEAN_SCORE_DEVIATION, max_deviation)
        
        interpretation = f"Maximum mean score deviation of {max_deviation:.1f} from overall average of {overall_mean:.1f}. "
        if severity == 'low':
            interpretation += "Scores are close to the average."
        elif severity == 'medium':
            interpretation += "Some sessions deviate from the average."
        elif severity == 'high':
            interpretation += "Significant deviations detected - review outlier sessions."
        else:
            interpretation += "Critical deviations - immediate investigation required."
        
        return DeviationAnalysis(
            metric=DeviationMetric.MEAN_SCORE_DEVIATION,
            value=max_deviation,
            interpretation=interpretation,
            severity=severity,
            details={'overall_mean': overall_mean, 'deviations': deviations, 'mean_scores': mean_scores}
        )
    
    def calculate_grade_distribution_shift(self, sessions: List[ExamSession]) -> DeviationAnalysis:
        """
        Calculate grade distribution shifts across exam sessions.
        
        Measures how much the distribution of grades changes between sessions.
        """
        if len(sessions) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.GRADE_DISTRIBUTION_SHIFT,
                value=0.0,
                interpretation="Insufficient data for grade distribution analysis",
                severity="low"
            )
        
        # Normalize distributions to percentages
        normalized_dists = []
        for session in sessions:
            total = sum(session.grade_distribution.values())
            if total > 0:
                normalized = {k: (v / total) * 100 for k, v in session.grade_distribution.items()}
                normalized_dists.append(normalized)
        
        if len(normalized_dists) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.GRADE_DISTRIBUTION_SHIFT,
                value=0.0,
                interpretation="Insufficient grade distribution data",
                severity="low"
            )
        
        # Calculate average shift across all grade categories
        grade_codes = ['EE', 'ME', 'AE', 'BE', 'IE']
        shifts = []
        
        for grade in grade_codes:
            values = [dist.get(grade, 0) for dist in normalized_dists]
            if values:
                max_val = max(values)
                min_val = min(values)
                shifts.append(max_val - min_val)
        
        avg_shift = statistics.mean(shifts) if shifts else 0.0
        severity = self._classify_severity(DeviationMetric.GRADE_DISTRIBUTION_SHIFT, avg_shift)
        
        interpretation = f"Average grade distribution shift of {avg_shift:.1f}%. "
        if severity == 'low':
            interpretation += "Grade distributions are consistent."
        elif severity == 'medium':
            interpretation += "Moderate grade distribution changes detected."
        elif severity == 'high':
            interpretation += "Significant grade distribution shifts - review grading standards."
        else:
            interpretation += "Critical grade distribution inconsistency - standardization required."
        
        return DeviationAnalysis(
            metric=DeviationMetric.GRADE_DISTRIBUTION_SHIFT,
            value=avg_shift,
            interpretation=interpretation,
            severity=severity,
            details={'shifts': shifts, 'grade_codes': grade_codes}
        )
    
    def calculate_performance_consistency(self, sessions: List[ExamSession]) -> DeviationAnalysis:
        """
        Calculate performance consistency score across exam sessions.
        
        Higher scores indicate more consistent performance patterns.
        """
        if len(sessions) < 2:
            return DeviationAnalysis(
                metric=DeviationMetric.PERFORMANCE_CONSISTENCY,
                value=1.0,
                interpretation="Insufficient data for consistency analysis",
                severity="low"
            )
        
        # Calculate coefficient of variation for each session
        cv_values = []
        for session in sessions:
            if session.mean_score > 0 and session.std_deviation > 0:
                cv = session.std_deviation / session.mean_score
                cv_values.append(cv)
        
        if not cv_values:
            return DeviationAnalysis(
                metric=DeviationMetric.PERFORMANCE_CONSISTENCY,
                value=1.0,
                interpretation="No valid coefficient of variation data",
                severity="low"
            )
        
        # Consistency score (inverse of average CV, normalized to 0-1)
        avg_cv = statistics.mean(cv_values)
        consistency_score = max(0, 1 - avg_cv)  # Lower CV = higher consistency
        
        severity = self._classify_severity(DeviationMetric.PERFORMANCE_CONSISTENCY, consistency_score)
        
        interpretation = f"Performance consistency score of {consistency_score:.3f}. "
        if severity == 'low':
            interpretation += "Performance is highly consistent."
        elif severity == 'medium':
            interpretation += "Moderate consistency in performance."
        elif severity == 'high':
            interpretation += "Low consistency - investigate performance variations."
        else:
            interpretation += "Critical inconsistency - immediate review required."
        
        return DeviationAnalysis(
            metric=DeviationMetric.PERFORMANCE_CONSISTENCY,
            value=consistency_score,
            interpretation=interpretation,
            severity=severity,
            details={'cv_values': cv_values, 'avg_cv': avg_cv}
        )
    
    # ── Statistical Comparison Methods ──────────────────────────────────────
    
    def perform_anova(self, sessions: List[ExamSession]) -> Dict[str, Any]:
        """
        Perform one-way ANOVA to test if there are significant differences
        between exam session mean scores.
        
        Returns:
            Dictionary with F-statistic, p-value, and interpretation
        """
        if len(sessions) < 2:
            return {
                'f_statistic': 0.0,
                'p_value': 1.0,
                'significant': False,
                'interpretation': "Insufficient data for ANOVA"
            }
        
        # Extract scores for each session
        session_scores = [s.scores for s in sessions if s.scores]
        
        if len(session_scores) < 2:
            return {
                'f_statistic': 0.0,
                'p_value': 1.0,
                'significant': False,
                'interpretation': "Insufficient score data for ANOVA"
            }
        
        # Perform one-way ANOVA
        try:
            stats = import_module('scipy.stats')
            f_statistic, p_value = stats.f_oneway(*session_scores)
            significant = p_value < 0.05
            
            interpretation = f"One-way ANOVA: F={f_statistic:.3f}, p={p_value:.4f}. "
            if significant:
                interpretation += "Significant differences detected between exam sessions (p < 0.05)."
            else:
                interpretation += "No significant differences between exam sessions (p >= 0.05)."
            
            return {
                'f_statistic': f_statistic,
                'p_value': p_value,
                'significant': significant,
                'interpretation': interpretation
            }
        except (ImportError, ModuleNotFoundError):
            # Fallback: manual ANOVA calculation when scipy is not available
            return self._manual_anova(session_scores)
    
    def _manual_anova(self, session_scores: List[List[float]]) -> Dict[str, Any]:
        """Manual ANOVA calculation when scipy is not available."""
        k = len(session_scores)  # Number of groups
        n_total = sum(len(scores) for scores in session_scores)  # Total observations
        
        if k < 2 or n_total < k:
            return {
                'f_statistic': 0.0,
                'p_value': 1.0,
                'significant': False,
                'interpretation': "Insufficient data for ANOVA"
            }
        
        # Calculate group means and overall mean
        group_means = [statistics.mean(scores) for scores in session_scores]
        all_scores = [score for scores in session_scores for score in scores]
        overall_mean = statistics.mean(all_scores)
        
        # Calculate sum of squares
        ss_between = sum(len(scores) * (mean - overall_mean) ** 2
                        for scores, mean in zip(session_scores, group_means))
        
        ss_within = sum((score - group_means[i]) ** 2
                       for i, scores in enumerate(session_scores)
                       for score in scores)
        
        # Degrees of freedom
        df_between = k - 1
        df_within = n_total - k
        
        if df_within == 0:
            return {
                'f_statistic': 0.0,
                'p_value': 1.0,
                'significant': False,
                'interpretation': "Insufficient degrees of freedom for ANOVA"
            }
        
        # Mean squares
        ms_between = ss_between / df_between
        ms_within = ss_within / df_within
        
        # F-statistic
        f_statistic = ms_between / ms_within if ms_within > 0 else 0.0
        
        # Approximate p-value using F-distribution
        # This is a simplified approximation
        p_value = 1.0 / (1.0 + f_statistic) if f_statistic > 0 else 1.0
        significant = p_value < 0.05
        
        interpretation = f"One-way ANOVA (manual): F={f_statistic:.3f}, p≈{p_value:.4f}. "
        if significant:
            interpretation += "Significant differences detected between exam sessions."
        else:
            interpretation += "No significant differences between exam sessions."
        
        return {
            'f_statistic': f_statistic,
            'p_value': p_value,
            'significant': significant,
            'interpretation': interpretation
        }
    
    def perform_time_series_analysis(self, sessions: List[ExamSession]) -> Dict[str, Any]:
        """
        Perform time-series analysis on exam sessions to identify trends.
        
        Returns:
            Dictionary with trend analysis results
        """
        if len(sessions) < 3:
            return {
                'trend': 'insufficient_data',
                'slope': 0.0,
                'r_squared': 0.0,
                'interpretation': "Insufficient data for time-series analysis"
            }
        
        # Sort sessions by date
        sorted_sessions = sorted([s for s in sessions if s.date], key=lambda x: x.date)
        
        if len(sorted_sessions) < 3:
            return {
                'trend': 'insufficient_data',
                'slope': 0.0,
                'r_squared': 0.0,
                'interpretation': "Insufficient dated sessions for time-series analysis"
            }
        
        # Prepare data for linear regression
        x_values = [(s.date - sorted_sessions[0].date).days for s in sorted_sessions]
        y_values = [s.mean_score for s in sorted_sessions]
        
        # Perform linear regression
        n = len(x_values)
        sum_x = sum(x_values)
        sum_y = sum(y_values)
        sum_xy = sum(x * y for x, y in zip(x_values, y_values))
        sum_x2 = sum(x ** 2 for x in x_values)
        sum_y2 = sum(y ** 2 for y in y_values)
        
        # Calculate slope and intercept
        denominator = n * sum_x2 - sum_x ** 2
        if denominator == 0:
            return {
                'trend': 'no_variation',
                'slope': 0.0,
                'r_squared': 0.0,
                'interpretation': "No variation in time points"
            }
        
        slope = (n * sum_xy - sum_x * sum_y) / denominator
        intercept = (sum_y - slope * sum_x) / n
        
        # Calculate R-squared
        y_mean = sum_y / n
        ss_total = sum((y - y_mean) ** 2 for y in y_values)
        ss_residual = sum((y - (slope * x + intercept)) ** 2
                         for x, y in zip(x_values, y_values))
        
        r_squared = 1 - (ss_residual / ss_total) if ss_total > 0 else 0.0
        
        # Determine trend
        if abs(slope) < 0.1:
            trend = 'stable'
            interpretation = "Scores are stable over time."
        elif slope > 0:
            trend = 'improving'
            interpretation = f"Scores are improving over time (slope: {slope:.3f} per day)."
        else:
            trend = 'declining'
            interpretation = f"Scores are declining over time (slope: {slope:.3f} per day)."
        
        interpretation += f" R² = {r_squared:.3f}, indicating {'strong' if r_squared > 0.7 else 'moderate' if r_squared > 0.4 else 'weak'} correlation."
        
        return {
            'trend': trend,
            'slope': slope,
            'intercept': intercept,
            'r_squared': r_squared,
            'interpretation': interpretation,
            'data_points': n,
            'date_range': f"{sorted_sessions[0].date.strftime('%Y-%m-%d')} to {sorted_sessions[-1].date.strftime('%Y-%m-%d')}"
        }
    
    # ── Comprehensive Comparison ────────────────────────────────────────────
    
    def compare_exam_sessions(self, sessions: List[ExamSession]) -> ExamComparison:
        """
        Perform comprehensive comparison of exam sessions.
        
        Returns:
            ExamComparison object with all analysis results
        """
        exam_type_summaries = self.summarize_exam_types(sessions)
        subject_deviation_rows = self.calculate_subject_deviation_rows(sessions)

        if len(sessions) < 2:
            return ExamComparison(
                exam_sessions=sessions,
                deviations=[],
                overall_similarity_score=100.0,
                anomalies=[],
                patterns=["Insufficient data for comparison"],
                recommendations=["Collect more exam data for meaningful analysis"],
                exam_type_summaries=exam_type_summaries,
                subject_deviation_rows=subject_deviation_rows,
            )
        
        # Calculate all deviation metrics
        deviations = [
            self.calculate_score_variance(sessions),
            self.calculate_pass_rate_fluctuation(sessions),
            self.calculate_difficulty_index_shift(sessions),
            self.calculate_mean_score_deviation(sessions),
            self.calculate_grade_distribution_shift(sessions),
            self.calculate_performance_consistency(sessions),
        ]
        
        # Perform statistical tests
        anova_results = self.perform_anova(sessions)
        time_series_results = self.perform_time_series_analysis(sessions)
        
        # Calculate overall similarity score (0-100)
        similarity_score = self._calculate_similarity_score(deviations)
        
        # Identify anomalies
        anomalies = self._identify_anomalies(sessions, deviations)
        
        # Identify patterns
        patterns = self._identify_patterns(sessions, time_series_results)
        
        # Generate recommendations
        recommendations = self._generate_recommendations(
            deviations,
            anova_results,
            time_series_results,
            subject_deviation_rows,
        )
        
        return ExamComparison(
            exam_sessions=sessions,
            deviations=deviations,
            overall_similarity_score=similarity_score,
            anomalies=anomalies,
            patterns=patterns,
            recommendations=recommendations,
            anova_results=anova_results,
            time_series_results=time_series_results,
            exam_type_summaries=exam_type_summaries,
            subject_deviation_rows=subject_deviation_rows,
        )
    
    def _calculate_similarity_score(self, deviations: List[DeviationAnalysis]) -> float:
        """Calculate overall similarity score from deviation metrics."""
        if not deviations:
            return 100.0
        
        # Weight different metrics
        weights = {
            DeviationMetric.SCORE_VARIANCE: 0.2,
            DeviationMetric.PASS_RATE_FLUCTUATION: 0.2,
            DeviationMetric.DIFFICULTY_INDEX_SHIFT: 0.15,
            DeviationMetric.MEAN_SCORE_DEVIATION: 0.2,
            DeviationMetric.GRADE_DISTRIBUTION_SHIFT: 0.15,
            DeviationMetric.PERFORMANCE_CONSISTENCY: 0.1,
        }
        
        # Normalize each deviation to 0-100 scale (lower deviation = higher score)
        scores = []
        for dev in deviations:
            if dev.metric == DeviationMetric.PERFORMANCE_CONSISTENCY:
                # Higher consistency = higher score
                score = dev.value * 100
            else:
                # Lower deviation = higher score
                max_threshold = self.SEVERITY_THRESHOLDS[dev.metric]['high']
                score = max(0, 100 - (dev.value / max_threshold) * 100)
            
            weight = weights.get(dev.metric, 0.1)
            scores.append(score * weight)
        
        return sum(scores) / sum(weights.get(d.metric, 0.1) for d in deviations)
    
    def _identify_anomalies(self, sessions: List[ExamSession],
                           deviations: List[DeviationAnalysis]) -> List[str]:
        """Identify anomalies in exam sessions."""
        anomalies = []
        
        # Check for critical deviations
        for dev in deviations:
            if dev.severity == 'critical':
                anomalies.append(f"Critical {dev.metric.value}: {dev.interpretation}")
        
        # Check for outlier sessions
        if len(sessions) >= 3:
            mean_scores = [s.mean_score for s in sessions if s.mean_score > 0]
            if mean_scores:
                overall_mean = statistics.mean(mean_scores)
                overall_std = statistics.stdev(mean_scores) if len(mean_scores) > 1 else 0
                
                for session in sessions:
                    if session.mean_score > 0 and overall_std > 0:
                        z_score = abs(session.mean_score - overall_mean) / overall_std
                        if z_score > 2:
                            anomalies.append(
                                f"Outlier session: {session.term} {session.exam_type} "
                                f"{session.subject} (z-score: {z_score:.2f})"
                            )

        for row in self.calculate_subject_deviation_rows(sessions)[:5]:
            if abs(row['score_deviation']) >= 10:
                direction = 'higher' if row['score_deviation'] > 0 else 'lower'
                anomalies.append(
                    f"{row['subject']} is {abs(row['score_deviation']):.1f} points {direction} in "
                    f"{row['comparison_exam_type']} than {row['baseline_exam_type']}"
                )
        
        return anomalies
    
    def _identify_patterns(self, sessions: List[ExamSession],
                          time_series_results: Dict[str, Any]) -> List[str]:
        """Identify patterns in exam sessions."""
        patterns = []
        exam_type_summaries = self.summarize_exam_types(sessions)
        subject_deviation_rows = self.calculate_subject_deviation_rows(sessions)
        
        # Time-based patterns
        trend = time_series_results.get('trend', 'unknown')
        if trend == 'improving':
            patterns.append("Overall improving trend in scores over time")
        elif trend == 'declining':
            patterns.append("Overall declining trend in scores over time")
        elif trend == 'stable':
            patterns.append("Stable performance over time")
        
        # Exam type patterns
        if len(exam_type_summaries) > 1:
            patterns.append(
                "Exam type performance ranking: "
                + ', '.join(
                    f"{item['exam_type']} ({item['mean_score']:.1f})"
                    for item in sorted(exam_type_summaries, key=lambda row: row['mean_score'], reverse=True)
                )
            )

        # Subject patterns
        subjects = set(s.subject for s in sessions)
        if len(subjects) > 1:
            subject_means = {}
            for subject in subjects:
                subject_sessions = [s for s in sessions if s.subject == subject]
                if subject_sessions:
                    subject_means[subject] = statistics.mean([s.mean_score for s in subject_sessions])
            
            if subject_means:
                sorted_subjects = sorted(subject_means.items(), key=lambda x: x[1], reverse=True)
                patterns.append(f"Subject performance ranking: {', '.join(f'{s[0]} ({s[1]:.1f})' for s in sorted_subjects[:5])}")

        if subject_deviation_rows:
            largest_gain = max(subject_deviation_rows, key=lambda row: row['score_deviation'])
            largest_drop = min(subject_deviation_rows, key=lambda row: row['score_deviation'])
            patterns.append(
                f"Largest positive exam-type shift: {largest_gain['subject']} "
                f"({largest_gain['baseline_exam_type']} -> {largest_gain['comparison_exam_type']}, "
                f"{largest_gain['score_deviation']:+.1f})"
            )
            patterns.append(
                f"Largest negative exam-type shift: {largest_drop['subject']} "
                f"({largest_drop['baseline_exam_type']} -> {largest_drop['comparison_exam_type']}, "
                f"{largest_drop['score_deviation']:+.1f})"
            )
        
        return patterns
    
    def _generate_recommendations(self, deviations: List[DeviationAnalysis],
                                 anova_results: Dict[str, Any],
                                 time_series_results: Dict[str, Any],
                                 subject_deviation_rows: Optional[List[Dict[str, Any]]] = None) -> List[str]:
        """Generate recommendations based on analysis results."""
        recommendations = []
        subject_deviation_rows = subject_deviation_rows or []
        
        # Check for critical deviations
        critical_deviations = [d for d in deviations if d.severity == 'critical']
        if critical_deviations:
            recommendations.append("Immediate review required for critical deviations in exam metrics")
        
        # Check for high deviations
        high_deviations = [d for d in deviations if d.severity == 'high']
        if high_deviations:
            recommendations.append("Review exam standards and grading criteria for high deviation metrics")
        
        # ANOVA recommendations
        if anova_results.get('significant', False):
            recommendations.append("Significant differences detected between exam sessions - standardize exam difficulty")
        
        # Time-series recommendations
        trend = time_series_results.get('trend', 'unknown')
        if trend == 'declining':
            recommendations.append("Declining trend detected - investigate causes and implement interventions")
        elif trend == 'improving':
            recommendations.append("Improving trend detected - identify and replicate successful practices")
        
        large_subject_shifts = [row for row in subject_deviation_rows if abs(row['score_deviation']) >= 10]
        if large_subject_shifts:
            recommendations.append(
                "Review subjects with large exam-type shifts to confirm difficulty balance and marking consistency"
            )
        
        # General recommendations
        if not recommendations:
            recommendations.append("Exam performance is within acceptable parameters")
        
        return recommendations
    
    # ── Utility Methods ─────────────────────────────────────────────────────
    
    def _classify_severity(self, metric: DeviationMetric, value: float) -> str:
        """Classify the severity of a deviation metric."""
        thresholds = self.SEVERITY_THRESHOLDS.get(metric, {})
        
        if not thresholds:
            return 'low'
        
        if metric == DeviationMetric.PERFORMANCE_CONSISTENCY:
            # Higher value = better for consistency
            if value >= thresholds['low']:
                return 'low'
            elif value >= thresholds['medium']:
                return 'medium'
            elif value >= thresholds['high']:
                return 'high'
            else:
                return 'critical'
        else:
            # Lower value = better for other metrics
            if value <= thresholds['low']:
                return 'low'
            elif value <= thresholds['medium']:
                return 'medium'
            elif value <= thresholds['high']:
                return 'high'
            else:
                return 'critical'
    
    def get_available_exam_sessions(self) -> List[Dict[str, str]]:
        """Get all available exam sessions for analysis."""
        return self.db.get_available_exam_sessions()
    
    def generate_comparison_report(self, comparison: ExamComparison) -> str:
        """Generate a text-based comparison report."""
        report = []
        report.append("=" * 80)
        report.append("EXAM COMPARISON ANALYSIS REPORT")
        report.append("=" * 80)
        report.append(f"\nAnalysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"Number of Exam Sessions: {len(comparison.exam_sessions)}")
        report.append(f"Overall Similarity Score: {comparison.overall_similarity_score:.1f}/100")
        
        report.append("\n" + "-" * 80)
        report.append("DEVIATION METRICS")
        report.append("-" * 80)
        
        for dev in comparison.deviations:
            report.append(f"\n{dev.metric.value.upper().replace('_', ' ')}:")
            report.append(f"  Value: {dev.value:.3f}")
            report.append(f"  Severity: {dev.severity.upper()}")
            report.append(f"  Interpretation: {dev.interpretation}")

        if comparison.anova_results:
            report.append("\n" + "-" * 80)
            report.append("STATISTICAL TESTS")
            report.append("-" * 80)
            report.append(f"ANOVA: {comparison.anova_results.get('interpretation', 'N/A')}")
        if comparison.time_series_results:
            report.append(f"Time Series: {comparison.time_series_results.get('interpretation', 'N/A')}")

        if comparison.exam_type_summaries:
            report.append("\n" + "-" * 80)
            report.append("EXAM TYPE SUMMARY")
            report.append("-" * 80)
            for summary in comparison.exam_type_summaries:
                report.append(
                    f"{summary['exam_type']}: mean={summary['mean_score']:.1f}, "
                    f"pass={summary['pass_rate']:.1f}%, difficulty={summary['difficulty_index']:.3f}, "
                    f"sessions={summary['session_count']}, subjects={summary['subject_count']}"
                )

        if comparison.subject_deviation_rows:
            report.append("\n" + "-" * 80)
            report.append("SUBJECT DEVIATION MATRIX")
            report.append("-" * 80)
            for row in comparison.subject_deviation_rows[:20]:
                report.append(
                    f"{row['subject']}: {row['baseline_exam_type']} {row['baseline_mean']:.1f} -> "
                    f"{row['comparison_exam_type']} {row['comparison_mean']:.1f} "
                    f"(dev {row['score_deviation']:+.1f}, pass {row['pass_rate_deviation']:+.1f}%, "
                    f"difficulty {row['difficulty_shift']:+.3f})"
                )
        
        if comparison.anomalies:
            report.append("\n" + "-" * 80)
            report.append("ANOMALIES DETECTED")
            report.append("-" * 80)
            for anomaly in comparison.anomalies:
                report.append(f"• {anomaly}")
        
        if comparison.patterns:
            report.append("\n" + "-" * 80)
            report.append("PATTERNS IDENTIFIED")
            report.append("-" * 80)
            for pattern in comparison.patterns:
                report.append(f"• {pattern}")
        
        if comparison.recommendations:
            report.append("\n" + "-" * 80)
            report.append("RECOMMENDATIONS")
            report.append("-" * 80)
            for i, rec in enumerate(comparison.recommendations, 1):
                report.append(f"{i}. {rec}")
        
        report.append("\n" + "=" * 80)
        report.append("END OF REPORT")
        report.append("=" * 80)
        
        return "\n".join(report)


# ── Visualization Functions ──────────────────────────────────────────────

def create_deviation_chart(comparison: ExamComparison, output_path: str = None) -> str:
    """
    Create a bar chart showing deviation metrics.
    
    Args:
        comparison: ExamComparison object with analysis results
        output_path: Optional path to save the chart image
    
    Returns:
        Path to the saved chart image or base64 encoded string
    """
    try:
        import matplotlib.pyplot as plt
        import matplotlib
        matplotlib.use('Agg')
        
        metrics = [d.metric.value.replace('_', ' ').title() for d in comparison.deviations]
        values = [d.value for d in comparison.deviations]
        severities = [d.severity for d in comparison.deviations]
        
        # Color mapping for severities
        severity_colors = {
            'low': '#2ecc71',      # Green
            'medium': '#f39c12',   # Orange
            'high': '#e74c3c',     # Red
            'critical': '#c0392b'  # Dark red
        }
        colors = [severity_colors.get(s, '#95a5a6') for s in severities]
        
        fig, ax = plt.subplots(figsize=(12, 6))
        bars = ax.barh(metrics, values, color=colors, edgecolor='white', linewidth=1)
        
        # Add value labels on bars
        for bar, value, severity in zip(bars, values, severities):
            ax.text(bar.get_width() + max(values) * 0.01, bar.get_y() + bar.get_height()/2,
                   f'{value:.2f} ({severity})', va='center', fontsize=9)
        
        ax.set_xlabel('Deviation Value', fontsize=11, fontweight='bold')
        ax.set_title('Exam Deviation Metrics Analysis', fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='x', alpha=0.3, linestyle='--')
        
        # Add legend
        from matplotlib.patches import Patch
        legend_elements = [Patch(facecolor=color, label=severity.capitalize())
                          for severity, color in severity_colors.items()]
        ax.legend(handles=legend_elements, loc='lower right', title='Severity')
        
        plt.tight_layout()
        
        if output_path:
            plt.savefig(output_path, dpi=150, bbox_inches='tight')
            plt.close()
            return output_path
        else:
            # Return base64 encoded image
            import io
            import base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buffer.seek(0)
            return base64.b64encode(buffer.getvalue()).decode('utf-8')
    except ImportError:
        return "Matplotlib not available for visualization"


def create_time_series_chart(sessions: List[ExamSession], output_path: str = None) -> str:
    """
    Create a time-series chart showing score trends over time.
    
    Args:
        sessions: List of ExamSession objects
        output_path: Optional path to save the chart image
    
    Returns:
        Path to the saved chart image or base64 encoded string
    """
    try:
        import matplotlib.pyplot as plt
        import matplotlib
        matplotlib.use('Agg')
        
        # Filter sessions with dates
        dated_sessions = [s for s in sessions if s.date]
        if len(dated_sessions) < 2:
            return "Insufficient data for time-series chart"
        
        # Sort by date
        dated_sessions.sort(key=lambda x: x.date)
        
        dates = [s.date for s in dated_sessions]
        mean_scores = [s.mean_score for s in dated_sessions]
        pass_rates = [s.pass_rate for s in dated_sessions]
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8), sharex=True)
        
        # Plot 1: Mean Scores
        ax1.plot(dates, mean_scores, marker='o', linewidth=2, markersize=8, color='#3498db')
        ax1.fill_between(dates, mean_scores, alpha=0.3, color='#3498db')
        ax1.set_ylabel('Mean Score', fontsize=11, fontweight='bold')
        ax1.set_title('Exam Performance Trends Over Time', fontsize=14, fontweight='bold', pad=20)
        ax1.grid(True, alpha=0.3, linestyle='--')
        ax1.axhline(y=50, color='#e74c3c', linestyle='--', linewidth=1, label='Pass Threshold (50%)')
        ax1.legend()
        
        # Plot 2: Pass Rates
        ax2.plot(dates, pass_rates, marker='s', linewidth=2, markersize=8, color='#2ecc71')
        ax2.fill_between(dates, pass_rates, alpha=0.3, color='#2ecc71')
        ax2.set_xlabel('Date', fontsize=11, fontweight='bold')
        ax2.set_ylabel('Pass Rate (%)', fontsize=11, fontweight='bold')
        ax2.grid(True, alpha=0.3, linestyle='--')
        ax2.axhline(y=80, color='#f39c12', linestyle='--', linewidth=1, label='Target (80%)')
        ax2.legend()
        
        # Format x-axis
        import matplotlib.dates as mdates
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        ax2.xaxis.set_major_locator(mdates.AutoDateLocator())
        plt.xticks(rotation=45, ha='right')
        
        plt.tight_layout()
        
        if output_path:
            plt.savefig(output_path, dpi=150, bbox_inches='tight')
            plt.close()
            return output_path
        else:
            import io
            import base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buffer.seek(0)
            return base64.b64encode(buffer.getvalue()).decode('utf-8')
    except ImportError:
        return "Matplotlib not available for visualization"


def create_grade_distribution_chart(sessions: List[ExamSession], output_path: str = None) -> str:
    """
    Create a stacked bar chart showing grade distribution across sessions.
    
    Args:
        sessions: List of ExamSession objects
        output_path: Optional path to save the chart image
    
    Returns:
        Path to the saved chart image or base64 encoded string
    """
    try:
        import matplotlib.pyplot as plt
        import matplotlib
        matplotlib.use('Agg')
        import numpy as np
        
        if len(sessions) < 2:
            return "Insufficient data for grade distribution chart"
        
        # Prepare data
        labels = [f"{s.term}\n{s.exam_type}" for s in sessions]
        grade_codes = ['EE', 'ME', 'AE', 'BE', 'IE']
        grade_colors = ['#2ecc71', '#3498db', '#f39c12', '#e67e22', '#e74c3c']
        
        # Calculate percentages
        data = {grade: [] for grade in grade_codes}
        for session in sessions:
            total = sum(session.grade_distribution.values())
            if total > 0:
                for grade in grade_codes:
                    percentage = (session.grade_distribution.get(grade, 0) / total) * 100
                    data[grade].append(percentage)
            else:
                for grade in grade_codes:
                    data[grade].append(0)
        
        # Create stacked bar chart
        fig, ax = plt.subplots(figsize=(12, 6))
        x = np.arange(len(labels))
        width = 0.6
        
        bottom = np.zeros(len(labels))
        for grade, color in zip(grade_codes, grade_colors):
            ax.bar(x, data[grade], width, label=grade, bottom=bottom, color=color, edgecolor='white')
            bottom += np.array(data[grade])
        
        ax.set_xlabel('Exam Session', fontsize=11, fontweight='bold')
        ax.set_ylabel('Percentage (%)', fontsize=11, fontweight='bold')
        ax.set_title('Grade Distribution Across Exam Sessions', fontsize=14, fontweight='bold', pad=20)
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.legend(title='Grade', bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        
        plt.tight_layout()
        
        if output_path:
            plt.savefig(output_path, dpi=150, bbox_inches='tight')
            plt.close()
            return output_path
        else:
            import io
            import base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buffer.seek(0)
            return base64.b64encode(buffer.getvalue()).decode('utf-8')
    except ImportError:
        return "Matplotlib not available for visualization"


def create_subject_deviation_chart(comparison: ExamComparison, output_path: str = None) -> str:
    """
    Create a chart showing subject-by-subject score deviations between exam types.
    """
    try:
        import matplotlib.pyplot as plt
        import matplotlib
        matplotlib.use('Agg')

        rows = comparison.subject_deviation_rows[:12]
        if not rows:
            return "Insufficient data for subject deviation chart"

        labels = [
            f"{row['subject']}\n{row['baseline_exam_type']}->{row['comparison_exam_type']}"
            for row in rows
        ]
        deviations = [row['score_deviation'] for row in rows]
        colors = ['#2ecc71' if value >= 0 else '#e74c3c' for value in deviations]

        fig, ax = plt.subplots(figsize=(12, 7))
        bars = ax.barh(labels, deviations, color=colors, edgecolor='white', linewidth=1)
        ax.axvline(0, color='#555555', linewidth=1)
        ax.set_xlabel('Mean Score Deviation', fontsize=11, fontweight='bold')
        ax.set_title('Subject Deviation by Exam Type', fontsize=14, fontweight='bold', pad=20)
        ax.grid(axis='x', alpha=0.3, linestyle='--')

        for bar, value in zip(bars, deviations):
            ax.text(
                value + (0.3 if value >= 0 else -0.3),
                bar.get_y() + bar.get_height() / 2,
                f'{value:+.1f}',
                va='center',
                ha='left' if value >= 0 else 'right',
                fontsize=9,
            )

        plt.tight_layout()

        if output_path:
            plt.savefig(output_path, dpi=150, bbox_inches='tight')
            plt.close()
            return output_path
        else:
            import io
            import base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buffer.seek(0)
            return base64.b64encode(buffer.getvalue()).decode('utf-8')
    except ImportError:
        return "Matplotlib not available for visualization"


def create_comparison_dashboard(comparison: ExamComparison, output_dir: str = '.') -> Dict[str, str]:
    """
    Create a comprehensive dashboard with multiple charts.
    
    Args:
        comparison: ExamComparison object with analysis results
        output_dir: Directory to save chart images
    
    Returns:
        Dictionary with chart names and their file paths
    """
    import os
    os.makedirs(output_dir, exist_ok=True)
    
    charts = {}
    
    # Deviation chart
    deviation_path = os.path.join(output_dir, 'deviation_metrics.png')
    charts['deviation'] = create_deviation_chart(comparison, deviation_path)
    
    # Time-series chart
    if any(s.date for s in comparison.exam_sessions):
        time_series_path = os.path.join(output_dir, 'time_series_trends.png')
        charts['time_series'] = create_time_series_chart(comparison.exam_sessions, time_series_path)
    
    # Grade distribution chart
    grade_dist_path = os.path.join(output_dir, 'grade_distribution.png')
    charts['grade_distribution'] = create_grade_distribution_chart(comparison.exam_sessions, grade_dist_path)

    # Subject deviation chart
    if comparison.subject_deviation_rows:
        subject_deviation_path = os.path.join(output_dir, 'subject_deviations.png')
        charts['subject_deviation'] = create_subject_deviation_chart(comparison, subject_deviation_path)
    
    return charts


# Global analytics instance
exam_analytics = ExamAnalytics()
