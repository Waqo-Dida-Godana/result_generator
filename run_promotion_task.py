#!/usr/bin/env python3
"""
Scheduled Promotion Task Runner
This script can be executed by Windows Task Scheduler, cron, or similar scheduling tools
to automatically run student promotions based on configured dates.

Usage:
    python run_promotion_task.py [--dry-run] [--class CLASS_NAME] [--verbose]

Examples:
    # Run automatic promotion check (respects settings)
    python run_promotion_task.py

    # Dry run to see what would happen
    python run_promotion_task.py --dry-run

    # Promote specific class only
    python run_promotion_task.py --class "Grade 1"

    # Verbose output
    python run_promotion_task.py --verbose
"""

import sys
import os
import argparse
from datetime import datetime

# Add the current directory to the path so we can import from the project
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from promotion import promotion_manager
from database import db


def setup_logging(verbose: bool = False):
    """Setup basic logging configuration."""
    import logging
    
    log_level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('promotion_task.log', mode='a')
        ]
    )
    return logging.getLogger(__name__)


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Run student promotion task'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Simulate promotion without making changes'
    )
    parser.add_argument(
        '--class',
        dest='class_name',
        help='Process only a specific class'
    )
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Enable verbose output'
    )
    parser.add_argument(
        '--force',
        action='store_true',
        help='Force promotion even if not due (bypasses date check)'
    )
    parser.add_argument(
        '--user',
        dest='user_id',
        help='User ID to attribute the promotion to'
    )
    
    return parser.parse_args()


def main():
    """Main entry point for the scheduled promotion task."""
    args = parse_arguments()
    logger = setup_logging(args.verbose)
    
    logger.info("=" * 60)
    logger.info("Student Promotion Task Started")
    logger.info(f"Timestamp: {datetime.now().isoformat()}")
    logger.info("=" * 60)
    
    try:
        # Get current settings
        settings = promotion_manager.get_settings()
        logger.info(f"Promotion Settings:")
        logger.info(f"  - Auto-promote enabled: {settings.get('auto_promote_enabled', 'false')}")
        logger.info(f"  - Promotion date: {settings.get('promotion_date', '12-01')}")
        logger.info(f"  - Min passing average: {settings.get('min_passing_average', '50.0')}%")
        logger.info(f"  - Current academic year: {promotion_manager.get_current_academic_year()}")
        
        # Check if promotion is due
        is_due = promotion_manager.is_promotion_due()
        logger.info(f"Promotion due: {is_due}")

        # Manual and dry-run invocations should still be usable before the trigger date.
        if args.class_name or args.dry_run or args.force:
            scope = args.class_name or 'all classes'
            logger.info(f"Processing manual promotion scope: {scope}")
            success, message, results = promotion_manager.execute_promotion(
                class_name=args.class_name,
                performed_by=args.user_id,
                dry_run=args.dry_run
            )
        else:
            if not is_due:
                logger.info("Promotion is not due yet. Use --force for a manual run.")
                return 0
            logger.info("Processing all classes")
            success, message, results = promotion_manager.check_and_execute_auto_promotion(
                performed_by=args.user_id
            )
        
        # Log results
        logger.info("-" * 60)
        logger.info(f"Result: {'SUCCESS' if success else 'FAILED'}")
        logger.info(f"Message: {message}")
        
        if results:
            logger.info(f"Statistics:")
            logger.info(f"  - Promoted: {results.get('promoted', 0)}")
            logger.info(f"  - Repeating: {results.get('repeating', 0)}")
            logger.info(f"  - Failed: {results.get('failed', 0)}")
            logger.info(f"  - Total: {results.get('total', 0)}")
            logger.info(f"  - No data: {results.get('no_data', 0)}")
            logger.info(f"  - Terminal classes: {results.get('terminal', 0)}")
            logger.info(f"  - Already processed: {results.get('already_processed', 0)}")
            if results.get('batch_id'):
                logger.info(f"  - Batch ID: {results.get('batch_id')}")
            
            if results.get('errors'):
                logger.warning(f"Errors encountered:")
                for error in results['errors']:
                    logger.warning(f"  - {error}")
        
        logger.info("=" * 60)
        logger.info("Student Promotion Task Completed")
        logger.info("=" * 60)
        
        return 0 if success else 1
        
    except Exception as e:
        logger.error(f"Fatal error during promotion task: {e}", exc_info=True)
        return 1


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
