from datetime import datetime, timedelta
import os
import sys

# Add the backend directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from connectwise_report.config.settings import IGNITE_COMPANY_ID, OUTPUT_DIR
from connectwise_report.utils.api import get_time_entries
from connectwise_report.reports.html_report import HTMLReport
from connectwise_report.reports.word_report import WordReport

def generate_customer_report(output_format='docx'):
    """Generate the main customer report"""
    # Create output directory
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Calculate date range
    current_date = datetime.now()
    # Get last week's Monday
    days_since_monday = current_date.weekday()
    start_date = current_date - timedelta(days=days_since_monday + 7)  # Go back an additional week
    # Set time to start of day (midnight)
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = start_date + timedelta(days=4)  # Go to Friday of that week
    end_date = end_date.replace(hour=23, minute=59, second=59)
    
    print(f"\nScript running on: {current_date.strftime('%A, %B %d, %Y')}")
    print(f"\nGenerating report for period:")
    print(f"Start Date: {start_date.strftime('%A, %B %d, %Y')}")
    print(f"End Date: {end_date.strftime('%A, %B %d, %Y')}")
    
    try:
        # Get time entries
        time_entries = get_time_entries(start_date, end_date, IGNITE_COMPANY_ID)
        if not time_entries:
            print("No time entries found for this period!")
            return
            
        # Filter out meetings
        time_entries = [entry for entry in time_entries if 'Meetings' not in entry.get('ticket_summary', '')]
        
        # Generate debug report
        html_report = HTMLReport()
        html_report.generate(time_entries, OUTPUT_DIR, IGNITE_COMPANY_ID)
        
        # Generate main report
        word_report = WordReport()
        return word_report.generate(time_entries, OUTPUT_DIR, IGNITE_COMPANY_ID)
        
    except Exception as e:
        print(f"Error generating report: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        report_file = generate_customer_report()
        print(f"\nScript completed successfully!")
    except Exception as e:
        print(f"\nScript failed: {str(e)}")
        exit(1)