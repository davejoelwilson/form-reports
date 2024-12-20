import os
from datetime import datetime
import pytz
import pdfkit
from connectwise_report.utils.formatting import format_detail

class HTMLReport:
    def __init__(self):
        # CSS styles for the HTML report
        self.css = """
            body{font-family:Calibri,sans-serif;margin:40px;background-color:#f5f5f5}
            .container{max-width:1200px;margin:0 auto;background-color:white;padding:20px;box-shadow:0 0 10px rgba(0,0,0,0.1);border-radius:5px}
            .header{background-color:#E5F3E2;padding:20px;margin-bottom:20px;border-radius:5px}
            .ticket{margin-bottom:30px;border:1px solid #ddd;border-radius:5px;overflow:hidden}
            .ticket-header{background-color:#E5F3E2;padding:15px;border-bottom:1px solid #ddd}
            .ticket-details{padding:15px;background-color:white}
            .time-entries{margin-top:15px}
            table{width:100%;border-collapse:collapse;margin-top:10px}
            th,td{padding:8px;text-align:left;border:1px solid #ddd}
            th{background-color:#f8f8f8}
            .totals{margin-top:15px;padding:10px;background-color:#f8f8f8;border-radius:3px}
            .meta-info{color:#666;font-size:0.9em}
        """

        # Add print-specific CSS
        self.css += """
            @media print {
                body { margin: 0; }
                .container { box-shadow: none; }
                .ticket { page-break-inside: avoid; }
                table { page-break-inside: auto; }
                tr { page-break-inside: avoid; page-break-after: auto; }
                thead { display: table-header-group; }
            }
        """

    def generate(self, time_entries, output_dir, company_id):
        """Generate both HTML and PDF reports"""
        tickets_data = self._process_entries(time_entries)
        html_content = self._generate_html(tickets_data, company_id)
        self._save_reports(html_content, output_dir, company_id)

    def _format_notes(self, notes):
        """Format notes to put each # item on its own line"""
        if not notes:
            return ""
            
        # Split by # but keep the # symbol
        parts = notes.split('#')
        formatted_lines = []
        
        for part in parts:
            part = part.strip()
            if part:  # Skip empty parts
                formatted_lines.append(f"# {part}")
        
        return '<br>'.join(formatted_lines)

    def _process_entries(self, time_entries):
        """Process time entries into ticket-grouped data"""
        tickets_data = {}
        
        def get_ticket_id(entry):
            ticket = entry.get('ticket', {})
            if isinstance(ticket, dict):
                return str(ticket.get('id', ''))  # Convert to string for consistent comparison
            return ''
        
        for entry in time_entries:
            ticket_id = get_ticket_id(entry)
            if not ticket_id:  # Skip entries without valid ticket IDs
                continue
            
            if ticket_id not in tickets_data:
                tickets_data[ticket_id] = {
                    'details': entry,
                    'entries': [],
                    'total_hours': 0,
                    'board': entry.get('ticketBoard', 'Unknown Board'),
                    'status': entry.get('ticketStatus', 'Unknown Status'),
                    'summary': entry.get('ticket', {}).get('summary', '')
                }
            
            # Convert UTC to NZ time for both start and end times
            start_date = datetime.fromisoformat(entry['timeStart'].replace('Z', '+00:00'))
            end_date = datetime.fromisoformat(entry['timeEnd'].replace('Z', '+00:00'))
            nz_start = start_date.astimezone(pytz.timezone('Pacific/Auckland'))
            nz_end = end_date.astimezone(pytz.timezone('Pacific/Auckland'))
            
            formatted_notes = format_detail(entry.get('notes', ''), entry.get('project', None))
            
            # Add entry details
            tickets_data[ticket_id]['entries'].append({
                'date': nz_start.strftime('%Y-%m-%d'),
                'start_time': nz_start.strftime('%I:%M %p'),
                'end_time': nz_end.strftime('%I:%M %p'),
                'hours': entry.get('actualHours', 0),
                'engineer': entry.get('member', {}).get('name'),
                'work_type': entry.get('workType', {}).get('name'),
                'notes': formatted_notes
            })
            tickets_data[ticket_id]['total_hours'] += entry.get('actualHours', 0)
        
        return tickets_data

    def _generate_html(self, tickets_data, company_id):
        """Generate the complete HTML document"""
        tickets_html = []
        for ticket_id, data in tickets_data.items():
            ticket_html = self._generate_ticket_html(ticket_id, data)
            tickets_html.append(ticket_html)

        return f"""<!DOCTYPE html>
            <html>
            <head>
                <title>ConnectWise Debug Report</title>
                <style>{self.css}</style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h1>ConnectWise Debug Report</h1>
                        <div class="meta-info">
                            <p>Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                            <p>Company ID: {company_id}</p>
                        </div>
                    </div>
                    {''.join(tickets_html)}
                </div>
            </body>
            </html>"""

    def _generate_ticket_html(self, ticket_id, data):
        """Generate HTML for a single ticket"""
        # Add ticket info section
        ticket_info = f"""
            <div class="ticket-info">
                <p><strong>Board:</strong> {data['board']}</p>
                <p><strong>Status:</strong> {data['status']}</p>
            </div>
        """

        entries_html = ''.join([
            f"""<tr>
                <td>{entry['date']}</td>
                <td>{entry['start_time']}</td>
                <td>{entry['end_time']}</td>
                <td>{entry['hours']}</td>
                <td>{entry['engineer']}</td>
                <td>{entry['work_type']}</td>
                <td>{entry['notes']}</td>
            </tr>""" for entry in data['entries']
        ])

        return f"""
            <div class="ticket">
                <div class="ticket-header">
                    <h2>Ticket #{ticket_id}</h2>
                    <p>{data['summary']}</p>
                    {ticket_info}
                </div>
                <div class="ticket-details">
                    <table>
                        <tr>
                            <th>Date</th>
                            <th>Start Time</th>
                            <th>End Time</th>
                            <th>Hours</th>
                            <th>Engineer</th>
                            <th>Work Type</th>
                            <th>Notes</th>
                        </tr>
                        {entries_html}
                    </table>
                    <div class="totals">
                        <p><strong>Total Hours:</strong> {data['total_hours']:.2f}</p>
                    </div>
                </div>
            </div>"""

    def _save_reports(self, html_content, output_dir, company_id):
        """Save both HTML and PDF versions of the report"""
        timestamp = datetime.now().strftime('%Y%m%d')
        
        # Save HTML
        html_file = os.path.join(output_dir, f"debug_report_{company_id}_{timestamp}.html")
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"HTML report saved to: {html_file}")
        
        # Try to save PDF if pdfkit is available
        try:
            import pdfkit
            pdf_file = os.path.join(output_dir, f"debug_report_{company_id}_{timestamp}.pdf")
            
            options = {
                'page-size': 'A4',
                'orientation': 'Landscape',
                'margin-top': '10mm',
                'margin-right': '10mm',
                'margin-bottom': '10mm',
                'margin-left': '10mm',
                'encoding': "UTF-8",
                'no-outline': None,
                'enable-local-file-access': None
            }
            
            pdfkit.from_string(html_content, pdf_file, options=options)
            print(f"PDF report saved to: {pdf_file}")
            
        except ImportError:
            print("\nNote: PDF generation is disabled. To enable:")
            print("1. Install pdfkit: pip install pdfkit")
            print("2. Install wkhtmltopdf:")
            print("   Mac: brew install wkhtmltopdf")
            print("   Ubuntu: sudo apt-get install wkhtmltopdf")
            print("   Windows: Download from https://wkhtmltopdf.org/downloads.html")
        except Exception as e:
            print(f"\nWarning: Could not generate PDF - {str(e)}") 