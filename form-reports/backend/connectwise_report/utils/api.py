import requests
import sys
from connectwise_report.config.settings import CW_URL, CW_HEADERS

def request(method, url, headers, params=None, data=None):
    """Generic function to handle HTTP requests"""
    try:
        if method.lower() == "get":
            response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error in API request: {str(e)}", file=sys.stderr)
        if hasattr(e.response, 'text'):
            print(f"Response content: {e.response.text}", file=sys.stderr)
        return None

def get_time_entries(start_date, end_date, company_id):
    """Get time entries for date range and company"""
    url = f"{CW_URL}time/entries"
    conditions = (f"company/id={company_id} "
                 f"and timeStart>=[{start_date.strftime('%Y-%m-%d')}T00:00:00Z] "
                 f"and timeStart<[{end_date.strftime('%Y-%m-%d')}T00:00:00Z]")
    
    params = {
        "conditions": conditions,
        "orderBy": "timeStart desc"
    }
    
    return request("get", url, CW_HEADERS, params=params) 