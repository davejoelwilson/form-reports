from fastapi import APIRouter
from datetime import datetime
from typing import List
from connectwise_report.utils.api import get_time_entries
from connectwise_report.config.settings import IGNITE_COMPANY_ID

router = APIRouter()

@router.get("/time-entries")
async def get_time_entries_route(startDate: str, endDate: str):
    start_date = datetime.fromisoformat(startDate.replace('Z', '+00:00'))
    end_date = datetime.fromisoformat(endDate.replace('Z', '+00:00'))
    
    entries = get_time_entries(start_date, end_date, IGNITE_COMPANY_ID)
    return entries 