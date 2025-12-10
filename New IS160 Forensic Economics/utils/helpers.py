"""
Helper utility functions
"""
from datetime import datetime
from typing import Dict, Any
import json


def calculate_age(dob: str, reference_date: str = None) -> int:
    """Calculate age from date of birth to reference date (or today)."""
    from dateutil.parser import parse
    
    birth_date = parse(dob)
    if reference_date:
        ref_date = parse(reference_date)
    else:
        ref_date = datetime.now()
    
    age = ref_date.year - birth_date.year
    if (ref_date.month, ref_date.day) < (birth_date.month, birth_date.day):
        age -= 1
    return age


def validate_date(date_string: str) -> bool:
    """Validate date string format."""
    try:
        from dateutil.parser import parse
        parse(date_string)
        return True
    except:
        return False


def format_currency(amount: float) -> str:
    """Format number as currency."""
    return f"${amount:,.2f}"


def save_json(data: Dict[Any, Any], filepath: str):
    """Save dictionary to JSON file."""
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=2, default=str)


def load_json(filepath: str) -> Dict[Any, Any]:
    """Load JSON file to dictionary."""
    with open(filepath, 'r') as f:
        return json.load(f)


