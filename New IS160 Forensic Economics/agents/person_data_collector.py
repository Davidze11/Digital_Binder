"""
PersonDataCollectorAgent: Gathers and validates base demographic and occupational data
"""
from typing import Dict, Any, Optional
from datetime import datetime
from dateutil.parser import parse
import json


class PersonDataCollectorAgent:
    """Collects and validates person data for forensic economic analysis."""
    
    REQUIRED_FIELDS = [
        'name', 'dob', 'dod', 'occupation', 'annual_salary',
        'sex', 'education_level', 'home_county', 'home_state', 'status'
    ]
    
    VALID_SEX = ['Male', 'Female', 'M', 'F']
    VALID_EDUCATION = [
        'Less than High School', 'High School', "Some College",
        "Bachelor's", "Master's", "Doctoral", "Professional"
    ]
    VALID_STATUS = ['Active', 'Inactive']
    
    def __init__(self):
        self.person_profile = {}
        self.validation_errors = []
    
    def collect(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Collect and validate person data.
        
        Args:
            input_data: Dictionary containing person information
            
        Returns:
            Structured person profile dictionary
        """
        self.validation_errors = []
        self.person_profile = {}
        
        # Validate required fields
        for field in self.REQUIRED_FIELDS:
            if field not in input_data:
                self.validation_errors.append(f"Missing required field: {field}")
            else:
                self.person_profile[field] = input_data[field]
        
        # Validate dates
        if 'dob' in input_data:
            if not self._validate_date(input_data['dob']):
                self.validation_errors.append("Invalid DOB format")
            else:
                self.person_profile['dob'] = input_data['dob']
        
        if 'dod' in input_data:
            if not self._validate_date(input_data['dod']):
                self.validation_errors.append("Invalid DOD format")
            else:
                self.person_profile['dod'] = input_data['dod']
        
        # Validate date logic (DOD should be after DOB)
        if 'dob' in self.person_profile and 'dod' in self.person_profile:
            try:
                dob = parse(self.person_profile['dob'])
                dod = parse(self.person_profile['dod'])
                if dod <= dob:
                    self.validation_errors.append("DOD must be after DOB")
            except:
                pass
        
        # Validate sex
        if 'sex' in input_data:
            sex = input_data['sex']
            if sex not in self.VALID_SEX:
                self.validation_errors.append(f"Invalid sex: {sex}. Must be one of {self.VALID_SEX}")
            else:
                # Normalize to Male/Female
                self.person_profile['sex'] = 'Male' if sex in ['Male', 'M'] else 'Female'
        
        # Validate education
        if 'education_level' in input_data:
            edu = input_data['education_level']
            if edu not in self.VALID_EDUCATION:
                self.validation_errors.append(f"Invalid education level: {edu}")
            else:
                self.person_profile['education_level'] = edu
        
        # Validate status
        if 'status' in input_data:
            status = input_data['status']
            if status not in self.VALID_STATUS:
                self.validation_errors.append(f"Invalid status: {status}")
            else:
                self.person_profile['status'] = status
        
        # Validate salary
        if 'annual_salary' in input_data:
            try:
                salary = float(input_data['annual_salary'])
                if salary < 0:
                    self.validation_errors.append("Annual salary must be positive")
                else:
                    self.person_profile['annual_salary'] = salary
            except (ValueError, TypeError):
                self.validation_errors.append("Annual salary must be a number")
        
        # Calculate age at death
        if 'dob' in self.person_profile and 'dod' in self.person_profile:
            try:
                dob = parse(self.person_profile['dob'])
                dod = parse(self.person_profile['dod'])
                age_at_death = dod.year - dob.year
                if (dod.month, dod.day) < (dob.month, dob.day):
                    age_at_death -= 1
                self.person_profile['age_at_death'] = age_at_death
            except:
                pass
        
        # Store raw input for reference
        self.person_profile['raw_input'] = input_data
        
        if self.validation_errors:
            raise ValueError(f"Validation errors: {', '.join(self.validation_errors)}")
        
        return self.person_profile
    
    def _validate_date(self, date_string: str) -> bool:
        """Validate date string format."""
        try:
            parse(date_string)
            return True
        except:
            return False
    
    def get_profile(self) -> Dict[str, Any]:
        """Get the collected person profile."""
        return self.person_profile.copy()
    
    def save_profile(self, filepath: str):
        """Save profile to JSON file."""
        with open(filepath, 'w') as f:
            json.dump(self.person_profile, f, indent=2, default=str)


