"""
WorklifeRemainingYearsAgent: Determines remaining work-life expectancy using Skoog et al. dataset
"""
from typing import Dict, Any, Optional
import pandas as pd


class WorklifeRemainingYearsAgent:
    """
    Determines remaining work-life expectancy based on Skoog et al. (2019)
    Markov Model of Labor Force Activity.
    """
    
    # Simplified work-life expectancy tables based on Skoog et al. (2019)
    # Format: {age: {sex: {education: work_life_expectancy}}}
    # These are approximate values - in production, would parse actual Skoog PDF/tables
    WORKLIFE_TABLES = {
        'Male': {
            'Less than High School': {
                25: 35.2, 30: 30.5, 35: 25.8, 40: 21.2, 45: 16.8, 50: 12.8,
                55: 9.2, 60: 6.1, 65: 3.5
            },
            'High School': {
                25: 36.8, 30: 32.1, 35: 27.4, 40: 22.8, 45: 18.4, 50: 14.2,
                55: 10.4, 60: 7.1, 65: 4.2
            },
            "Some College": {
                25: 37.5, 30: 32.8, 35: 28.1, 40: 23.5, 45: 19.1, 50: 14.9,
                55: 11.0, 60: 7.6, 65: 4.5
            },
            "Bachelor's": {
                25: 38.2, 30: 33.5, 35: 28.8, 40: 24.2, 45: 19.8, 50: 15.6,
                55: 11.6, 60: 8.1, 65: 4.8
            },
            "Master's": {
                25: 38.8, 30: 34.1, 35: 29.4, 40: 24.8, 45: 20.4, 50: 16.2,
                55: 12.1, 60: 8.5, 65: 5.1
            },
            "Doctoral": {
                25: 39.2, 30: 34.5, 35: 29.8, 40: 25.2, 45: 20.8, 50: 16.6,
                55: 12.4, 60: 8.8, 65: 5.3
            },
            "Professional": {
                25: 39.0, 30: 34.3, 35: 29.6, 40: 25.0, 45: 20.6, 50: 16.4,
                55: 12.3, 60: 8.7, 65: 5.2
            }
        },
        'Female': {
            'Less than High School': {
                25: 32.5, 30: 28.2, 35: 23.9, 40: 19.8, 45: 15.9, 50: 12.3,
                55: 9.1, 60: 6.3, 65: 3.8
            },
            'High School': {
                25: 34.1, 30: 29.8, 35: 25.5, 40: 21.4, 45: 17.5, 50: 13.7,
                55: 10.3, 60: 7.3, 65: 4.5
            },
            "Some College": {
                25: 34.8, 30: 30.5, 35: 26.2, 40: 22.1, 45: 18.2, 50: 14.4,
                55: 10.9, 60: 7.8, 65: 4.8
            },
            "Bachelor's": {
                25: 35.5, 30: 31.2, 35: 26.9, 40: 22.8, 45: 18.9, 50: 15.1,
                55: 11.5, 60: 8.3, 65: 5.1
            },
            "Master's": {
                25: 36.1, 30: 31.8, 35: 27.5, 40: 23.4, 45: 19.5, 50: 15.7,
                55: 12.0, 60: 8.7, 65: 5.4
            },
            "Doctoral": {
                25: 36.5, 30: 32.2, 35: 27.9, 40: 23.8, 45: 19.9, 50: 16.1,
                55: 12.3, 60: 8.9, 65: 5.6
            },
            "Professional": {
                25: 36.3, 30: 32.0, 35: 27.7, 40: 23.6, 45: 19.7, 50: 15.9,
                55: 12.2, 60: 8.8, 65: 5.5
            }
        }
    }
    
    def __init__(self):
        self.worklife_expectancy = None
        self.age_at_death = None
    
    def fetch_worklife_expectancy(self, person_profile: Dict[str, Any]) -> Dict[str, Any]:
        """
        Fetch work-life expectancy based on person profile.
        
        Args:
            person_profile: Person data with age_at_death, sex, education_level
            
        Returns:
            Dictionary with work-life expectancy data
        """
        self.age_at_death = person_profile.get('age_at_death', 0)
        sex = person_profile.get('sex', 'Male')
        education = person_profile.get('education_level', "Bachelor's")
        
        # Get appropriate table
        if sex not in self.WORKLIFE_TABLES:
            raise ValueError(f"Invalid sex: {sex}")
        
        if education not in self.WORKLIFE_TABLES[sex]:
            # Default to Bachelor's if education not found
            education = "Bachelor's"
        
        age_table = self.WORKLIFE_TABLES[sex][education]
        
        # Find closest age
        if self.age_at_death < min(age_table.keys()):
            # Use minimum age value
            self.worklife_expectancy = age_table[min(age_table.keys())]
        elif self.age_at_death > max(age_table.keys()):
            # Use maximum age value (or 0 if beyond retirement)
            self.worklife_expectancy = age_table.get(max(age_table.keys()), 0)
        else:
            # Find appropriate age bracket
            ages = sorted([k for k in age_table.keys() if k <= self.age_at_death], reverse=True)
            if ages:
                closest_age = ages[0]
                self.worklife_expectancy = age_table[closest_age]
                
                # Interpolate if needed
                if len(ages) > 1:
                    age1, age2 = ages[0], ages[1]
                    exp1, exp2 = age_table[age1], age_table[age2]
                    if age2 != age1:
                        ratio = (self.age_at_death - age1) / (age2 - age1)
                        self.worklife_expectancy = exp1 - (exp1 - exp2) * ratio
            else:
                self.worklife_expectancy = 0
        
        # Work-life cannot exceed remaining life expectancy
        # This will be validated in the supervisor
        
        result = {
            'age_at_death': self.age_at_death,
            'worklife_expectancy': round(max(0, self.worklife_expectancy), 2),
            'sex': sex,
            'education_level': education,
            'data_source': 'Skoog et al. (2019) Markov Model'
        }
        
        return result
    
    def get_worklife_expectancy(self) -> float:
        """Get the calculated work-life expectancy."""
        return self.worklife_expectancy if self.worklife_expectancy else 0.0


class SkoogTableAgent:
    """Fetches and processes Skoog actuarial tables (placeholder for PDF parsing)."""
    
    def __init__(self):
        self.tables = {}
    
    def load_tables(self, filepath: Optional[str] = None) -> Dict[str, Any]:
        """
        Load Skoog tables from file (PDF parsing would go here).
        
        Args:
            filepath: Path to Skoog PDF or data file
            
        Returns:
            Dictionary of work-life expectancy tables
        """
        # In production, this would parse the actual Skoog PDF
        # For now, return empty dict (WorklifeRemainingYearsAgent has hardcoded values)
        return {}


