"""
LifeExpectancyAgent: Retrieves life expectancy based on sex, age, DOB/DOD from CDC Life Tables
"""
from typing import Dict, Any, Optional
import requests
import pandas as pd
from dateutil.parser import parse


class LifeExpectancyAgent:
    """Fetches life expectancy data from CDC Life Tables."""
    
    CDC_LIFE_TABLES_BASE_URL = "https://www.cdc.gov/nchs/data/nvsr/nvsr70/nvsr70-17.pdf"
    
    # Approximate life expectancy tables (simplified - in production, would fetch from CDC API)
    # These are based on 2019 US Life Tables
    LIFE_EXPECTANCY_MALE = {
        # Age: Remaining life expectancy
        0: 76.3, 5: 71.8, 10: 66.9, 15: 62.0, 20: 57.1, 25: 52.3,
        30: 47.5, 35: 42.7, 40: 37.9, 45: 33.2, 50: 28.6, 55: 24.2,
        60: 20.0, 65: 16.1, 70: 12.6, 75: 9.6, 80: 7.1, 85: 5.1
    }
    
    LIFE_EXPECTANCY_FEMALE = {
        0: 81.4, 5: 76.8, 10: 71.9, 15: 67.0, 20: 62.1, 25: 57.2,
        30: 52.3, 35: 47.5, 40: 42.7, 45: 38.0, 50: 33.4, 55: 28.9,
        60: 24.5, 65: 20.3, 70: 16.3, 75: 12.7, 80: 9.6, 85: 6.9
    }
    
    def __init__(self):
        self.life_expectancy = None
        self.age_at_death = None
        self.sex = None
    
    def fetch_life_expectancy(self, person_profile: Dict[str, Any]) -> Dict[str, Any]:
        """
        Fetch life expectancy based on person profile.
        
        Args:
            person_profile: Person data dictionary with dob, dod, sex
            
        Returns:
            Dictionary with life expectancy data
        """
        # Extract age at death
        if 'age_at_death' in person_profile:
            self.age_at_death = person_profile['age_at_death']
        else:
            # Calculate from dob and dod
            dob = parse(person_profile['dob'])
            dod = parse(person_profile['dod'])
            self.age_at_death = dod.year - dob.year
            if (dod.month, dod.day) < (dob.month, dob.day):
                self.age_at_death -= 1
        
        self.sex = person_profile.get('sex', 'Male')
        
        # Get life expectancy table based on sex
        if self.sex == 'Male':
            table = self.LIFE_EXPECTANCY_MALE
        else:
            table = self.LIFE_EXPECTANCY_FEMALE
        
        # Find closest age in table
        closest_age = min(table.keys(), key=lambda x: abs(x - self.age_at_death))
        if self.age_at_death > closest_age:
            # Interpolate if needed (simplified linear interpolation)
            ages = sorted([k for k in table.keys() if k <= self.age_at_death], reverse=True)
            if len(ages) >= 2:
                age1, age2 = ages[0], ages[1]
                exp1, exp2 = table[age1], table[age2]
                # Simple interpolation
                ratio = (self.age_at_death - age1) / (age2 - age1) if age2 != age1 else 0
                self.life_expectancy = exp1 - (exp1 - exp2) * ratio
            else:
                self.life_expectancy = table[closest_age]
        else:
            self.life_expectancy = table[closest_age]
        
        # Calculate total expected lifespan
        total_lifespan = self.age_at_death + self.life_expectancy
        
        result = {
            'age_at_death': self.age_at_death,
            'remaining_life_expectancy': round(self.life_expectancy, 2),
            'total_expected_lifespan': round(total_lifespan, 2),
            'sex': self.sex,
            'data_source': 'CDC Life Tables (2019)'
        }
        
        return result
    
    def get_life_expectancy(self) -> float:
        """Get the calculated life expectancy."""
        return self.life_expectancy if self.life_expectancy else 0.0


