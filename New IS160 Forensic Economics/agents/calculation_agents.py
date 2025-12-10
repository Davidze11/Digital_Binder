"""
Calculation Agents: Perform earnings projections and present value calculations
"""
from typing import Dict, Any, List
import pandas as pd
from dateutil.parser import parse
from datetime import datetime


class PersonLIFEyrAgentEnhanced:
    """Builds age timeline from DOB to life expectancy."""
    
    def __init__(self):
        self.age_timeline = []
        self.year_timeline = []
    
    def build_timeline(self, person_profile: Dict[str, Any], 
                      life_expectancy_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Build age timeline from DOB through expected lifespan.
        
        Args:
            person_profile: Person data with dob, dod
            life_expectancy_data: Life expectancy data with remaining_life_expectancy
            
        Returns:
            List of dictionaries with year, age, and date information
        """
        dob = parse(person_profile['dob'])
        dod = parse(person_profile['dod'])
        age_at_death = person_profile.get('age_at_death', 0)
        remaining_life = life_expectancy_data.get('remaining_life_expectancy', 0)
        
        timeline = []
        
        # Start from year of death
        current_year = dod.year
        current_age = age_at_death
        
        # Calculate exact age at death with decimal
        age_delta_days = (dod - dob).days
        age_at_death_decimal = age_delta_days / 365.25
        
        # Add years from death to end of life expectancy
        for year_offset in range(int(remaining_life) + 1):
            year = current_year + year_offset
            # Calculate age with decimal precision
            age = age_at_death_decimal + year_offset
            
            timeline.append({
                'year': year,
                'age': age,
                'age_integer': int(age),
                'year_offset': year_offset,  # Years from death
                'date': f"{year}-01-01"
            })
        
        self.age_timeline = timeline
        return timeline


class CalcFullActualCumAgent:
    """Projects full earning capacity with salary growth."""
    
    def __init__(self):
        self.projected_earnings = []
        self.base_salary = None
        self.growth_rate = None
    
    def calculate_earnings(self, person_profile: Dict[str, Any],
                          wage_data: Dict[str, Any],
                          worklife_data: Dict[str, Any],
                          timeline: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Calculate projected earnings for each year.
        
        Args:
            person_profile: Person data with annual_salary
            wage_data: Wage growth data with annual_growth_rate
            worklife_data: Work-life expectancy data
            timeline: Age timeline
            
        Returns:
            List of dictionaries with year, age, projected earnings
        """
        self.base_salary = person_profile.get('annual_salary', 0)
        self.growth_rate = wage_data.get('annual_growth_rate', 0.025)
        worklife_years = worklife_data.get('worklife_expectancy', 0)
        age_at_death = person_profile.get('age_at_death', 0)
        
        earnings_table = []
        cumulative_earnings = 0
        
        for entry in timeline:
            year_offset = entry['year_offset']
            age = entry['age']
            
            # Check if still in work-life period
            if year_offset <= worklife_years:
                # Calculate projected salary for this year
                # Salary grows from base salary at death
                projected_salary = self.base_salary * ((1 + self.growth_rate) ** year_offset)
                cumulative_earnings += projected_salary
            else:
                # Beyond work-life expectancy - no earnings
                projected_salary = 0
            
            earnings_table.append({
                'year': entry['year'],
                'age': age,
                'year_offset': year_offset,
                'projected_earnings': round(projected_salary, 2),
                'cumulative_earnings': round(cumulative_earnings, 2)
            })
        
        self.projected_earnings = earnings_table
        return earnings_table


class CalcPresentCumulPresentValueAgent:
    """Calculates present value and cumulative present value."""
    
    def __init__(self):
        self.present_values = []
        self.discount_rate = None
    
    def calculate_present_values(self, earnings_table: List[Dict[str, Any]],
                                discount_rate: float) -> List[Dict[str, Any]]:
        """
        Calculate present value for each year's earnings.
        
        Args:
            earnings_table: Table with projected earnings
            discount_rate: Annual discount rate
            
        Returns:
            List of dictionaries with present values
        """
        self.discount_rate = discount_rate
        pv_table = []
        cumulative_pv = 0
        
        for entry in earnings_table:
            year_offset = entry['year_offset']
            projected_earnings = entry.get('projected_earnings', 0)
            
            # Calculate discount factor
            discount_factor = 1.0 / ((1 + discount_rate) ** year_offset)
            
            # Calculate present value
            present_value = projected_earnings * discount_factor
            cumulative_pv += present_value
            
            pv_entry = {
                'year': entry['year'],
                'age': entry['age'],
                'year_offset': year_offset,
                'projected_earnings': projected_earnings,
                'discount_factor': round(discount_factor, 6),
                'present_value': round(present_value, 2),
                'cumulative_present_value': round(cumulative_pv, 2)
            }
            
            pv_table.append(pv_entry)
        
        self.present_values = pv_table
        return pv_table
    
    def get_total_economic_loss(self) -> float:
        """Get total economic loss (final cumulative present value)."""
        if self.present_values:
            return self.present_values[-1].get('cumulative_present_value', 0)
        return 0.0

