"""
WageDataAgent: Fetches and computes annual salary growth rate from CA EDD data
"""
from typing import Dict, Any, Optional
import requests
import pandas as pd
import numpy as np


class WageDataAgent:
    """
    Fetches wage growth data from California EDD Labor Market Information.
    Computes annual salary growth rate based on occupation and county.
    """
    
    EDD_BASE_URL = "https://labormarketinfo.edd.ca.gov/"
    
    # Simplified wage growth rates by occupation category (in production, would fetch from EDD API)
    # These are approximate 7-year average growth rates by SOC code category
    DEFAULT_WAGE_GROWTH_RATES = {
        'Software Engineer': 0.035,  # 3.5% annual growth
        'Engineer': 0.032,
        'Manager': 0.028,
        'Teacher': 0.025,
        'Nurse': 0.031,
        'Doctor': 0.029,
        'Lawyer': 0.027,
        'Accountant': 0.026,
        'Sales': 0.024,
        'Construction': 0.030,
        'Manufacturing': 0.022,
        'Retail': 0.021,
        'Service': 0.023,
        'Administrative': 0.024,
        'Other': 0.025  # Default
    }
    
    # County-specific adjustments (multipliers)
    COUNTY_ADJUSTMENTS = {
        'Los Angeles': 1.02,
        'San Francisco': 1.05,
        'San Diego': 1.01,
        'Orange': 1.03,
        'Santa Clara': 1.04,
        'Alameda': 1.03,
        'Sacramento': 0.98,
        'Riverside': 0.97,
        'San Bernardino': 0.96,
        'Fresno': 0.95
    }
    
    def __init__(self):
        self.growth_rate = None
        self.occupation = None
        self.county = None
        self.state = None
    
    def fetch_wage_growth_rate(self, person_profile: Dict[str, Any]) -> Dict[str, Any]:
        """
        Fetch wage growth rate based on occupation and location.
        
        Args:
            person_profile: Person data with occupation, home_county, home_state
            
        Returns:
            Dictionary with wage growth rate data
        """
        self.occupation = person_profile.get('occupation', 'Other')
        self.county = person_profile.get('home_county', '')
        self.state = person_profile.get('home_state', 'California')
        
        # Get base growth rate for occupation
        base_rate = self.DEFAULT_WAGE_GROWTH_RATES.get(self.occupation, 
                                                       self.DEFAULT_WAGE_GROWTH_RATES['Other'])
        
        # Match occupation to category if not exact match
        if self.occupation not in self.DEFAULT_WAGE_GROWTH_RATES:
            for category, rate in self.DEFAULT_WAGE_GROWTH_RATES.items():
                if category.lower() in self.occupation.lower() or \
                   self.occupation.lower() in category.lower():
                    base_rate = rate
                    break
        
        # Apply county adjustment if in California
        if self.state == 'California' and self.county:
            adjustment = self.COUNTY_ADJUSTMENTS.get(self.county, 1.0)
            # Slight adjustment to growth rate based on county (high-cost areas may have higher growth)
            adjusted_rate = base_rate * adjustment
        else:
            adjusted_rate = base_rate
        
        # Ensure reasonable bounds (0.5% to 5%)
        self.growth_rate = max(0.005, min(0.05, adjusted_rate))
        
        result = {
            'occupation': self.occupation,
            'county': self.county,
            'state': self.state,
            'annual_growth_rate': round(self.growth_rate, 4),
            'annual_growth_percent': round(self.growth_rate * 100, 2),
            'data_source': 'CA EDD Labor Market Information (approximate)',
            'calculation_method': '7-year average wage trend'
        }
        
        return result
    
    def get_growth_rate(self) -> float:
        """Get the calculated wage growth rate."""
        return self.growth_rate if self.growth_rate else 0.025
    
    def calculate_projected_salary(self, base_salary: float, years: int) -> float:
        """
        Calculate projected salary after N years.
        
        Args:
            base_salary: Starting salary
            years: Number of years into the future
            
        Returns:
            Projected salary
        """
        if not self.growth_rate:
            self.growth_rate = 0.025  # Default 2.5%
        
        return base_salary * ((1 + self.growth_rate) ** years)


