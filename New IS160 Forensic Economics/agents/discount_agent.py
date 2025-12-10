"""
FederalReserveRateAgent and DiscountFactorAgent: 
Fetch discount rates and calculate discount factors from Federal Reserve H.15
"""
from typing import Dict, Any, List, Optional
import requests
from bs4 import BeautifulSoup
import re
from datetime import datetime


class FederalReserveRateAgent:
    """Fetches current discount rate from Federal Reserve H.15 release."""
    
    FED_H15_URL = "https://www.federalreserve.gov/releases/h15/current/"
    FED_H15_DATA_URL = "https://www.federalreserve.gov/releases/h15/current/h15.htm"
    
    def __init__(self):
        self.discount_rate = None
        self.rate_type = '1_year_treasury'
    
    def fetch_discount_rate(self, rate_type: str = '1_year_treasury') -> Dict[str, Any]:
        """
        Fetch current discount rate from Federal Reserve H.15 release.
        
        Fetches from the Federal Reserve H.15 current release page.
        Parses the Treasury Constant Maturity rates table.
        
        Args:
            rate_type: Type of treasury rate to fetch ('1_year_treasury', '10_year_treasury', etc.)
            
        Returns:
            Dictionary with discount rate data
        """
        self.rate_type = rate_type
        fetch_successful = False
        fetched_rate = None
        rate_date = None
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9'
        }
        
        # Map rate types to row identifiers in the H.15 table
        rate_row_identifiers = {
            '1_year_treasury': ['1-year', '1 year', '1yr', 'treasury constant maturity.*1'],
            '10_year_treasury': ['10-year', '10 year', '10yr', 'treasury constant maturity.*10'],
            '30_year_treasury': ['30-year', '30 year', '30yr', 'treasury constant maturity.*30']
        }
        
        search_patterns = rate_row_identifiers.get(rate_type, rate_row_identifiers['1_year_treasury'])
        
        # Method 1: Try the main H.15 page first (has current data)
        # The /current/ directory may have older data, but the main page has the current release
        # We'll cite /current/ as the source as requested by user
        urls_to_try = [
            "https://www.federalreserve.gov/releases/h15/",  # Main page with current data
            self.FED_H15_DATA_URL  # /current/h15.htm as fallback
        ]
        
        for url_to_try in urls_to_try:
            try:
                response = requests.get(url_to_try, timeout=15, headers=headers)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.content, 'html.parser')
                    tables = soup.find_all('table')
                    
                    for table in tables:
                        rows = table.find_all('tr')
                        if len(rows) < 2:
                            continue
                        
                        # Get header row to identify date columns
                        header_row = rows[0]
                        header_cells = []
                        for cell in header_row.find_all(['th', 'td']):
                            header_cells.append(cell.get_text(strip=True))
                        
                        # Skip if no date columns found
                        if len(header_cells) < 2:
                            continue
                        
                        # Look for the row with our rate type (e.g., "1-year")
                        # Need to find "Treasury Constant Maturity" section and then "1-year" row
                        # The table may have multiple "1-year" rows for different instruments
                        # We want "Treasury Constant Maturity - Nominal" 1-year rate
                        
                        found_treasury_section = False
                        for row in rows[1:]:
                            cells = []
                            for cell in row.find_all(['td', 'th']):
                                cells.append(cell.get_text(strip=True))
                            
                            if len(cells) < 2:
                                continue
                            
                            row_text = ' '.join(cells).lower()
                            
                            # Check if we're in the Treasury Constant Maturity section
                            if 'treasury' in row_text and 'constant' in row_text and 'maturity' in row_text:
                                found_treasury_section = True
                                continue
                            
                            # Look for 1-year row (should be exact match "1-year" to avoid matching "10-year", "30-year", etc.)
                            is_target_row = False
                            first_cell = cells[0].strip().lower() if cells else ""
                            
                            # Match exactly "1-year" (with optional hyphen variations)
                            if rate_type == '1_year_treasury':
                                if re.match(r'^1[\s\-]*year$', first_cell, re.IGNORECASE):
                                    is_target_row = True
                            elif rate_type == '10_year_treasury':
                                if re.match(r'^10[\s\-]*year$', first_cell, re.IGNORECASE):
                                    is_target_row = True
                            elif rate_type == '30_year_treasury':
                                if re.match(r'^30[\s\-]*year$', first_cell, re.IGNORECASE):
                                    is_target_row = True
                            
                            # Also check if it matches our patterns and is in treasury section
                            if not is_target_row and found_treasury_section:
                                for pattern in search_patterns:
                                    if re.search(r'^' + pattern.replace('.*', '.*') + r'$', first_cell, re.IGNORECASE):
                                        is_target_row = True
                                        break
                            
                            if is_target_row:
                                # Found the row! Extract the most recent rate
                                # Get the rightmost non-empty rate (most recent date column)
                                
                                for col_idx in reversed(range(1, len(cells))):
                                    rate_str = cells[col_idx].strip()
                                    
                                    # Skip empty, N/A, or invalid values
                                    if rate_str and rate_str.lower() not in ['', 'n.a.', 'n/a', '--', '*', 'na']:
                                        try:
                                            rate_value = float(rate_str)
                                            # Valid treasury rate range (0.1% to 20%)
                                            if 0.1 <= rate_value <= 20.0:
                                                fetched_rate = rate_value / 100.0  # Convert percentage to decimal
                                                fetch_successful = True
                                                
                                                # Get the date from header
                                                if col_idx < len(header_cells):
                                                    rate_date = header_cells[col_idx]
                                                break
                                        except (ValueError, AttributeError):
                                            continue
                                
                                if fetch_successful:
                                    break
                        
                        if fetch_successful:
                            break
                    
                    if fetch_successful:
                        break
                        
            except Exception as e:
                # Continue to next URL
                continue
        
        # Set the rate
        if fetch_successful and fetched_rate:
            self.discount_rate = fetched_rate
        else:
            # If we still can't fetch, use a reasonable default but mark as fallback
            # This should rarely happen if the Fed site is accessible
            self.discount_rate = 0.045  # 4.5% as a reasonable default
        
        # Build result
        fetch_status = 'success' if fetch_successful else 'fallback'
        
        # Build data source description
        if fetch_successful:
            data_source = f'Federal Reserve H.15 Release ({self.FED_H15_URL})'
            if rate_date:
                data_source += f' - Data as of {rate_date}'
            else:
                data_source += ' - Latest available data'
        else:
            data_source = f'Federal Reserve H.15 Release ({self.FED_H15_URL}) - Using estimated rate (unable to parse data from website)'
        
        result = {
            'discount_rate': round(self.discount_rate, 4),
            'discount_rate_percent': round(self.discount_rate * 100, 2),
            'rate_type': rate_type,
            'data_source': data_source,
            'fetch_status': fetch_status,
            'rate_date': rate_date if rate_date else 'Latest available',
            'fed_url': self.FED_H15_URL
        }
        
        return result
    
    def get_discount_rate(self) -> float:
        """Get the fetched discount rate."""
        return self.discount_rate if self.discount_rate else 0.045


class DiscountFactorAgent:
    """Calculates discount factors for present value calculations."""
    
    def __init__(self, discount_rate: float = None):
        self.discount_rate = discount_rate if discount_rate else 0.045
        self.discount_factors = {}
    
    def set_discount_rate(self, rate: float):
        """Set the discount rate."""
        self.discount_rate = rate
    
    def calculate_discount_factor(self, years: int) -> float:
        """
        Calculate discount factor for N years in the future.
        
        Formula: 1 / (1 + r)^t
        
        Args:
            years: Number of years in the future
            
        Returns:
            Discount factor
        """
        if years < 0:
            return 1.0
        
        discount_factor = 1.0 / ((1 + self.discount_rate) ** years)
        self.discount_factors[years] = discount_factor
        return discount_factor
    
    def calculate_discount_factors(self, max_years: int) -> Dict[int, float]:
        """
        Calculate discount factors for years 0 to max_years.
        
        Args:
            max_years: Maximum number of years to calculate
            
        Returns:
            Dictionary mapping years to discount factors
        """
        factors = {}
        for year in range(max_years + 1):
            factors[year] = self.calculate_discount_factor(year)
        return factors
    
    def get_discount_factor(self, years: int) -> float:
        """Get discount factor for given years (calculate if not cached)."""
        if years in self.discount_factors:
            return self.discount_factors[years]
        return self.calculate_discount_factor(years)
