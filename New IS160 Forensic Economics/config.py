"""
Configuration file for Forensic Economics AI Project
"""

# Default values
DEFAULT_DISCOUNT_RATE = 0.045  # 4.5%
DEFAULT_WAGE_GROWTH_RATE = 0.025  # 2.5%

# Data source URLs
CDC_LIFE_TABLES_URL = "https://www.cdc.gov/nchs/data/nvsr/nvsr70/nvsr70-17.pdf"
FED_H15_URL = "https://www.federalreserve.gov/releases/h15/current/"
CA_EDD_URL = "https://labormarketinfo.edd.ca.gov/"

# Output settings
OUTPUT_DIRECTORY = "output"
LOG_FILE = "forensic_economics.log"

# Validation settings
MIN_AGE = 0
MAX_AGE = 120
MIN_SALARY = 0
MAX_SALARY = 10000000

# Work-life expectancy settings
MIN_WORKLIFE_YEARS = 0
MAX_WORKLIFE_YEARS = 50


