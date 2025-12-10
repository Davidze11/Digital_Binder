[README.md](https://github.com/user-attachments/files/24067826/README.md)
# Forensic Economics AI Project

Automated forensic economic loss calculation system for wrongful death cases using multi-agent AI architecture.

## Overview

This system automates the workflow of a forensic economist by calculating financial damages from premature death, considering:
- **Life expectancy** (CDC Life Tables)
- **Work-life expectancy** (Skoog et al., 2019 Markov Model)
- **Wage growth** (CA EDD Labor Market Information)
- **Discount factors** (Federal Reserve H.15 Treasury Rates)

The system uses a multi-agent architecture where specialized agents handle different aspects of the analysis, coordinated by a SupervisorAgent.

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Setup

```bash
# Clone or download the project
cd "New IS160 Forensic Economics"

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Command Line

```bash
# Using example input
python main.py

# Using custom JSON input file
python main.py example_input.json
```

### Python API

```python
from agents.supervisor_agent import SupervisorAgent

# Initialize supervisor
supervisor = SupervisorAgent()

# Prepare input data
input_data = {
    "name": "John Doe",
    "dob": "1980-01-15",
    "dod": "2024-03-20",
    "occupation": "Software Engineer",
    "annual_salary": 120000,
    "sex": "Male",
    "education_level": "Bachelor's",
    "home_county": "Los Angeles",
    "home_state": "California",
    "status": "Active"
}

# Run analysis
result = supervisor.run_analysis(input_data)

# Check results
if result['status'] == 'success':
    print(f"Total Economic Loss: {result['total_economic_loss_formatted']}")
    print(f"Excel report: {result['output_file']}")
else:
    print(f"Error: {result['error_message']}")
```

### Testing

```bash
# Run test script
python test_system.py
```

## Project Structure

```
.
├── agents/
│   ├── __init__.py
│   ├── supervisor_agent.py
│   ├── person_data_collector.py
│   ├── life_expectancy_agent.py
│   ├── worklife_agent.py
│   ├── wage_data_agent.py
│   ├── calculation_agents.py
│   ├── discount_agent.py
│   └── excel_generator.py
├── data/
│   └── skoog_tables/  # Skoog work-life expectancy tables
├── utils/
│   ├── __init__.py
│   └── helpers.py
├── main.py
├── requirements.txt
└── README.md
```

## System Architecture

### Agents

| Agent | Purpose | Key Functions |
|-------|---------|---------------|
| **SupervisorAgent** | Coordinates all operations | Orchestrates workflow, error handling, result aggregation |
| **PersonDataCollectorAgent** | Validates and structures input data | Data validation, format conversion, age calculation |
| **LifeExpectancyAgent** | Fetches CDC life expectancy data | Retrieves life tables, calculates remaining life expectancy |
| **WorklifeRemainingYearsAgent** | Calculates work-life expectancy | Uses Skoog et al. tables, matches by age/sex/education |
| **WageDataAgent** | Retrieves salary growth rates | Fetches CA EDD data, computes annual growth rates |
| **FederalReserveRateAgent** | Gets current discount rates | Fetches Treasury rates from Fed H.15 |
| **DiscountFactorAgent** | Calculates discount factors | Computes PV discount factors for future years |
| **PersonLIFEyrAgentEnhanced** | Builds age timeline | Creates year-by-year age progression |
| **CalcFullActualCumAgent** | Projects earnings | Calculates annual and cumulative earnings |
| **CalcPresentCumulPresentValueAgent** | Calculates present values | Computes PV and cumulative PV |
| **ComprehensiveExcelGenerator** | Creates Excel reports | Generates formatted workbook with all data |

### Workflow

1. **Data Collection**: PersonDataCollectorAgent validates and structures input
2. **Life Expectancy**: LifeExpectancyAgent retrieves remaining life expectancy
3. **Work-Life Expectancy**: WorklifeRemainingYearsAgent calculates work years
4. **Wage Growth**: WageDataAgent fetches salary growth rates
5. **Discount Rate**: FederalReserveRateAgent gets current Treasury rates
6. **Timeline**: PersonLIFEyrAgentEnhanced builds age timeline
7. **Earnings Projection**: CalcFullActualCumAgent projects future earnings
8. **Present Value**: CalcPresentCumulPresentValueAgent calculates discounted values
9. **Report Generation**: ComprehensiveExcelGenerator creates Excel workbook

## Output

The system generates a comprehensive Excel workbook (`forensic_economic_report_<name>_<timestamp>.xlsx`) with the following worksheets:

1. **Summary**: Key metrics and total economic loss
2. **Personal Data**: Input person information
3. **Life & Work-Life Expectancy**: Life expectancy and work-life calculations
4. **Wage Growth**: Salary growth rate analysis
5. **Discount Rate**: Discount rate and data source information
6. **Earnings Projections**: Year-by-year projected earnings
7. **Present Value Calculations**: Discounted present values and cumulative totals

### Key Outputs

- **Total Economic Loss**: Present value of all projected earnings (in USD)
- **Life Expectancy**: Remaining life expectancy in years
- **Work-Life Expectancy**: Remaining work years
- **Annual Growth Rate**: Salary growth percentage
- **Discount Rate**: Treasury rate used for discounting
- **Detailed Tables**: Year-by-year breakdown of earnings and present values

## Input Data Format

### Required Fields

```json
{
  "name": "Person's full name",
  "dob": "YYYY-MM-DD",
  "dod": "YYYY-MM-DD",
  "occupation": "Job title or occupation",
  "annual_salary": 120000,
  "sex": "Male" or "Female",
  "education_level": "Less than High School" | "High School" | "Some College" | "Bachelor's" | "Master's" | "Doctoral" | "Professional",
  "home_county": "County name",
  "home_state": "State name",
  "status": "Active" or "Inactive"
}
```

### Example Input

See `example_input.json` for a complete example.

## Data Sources

- **CDC Life Tables**: https://www.cdc.gov/nchs/data/nvsr/nvsr70/nvsr70-17.pdf
- **CA EDD Labor Market Info**: https://labormarketinfo.edd.ca.gov/
- **Federal Reserve H.15**: https://www.federalreserve.gov/releases/h15/current/
- **Skoog et al. (2019)**: Markov Model of Labor Force Activity

## Key Formulas

- **Annual Growth Rate**: `(CurrentYear - PrevYear) / PrevYear × 100`
- **Projected Earnings**: `BaseSalary × (1 + GrowthRate)^Years`
- **Discount Factor**: `1 / (1 + DiscountRate)^Years`
- **Present Value**: `ProjectedEarnings × DiscountFactor`
- **Cumulative PV**: Sum of all present values up to current year

## Limitations and Notes

1. **Data Sources**: The system uses approximate/simplified data tables. In production, you should:
   - Integrate with actual CDC API for life tables
   - Parse actual Skoog PDF for work-life tables
   - Connect to CA EDD API for real wage data
   - Scrape Federal Reserve website for current rates

2. **Assumptions**:
   - Salary grows at constant rate
   - Work-life ends at calculated work-life expectancy
   - No consideration for benefits, retirement contributions, etc.
   - Simplified discounting (constant rate)

3. **Future Enhancements**:
   - Add Chat-based supervisor for natural language queries
   - Integrate MongoDB/SQLite for historical case storage
   - Implement PDF-to-Excel ingestion for automated table extraction
   - Add API endpoints (FastAPI/Flask) for dashboard integration
   - Support for benefits, retirement, and other compensation

## Error Handling

The system includes comprehensive error handling:
- Input validation with detailed error messages
- Graceful fallbacks for external data sources
- Logging to `forensic_economics.log`
- Status reporting in results dictionary

## License

This project is for educational and research purposes.

## Contributing

Contributions are welcome! Please ensure:
- Code follows PEP 8 style guidelines
- All tests pass
- Documentation is updated
- New features include appropriate error handling

