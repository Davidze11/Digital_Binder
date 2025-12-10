"""
Main entry point for Forensic Economics AI Project
"""
import json
import sys
from agents.supervisor_agent import SupervisorAgent


def main():
    """Main function to run forensic economic analysis."""
    
    # Example input data
    example_input = {
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
    
    # Check if input file provided
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        try:
            with open(input_file, 'r') as f:
                input_data = json.load(f)
        except Exception as e:
            print(f"Error reading input file: {e}")
            print("Using example input data instead.")
            input_data = example_input
    else:
        print("No input file provided. Using example data.")
        print("Usage: python main.py <input_json_file>")
        print("\nUsing example input:")
        input_data = example_input
    
    # Print input data
    print("\n" + "="*60)
    print("FORENSIC ECONOMICS AI - ECONOMIC LOSS ANALYSIS")
    print("="*60)
    print(f"\nAnalyzing case for: {input_data.get('name', 'Unknown')}")
    print(f"Date of Birth: {input_data.get('dob', 'N/A')}")
    print(f"Date of Death: {input_data.get('dod', 'N/A')}")
    print(f"Occupation: {input_data.get('occupation', 'N/A')}")
    print(f"Annual Salary: ${input_data.get('annual_salary', 0):,.2f}")
    print("\n" + "-"*60)
    
    # Initialize supervisor
    supervisor = SupervisorAgent()
    
    # Validate inputs
    is_valid, errors = supervisor.validate_inputs(input_data)
    if not is_valid:
        print("Input validation errors:")
        for error in errors:
            print(f"  - {error}")
        return
    
    # Run analysis
    print("Starting analysis...\n")
    results = supervisor.run_analysis(input_data)
    
    # Print results
    print("\n" + "="*60)
    print("ANALYSIS RESULTS")
    print("="*60)
    
    if results['status'] == 'success':
        print(f"\n[SUCCESS] Analysis completed successfully")
        print(f"  Execution time: {results['execution_time_seconds']} seconds")
        print(f"\n  Total Economic Loss: {results['total_economic_loss_formatted']}")
        print(f"\n  Key Metrics:")
        summary = results['summary']
        print(f"    - Age at Death: {summary['age_at_death']} years")
        print(f"    - Remaining Life Expectancy: {summary['remaining_life_years']:.2f} years")
        print(f"    - Work-Life Expectancy: {summary['worklife_years']:.2f} years")
        print(f"    - Base Annual Salary: ${summary['base_salary']:,.2f}")
        
        print(f"\n  Excel Report: {results['output_file']}")
        print("\n" + "="*60)
    else:
        print(f"\n[FAILED] Analysis failed")
        print(f"  Error: {results.get('error_message', 'Unknown error')}")
        if 'error_traceback' in results:
            print(f"\n  Traceback:\n{results['error_traceback']}")
        print("\n" + "="*60)


if __name__ == "__main__":
    main()

