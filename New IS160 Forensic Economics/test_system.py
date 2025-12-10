"""
Test script for Forensic Economics AI System
"""
import json
from agents.supervisor_agent import SupervisorAgent


def test_basic_analysis():
    """Test basic forensic economic analysis."""
    print("Testing Forensic Economics AI System...")
    print("="*60)
    
    # Test input data
    test_input = {
        "name": "Jane Smith",
        "dob": "1975-06-10",
        "dod": "2023-12-15",
        "occupation": "Nurse",
        "annual_salary": 85000,
        "sex": "Female",
        "education_level": "Bachelor's",
        "home_county": "San Francisco",
        "home_state": "California",
        "status": "Active"
    }
    
    # Initialize supervisor
    supervisor = SupervisorAgent()
    
    # Validate inputs
    is_valid, errors = supervisor.validate_inputs(test_input)
    if not is_valid:
        print("Validation errors:")
        for error in errors:
            print(f"  - {error}")
        return False
    
    print("Input validation: PASSED")
    
    # Run analysis
    print("\nRunning analysis...")
    results = supervisor.run_analysis(test_input)
    
    if results['status'] == 'success':
        print("\n[SUCCESS] Analysis completed successfully!")
        print(f"  Total Economic Loss: {results['total_economic_loss_formatted']}")
        print(f"  Output file: {results['output_file']}")
        print(f"  Execution time: {results['execution_time_seconds']} seconds")
        return True
    else:
        print(f"\n[FAILED] Analysis failed: {results.get('error_message', 'Unknown error')}")
        return False


if __name__ == "__main__":
    test_basic_analysis()

