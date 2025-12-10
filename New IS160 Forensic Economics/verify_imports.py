"""
Quick script to verify all imports work correctly
"""
import sys

def verify_imports():
    """Verify all module imports."""
    print("Verifying imports...")
    errors = []
    
    try:
        from agents.supervisor_agent import SupervisorAgent
        print("[OK] SupervisorAgent imported successfully")
    except Exception as e:
        errors.append(f"SupervisorAgent: {e}")
        print(f"[ERROR] SupervisorAgent import failed: {e}")
    
    try:
        from agents.person_data_collector import PersonDataCollectorAgent
        print("[OK] PersonDataCollectorAgent imported successfully")
    except Exception as e:
        errors.append(f"PersonDataCollectorAgent: {e}")
        print(f"[ERROR] PersonDataCollectorAgent import failed: {e}")
    
    try:
        from agents.life_expectancy_agent import LifeExpectancyAgent
        print("[OK] LifeExpectancyAgent imported successfully")
    except Exception as e:
        errors.append(f"LifeExpectancyAgent: {e}")
        print(f"[ERROR] LifeExpectancyAgent import failed: {e}")
    
    try:
        from agents.worklife_agent import WorklifeRemainingYearsAgent
        print("[OK] WorklifeRemainingYearsAgent imported successfully")
    except Exception as e:
        errors.append(f"WorklifeRemainingYearsAgent: {e}")
        print(f"[ERROR] WorklifeRemainingYearsAgent import failed: {e}")
    
    try:
        from agents.wage_data_agent import WageDataAgent
        print("[OK] WageDataAgent imported successfully")
    except Exception as e:
        errors.append(f"WageDataAgent: {e}")
        print(f"[ERROR] WageDataAgent import failed: {e}")
    
    try:
        from agents.discount_agent import FederalReserveRateAgent, DiscountFactorAgent
        print("[OK] Discount agents imported successfully")
    except Exception as e:
        errors.append(f"Discount agents: {e}")
        print(f"[ERROR] Discount agents import failed: {e}")
    
    try:
        from agents.calculation_agents import (
            PersonLIFEyrAgentEnhanced,
            CalcFullActualCumAgent,
            CalcPresentCumulPresentValueAgent
        )
        print("[OK] Calculation agents imported successfully")
    except Exception as e:
        errors.append(f"Calculation agents: {e}")
        print(f"[ERROR] Calculation agents import failed: {e}")
    
    try:
        from agents.excel_generator import ComprehensiveExcelGenerator
        print("[OK] ComprehensiveExcelGenerator imported successfully")
    except Exception as e:
        errors.append(f"ComprehensiveExcelGenerator: {e}")
        print(f"[ERROR] ComprehensiveExcelGenerator import failed: {e}")
    
    if errors:
        print(f"\n[FAILED] {len(errors)} import error(s) found")
        return False
    else:
        print("\n[SUCCESS] All imports successful!")
        return True

if __name__ == "__main__":
    success = verify_imports()
    sys.exit(0 if success else 1)
