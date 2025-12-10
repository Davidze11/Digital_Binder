"""
SupervisorAgent: Oversees all operations and coordinates data flow between agents
"""
from typing import Dict, Any, Optional
import logging
from datetime import datetime
import traceback

from agents.person_data_collector import PersonDataCollectorAgent
from agents.life_expectancy_agent import LifeExpectancyAgent
from agents.worklife_agent import WorklifeRemainingYearsAgent, SkoogTableAgent
from agents.wage_data_agent import WageDataAgent
from agents.discount_agent import FederalReserveRateAgent, DiscountFactorAgent
from agents.calculation_agents import (
    PersonLIFEyrAgentEnhanced,
    CalcFullActualCumAgent,
    CalcPresentCumulPresentValueAgent
)
from agents.excel_generator import ComprehensiveExcelGenerator


class SupervisorAgent:
    """Supervisor agent that coordinates all forensic economic analysis operations."""
    
    def __init__(self, log_level: int = logging.INFO):
        """Initialize supervisor with logging."""
        self.setup_logging(log_level)
        
        # Initialize all agents
        self.person_collector = PersonDataCollectorAgent()
        self.life_expectancy_agent = LifeExpectancyAgent()
        self.worklife_agent = WorklifeRemainingYearsAgent()
        self.skoog_agent = SkoogTableAgent()
        self.wage_agent = WageDataAgent()
        self.fed_rate_agent = FederalReserveRateAgent()
        self.discount_agent = DiscountFactorAgent()
        self.timeline_agent = PersonLIFEyrAgentEnhanced()
        self.earnings_agent = CalcFullActualCumAgent()
        self.pv_agent = CalcPresentCumulPresentValueAgent()
        self.excel_generator = ComprehensiveExcelGenerator()
        
        # Storage for intermediate results
        self.person_profile = None
        self.life_expectancy_data = None
        self.worklife_data = None
        self.wage_data = None
        self.discount_rate_data = None
        self.timeline = None
        self.earnings_table = None
        self.pv_table = None
        
        self.logger = logging.getLogger(__name__)
    
    def setup_logging(self, log_level: int):
        """Setup logging configuration."""
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('forensic_economics.log'),
                logging.StreamHandler()
            ]
        )
    
    def run_analysis(self, input_data: Dict[str, Any], 
                    output_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Run complete forensic economic analysis.
        
        Args:
            input_data: Person data dictionary
            output_path: Optional path for Excel output
            
        Returns:
            Dictionary with analysis results and output path
        """
        try:
            self.logger.info("Starting forensic economic analysis")
            start_time = datetime.now()
            
            # Step 1: Collect and validate person data
            self.logger.info("Step 1: Collecting person data")
            self.person_profile = self.person_collector.collect(input_data)
            self.logger.info(f"Person data collected: {self.person_profile.get('name', 'Unknown')}")
            
            # Step 2: Fetch life expectancy
            self.logger.info("Step 2: Fetching life expectancy")
            self.life_expectancy_data = self.life_expectancy_agent.fetch_life_expectancy(
                self.person_profile
            )
            self.logger.info(f"Life expectancy: {self.life_expectancy_data['remaining_life_expectancy']} years")
            
            # Step 3: Fetch work-life expectancy
            self.logger.info("Step 3: Fetching work-life expectancy")
            self.worklife_data = self.worklife_agent.fetch_worklife_expectancy(
                self.person_profile
            )
            self.logger.info(f"Work-life expectancy: {self.worklife_data['worklife_expectancy']} years")
            
            # Validate work-life doesn't exceed life expectancy
            if self.worklife_data['worklife_expectancy'] > self.life_expectancy_data['remaining_life_expectancy']:
                self.logger.warning(
                    f"Work-life expectancy ({self.worklife_data['worklife_expectancy']}) "
                    f"exceeds remaining life expectancy ({self.life_expectancy_data['remaining_life_expectancy']}). "
                    f"Adjusting work-life to match life expectancy."
                )
                self.worklife_data['worklife_expectancy'] = self.life_expectancy_data['remaining_life_expectancy']
            
            # Step 4: Fetch wage growth data
            self.logger.info("Step 4: Fetching wage growth data")
            self.wage_data = self.wage_agent.fetch_wage_growth_rate(self.person_profile)
            self.logger.info(f"Wage growth rate: {self.wage_data['annual_growth_percent']}%")
            
            # Step 5: Fetch discount rate
            self.logger.info("Step 5: Fetching discount rate")
            self.discount_rate_data = self.fed_rate_agent.fetch_discount_rate('1_year_treasury')
            discount_rate = self.discount_rate_data['discount_rate']
            self.discount_agent.set_discount_rate(discount_rate)
            self.logger.info(f"Discount rate: {self.discount_rate_data['discount_rate_percent']}%")
            
            # Step 6: Build age timeline
            self.logger.info("Step 6: Building age timeline")
            self.timeline = self.timeline_agent.build_timeline(
                self.person_profile,
                self.life_expectancy_data
            )
            self.logger.info(f"Timeline created: {len(self.timeline)} years")
            
            # Step 7: Calculate earnings projections
            self.logger.info("Step 7: Calculating earnings projections")
            self.earnings_table = self.earnings_agent.calculate_earnings(
                self.person_profile,
                self.wage_data,
                self.worklife_data,
                self.timeline
            )
            total_earnings = sum(e['projected_earnings'] for e in self.earnings_table)
            self.logger.info(f"Total projected earnings: ${total_earnings:,.2f}")
            
            # Step 8: Calculate present values
            self.logger.info("Step 8: Calculating present values")
            self.pv_table = self.pv_agent.calculate_present_values(
                self.earnings_table,
                discount_rate
            )
            total_pv = self.pv_agent.get_total_economic_loss()
            self.logger.info(f"Total economic loss (PV): ${total_pv:,.2f}")
            
            # Step 9: Generate Excel report
            self.logger.info("Step 9: Generating Excel report")
            excel_path = self.excel_generator.generate_excel_report(
                self.person_profile,
                self.life_expectancy_data,
                self.worklife_data,
                self.wage_data,
                self.discount_rate_data,
                self.timeline,
                self.earnings_table,
                self.pv_table,
                output_path
            )
            self.logger.info(f"Excel report generated: {excel_path}")
            
            # Calculate execution time
            end_time = datetime.now()
            execution_time = (end_time - start_time).total_seconds()
            
            # Prepare results
            results = {
                'status': 'success',
                'execution_time_seconds': round(execution_time, 2),
                'output_file': excel_path,
                'total_economic_loss': total_pv,
                'total_economic_loss_formatted': f"${total_pv:,.2f}",
                'person_profile': self.person_profile,
                'life_expectancy': self.life_expectancy_data,
                'worklife_expectancy': self.worklife_data,
                'wage_growth': self.wage_data,
                'discount_rate': self.discount_rate_data,
                'summary': {
                    'name': self.person_profile.get('name'),
                    'age_at_death': self.person_profile.get('age_at_death'),
                    'remaining_life_years': self.life_expectancy_data['remaining_life_expectancy'],
                    'worklife_years': self.worklife_data['worklife_expectancy'],
                    'base_salary': self.person_profile.get('annual_salary'),
                    'total_economic_loss': total_pv
                }
            }
            
            self.logger.info(f"Analysis completed successfully in {execution_time:.2f} seconds")
            return results
            
        except Exception as e:
            self.logger.error(f"Error in analysis: {str(e)}")
            self.logger.error(traceback.format_exc())
            return {
                'status': 'error',
                'error_message': str(e),
                'error_traceback': traceback.format_exc()
            }
    
    def get_results(self) -> Dict[str, Any]:
        """Get current analysis results."""
        return {
            'person_profile': self.person_profile,
            'life_expectancy_data': self.life_expectancy_data,
            'worklife_data': self.worklife_data,
            'wage_data': self.wage_data,
            'discount_rate_data': self.discount_rate_data,
            'timeline': self.timeline,
            'earnings_table': self.earnings_table,
            'pv_table': self.pv_table
        }
    
    def validate_inputs(self, input_data: Dict[str, Any]):
        """
        Validate input data before running analysis.
        
        Args:
            input_data: Input data dictionary
            
        Returns:
            Tuple of (is_valid, error_messages) - bool, list
        """
        errors = []
        
        # Check required fields
        required_fields = PersonDataCollectorAgent.REQUIRED_FIELDS
        for field in required_fields:
            if field not in input_data:
                errors.append(f"Missing required field: {field}")
        
        return len(errors) == 0, errors

