"""
Flask web application for Forensic Economics AI System
Runs on localhost:8000 with real-time dashboard
"""
from flask import Flask, render_template, request, jsonify, send_file, Response
import json
import os
import time
import threading
from datetime import datetime
from agents.supervisor_agent import SupervisorAgent
import traceback
import queue

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.getcwd(), 'output')

# Create necessary directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Global storage for analysis progress
analysis_progress = {}
progress_queues = {}


class ProgressTracker:
    """Tracks progress of analysis for real-time updates."""
    
    def __init__(self, session_id):
        self.session_id = session_id
        self.status = "PENDING"
        self.current_step = "Initializing..."
        self.progress = 0
        self.total_agents = 8
        self.agents = {
            "Person Investigation": {"status": "PENDING", "message": "", "output": ""},
            "Federal Reserve": {"status": "PENDING", "message": "", "output": ""},
            "Life Expectancy": {"status": "PENDING", "message": "", "output": ""},
            "Skoog Table": {"status": "PENDING", "message": "", "output": ""},
            "Annual Growth": {"status": "PENDING", "message": "", "output": ""},
            "Present Value": {"status": "PENDING", "message": "", "output": ""},
            "Excel Report": {"status": "PENDING", "message": "", "output": ""},
            "Summary Report": {"status": "PENDING", "message": "", "output": ""}
        }
        self.person_name = ""
        self.generated_files = []
        self.errors = []
    
    def update_agent(self, agent_name, status, message="", output=""):
        """Update agent status."""
        if agent_name in self.agents:
            self.agents[agent_name]["status"] = status
            self.agents[agent_name]["message"] = message
            self.agents[agent_name]["output"] = output
            self._update_progress()
    
    def _update_progress(self):
        """Calculate overall progress."""
        completed = sum(1 for agent in self.agents.values() if agent["status"] == "COMPLETED")
        self.progress = completed
        if completed == self.total_agents:
            self.status = "COMPLETED"
            self.current_step = "Analysis completed successfully"
        elif completed > 0:
            self.status = "IN_PROGRESS"
    
    def to_dict(self):
        """Convert to dictionary for JSON response."""
        return {
            "session_id": self.session_id,
            "person_name": self.person_name,
            "status": self.status,
            "progress": self.progress,
            "total_agents": self.total_agents,
            "current_step": self.current_step,
            "agents": self.agents,
            "generated_files": self.generated_files,
            "errors": self.errors
        }


@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')


@app.route('/dashboard')
def dashboard():
    """Render the dashboard page."""
    return render_template('dashboard.html')


@app.route('/api/progress/<session_id>')
def get_progress(session_id):
    """Get current progress for a session."""
    if session_id in analysis_progress:
        return jsonify(analysis_progress[session_id].to_dict())
    return jsonify({"error": "Session not found"}), 404


@app.route('/api/analyze', methods=['POST'])
def analyze():
    """API endpoint to run forensic economic analysis."""
    try:
        # Get JSON data from request
        data = request.get_json()
        
        if not data:
            return jsonify({
                'status': 'error',
                'message': 'No data provided'
            }), 400
        
        # Create session ID
        session_id = f"{data.get('name', 'unknown').lower().replace(' ', '-')}-{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # Create progress tracker
        tracker = ProgressTracker(session_id)
        tracker.person_name = data.get('name', 'Unknown')
        analysis_progress[session_id] = tracker
        
        # Validate inputs
        try:
            is_valid, errors = supervisor.validate_inputs(data)
            if not is_valid:
                tracker.status = "ERROR"
                tracker.errors = errors
                tracker.current_step = "Validation failed"
                analysis_progress[session_id] = tracker
                return jsonify({
                    'status': 'error',
                    'message': 'Validation failed',
                    'errors': errors,
                    'session_id': session_id
                }), 400
        except Exception as e:
            tracker.status = "ERROR"
            tracker.errors = [str(e)]
            tracker.current_step = "Validation error"
            analysis_progress[session_id] = tracker
            return jsonify({
                'status': 'error',
                'message': 'Validation error',
                'errors': [str(e)],
                'session_id': session_id
            }), 400
        
        # Run analysis in background thread
        thread = threading.Thread(
            target=run_analysis_with_progress,
            args=(data, session_id, tracker)
        )
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'status': 'started',
            'session_id': session_id,
            'message': 'Analysis started'
        })
    
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500


def run_analysis_with_progress(data, session_id, tracker):
    """Run analysis with progress tracking."""
    try:
        # Initialize supervisor
        supervisor = SupervisorAgent()
        
        # Step 1: Person Data Collection
        tracker.update_agent("Person Investigation", "IN_PROGRESS", "Collecting person data...")
        tracker.current_step = "Step 1: Collecting person data"
        person_profile = supervisor.person_collector.collect(data)
        tracker.update_agent(
            "Person Investigation",
            "COMPLETED",
            "Person data validated successfully",
            f"Validated: {person_profile.get('name', 'Unknown')}, Age {person_profile.get('age_at_death', 'N/A')}, {person_profile.get('occupation', 'N/A')}, Salary ${person_profile.get('annual_salary', 0):,.0f}"
        )
        time.sleep(0.5)  # Small delay for visibility
        
        # Step 2: Life Expectancy
        tracker.update_agent("Life Expectancy", "IN_PROGRESS", "Fetching life expectancy...")
        tracker.current_step = "Step 2: Fetching life expectancy"
        life_expectancy_data = supervisor.life_expectancy_agent.fetch_life_expectancy(person_profile)
        tracker.update_agent(
            "Life Expectancy",
            "COMPLETED",
            f"Life expectancy: {life_expectancy_data['remaining_life_expectancy']:.1f} years",
            f"Life Expectancy: {life_expectancy_data['remaining_life_expectancy']:.1f} years ({person_profile.get('age_at_death', 'N/A')}-year-old {person_profile.get('sex', 'N/A')}, CDC 2019)"
        )
        time.sleep(0.5)
        
        # Step 3: Work-Life Expectancy
        tracker.update_agent("Skoog Table", "IN_PROGRESS", "Fetching work-life expectancy...")
        tracker.current_step = "Step 3: Fetching work-life expectancy"
        worklife_data = supervisor.worklife_agent.fetch_worklife_expectancy(person_profile)
        if worklife_data['worklife_expectancy'] > life_expectancy_data['remaining_life_expectancy']:
            worklife_data['worklife_expectancy'] = life_expectancy_data['remaining_life_expectancy']
        tracker.update_agent(
            "Skoog Table",
            "COMPLETED",
            f"Worklife expectancy: {worklife_data['worklife_expectancy']:.1f} years",
            f"Worklife Expectancy: {worklife_data['worklife_expectancy']:.1f} years (median for age {person_profile.get('age_at_death', 'N/A')}, {person_profile.get('education_level', 'N/A')})"
        )
        time.sleep(0.5)
        
        # Step 4: Discount Rate
        tracker.update_agent("Federal Reserve", "IN_PROGRESS", "Fetching discount rate from Federal Reserve H.15...")
        tracker.current_step = "Step 4: Fetching discount rate from Federal Reserve H.15"
        discount_rate_data = supervisor.fed_rate_agent.fetch_discount_rate('1_year_treasury')
        discount_rate = discount_rate_data['discount_rate']
        supervisor.discount_agent.set_discount_rate(discount_rate)
        
        # Build output message with source
        fed_url = discount_rate_data.get('fed_url', 'https://www.federalreserve.gov/releases/h15/current/')
        fetch_status = discount_rate_data.get('fetch_status', 'unknown')
        source_note = f"Source: {fed_url}"
        if fetch_status == 'fallback':
            source_note += " (using estimated rate - parsing limited)"
        
        tracker.update_agent(
            "Federal Reserve",
            "COMPLETED",
            f"Current Treasury rate: {discount_rate_data['discount_rate_percent']:.2f}%",
            f"Treasury Rate: {discount_rate_data['discount_rate_percent']:.2f}% (1-Year Constant Maturity, {datetime.now().strftime('%B %Y')}) - {source_note}"
        )
        time.sleep(0.5)
        
        # Step 5: Wage Growth
        tracker.update_agent("Annual Growth", "IN_PROGRESS", "Fetching wage growth data...")
        tracker.current_step = "Step 5: Fetching wage growth data"
        wage_data = supervisor.wage_agent.fetch_wage_growth_rate(person_profile)
        
        # Build output message with source information
        growth_source = wage_data.get('data_source', 'N/A')
        growth_method = wage_data.get('calculation_method', '')
        if growth_method:
            source_text = f" ({growth_source}, {growth_method})"
        else:
            source_text = f" ({growth_source})"
        
        tracker.update_agent(
            "Annual Growth",
            "COMPLETED",
            "Annual growth rate applied",
            f"Annual Growth Rate: {wage_data['annual_growth_percent']:.1f}% applied for {person_profile.get('occupation', 'N/A')} in {person_profile.get('home_county', 'N/A')}, {person_profile.get('home_state', 'N/A')}{source_text}"
        )
        time.sleep(0.5)
        
        # Step 6: Timeline
        tracker.current_step = "Step 6: Building age timeline"
        timeline = supervisor.timeline_agent.build_timeline(person_profile, life_expectancy_data)
        time.sleep(0.3)
        
        # Step 7: Earnings Projection
        tracker.current_step = "Step 7: Calculating earnings projections"
        earnings_table = supervisor.earnings_agent.calculate_earnings(
            person_profile, wage_data, worklife_data, timeline
        )
        time.sleep(0.3)
        
        # Step 8: Present Value
        tracker.update_agent("Present Value", "IN_PROGRESS", "Calculating present values...")
        tracker.current_step = "Step 8: Calculating present values"
        pv_table = supervisor.pv_agent.calculate_present_values(earnings_table, discount_rate)
        total_pv = supervisor.pv_agent.get_total_economic_loss()
        tracker.update_agent(
            "Present Value",
            "COMPLETED",
            "Present value calculations complete",
            f"Present Value: Calculated for {worklife_data['worklife_expectancy']:.1f}-year projection with {discount_rate_data['discount_rate_percent']:.2f}% discount rate"
        )
        time.sleep(0.5)
        
        # Step 9: Excel Report
        tracker.update_agent("Excel Report", "IN_PROGRESS", "Generating Excel report...")
        tracker.current_step = "Step 9: Generating Excel report"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = person_profile.get('name', 'Person').replace(' ', '_')
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"forensic_economic_report_{name}_{timestamp}.xlsx")
        
        excel_path = supervisor.excel_generator.generate_excel_report(
            person_profile,
            life_expectancy_data,
            worklife_data,
            wage_data,
            discount_rate_data,
            timeline,
            earnings_table,
            pv_table,
            output_path
        )
        filename = os.path.basename(excel_path)
        tracker.generated_files.append(filename)
        tracker.update_agent(
            "Excel Report",
            "COMPLETED",
            "Excel report generated successfully",
            f"Excel Report: {filename}"
        )
        time.sleep(0.5)
        
        # Step 10: Summary Report
        tracker.update_agent("Summary Report", "IN_PROGRESS", "Creating summary report...")
        tracker.current_step = "Step 10: Creating summary report"
        summary_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{name}_Analysis_{timestamp}.txt")
        create_summary_report(summary_path, person_profile, total_pv, life_expectancy_data, worklife_data, wage_data, discount_rate_data)
        summary_filename = os.path.basename(summary_path)
        tracker.generated_files.append(summary_filename)
        tracker.update_agent(
            "Summary Report",
            "COMPLETED",
            "Summary report created successfully",
            f"Summary Report: {summary_filename}"
        )
        
        # Mark as completed
        tracker.status = "COMPLETED"
        tracker.current_step = "Analysis completed successfully"
        
        # Store final results
        analysis_progress[session_id].final_results = {
            'total_economic_loss': total_pv,
            'total_economic_loss_formatted': f"${total_pv:,.2f}",
            'output_file': excel_path,
            'output_filename': filename
        }
        
    except Exception as e:
        tracker.status = "ERROR"
        tracker.current_step = f"Error: {str(e)}"
        tracker.errors.append(str(e))
        tracker.update_agent("Summary Report", "ERROR", f"Analysis failed: {str(e)}")


def create_summary_report(filepath, person_profile, total_pv, life_expectancy_data, worklife_data, wage_data, discount_rate_data):
    """Create a text summary report."""
    with open(filepath, 'w') as f:
        f.write("=" * 60 + "\n")
        f.write("FORENSIC ECONOMIC LOSS ANALYSIS - SUMMARY REPORT\n")
        f.write("=" * 60 + "\n\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(f"Person: {person_profile.get('name', 'Unknown')}\n")
        f.write(f"Age at Death: {person_profile.get('age_at_death', 'N/A')} years\n")
        f.write(f"Occupation: {person_profile.get('occupation', 'N/A')}\n")
        f.write(f"Annual Salary: ${person_profile.get('annual_salary', 0):,.2f}\n\n")
        f.write(f"Life Expectancy: {life_expectancy_data['remaining_life_expectancy']:.2f} years\n")
        f.write(f"Work-Life Expectancy: {worklife_data['worklife_expectancy']:.2f} years\n")
        f.write(f"Annual Growth Rate: {wage_data['annual_growth_percent']:.2f}%\n")
        f.write(f"Discount Rate: {discount_rate_data['discount_rate_percent']:.2f}%\n\n")
        f.write(f"TOTAL ECONOMIC LOSS: {total_pv:,.2f}\n")
        f.write("=" * 60 + "\n")


@app.route('/api/download/<filename>')
def download_file(filename):
    """Download generated Excel file."""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({
                'status': 'error',
                'message': 'File not found'
            }), 404
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


@app.route('/api/example', methods=['GET'])
def get_example():
    """Get example input data."""
    example_data = {
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
    return jsonify(example_data)


@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        'status': 'healthy',
        'service': 'Forensic Economics AI',
        'timestamp': datetime.now().isoformat()
    })


# Initialize supervisor (global)
supervisor = SupervisorAgent()

if __name__ == '__main__':
    port = 8000
    print("=" * 60)
    print("Forensic Economics AI - Web Application")
    print("=" * 60)
    print(f"Server starting on http://localhost:{port}")
    print(f"Dashboard: http://localhost:{port}/dashboard")
    print(f"Press CTRL+C to stop the server")
    print("=" * 60)
    app.run(host='localhost', port=port, debug=True, threaded=True)
