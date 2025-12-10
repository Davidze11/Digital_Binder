# Real-Time Dashboard Guide

## 🎯 Dashboard Features

The Forensic Economics AI system now includes a **real-time progress dashboard** that shows live updates as agents process your analysis.

### Access the Dashboard

**URL:** `http://localhost:8000/dashboard`

### Features

1. **Analysis Overview Panel**
   - Session ID tracking
   - Person name
   - Current status (PENDING, IN_PROGRESS, COMPLETED, ERROR)
   - Progress indicator (X/8 agents)
   - Visual progress bar
   - Current step description

2. **Agent Flow Table**
   - Real-time status for each agent:
     - Person Investigation
     - Federal Reserve
     - Life Expectancy
     - Skoog Table
     - Annual Growth
     - Present Value
     - Excel Report
     - Summary Report
   - Status indicators (PENDING, IN_PROGRESS, COMPLETED, ERROR)
   - Agent messages
   - Agent output details

3. **Generated Files Section**
   - List of all generated files
   - Excel reports
   - Summary reports

4. **Errors & Issues Section**
   - Any errors that occur during analysis
   - Detailed error messages

### How It Works

1. **Start Analysis**: Fill in the form on the dashboard and click "Start Analysis"
2. **Real-Time Updates**: The dashboard automatically refreshes every 3 seconds
3. **Progress Tracking**: Watch each agent complete in real-time
4. **Completion**: When all agents complete, view the final results

### Agent Status Colors

- **Yellow (PENDING)**: Agent hasn't started yet
- **Blue (IN_PROGRESS)**: Agent is currently processing
- **Green (COMPLETED)**: Agent finished successfully
- **Red (ERROR)**: Agent encountered an error

### Example Workflow

1. Open dashboard: `http://localhost:8000/dashboard`
2. Click "Load Example" to populate form
3. Click "Start Analysis"
4. Watch agents process in real-time:
   - Person Investigation → Validates input data
   - Federal Reserve → Fetches discount rate
   - Life Expectancy → Calculates life expectancy
   - Skoog Table → Determines work-life expectancy
   - Annual Growth → Applies salary growth rate
   - Present Value → Calculates discounted values
   - Excel Report → Generates Excel file
   - Summary Report → Creates summary text file
5. View generated files when complete

### Technical Details

- **Update Frequency**: Every 3 seconds
- **API Endpoint**: `/api/progress/<session_id>`
- **Session Management**: Each analysis gets a unique session ID
- **Threading**: Analysis runs in background thread
- **Progress Tracking**: Real-time status updates for all agents

### Status Messages

Each agent provides:
- **Status**: Current state (PENDING, IN_PROGRESS, COMPLETED, ERROR)
- **Message**: Brief description of current action
- **Output**: Detailed output information

### Example Output

**Person Investigation:**
- Status: COMPLETED
- Message: "Person data validated successfully"
- Output: "Validated: John Doe, Age 44, Software Engineer, Salary $120,000"

**Federal Reserve:**
- Status: COMPLETED
- Message: "Current Treasury rate: 4.50%"
- Output: "Treasury Rate: 4.50% (1-Year Constant Maturity, November 2025)"

**Life Expectancy:**
- Status: COMPLETED
- Message: "Life expectancy: 33.2 years"
- Output: "Life Expectancy: 33.2 years (44-year-old Male, CDC 2019)"

### Troubleshooting

**Dashboard not updating?**
- Check browser console for errors
- Verify server is running
- Check network tab for API calls

**Analysis stuck?**
- Check server logs
- Verify all required fields are filled
- Check for errors in the "Errors & Issues" section

**Session not found?**
- Start a new analysis
- Session IDs are unique per analysis
- Sessions are stored in memory (lost on server restart)

---

**Ready to use?** Open `http://localhost:8000/dashboard` and start an analysis!


