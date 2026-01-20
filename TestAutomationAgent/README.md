# Test Automation Agent

**Intelligent agent system for automated test case execution using Chrome DevTools MCP**

---

## Overview

The Test Automation Agent is an intelligent system that:
1. Reads test cases from Excel
2. Guides you through manual execution using Chrome DevTools MCP
3. Captures screenshots at every step
4. Generates comprehensive HTML reports
5. Creates Python automation scripts (standalone + integrated)
6. Integrates new test cases into main automation framework

---

## Features

✓ **Excel-Driven**: Reads test cases directly from your Excel file
✓ **Interactive Execution**: Step-by-step guided manual testing with Chrome DevTools MCP
✓ **Screenshot Capture**: Automatically captures screenshots for each step
✓ **HTML Reports**: Beautiful, comprehensive reports with all execution details
✓ **Script Generation**: Converts manual actions into Python automation scripts
✓ **Framework Integration**: Automatically integrates into existing automation framework
✓ **Retry Logic**: Generated scripts include robust retry logic and error handling
✓ **Documentation**: Auto-generates README for each test case

---

## Installation

### Prerequisites

- Python 3.8+
- Playwright
- Excel file with test cases
- Chrome browser

### Install Dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

---

## Quick Start

### 1. List Available Test Cases

```bash
cd D:\Vishesh\OmneyBusiness\TestAutomationAgent
python agent.py --list
```

**Output:**
```
Available Test Cases:

  TC_01 [High]: To check if URL is working
  TC_02 [High]: To check if user is able to Login using valid credentials
  TC_03 [High]: To check if user can navigate to Raise Invoice page and Create a Invoice
  ...
  TC_09 [High]: Your new test case

Total: 9 test cases
```

### 2. Run Agent for Specific Test Case

```bash
python agent.py --testcase TC_09
```

### 3. Follow Interactive Prompts

The agent will:
1. Display test case details
2. Ask for confirmation to proceed
3. Guide you through each step
4. Capture screenshots
5. Generate report
6. Create automation scripts
7. Integrate into framework (with your approval)

---

## Usage

### Basic Command

```bash
python agent.py --testcase <TC_ID>
```

### With Options

```bash
# Auto-approve all prompts (non-interactive)
python agent.py --testcase TC_09 --auto-approve

# Use custom configuration file
python agent.py --testcase TC_09 --config custom_config.json

# List all test cases
python agent.py --list
```

### Command Line Arguments

| Argument | Short | Description |
|----------|-------|-------------|
| `--testcase` | `-t` | Test case ID to automate (e.g., TC_09) |
| `--list` | `-l` | List all available test cases |
| `--auto-approve` | `-a` | Auto-approve all prompts (non-interactive) |
| `--config` | `-c` | Path to configuration file |
| `--help` | `-h` | Show help message |

---

## Interactive Workflow

### Phase 1: Test Case Reading

```
[1/5] Reading Test Case from Excel...
✓ Test Case ID: TC_09
✓ Scenario: To verify invoice payment workflow
✓ Priority: High
✓ Steps: 8

Test Case Details:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Step 1: Login with Client credentials
Step 2: Navigate to Pending Payables
Step 3: Click Approve on invoice
Step 4: Click Pay Now
Step 5: Verify form data
Step 6: Submit payment
Step 7: Capture transaction details
Step 8: Close success popup
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Proceed with execution? (yes/no): yes
```

### Phase 2: Manual Execution (Chrome DevTools MCP)

```
[2/5] Executing Test Case (Chrome DevTools MCP)...
───────────────────────────────────────────────────────────────────
⚠ IMPORTANT: This phase requires Chrome DevTools MCP integration
⚠ The agent will guide you through manual execution
⚠ Confirm each step after completion
───────────────────────────────────────────────────────────────────

[Step 1/8] Login with Client credentials
  ⚙ Navigating to https://qaoneob.remit.in/login
  ⚙ Entering email: ganeshthakurpm@yopmail.com
  ⚙ Entering password: ********
  ⚙ Clicking login button
  📸 Screenshot: TC_09_01_Login_Success.png
  ✓ Step completed successfully

Continue to next step? (yes/no): yes

[Step 2/8] Navigate to Pending Payables
  ⚙ Scrolling to Pending Payables section
  📸 Screenshot: TC_09_02_Pending_Payables.png
  ✓ Step completed successfully

Continue to next step? (yes/no): yes

... [continues for all steps] ...
```

### Phase 3: Report Generation

```
[3/5] Generating HTML Report...
✓ Report saved: D:\Vishesh\OmneyBusiness\Reports\TC_09_DevTools\TC_09_Report.html
✓ Screenshots: 8 files saved
```

### Phase 4: Script Generation

```
[4/5] Generating Python Scripts...
⚠ This requires your approval. Review execution results above.

Generate automation scripts? (yes/no): yes

  ⚙ Analyzing recorded actions...
  ⚙ Generating standalone script...
  ✓ Created: tc_09_automation.py (423 lines)

  ⚙ Generating integration code...
  ✓ Generated method: tc_09()
```

### Phase 5: Integration

```
[5/5] Integration (Requires Approval)...

Generated files for your review:
  • Scripts/tc_09_automation.py
  • Integration code ready (not yet applied)

Review the generated code. Integrate into main framework? (yes/no): yes

Integrating into main automation framework...
[BACKUP] Created: omney_business_automation.py.backup
[SUCCESS] Method added to OmneyBusinessAutomation
[SUCCESS] Init variables added for TC_09
[SUCCESS] Changes saved to omney_business_automation.py
✓ Integration complete!
```

### Final Summary

```
╔════════════════════════════════════════════════════════════════╗
║                    AUTOMATION COMPLETE                          ║
╚════════════════════════════════════════════════════════════════╝

Summary:
  ✓ Test Case: TC_09
  ✓ Status: PASSED
  ✓ Duration: 4.32 seconds
  ✓ Screenshots: 8
  ✓ Report: TC_09_Report_20260112_210530.html
  ✓ Standalone Script: tc_09_automation.py
  ✓ Integration: COMPLETED

Next Steps:
  1. Review report: D:\Vishesh\OmneyBusiness\Reports\TC_09_DevTools\TC_09_Report.html
  2. Test standalone: python Scripts/tc_09_automation.py
  3. Run full suite: python Scripts/omney_business_automation.py
```

---

## Directory Structure

```
TestAutomationAgent/
├── agent.py                      # Main orchestrator
├── modules/
│   ├── __init__.py
│   ├── testcase_reader.py       # Excel test case parser
│   ├── devtools_executor.py     # Chrome DevTools automation
│   ├── report_generator.py      # HTML report generation
│   ├── script_generator.py      # Python script creation
│   └── integrator.py            # Framework integration
├── config/
│   ├── config.json              # Agent configuration
│   └── selectors.json           # Common UI selectors
├── AGENT_DESIGN.md              # Design document
└── README.md                     # This file
```

---

## Configuration

### config/config.json

```json
{
  "excel_path": "D:\\Vishesh\\OmneyBusiness\\Testcase\\OB_Automation.xlsx",
  "base_url": "https://qaoneob.remit.in",
  "reports_base_dir": "D:\\Vishesh\\OmneyBusiness\\Reports",
  "scripts_dir": "D:\\Vishesh\\OmneyBusiness\\Scripts",
  "main_script": "omney_business_automation.py",
  "browser": {
    "headless": false,
    "slow_mo": 500,
    "timeout": 30000
  },
  "retry_logic": {
    "max_attempts": 3,
    "delay_ms": 2000
  }
}
```

### Customizing Configuration

1. Copy `config/config.json` to `config/my_config.json`
2. Modify settings as needed
3. Run with: `python agent.py --testcase TC_09 --config config/my_config.json`

---

## Generated Files

### For Each Test Case, Agent Generates:

1. **HTML Report**
   - Location: `Reports/<TC_ID>_DevTools/`
   - Contains: Execution details, screenshots, verification tables

2. **Standalone Python Script**
   - Location: `Scripts/<tc_id>_automation.py`
   - Can run independently
   - Includes retry logic and error handling

3. **README Documentation**
   - Location: `Scripts/<TC_ID>_README.md`
   - Usage instructions, dependencies, troubleshooting

4. **Integration into Main Framework**
   - Updated: `Scripts/omney_business_automation.py`
   - Backup created automatically
   - Method added to class
   - Integrated into test execution flow

5. **Recorded Actions JSON**
   - Location: `Reports/<TC_ID>_DevTools/<TC_ID>_recorded_actions.json`
   - Contains all actions, selectors, and metadata

---

## How It Works

### 1. Test Case Reading

The agent reads test cases from Excel using `TestCaseReader`:
- Parses test steps
- Infers action types (click, input, select, verify, etc.)
- Loads test data and credentials

### 2. Interactive Execution

`DevToolsExecutor` guides manual execution:
- Displays step descriptions
- Records all actions and selectors tried
- Captures screenshots automatically
- Allows retry on failures

### 3. Report Generation

`ReportGenerator` creates comprehensive HTML reports:
- Test case header with status banner
- Step-by-step execution results
- Data verification tables
- Screenshot gallery
- Transaction details (if applicable)

### 4. Script Generation

`ScriptGenerator` converts manual actions to code:
- Analyzes recorded actions
- Generates Python code with retry logic
- Creates both standalone and integration versions
- Includes error handling and timeouts

### 5. Framework Integration

`Integrator` adds new test case to framework:
- Creates backup of main script
- Adds method to automation class
- Updates __init__ with data storage variables
- Integrates into test execution flow
- Generates README documentation

---

## Best Practices

### When Using the Agent:

1. **Start with Simple Test Cases**: Test the agent with straightforward test cases first

2. **Review Generated Code**: Always review generated scripts before running

3. **Test Standalone First**: Test standalone script before relying on integration

4. **Keep Backups**: Agent creates backups automatically, but keep your own too

5. **Update Selectors**: Add common selectors to `config/selectors.json` for reuse

6. **Document Manually**: Add comments to generated code for complex logic

### Excel File Format:

Ensure your Excel file has these columns:
- **Test Case ID** (e.g., TC_09)
- **Test Scenario** (description)
- **Test Steps** (numbered steps, one per line)
- **Expected Result**
- **Priority** (High/Medium/Low)

---

## Troubleshooting

### Issue: "Excel file not found"
**Solution**: Check `excel_path` in config.json points to correct file

### Issue: "Test case not found"
**Solution**: Run `python agent.py --list` to see available test cases

### Issue: "Module not found"
**Solution**: Install dependencies with `pip install -r requirements.txt`

### Issue: "Integration failed"
**Solution**: Backup is automatically created. Restore with:
```bash
copy Scripts\omney_business_automation.py.backup Scripts\omney_business_automation.py
```

### Issue: "Generated script has syntax errors"
**Solution**: Review recorded actions JSON and manually fix generated code

---

## Advanced Usage

### Custom Action Recording

Edit `modules/devtools_executor.py` to add custom action types:

```python
async def _execute_custom_action(self, step: Dict, result: Dict):
    """Execute your custom action."""
    # Your implementation
```

### Custom Report Sections

Edit `modules/report_generator.py` to add custom report sections:

```python
def _create_custom_section(self, data: Dict) -> str:
    """Generate custom HTML section."""
    # Your implementation
```

### Custom Script Templates

Modify `modules/script_generator.py` to customize generated code:

```python
def _convert_action_to_code(self, action: Dict, step_no: int) -> str:
    """Customize code generation."""
    # Your implementation
```

---

## Limitations

### Current Limitations:

1. **Chrome DevTools MCP Integration**: Manual execution guidance only (not fully automated)
2. **Action Inference**: May not correctly infer all action types from descriptions
3. **Selector Generation**: Limited to common patterns
4. **Complex Workflows**: May need manual code adjustments for complex scenarios

### Future Enhancements:

- [ ] Full Chrome DevTools MCP automation (remove manual steps)
- [ ] AI-powered selector suggestion
- [ ] Visual regression testing
- [ ] Multi-browser support
- [ ] Cloud execution (Selenium Grid)

---

## Support

### Documentation:
- **Design Document**: See `AGENT_DESIGN.md` for architecture details
- **API Reference**: See inline docstrings in each module

### Getting Help:
- Check `Reports/<TC_ID>_DevTools/<TC_ID>_recorded_actions.json` for debugging
- Review generated code for issues
- Consult backup files if integration fails

---

## Version History

### v1.0.0 (2026-01-12)
- Initial release
- Excel test case reading
- Interactive execution with Chrome DevTools MCP
- HTML report generation
- Python script generation (standalone + integration)
- Framework integration
- Comprehensive documentation

---

## License

Proprietary - Omney Business Test Automation Framework

---

## Author

**Framework**: Omney Business Test Automation Team
**Date**: 2026-01-12
**Contact**: automation-team@omneyBusiness.com

---

## Acknowledgments

- Chrome DevTools MCP for browser automation
- Playwright for Python automation framework
- Colorama for colored terminal output

---

**Happy Automating! 🚀**
