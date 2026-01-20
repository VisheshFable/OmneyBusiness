# Test Automation Agent - Design Document

## Overview

An intelligent agent system that automates test case execution using Chrome DevTools MCP, captures screenshots, generates comprehensive reports, and creates Python automation scripts.

---

## Architecture

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
├── templates/
│   ├── report_template.html     # HTML report template
│   └── script_template.py       # Python script template
├── config/
│   ├── config.json              # Agent configuration
│   └── selectors.json           # Common UI selectors
└── README.md                     # Usage documentation
```

---

## Workflow

### Phase 1: Test Case Input & Parsing
```
User Input: "TC_09"
    ↓
1. Read TC_09 from Excel (D:\Vishesh\OmneyBusiness\Testcase\OB_Automation.xlsx)
2. Parse test steps, expected results, and test data
3. Display test case details to user for confirmation
```

### Phase 2: Interactive Execution (Chrome DevTools MCP)
```
4. Launch browser using Chrome DevTools MCP
5. For each test step:
   a. Display step description
   b. Execute step interactively
   c. Capture screenshot
   d. Wait for user confirmation/retry
   e. Capture any data (forms, popups, tables)
6. Record all actions, selectors, and data
```

### Phase 3: Report Generation
```
7. Generate comprehensive HTML report:
   - Test case details
   - Step-by-step execution log
   - Screenshots gallery
   - Data verification tables
   - Transaction details (if applicable)
8. Save report to: D:\Vishesh\OmneyBusiness\Reports\TC_XX_DevTools/
9. Display report path to user
```

### Phase 4: Script Generation (After User Approval)
```
10. Analyze recorded actions and generate:
    - Standalone Python script (tc_XX_automation.py)
    - Integration code for omney_business_automation.py
11. Include:
    - Retry logic for dropdowns
    - Multiple selector strategies
    - Error handling
    - Data verification
    - Screenshot capture
12. Show generated code to user for review
```

### Phase 5: Integration (After User Approval)
```
13. Create standalone script:
    Scripts/tc_XX_automation.py
14. Integrate into main framework:
    - Add method to OmneyBusinessAutomation class
    - Update run_all_tests() workflow
    - Add data storage variables
15. Generate README documentation
```

---

## Core Components

### 1. TestCaseReader (`testcase_reader.py`)

**Purpose**: Read and parse test cases from Excel file

**Key Methods**:
```python
class TestCaseReader:
    def __init__(self, excel_path: str)
    def get_testcase(self, tc_id: str) -> dict
    def parse_steps(self, tc_data: dict) -> list[dict]
    def get_credentials(self, user_type: str) -> dict
    def get_test_data(self, tc_id: str) -> dict
```

**Returns**:
```python
{
    "tc_id": "TC_09",
    "test_scenario": "...",
    "steps": [
        {"step_no": 1, "description": "...", "action": "..."},
        ...
    ],
    "expected_result": "...",
    "priority": "High",
    "credentials": {"email": "...", "password": "..."},
    "test_data": {...}
}
```

---

### 2. DevToolsExecutor (`devtools_executor.py`)

**Purpose**: Execute test steps using Chrome DevTools MCP

**Key Methods**:
```python
class DevToolsExecutor:
    def __init__(self, base_url: str, reports_dir: str)

    async def execute_step(self, step: dict) -> dict:
        """Execute a single test step"""

    async def capture_screenshot(self, name: str) -> str:
        """Take screenshot and save"""

    async def capture_form_data(self) -> dict:
        """Capture all form field values"""

    async def capture_table_data(self, table_selector: str) -> list:
        """Capture table content"""

    async def find_element(self, strategies: list) -> str:
        """Try multiple selector strategies"""

    async def wait_for_element(self, selector: str, timeout: int) -> bool:
        """Wait for element with timeout"""
```

**Action Recording**:
```python
{
    "step_no": 4,
    "action": "click",
    "element": "Pay Invoice button",
    "selectors_tried": [
        "text=Pay Invoice",
        "button:has-text('Pay Invoice')",
        "[aria-label='Pay Invoice']"
    ],
    "selector_used": "text=Pay Invoice",
    "screenshot": "TC_09_04_Pay_Invoice_Button.png",
    "duration_ms": 1234,
    "success": True,
    "data_captured": {...}
}
```

---

### 3. ReportGenerator (`report_generator.py`)

**Purpose**: Generate comprehensive HTML reports

**Key Methods**:
```python
class ReportGenerator:
    def __init__(self, template_path: str)

    def generate_report(
        self,
        tc_data: dict,
        execution_results: list[dict],
        screenshots: list[str],
        verification_data: dict = None,
        transaction_data: dict = None
    ) -> str:
        """Generate complete HTML report"""

    def create_verification_table(self, expected: dict, actual: dict) -> str:
        """Create data verification comparison table"""

    def create_screenshot_gallery(self, screenshots: list[str]) -> str:
        """Create screenshot gallery section"""
```

**Report Sections**:
- Test Case Header (ID, Scenario, Priority)
- Execution Summary (Steps, Duration, Status)
- Step-by-Step Results (with screenshots)
- Data Verification Table (if applicable)
- Transaction Success Details (if applicable)
- Screenshots Gallery
- Execution Environment Details

---

### 4. ScriptGenerator (`script_generator.py`)

**Purpose**: Generate Python automation scripts from recorded actions

**Key Methods**:
```python
class ScriptGenerator:
    def __init__(self, template_path: str)

    def generate_standalone_script(
        self,
        tc_data: dict,
        recorded_actions: list[dict],
        output_path: str
    ) -> str:
        """Generate standalone tc_XX_automation.py"""

    def generate_integration_code(
        self,
        tc_data: dict,
        recorded_actions: list[dict]
    ) -> str:
        """Generate code for omney_business_automation.py"""

    def _convert_action_to_code(self, action: dict) -> str:
        """Convert recorded action to Python code"""

    def _add_retry_logic(self, code: str, action: dict) -> str:
        """Wrap code with retry logic"""

    def _add_error_handling(self, code: str) -> str:
        """Add try-except blocks"""
```

**Code Generation Strategy**:
```python
# For click actions
action = {
    "action": "click",
    "selector_used": "text=Pay Invoice",
    "selectors_tried": ["text=Pay Invoice", "button:has-text('Pay')"]
}

# Generates:
"""
# Click Pay Invoice button
pay_invoice_btn = None
for selector in ["text=Pay Invoice", "button:has-text('Pay')"]:
    try:
        pay_invoice_btn = self.page.locator(selector).first
        if pay_invoice_btn.is_visible(timeout=5000):
            break
    except:
        continue

if pay_invoice_btn:
    pay_invoice_btn.click()
    self.page.wait_for_timeout(2000)
else:
    raise Exception("Pay Invoice button not found")
"""
```

---

### 5. Integrator (`integrator.py`)

**Purpose**: Integrate new test case into main automation framework

**Key Methods**:
```python
class Integrator:
    def __init__(self, main_script_path: str)

    def add_method_to_class(
        self,
        method_code: str,
        class_name: str = "OmneyBusinessAutomation"
    ) -> bool:
        """Add new tc_XX method to class"""

    def update_init_method(self, tc_id: str) -> bool:
        """Add data storage variables to __init__"""

    def update_run_all_tests(self, tc_id: str, dependencies: list) -> bool:
        """Add test case to execution flow"""

    def create_readme(self, tc_id: str, tc_data: dict) -> str:
        """Generate TC_XX_README.md"""
```

---

## Configuration

### `config/config.json`
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
    "screenshots": {
        "format": "png",
        "full_page": false
    },
    "retry_logic": {
        "max_attempts": 3,
        "delay_ms": 2000
    }
}
```

### `config/selectors.json`
```json
{
    "common": {
        "login_button": ["text=Log in", "button:has-text('Login')"],
        "email_input": ["input[type='email']", "input[type='text']"],
        "password_input": ["input[type='password']"],
        "dropdown": ["[role='combobox']", "select", "button:has-text('Choose')"],
        "close_button": ["text=Close", "button:has-text('Close')"]
    },
    "invoice": {
        "raise_invoice_btn": ["text=Raise Invoice", "button:has-text('Raise')"],
        "invoice_number": ["input[placeholder*='Invoice']"],
        "client_dropdown": ["text=Choose a client"],
        "approve_button": ["button:has-text('Approve')"]
    },
    "payment": {
        "pay_now_btn": ["text=Pay Now", "button:has-text('Pay')"],
        "pay_invoice_btn": ["text=Pay Invoice"],
        "choose_invoice_dropdown": ["text=Choose invoice"]
    }
}
```

---

## Usage

### Command Line Interface

```bash
# Basic usage
python TestAutomationAgent/agent.py --testcase TC_09

# With options
python TestAutomationAgent/agent.py --testcase TC_09 \
    --headless false \
    --auto-approve false \
    --output-dir custom_reports/

# List available test cases
python TestAutomationAgent/agent.py --list

# Dry run (parse only, no execution)
python TestAutomationAgent/agent.py --testcase TC_09 --dry-run
```

### Interactive Mode

```
$ python TestAutomationAgent/agent.py --testcase TC_09

╔════════════════════════════════════════════════════════════════╗
║          TEST AUTOMATION AGENT - TC_09                         ║
╚════════════════════════════════════════════════════════════════╝

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

[2/5] Executing Test Case (Chrome DevTools MCP)...

[Step 1/8] Login with Client credentials
  ⚙ Navigating to https://qaoneob.remit.in/login
  ⚙ Entering email: ganeshthakurpm@yopmail.com
  ⚙ Entering password: ********
  ⚙ Clicking login button
  📸 Screenshot: TC_09_01_Login_Success.png
  ✓ Step completed successfully

Continue to next step? (yes/retry/skip): yes

[Step 2/8] Navigate to Pending Payables
  ⚙ Scrolling to Pending Payables section
  📸 Screenshot: TC_09_02_Pending_Payables.png
  ✓ Step completed successfully

Continue to next step? (yes/retry/skip): yes

... [continues for all steps] ...

[3/5] Generating HTML Report...
✓ Report saved: D:\Vishesh\OmneyBusiness\Reports\TC_09_DevTools\TC_09_Report.html
✓ Screenshots: 8 files saved

[4/5] Generating Python Scripts...
⚠ This requires your approval. Review execution results above.

Generate automation scripts? (yes/no): yes

  ⚙ Analyzing recorded actions...
  ⚙ Generating standalone script...
  ✓ Created: Scripts/tc_09_automation.py (423 lines)

  ⚙ Generating integration code...
  ✓ Generated method: tc_09_verify_payment_workflow()

[5/5] Integration (Requires Approval)...

Generated files for your review:
  • Scripts/tc_09_automation.py
  • Integration code ready (not yet applied)

Review the generated code. Integrate into main framework? (yes/no): yes

  ⚙ Adding method to OmneyBusinessAutomation class...
  ⚙ Updating __init__ with data storage variables...
  ⚙ Updating run_all_tests() workflow...
  ⚙ Generating TC_09_README.md...
  ✓ Integration complete!

╔════════════════════════════════════════════════════════════════╗
║                    AUTOMATION COMPLETE                          ║
╚════════════════════════════════════════════════════════════════╝

Summary:
  ✓ Test Case: TC_09
  ✓ Status: PASSED
  ✓ Duration: 4 minutes 32 seconds
  ✓ Screenshots: 8
  ✓ Report: TC_09_Report.html
  ✓ Standalone Script: tc_09_automation.py
  ✓ Integration: COMPLETED

Next Steps:
  1. Review report: D:\Vishesh\OmneyBusiness\Reports\TC_09_DevTools\TC_09_Report.html
  2. Test standalone: python Scripts/tc_09_automation.py
  3. Run full suite: python Scripts/omney_business_automation.py
```

---

## Key Features

### 1. Intelligent Element Detection
- Tries multiple selector strategies
- Records which selectors work
- Generates fallback logic in code

### 2. Automatic Retry Logic
- Detects timing-sensitive operations (dropdowns, popups)
- Adds appropriate wait times
- Implements retry loops with exponential backoff

### 3. Data Verification
- Compares form data with test data
- Generates verification tables in reports
- Creates assertion code in scripts

### 4. Screenshot Management
- Captures at every major step
- Names descriptively (TC_XX_01_Action_Description.png)
- Organizes in report-specific folders

### 5. Error Recovery
- Records failures and recovery attempts
- Suggests code improvements
- Provides debugging information

---

## Error Handling Strategy

### During Execution (Chrome DevTools)
```python
try:
    # Execute step
    result = await execute_step(step)
except ElementNotFoundError:
    # Offer options: retry, skip, modify selector, abort
    user_choice = await prompt_user_action()
except TimeoutError:
    # Suggest increasing timeout
    # Retry with longer wait
except Exception as e:
    # Log error, capture screenshot
    # Continue or abort based on step criticality
```

### During Script Generation
- Warn about steps that required manual intervention
- Add TODO comments for manual verification
- Generate conservative timeouts

### During Integration
- Validate syntax before modifying main script
- Create backup before integration
- Rollback on failure

---

## Extensibility

### Adding New Action Types

**1. Add to DevToolsExecutor:**
```python
async def execute_drag_drop(self, from_selector: str, to_selector: str):
    """New drag-drop action"""
    # Implementation
```

**2. Add to ScriptGenerator:**
```python
def _generate_drag_drop_code(self, action: dict) -> str:
    """Generate drag-drop code"""
    # Template
```

### Adding New Report Sections

**1. Add to ReportGenerator:**
```python
def create_custom_section(self, data: dict) -> str:
    """Generate custom report section"""
    # HTML generation
```

### Adding New Integrations

**1. Create new module:**
```python
# modules/jenkins_integrator.py
class JenkinsIntegrator:
    """Integrate with Jenkins CI/CD"""
```

---

## Testing Strategy

### Unit Tests
```
tests/
├── test_testcase_reader.py
├── test_devtools_executor.py
├── test_report_generator.py
├── test_script_generator.py
└── test_integrator.py
```

### Integration Tests
```
tests/integration/
├── test_end_to_end.py
└── test_agent_workflow.py
```

### Test Data
```
tests/fixtures/
├── sample_testcase.xlsx
├── sample_recorded_actions.json
└── expected_outputs/
```

---

## Security Considerations

1. **Credentials**: Never hardcode in generated scripts
2. **File Paths**: Use environment variables for sensitive paths
3. **Generated Code**: Validate and sanitize before execution
4. **Browser Data**: Clear cookies/cache after execution

---

## Performance Optimization

1. **Parallel Execution**: Support running multiple test cases in parallel
2. **Caching**: Cache common selectors and patterns
3. **Lazy Loading**: Load templates and configs on-demand
4. **Screenshot Optimization**: Compress images without quality loss

---

## Future Enhancements

### Phase 2 (Next Release)
- [ ] AI-powered selector suggestion
- [ ] Visual regression testing
- [ ] Test data generation
- [ ] Multi-browser support (Firefox, Safari)
- [ ] Cloud execution (Selenium Grid)

### Phase 3 (Future)
- [ ] Natural language test case input
- [ ] Self-healing test scripts
- [ ] Continuous learning from failures
- [ ] Integration with test management tools (Jira, TestRail)

---

## Dependencies

```
playwright>=1.40.0
pandas>=2.0.0
openpyxl>=3.1.0
jinja2>=3.1.0
click>=8.1.0
colorama>=0.4.6
```

---

## Maintenance

### Updating Templates
- Report template: `templates/report_template.html`
- Script template: `templates/script_template.py`

### Updating Selectors
- Common selectors: `config/selectors.json`
- Add new patterns as application evolves

### Logs
- Agent logs: `logs/agent_YYYYMMDD_HHMMSS.log`
- Execution logs: `logs/execution_TC_XX_YYYYMMDD_HHMMSS.log`

---

## Support & Documentation

- Full documentation: `TestAutomationAgent/README.md`
- API reference: `docs/api_reference.md`
- Troubleshooting: `docs/troubleshooting.md`
- Examples: `examples/`

---

**Version**: 1.0.0
**Last Updated**: 2026-01-12
**Author**: Test Automation Framework Team
