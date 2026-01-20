"""
Integrator Module
================
Integrates new test cases into main automation framework.
"""

import re
from pathlib import Path
from typing import Dict, List


class Integrator:
    """Integrates generated test cases into main automation framework."""

    def __init__(self, main_script_path: str):
        """
        Initialize Integrator.

        Args:
            main_script_path: Path to main automation script
        """
        self.main_script_path = Path(main_script_path)
        if not self.main_script_path.exists():
            raise FileNotFoundError(f"Main script not found: {main_script_path}")

        with open(self.main_script_path, 'r', encoding='utf-8') as f:
            self.script_content = f.read()

        # Create backup
        self.backup_path = self.main_script_path.with_suffix('.py.backup')

    def create_backup(self):
        """Create backup of main script before modification."""
        with open(self.backup_path, 'w', encoding='utf-8') as f:
            f.write(self.script_content)
        print(f"[BACKUP] Created: {self.backup_path.name}")

    def add_method_to_class(
        self,
        method_code: str,
        class_name: str = "OmneyBusinessAutomation"
    ) -> bool:
        """
        Add new test method to automation class.

        Args:
            method_code: Complete method code to add
            class_name: Name of class to add method to

        Returns:
            True if successful, False otherwise
        """
        try:
            # Find the class definition
            class_pattern = rf'class {class_name}.*?:'
            class_match = re.search(class_pattern, self.script_content)

            if not class_match:
                print(f"[ERROR] Class {class_name} not found")
                return False

            # Find the last method in the class
            # Look for the last "def " before the next "class" or end of file
            class_start = class_match.end()

            # Find position to insert (before generate_report method or at end)
            insert_pattern = r'\n    def generate_report\('
            insert_match = re.search(insert_pattern, self.script_content[class_start:])

            if insert_match:
                insert_pos = class_start + insert_match.start()
            else:
                # Insert before the last method
                # Find all method definitions
                method_pattern = r'\n    def \w+\('
                methods = list(re.finditer(method_pattern, self.script_content[class_start:]))
                if methods:
                    # Insert before last method
                    insert_pos = class_start + methods[-1].start()
                else:
                    print("[ERROR] Could not find insertion point")
                    return False

            # Insert the new method
            self.script_content = (
                self.script_content[:insert_pos] +
                '\n' + method_code + '\n' +
                self.script_content[insert_pos:]
            )

            print(f"[SUCCESS] Method added to {class_name}")
            return True

        except Exception as e:
            print(f"[ERROR] Failed to add method: {e}")
            return False

    def update_init_method(self, tc_id: str, init_code: str) -> bool:
        """
        Add data storage variables to __init__ method.

        Args:
            tc_id: Test case ID
            init_code: Initialization code to add

        Returns:
            True if successful, False otherwise
        """
        try:
            # Find __init__ method
            init_pattern = r'def __init__\(self.*?\):'
            init_match = re.search(init_pattern, self.script_content)

            if not init_match:
                print("[ERROR] __init__ method not found")
                return False

            # Find the end of __init__ (next method definition)
            init_start = init_match.end()
            next_method_pattern = r'\n    def \w+\('
            next_method_match = re.search(next_method_pattern, self.script_content[init_start:])

            if next_method_match:
                # Insert before next method
                insert_pos = init_start + next_method_match.start()

                self.script_content = (
                    self.script_content[:insert_pos] +
                    '\n' + init_code + '\n' +
                    self.script_content[insert_pos:]
                )

                print(f"[SUCCESS] Init variables added for {tc_id}")
                return True
            else:
                print("[ERROR] Could not find insertion point in __init__")
                return False

        except Exception as e:
            print(f"[ERROR] Failed to update __init__: {e}")
            return False

    def update_run_all_tests(
        self,
        tc_id: str,
        method_name: str,
        dependencies: List[str] = None
    ) -> bool:
        """
        Add test case to run_all_tests execution flow.

        Args:
            tc_id: Test case ID
            method_name: Name of method to call
            dependencies: List of prerequisite test case IDs

        Returns:
            True if successful, False otherwise
        """
        try:
            # Find run_all_tests method
            run_all_pattern = r'def run_all_tests\(self\):'
            run_all_match = re.search(run_all_pattern, self.script_content)

            if not run_all_match:
                print("[ERROR] run_all_tests method not found")
                return False

            # Find where to insert (before report generation or at end)
            method_start = run_all_match.end()
            report_gen_pattern = r'# Generate final report'
            report_match = re.search(report_gen_pattern, self.script_content[method_start:])

            if report_match:
                insert_pos = method_start + report_match.start()
            else:
                # Find end of method
                next_method_pattern = r'\n    def \w+\('
                next_method = re.search(next_method_pattern, self.script_content[method_start:])
                if next_method:
                    insert_pos = method_start + next_method.start()
                else:
                    print("[ERROR] Could not find insertion point in run_all_tests")
                    return False

            # Generate execution code
            dependency_code = ""
            if dependencies:
                dep_list = ", ".join(dependencies)
                dependency_code = f"        # Requires: {dep_list}\n"

            execution_code = f'''
        # Execute {tc_id}
{dependency_code}        print(f"\\n[EXECUTING] {tc_id}")
        try:
            result = self.{method_name}
            self.test_results.append({{
                'tc_id': '{tc_id}',
                'status': 'PASSED' if result else 'FAILED',
                'timestamp': datetime.now().isoformat()
            }})
        except Exception as e:
            print(f"[ERROR] {tc_id} failed: {{e}}")
            self.test_results.append({{
                'tc_id': '{tc_id}',
                'status': 'FAILED',
                'timestamp': datetime.now().isoformat()
            }})

'''

            self.script_content = (
                self.script_content[:insert_pos] +
                execution_code +
                self.script_content[insert_pos:]
            )

            print(f"[SUCCESS] {tc_id} added to run_all_tests")
            return True

        except Exception as e:
            print(f"[ERROR] Failed to update run_all_tests: {e}")
            return False

    def save_changes(self) -> bool:
        """
        Save changes to main script.

        Returns:
            True if successful, False otherwise
        """
        try:
            with open(self.main_script_path, 'w', encoding='utf-8') as f:
                f.write(self.script_content)

            print(f"[SUCCESS] Changes saved to {self.main_script_path.name}")
            return True

        except Exception as e:
            print(f"[ERROR] Failed to save changes: {e}")
            return False

    def rollback(self) -> bool:
        """
        Rollback changes using backup.

        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.backup_path.exists():
                print("[ERROR] Backup file not found")
                return False

            with open(self.backup_path, 'r', encoding='utf-8') as f:
                backup_content = f.read()

            with open(self.main_script_path, 'w', encoding='utf-8') as f:
                f.write(backup_content)

            print(f"[SUCCESS] Rolled back to backup")
            return True

        except Exception as e:
            print(f"[ERROR] Failed to rollback: {e}")
            return False

    def create_readme(self, tc_id: str, tc_data: Dict, output_path: str) -> str:
        """
        Generate README documentation for test case.

        Args:
            tc_id: Test case ID
            tc_data: Test case data dictionary
            output_path: Path to save README

        Returns:
            Path to generated README
        """
        test_scenario = tc_data.get('test_scenario', '')
        priority = tc_data.get('priority', 'Medium')
        steps = tc_data.get('test_steps', '')
        expected_result = tc_data.get('expected_result', '')

        readme_content = f'''# {tc_id}: {test_scenario}

## Overview

Automated test case for: {test_scenario}

## Test Case Details

**Test Case ID**: {tc_id}
**Priority**: {priority}
**Expected Result**: {expected_result}

## Test Steps

{steps}

## Files Created

### 1. Standalone Script
**File**: `{tc_id.lower()}_automation.py`
**Location**: `D:\\Vishesh\\OmneyBusiness\\Scripts\\`

This is a standalone Python script that can run {tc_id} independently.

#### Usage:
```bash
python Scripts/{tc_id.lower()}_automation.py
```

### 2. Integrated Script
**File**: `omney_business_automation.py` (updated)
**Location**: `D:\\Vishesh\\OmneyBusiness\\Scripts\\`

{tc_id} has been integrated into the main automation framework.

#### Usage:
```bash
# Run all test cases including {tc_id}
python Scripts/omney_business_automation.py

# Run only specific test cases
python Scripts/omney_business_automation.py --tc03-only
```

## Reports Generated

### Standalone Script Reports
- Location: `D:\\Vishesh\\OmneyBusiness\\Reports\\{tc_id}_Standalone\\`
- Files:
  - `{tc_id}_*_<timestamp>.png` (screenshots)
  - `{tc_id}_Result_<timestamp>.json` (test results)

### Integrated Script Reports
- Location: `D:\\Vishesh\\OmneyBusiness\\Reports\\Python_Automation\\`
- Files:
  - All screenshots with `{tc_id}_` prefix
  - Consolidated HTML report with all test cases

## Prerequisites

Before running {tc_id}:
1. Install dependencies: `pip install playwright pandas openpyxl`
2. Install browser: `playwright install chromium`
3. Ensure test data is available in Excel file
4. Valid credentials must be configured

## Dependencies

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

## Troubleshooting

### Issue: "Module not found"
**Solution**: Install dependencies using pip command above

### Issue: "Browser not found"
**Solution**: Run `playwright install chromium`

## Version History

- **v1.0** ({datetime.now().strftime('%Y-%m-%d')}): Initial implementation
  - Created standalone script
  - Integrated into main automation framework
  - Generated comprehensive reports

## Author

**Automation Framework**: Omney Business Test Automation
**{tc_id} Implementation**: {datetime.now().strftime('%Y-%m-%d')}
**Automation Method**: Chrome DevTools MCP + Python Playwright

---

For questions or issues, please refer to the main automation documentation.
'''

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(readme_content)

        return output_path
