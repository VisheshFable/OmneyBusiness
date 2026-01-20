"""
TestCaseReader Module
====================
Reads and parses test cases from Excel file.
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional


class TestCaseReader:
    """Reads test cases from Excel and provides structured data."""

    def __init__(self, excel_path: str):
        """
        Initialize TestCaseReader.

        Args:
            excel_path: Path to Excel file containing test cases
        """
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        self.df = pd.read_excel(self.excel_path)
        self._validate_excel_structure()

    def _validate_excel_structure(self):
        """Validate that Excel has required columns."""
        required_columns = ['TC_ID', 'Test Scenario']
        for col in required_columns:
            if col not in self.df.columns:
                raise ValueError(f"Missing required column: {col}")

    def list_testcases(self) -> List[str]:
        """
        Get list of all available test case IDs.

        Returns:
            List of test case IDs
        """
        return self.df['TC_ID'].dropna().unique().tolist()

    def get_testcase(self, tc_id: str) -> Optional[Dict]:
        """
        Get test case data by ID.

        Args:
            tc_id: Test case ID (e.g., 'TC_09')

        Returns:
            Dictionary containing test case data, or None if not found
        """
        tc_row = self.df[self.df['TC_ID'] == tc_id]

        if tc_row.empty:
            return None

        tc_row = tc_row.iloc[0]

        # Handle "Test Steps " with trailing space
        test_steps_col = 'Test Steps ' if 'Test Steps ' in self.df.columns else 'Test Steps'

        testcase_data = {
            'tc_id': tc_id,
            'test_scenario': str(tc_row.get('Test Scenario', '')),
            'test_steps': str(tc_row.get(test_steps_col, '')),
            'expected_result': str(tc_row.get('Expected Result', '')),
            'priority': str(tc_row.get('Priority', 'Medium')),
            'test_data': {}
        }

        # Parse additional data columns
        excluded_cols = ['TC_ID', 'Test Scenario', 'Test Steps', 'Test Steps ', 'Expected Result', 'Priority', 'Test Data']
        for col in self.df.columns:
            if col not in excluded_cols:
                value = tc_row.get(col)
                if pd.notna(value):
                    testcase_data['test_data'][col] = str(value)

        return testcase_data

    def parse_steps(self, tc_data: Dict) -> List[Dict]:
        """
        Parse test steps from test case data.

        Args:
            tc_data: Test case data dictionary

        Returns:
            List of step dictionaries
        """
        steps_text = tc_data.get('test_steps', '')
        if not steps_text or steps_text == 'nan':
            return []

        steps = []
        lines = steps_text.split('\n')

        step_no = 1
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Try to extract step number if present
            if line[0].isdigit() and ('.' in line or ')' in line):
                parts = line.split('.', 1) if '.' in line else line.split(')', 1)
                if len(parts) == 2:
                    try:
                        step_no = int(parts[0].strip())
                        description = parts[1].strip()
                    except ValueError:
                        description = line
                else:
                    description = line
            else:
                description = line

            steps.append({
                'step_no': step_no,
                'description': description,
                'action': self._infer_action_type(description)
            })
            step_no += 1

        return steps

    def _infer_action_type(self, description: str) -> str:
        """
        Infer action type from step description.

        Args:
            description: Step description text

        Returns:
            Action type string
        """
        description_lower = description.lower()

        if any(word in description_lower for word in ['login', 'log in', 'sign in']):
            return 'login'
        elif any(word in description_lower for word in ['click', 'press', 'tap']):
            return 'click'
        elif any(word in description_lower for word in ['enter', 'fill', 'type', 'input']):
            return 'input'
        elif any(word in description_lower for word in ['select', 'choose', 'pick']):
            return 'select'
        elif any(word in description_lower for word in ['verify', 'check', 'compare', 'validate']):
            return 'verify'
        elif any(word in description_lower for word in ['capture', 'fetch', 'get']):
            return 'capture'
        elif any(word in description_lower for word in ['navigate', 'go to', 'open']):
            return 'navigate'
        elif any(word in description_lower for word in ['close', 'dismiss']):
            return 'close'
        elif any(word in description_lower for word in ['scroll']):
            return 'scroll'
        elif any(word in description_lower for word in ['wait']):
            return 'wait'
        else:
            return 'custom'

    def get_credentials(self, user_type: str = 'Vendor_Individual') -> Dict:
        """
        Get credentials for specified user type.

        Args:
            user_type: Type of user (Vendor_Individual, Client_Business, etc.)

        Returns:
            Dictionary with email and password
        """
        credentials = {
            'Vendor_Individual': {
                'email': 'visheshindindia@yopmail.com',
                'password': 'Password@2'
            },
            'Client_Business': {
                'email': 'ganeshthakurpm@yopmail.com',
                'password': 'Password@2'
            }
        }

        return credentials.get(user_type, {})

    def get_test_data(self, tc_id: str) -> Dict:
        """
        Get test data for specified test case.

        Args:
            tc_id: Test case ID

        Returns:
            Dictionary containing test data
        """
        tc_data = self.get_testcase(tc_id)
        if tc_data:
            return tc_data.get('test_data', {})
        return {}
