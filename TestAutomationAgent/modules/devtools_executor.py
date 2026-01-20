"""
DevToolsExecutor Module
=======================
Executes test steps using Chrome DevTools MCP tools.
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any


class DevToolsExecutor:
    """Executes test steps interactively using Chrome DevTools MCP."""

    def __init__(self, base_url: str, reports_dir: str, mcp_tools: Dict):
        """
        Initialize DevToolsExecutor.

        Args:
            base_url: Base URL of application under test
            reports_dir: Directory to save screenshots and reports
            mcp_tools: Dictionary of Chrome DevTools MCP tool functions
        """
        self.base_url = base_url
        self.reports_dir = Path(reports_dir)
        self.reports_dir.mkdir(parents=True, exist_ok=True)

        self.mcp_tools = mcp_tools
        self.recorded_actions = []
        self.screenshots = []
        self.execution_start_time = None

    def start_execution(self):
        """Start execution tracking."""
        self.execution_start_time = datetime.now()
        self.recorded_actions = []
        self.screenshots = []

    async def execute_step(self, step: Dict, tc_id: str) -> Dict:
        """
        Execute a single test step.

        Args:
            step: Step dictionary with description and action type
            tc_id: Test case ID for naming

        Returns:
            Dictionary with execution results
        """
        step_no = step['step_no']
        description = step['description']
        action_type = step['action']

        print(f"\n[Step {step_no}] {description}")
        print(f"  Action Type: {action_type}")

        start_time = time.time()

        result = {
            'step_no': step_no,
            'description': description,
            'action_type': action_type,
            'selectors_tried': [],
            'selector_used': None,
            'screenshot': None,
            'duration_ms': 0,
            'success': False,
            'data_captured': {},
            'error': None
        }

        try:
            # Execute based on action type
            if action_type == 'navigate':
                await self._execute_navigate(step, result)
            elif action_type == 'login':
                await self._execute_login(step, result)
            elif action_type == 'click':
                await self._execute_click(step, result)
            elif action_type == 'input':
                await self._execute_input(step, result)
            elif action_type == 'select':
                await self._execute_select(step, result)
            elif action_type == 'verify':
                await self._execute_verify(step, result)
            elif action_type == 'capture':
                await self._execute_capture(step, result)
            elif action_type == 'scroll':
                await self._execute_scroll(step, result)
            elif action_type == 'wait':
                await self._execute_wait(step, result)
            elif action_type == 'close':
                await self._execute_close(step, result)
            else:
                # Custom action - user guided
                await self._execute_custom(step, result)

            result['success'] = True

            # Capture screenshot after successful execution
            screenshot_name = f"{tc_id}_{step_no:02d}_{self._sanitize_filename(description)}"
            screenshot_path = await self.capture_screenshot(screenshot_name)
            result['screenshot'] = screenshot_path

        except Exception as e:
            result['error'] = str(e)
            print(f"  ✗ Error: {e}")

            # Capture error screenshot
            screenshot_name = f"{tc_id}_{step_no:02d}_ERROR"
            screenshot_path = await self.capture_screenshot(screenshot_name)
            result['screenshot'] = screenshot_path

        result['duration_ms'] = int((time.time() - start_time) * 1000)
        self.recorded_actions.append(result)

        return result

    async def _execute_navigate(self, step: Dict, result: Dict):
        """Execute navigation action."""
        # This is a placeholder - actual implementation will use MCP tools
        print(f"  ⚙ Navigate action - requires Chrome DevTools MCP integration")
        print(f"  ⚙ URL: {self.base_url}")
        result['action_details'] = {'url': self.base_url}

    async def _execute_login(self, step: Dict, result: Dict):
        """Execute login action."""
        print(f"  ⚙ Login action - requires Chrome DevTools MCP integration")
        result['action_details'] = {'action': 'login'}

    async def _execute_click(self, step: Dict, result: Dict):
        """Execute click action."""
        print(f"  ⚙ Click action - requires Chrome DevTools MCP integration")

        # Try common button selectors
        selectors = self._generate_click_selectors(step['description'])
        result['selectors_tried'] = selectors
        result['selector_used'] = selectors[0] if selectors else None

    async def _execute_input(self, step: Dict, result: Dict):
        """Execute input action."""
        print(f"  ⚙ Input action - requires Chrome DevTools MCP integration")

    async def _execute_select(self, step: Dict, result: Dict):
        """Execute select/dropdown action."""
        print(f"  ⚙ Select action - requires Chrome DevTools MCP integration")

    async def _execute_verify(self, step: Dict, result: Dict):
        """Execute verification action."""
        print(f"  ⚙ Verify action - requires Chrome DevTools MCP integration")

    async def _execute_capture(self, step: Dict, result: Dict):
        """Execute data capture action."""
        print(f"  ⚙ Capture action - requires Chrome DevTools MCP integration")

    async def _execute_scroll(self, step: Dict, result: Dict):
        """Execute scroll action."""
        print(f"  ⚙ Scroll action - requires Chrome DevTools MCP integration")

    async def _execute_wait(self, step: Dict, result: Dict):
        """Execute wait action."""
        print(f"  ⚙ Wait action - requires Chrome DevTools MCP integration")

    async def _execute_close(self, step: Dict, result: Dict):
        """Execute close action."""
        print(f"  ⚙ Close action - requires Chrome DevTools MCP integration")

    async def _execute_custom(self, step: Dict, result: Dict):
        """Execute custom user-guided action."""
        print(f"  ⚙ Custom action - requires manual guidance")
        print(f"  📝 Please perform this step manually and confirm")

    def _generate_click_selectors(self, description: str) -> List[str]:
        """
        Generate potential selectors for click actions.

        Args:
            description: Step description

        Returns:
            List of selector strings to try
        """
        selectors = []
        desc_lower = description.lower()

        # Extract button/element text from description
        if 'button' in desc_lower:
            # Try to find quoted text or text after "button"
            if '"' in description:
                text = description.split('"')[1]
                selectors.extend([
                    f"text={text}",
                    f"button:has-text('{text}')",
                    f"[aria-label='{text}']"
                ])

        # Common patterns
        if 'approve' in desc_lower:
            selectors.extend([
                "text=Approve",
                "button:has-text('Approve')",
                "[aria-label='Approve']"
            ])
        elif 'pay' in desc_lower and 'now' in desc_lower:
            selectors.extend([
                "text=Pay Now",
                "button:has-text('Pay')",
                "[aria-label='Pay Now']"
            ])
        elif 'invoice' in desc_lower and 'raise' in desc_lower:
            selectors.extend([
                "text=Raise Invoice",
                "button:has-text('Raise')",
                "[aria-label='Raise Invoice']"
            ])

        return selectors if selectors else ["button"]

    async def capture_screenshot(self, name: str) -> str:
        """
        Capture screenshot and save to reports directory.

        Args:
            name: Base name for screenshot file

        Returns:
            Path to saved screenshot
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{name}_{timestamp}.png"
        filepath = self.reports_dir / filename

        # Placeholder - actual implementation will use MCP tool
        print(f"  📸 Screenshot: {filename}")

        self.screenshots.append(str(filepath))
        return str(filepath)

    async def capture_form_data(self) -> Dict:
        """
        Capture all form field values from current page.

        Returns:
            Dictionary of field names and values
        """
        # Placeholder - actual implementation will use MCP tool
        print(f"  📋 Capturing form data...")
        return {}

    async def capture_table_data(self, table_selector: str = "table") -> List[Dict]:
        """
        Capture table content from page.

        Args:
            table_selector: CSS selector for table

        Returns:
            List of row dictionaries
        """
        # Placeholder - actual implementation will use MCP tool
        print(f"  📊 Capturing table data...")
        return []

    def _sanitize_filename(self, text: str) -> str:
        """
        Sanitize text for use in filename.

        Args:
            text: Text to sanitize

        Returns:
            Sanitized filename string
        """
        # Remove special characters, limit length
        sanitized = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in text)
        sanitized = sanitized.replace(' ', '_')
        return sanitized[:50]

    def get_execution_summary(self) -> Dict:
        """
        Get summary of execution.

        Returns:
            Dictionary with execution statistics
        """
        total_steps = len(self.recorded_actions)
        successful_steps = sum(1 for a in self.recorded_actions if a['success'])
        failed_steps = total_steps - successful_steps

        total_duration = sum(a['duration_ms'] for a in self.recorded_actions)

        return {
            'total_steps': total_steps,
            'successful_steps': successful_steps,
            'failed_steps': failed_steps,
            'total_duration_ms': total_duration,
            'total_screenshots': len(self.screenshots),
            'start_time': self.execution_start_time.isoformat() if self.execution_start_time else None,
            'end_time': datetime.now().isoformat()
        }

    def save_recorded_actions(self, filepath: str):
        """
        Save recorded actions to JSON file.

        Args:
            filepath: Path to save JSON file
        """
        with open(filepath, 'w') as f:
            json.dump({
                'actions': self.recorded_actions,
                'summary': self.get_execution_summary()
            }, f, indent=2)
