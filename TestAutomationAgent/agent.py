"""
Test Automation Agent
====================
Main orchestrator for automated test case execution and script generation.

Usage:
    python agent.py --testcase TC_09
    python agent.py --list
    python agent.py --help
"""

import argparse
import asyncio
import json
import sys
from pathlib import Path
from datetime import datetime
from colorama import init, Fore, Style

# Initialize colorama for colored output
init(autoreset=True)

# Add modules to path
sys.path.append(str(Path(__file__).parent))

from modules import (
    TestCaseReader,
    DevToolsExecutor,
    ReportGenerator,
    ScriptGenerator,
    Integrator
)


class TestAutomationAgent:
    """Main agent orchestrator for test automation."""

    def __init__(self, config_path: str = None):
        """
        Initialize Test Automation Agent.

        Args:
            config_path: Path to configuration file
        """
        self.config = self._load_config(config_path)

        # Initialize components
        self.reader = TestCaseReader(self.config['excel_path'])
        self.executor = None  # Initialized per test case
        self.report_generator = ReportGenerator()
        self.script_generator = ScriptGenerator()
        self.integrator = Integrator(self.config['scripts_dir'] / self.config['main_script'])

        print(f"{Fore.CYAN}{'='*70}")
        print(f"{Fore.CYAN}TEST AUTOMATION AGENT")
        print(f"{Fore.CYAN}{'='*70}{Style.RESET_ALL}\n")

    def _load_config(self, config_path: str = None) -> dict:
        """Load configuration."""
        if config_path and Path(config_path).exists():
            with open(config_path, 'r') as f:
                return json.load(f)

        # Default configuration
        base_dir = Path(__file__).parent.parent
        return {
            'excel_path': str(base_dir / 'Testcase' / 'OB_Automation.xlsx'),
            'base_url': 'https://qaoneob.remit.in',
            'reports_base_dir': str(base_dir / 'Reports'),
            'scripts_dir': base_dir / 'Scripts',
            'main_script': 'omney_business_automation.py',
            'browser': {
                'headless': False,
                'slow_mo': 500,
                'timeout': 30000
            }
        }

    def list_testcases(self):
        """List all available test cases."""
        print(f"{Fore.YELLOW}Available Test Cases:{Style.RESET_ALL}\n")

        testcases = self.reader.list_testcases()
        for tc_id in testcases:
            tc_data = self.reader.get_testcase(tc_id)
            if tc_data:
                priority = tc_data.get('priority', 'Medium')
                scenario = tc_data.get('test_scenario', 'N/A')
                print(f"  {Fore.GREEN}{tc_id}{Style.RESET_ALL} [{priority}]: {scenario}")

        print(f"\n{Fore.CYAN}Total: {len(testcases)} test cases{Style.RESET_ALL}\n")

    async def run_testcase(self, tc_id: str, auto_approve: bool = False) -> bool:
        """
        Execute complete automation workflow for test case.

        Args:
            tc_id: Test case ID to execute
            auto_approve: If True, auto-approve script generation

        Returns:
            True if successful, False otherwise
        """
        print(f"{Fore.CYAN}╔{'═'*68}╗")
        print(f"║{f'TEST AUTOMATION AGENT - {tc_id}'.center(68)}║")
        print(f"╚{'═'*68}╝{Style.RESET_ALL}\n")

        # Phase 1: Read Test Case
        print(f"{Fore.YELLOW}[1/5] Reading Test Case from Excel...{Style.RESET_ALL}")

        tc_data = self.reader.get_testcase(tc_id)
        if not tc_data:
            print(f"{Fore.RED}✗ Test case {tc_id} not found!{Style.RESET_ALL}")
            return False

        print(f"{Fore.GREEN}✓ Test Case ID: {tc_data['tc_id']}")
        print(f"✓ Scenario: {tc_data['test_scenario']}")
        print(f"✓ Priority: {tc_data['priority']}{Style.RESET_ALL}")

        steps = self.reader.parse_steps(tc_data)
        print(f"{Fore.GREEN}✓ Steps: {len(steps)}{Style.RESET_ALL}\n")

        # Display test case details
        print(f"{Fore.CYAN}Test Case Details:")
        print(f"{'━'*70}{Style.RESET_ALL}")
        for step in steps:
            print(f"Step {step['step_no']}: {step['description']}")
        print(f"{Fore.CYAN}{'━'*70}{Style.RESET_ALL}\n")

        if not auto_approve:
            response = input(f"{Fore.YELLOW}Proceed with execution? (yes/no): {Style.RESET_ALL}")
            if response.lower() not in ['yes', 'y']:
                print(f"{Fore.RED}Execution cancelled by user.{Style.RESET_ALL}")
                return False

        # Phase 2: Execute Test Case (Manual with Chrome DevTools MCP)
        print(f"\n{Fore.YELLOW}[2/5] Executing Test Case (Chrome DevTools MCP)...{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'─'*70}{Style.RESET_ALL}")
        print(f"{Fore.RED}⚠ IMPORTANT: This phase requires Chrome DevTools MCP integration{Style.RESET_ALL}")
        print(f"{Fore.RED}⚠ The agent will guide you through manual execution{Style.RESET_ALL}")
        print(f"{Fore.RED}⚠ Confirm each step after completion{Style.RESET_ALL}")
        print(f"{Fore.CYAN}{'─'*70}{Style.RESET_ALL}\n")

        # Create reports directory
        reports_dir = Path(self.config['reports_base_dir']) / f"{tc_id}_DevTools"
        reports_dir.mkdir(parents=True, exist_ok=True)

        # Initialize executor
        self.executor = DevToolsExecutor(
            base_url=self.config['base_url'],
            reports_dir=str(reports_dir),
            mcp_tools={}  # Placeholder for MCP tools
        )

        self.executor.start_execution()

        # Execute each step
        credentials = self.reader.get_credentials('Client_Business')  # Default
        execution_results = []

        for step in steps:
            result = await self.executor.execute_step(step, tc_id)
            execution_results.append(result)

            if not result['success']:
                print(f"{Fore.RED}✗ Step {step['step_no']} failed{Style.RESET_ALL}")
                response = input(f"{Fore.YELLOW}Continue? (yes/retry/abort): {Style.RESET_ALL}")
                if response.lower() == 'abort':
                    break
                elif response.lower() == 'retry':
                    result = await self.executor.execute_step(step, tc_id)
                    execution_results[-1] = result
            else:
                print(f"{Fore.GREEN}✓ Step {step['step_no']} completed{Style.RESET_ALL}")

            if not auto_approve:
                response = input(f"{Fore.YELLOW}Continue to next step? (yes/no): {Style.RESET_ALL}")
                if response.lower() not in ['yes', 'y']:
                    break

        # Phase 3: Generate Report
        print(f"\n{Fore.YELLOW}[3/5] Generating HTML Report...{Style.RESET_ALL}")

        summary = self.executor.get_execution_summary()
        report_path = reports_dir / f"{tc_id}_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        generated_report = self.report_generator.generate_report(
            tc_data=tc_data,
            execution_results=execution_results,
            screenshots=self.executor.screenshots,
            summary=summary,
            output_path=str(report_path)
        )

        print(f"{Fore.GREEN}✓ Report saved: {report_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}✓ Screenshots: {len(self.executor.screenshots)} files saved{Style.RESET_ALL}\n")

        # Save recorded actions
        actions_path = reports_dir / f"{tc_id}_recorded_actions.json"
        self.executor.save_recorded_actions(str(actions_path))

        # Phase 4: Generate Python Scripts
        print(f"\n{Fore.YELLOW}[4/5] Generating Python Scripts...{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}⚠ This requires your approval. Review execution results above.{Style.RESET_ALL}\n")

        if not auto_approve:
            response = input(f"{Fore.YELLOW}Generate automation scripts? (yes/no): {Style.RESET_ALL}")
            if response.lower() not in ['yes', 'y']:
                print(f"{Fore.YELLOW}Script generation skipped.{Style.RESET_ALL}")
                return True

        # Generate standalone script
        standalone_path = self.config['scripts_dir'] / f"{tc_id.lower()}_automation.py"
        generated_script = self.script_generator.generate_standalone_script(
            tc_data=tc_data,
            recorded_actions=execution_results,
            credentials=credentials,
            output_path=str(standalone_path)
        )

        print(f"{Fore.GREEN}✓ Created: {standalone_path.name} ({len(open(generated_script).readlines())} lines){Style.RESET_ALL}")

        # Generate integration code
        integration_code = self.script_generator.generate_integration_code(
            tc_data=tc_data,
            recorded_actions=execution_results
        )

        print(f"{Fore.GREEN}✓ Generated method: {integration_code['method_name']}{Style.RESET_ALL}\n")

        # Phase 5: Integration
        print(f"\n{Fore.YELLOW}[5/5] Integration (Requires Approval)...{Style.RESET_ALL}\n")

        print(f"Generated files for your review:")
        print(f"  • {standalone_path}")
        print(f"  • Integration code ready (not yet applied)\n")

        if not auto_approve:
            response = input(f"{Fore.YELLOW}Review the generated code. Integrate into main framework? (yes/no): {Style.RESET_ALL}")
            if response.lower() not in ['yes', 'y']:
                print(f"{Fore.YELLOW}Integration skipped. Scripts saved but not integrated.{Style.RESET_ALL}")
                return True

        # Perform integration
        print(f"\n{Fore.CYAN}Integrating into main automation framework...{Style.RESET_ALL}")

        # Create backup
        self.integrator.create_backup()

        # Add method
        if not self.integrator.add_method_to_class(integration_code['method_code']):
            print(f"{Fore.RED}✗ Failed to add method{Style.RESET_ALL}")
            return False

        # Update __init__
        if not self.integrator.update_init_method(tc_id, integration_code['init_code']):
            print(f"{Fore.RED}✗ Failed to update __init__{Style.RESET_ALL}")
            return False

        # Save changes
        if not self.integrator.save_changes():
            print(f"{Fore.RED}✗ Failed to save changes{Style.RESET_ALL}")
            return False

        # Generate README
        readme_path = self.config['scripts_dir'] / f"{tc_id}_README.md"
        self.integrator.create_readme(tc_id, tc_data, str(readme_path))

        print(f"{Fore.GREEN}✓ Integration complete!{Style.RESET_ALL}\n")

        # Final summary
        print(f"{Fore.CYAN}╔{'═'*68}╗")
        print(f"║{'AUTOMATION COMPLETE'.center(68)}║")
        print(f"╚{'═'*68}╝{Style.RESET_ALL}\n")

        print(f"{Fore.GREEN}Summary:")
        print(f"  ✓ Test Case: {tc_id}")
        print(f"  ✓ Status: {'PASSED' if summary['failed_steps'] == 0 else 'FAILED'}")
        print(f"  ✓ Duration: {summary['total_duration_ms'] / 1000:.2f} seconds")
        print(f"  ✓ Screenshots: {summary['total_screenshots']}")
        print(f"  ✓ Report: {report_path.name}")
        print(f"  ✓ Standalone Script: {standalone_path.name}")
        print(f"  ✓ Integration: COMPLETED{Style.RESET_ALL}\n")

        print(f"{Fore.YELLOW}Next Steps:")
        print(f"  1. Review report: {report_path}")
        print(f"  2. Test standalone: python {standalone_path}")
        print(f"  3. Run full suite: python Scripts/{self.config['main_script']}{Style.RESET_ALL}\n")

        return True


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Test Automation Agent - Automate test cases with Chrome DevTools MCP'
    )
    parser.add_argument(
        '--testcase', '-t',
        help='Test case ID to automate (e.g., TC_09)'
    )
    parser.add_argument(
        '--list', '-l',
        action='store_true',
        help='List all available test cases'
    )
    parser.add_argument(
        '--auto-approve', '-a',
        action='store_true',
        help='Auto-approve all prompts (non-interactive mode)'
    )
    parser.add_argument(
        '--config', '-c',
        help='Path to configuration file'
    )

    args = parser.parse_args()

    try:
        # Initialize agent
        agent = TestAutomationAgent(config_path=args.config)

        if args.list:
            agent.list_testcases()
            return 0

        if not args.testcase:
            print(f"{Fore.RED}Error: --testcase is required (or use --list to see available cases){Style.RESET_ALL}")
            parser.print_help()
            return 1

        # Run test case
        result = asyncio.run(agent.run_testcase(args.testcase, auto_approve=args.auto_approve))

        return 0 if result else 1

    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}Interrupted by user{Style.RESET_ALL}")
        return 130
    except Exception as e:
        print(f"\n{Fore.RED}[CRITICAL ERROR] {str(e)}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
