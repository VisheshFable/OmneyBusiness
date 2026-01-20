"""
ReportGenerator Module
=====================
Generates comprehensive HTML reports for test execution.
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional


class ReportGenerator:
    """Generates HTML reports from test execution data."""

    def __init__(self, template_path: Optional[str] = None):
        """
        Initialize ReportGenerator.

        Args:
            template_path: Path to HTML template file (optional)
        """
        self.template_path = Path(template_path) if template_path else None

    def generate_report(
        self,
        tc_data: Dict,
        execution_results: List[Dict],
        screenshots: List[str],
        summary: Dict,
        verification_data: Optional[Dict] = None,
        transaction_data: Optional[Dict] = None,
        output_path: str = None
    ) -> str:
        """
        Generate complete HTML report.

        Args:
            tc_data: Test case data dictionary
            execution_results: List of step execution results
            screenshots: List of screenshot paths
            summary: Execution summary dictionary
            verification_data: Optional data verification results
            transaction_data: Optional transaction success data
            output_path: Path to save report

        Returns:
            Path to generated report
        """
        tc_id = tc_data['tc_id']
        test_scenario = tc_data['test_scenario']
        priority = tc_data.get('priority', 'Medium')

        # Determine overall status
        total_steps = summary['total_steps']
        failed_steps = summary['failed_steps']
        status = "PASSED" if failed_steps == 0 else "FAILED"
        status_class = "passed" if status == "PASSED" else "failed"

        # Generate test steps HTML
        steps_html = self._generate_steps_html(execution_results)

        # Generate verification table if provided
        verification_html = ""
        if verification_data:
            verification_html = self._create_verification_table(verification_data)

        # Generate transaction details if provided
        transaction_html = ""
        if transaction_data:
            transaction_html = self._create_transaction_details(transaction_data, tc_id)

        # Generate screenshot gallery
        screenshots_html = self._create_screenshot_gallery(screenshots)

        # Generate full HTML
        html_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Test Report - {tc_id}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: white; padding: 30px 40px; text-align: center; }}
        .header h1 {{ font-size: 2.5rem; margin-bottom: 10px; }}
        .header .tc-id {{ font-size: 1.8rem; color: #20c997; margin-bottom: 10px; }}
        .header .scenario {{ opacity: 0.9; font-size: 1.1rem; margin-bottom: 20px; }}
        .meta-info {{ display: flex; justify-content: center; gap: 40px; margin-top: 20px; flex-wrap: wrap; }}
        .meta-item {{ text-align: center; }}
        .meta-item label {{ display: block; font-size: 0.8rem; opacity: 0.7; text-transform: uppercase; }}
        .meta-item span {{ font-size: 1rem; font-weight: 600; }}
        .status-banner {{ padding: 30px; text-align: center; font-size: 2rem; font-weight: bold; }}
        .status-banner.passed {{ background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; }}
        .status-banner.failed {{ background: linear-gradient(135deg, #dc3545 0%, #e83e8c 100%); color: white; }}
        .summary {{ display: flex; justify-content: space-around; padding: 30px; background: #f8f9fa; border-bottom: 1px solid #e9ecef; flex-wrap: wrap; gap: 20px; }}
        .summary-card {{ text-align: center; padding: 20px 40px; border-radius: 10px; background: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1); min-width: 150px; }}
        .summary-card h2 {{ font-size: 2.5rem; margin-bottom: 5px; color: #667eea; }}
        .summary-card p {{ color: #6c757d; font-size: 0.9rem; text-transform: uppercase; }}
        .content {{ padding: 40px; }}
        .section-title {{ font-size: 1.5rem; color: #1a1a2e; margin: 30px 0 20px 0; padding-bottom: 10px; border-bottom: 2px solid #667eea; }}
        .step-item {{ background: #f8f9fa; border-radius: 10px; padding: 20px; margin-bottom: 20px; border-left: 4px solid #667eea; }}
        .step-item.success {{ border-left-color: #28a745; }}
        .step-item.error {{ border-left-color: #dc3545; }}
        .step-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }}
        .step-number {{ font-size: 1.2rem; font-weight: bold; color: #667eea; }}
        .step-status {{ padding: 5px 15px; border-radius: 15px; font-size: 0.85rem; font-weight: 600; }}
        .step-status.success {{ background: #28a745; color: white; }}
        .step-status.error {{ background: #dc3545; color: white; }}
        .step-description {{ font-size: 1rem; color: #1a1a2e; margin-bottom: 10px; }}
        .step-details {{ font-size: 0.9rem; color: #6c757d; }}
        .step-details span {{ display: inline-block; margin-right: 20px; }}
        .data-table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        .data-table th, .data-table td {{ padding: 12px 15px; text-align: left; border-bottom: 1px solid #e9ecef; }}
        .data-table th {{ background: #f8f9fa; font-weight: 600; color: #495057; font-size: 0.9rem; text-transform: uppercase; }}
        .data-table td {{ color: #1a1a2e; }}
        .screenshot-gallery {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-top: 20px; }}
        .screenshot-item {{ background: #f8f9fa; border-radius: 8px; padding: 15px; text-align: center; }}
        .screenshot-item img {{ max-width: 100%; border-radius: 5px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); cursor: pointer; transition: transform 0.3s; }}
        .screenshot-item img:hover {{ transform: scale(1.02); }}
        .screenshot-item p {{ margin-top: 10px; font-size: 0.85rem; color: #6c757d; }}
        .transaction-box {{ margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #6f42c1 0%, #9561e2 100%); color: white; border-radius: 10px; }}
        .transaction-box h3 {{ text-align: center; margin-bottom: 15px; }}
        .transaction-box table {{ width: 100%; background: rgba(255,255,255,0.1); border-radius: 5px; }}
        .transaction-box td {{ padding: 10px; color: white; }}
        .footer {{ background: #1a1a2e; color: white; padding: 20px 40px; text-align: center; }}
        .footer p {{ opacity: 0.7; font-size: 0.9rem; }}
        .modal {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); z-index: 1000; justify-content: center; align-items: center; }}
        .modal.active {{ display: flex; }}
        .modal img {{ max-width: 90%; max-height: 90%; border-radius: 10px; }}
        .modal-close {{ position: absolute; top: 20px; right: 30px; color: white; font-size: 2rem; cursor: pointer; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Test Execution Report</h1>
            <div class="tc-id">{tc_id}</div>
            <div class="scenario">{test_scenario}</div>
            <div class="meta-info">
                <div class="meta-item">
                    <label>Priority</label>
                    <span>{priority}</span>
                </div>
                <div class="meta-item">
                    <label>Execution Date</label>
                    <span>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</span>
                </div>
                <div class="meta-item">
                    <label>Framework</label>
                    <span>Chrome DevTools MCP</span>
                </div>
            </div>
        </div>

        <div class="status-banner {status_class}">
            {'✓ TEST PASSED' if status == 'PASSED' else '✗ TEST FAILED'}
        </div>

        <div class="summary">
            <div class="summary-card">
                <h2>{total_steps}</h2>
                <p>Total Steps</p>
            </div>
            <div class="summary-card">
                <h2>{summary['successful_steps']}</h2>
                <p>Successful</p>
            </div>
            <div class="summary-card">
                <h2>{failed_steps}</h2>
                <p>Failed</p>
            </div>
            <div class="summary-card">
                <h2>{summary['total_screenshots']}</h2>
                <p>Screenshots</p>
            </div>
        </div>

        <div class="content">
            <h2 class="section-title">Test Execution Steps</h2>
            {steps_html}

            {verification_html}

            {transaction_html}

            <h2 class="section-title">Screenshots Gallery</h2>
            {screenshots_html}
        </div>

        <div class="footer">
            <p>Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Test Automation Agent</p>
        </div>
    </div>

    <div class="modal" id="imageModal">
        <span class="modal-close" onclick="closeModal()">&times;</span>
        <img id="modalImage" src="" alt="Full size screenshot">
    </div>

    <script>
        function openModal(img) {{
            document.getElementById('imageModal').classList.add('active');
            document.getElementById('modalImage').src = img.src;
        }}
        function closeModal() {{
            document.getElementById('imageModal').classList.remove('active');
        }}
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'Escape') closeModal();
        }});
        document.getElementById('imageModal').addEventListener('click', function(e) {{
            if (e.target === this) closeModal();
        }});
    </script>
</body>
</html>'''

        # Save report
        if not output_path:
            output_path = f"{tc_id}_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        return output_path

    def _generate_steps_html(self, execution_results: List[Dict]) -> str:
        """Generate HTML for test steps."""
        steps_html = ""

        for result in execution_results:
            step_no = result['step_no']
            description = result['description']
            success = result['success']
            duration_ms = result['duration_ms']
            action_type = result['action_type']

            status_class = "success" if success else "error"
            status_text = "SUCCESS" if success else "ERROR"

            error_html = ""
            if not success and result.get('error'):
                error_html = f'<div style="color: #dc3545; margin-top: 10px;">Error: {result["error"]}</div>'

            steps_html += f'''
            <div class="step-item {status_class}">
                <div class="step-header">
                    <div class="step-number">Step {step_no}</div>
                    <div class="step-status {status_class}">{status_text}</div>
                </div>
                <div class="step-description">{description}</div>
                <div class="step-details">
                    <span><strong>Action:</strong> {action_type}</span>
                    <span><strong>Duration:</strong> {duration_ms}ms</span>
                </div>
                {error_html}
            </div>'''

        return steps_html

    def _create_verification_table(self, verification_data: Dict) -> str:
        """Create data verification comparison table."""
        if not verification_data:
            return ""

        expected = verification_data.get('expected', {})
        actual = verification_data.get('actual', {})

        rows_html = ""
        for field in expected.keys():
            exp_val = expected.get(field, '')
            act_val = actual.get(field, '')
            match = str(exp_val).strip().upper() == str(act_val).strip().upper()

            status_color = "green" if match else "red"
            status_text = "MATCH" if match else "MISMATCH"

            rows_html += f'''
                <tr>
                    <td>{field}</td>
                    <td>{exp_val}</td>
                    <td>{act_val}</td>
                    <td style="color: {status_color}; font-weight: bold;">{status_text}</td>
                </tr>'''

        html = f'''
            <h2 class="section-title">Data Verification Results</h2>
            <table class="data-table">
                <tr>
                    <th>Field</th>
                    <th>Expected</th>
                    <th>Actual</th>
                    <th>Status</th>
                </tr>
                {rows_html}
            </table>'''

        return html

    def _create_transaction_details(self, transaction_data: Dict, tc_id: str) -> str:
        """Create transaction success details box."""
        if not transaction_data:
            return ""

        rows_html = ""
        for key, value in transaction_data.items():
            style = "font-weight: bold;" if key == "Booking ID" else ""
            rows_html += f'<tr><td>{key}</td><td style="{style}">{value}</td></tr>'

        html = f'''
            <div class="transaction-box">
                <h3>{tc_id} Transaction Success Details</h3>
                <table>
                    {rows_html}
                </table>
            </div>'''

        return html

    def _create_screenshot_gallery(self, screenshots: List[str]) -> str:
        """Create screenshot gallery section."""
        if not screenshots:
            return "<p>No screenshots captured.</p>"

        gallery_html = ""
        for screenshot in screenshots:
            filename = Path(screenshot).name
            gallery_html += f'''
            <div class="screenshot-item">
                <img src="{filename}" alt="{filename}" onclick="openModal(this)">
                <p>{filename}</p>
            </div>'''

        return f'<div class="screenshot-gallery">{gallery_html}</div>'
