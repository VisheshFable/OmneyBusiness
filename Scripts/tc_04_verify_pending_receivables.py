"""
TC_04 - Verify Pending Receivables
==================================
Verifies invoice data in Pending Receivables section by finding the invoice
using the Request ID captured during TC_03.

Requirements:
    pip install playwright pandas openpyxl
    playwright install chromium

Usage:
    python tc_04_verify_pending_receivables.py
    python tc_04_verify_pending_receivables.py --request-id YOUR_REQUEST_ID
    python tc_04_verify_pending_receivables.py --invoice-number INV-20260108XXXX
"""

import os
import sys
import io
import argparse
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, expect, TimeoutError as PlaywrightTimeout

# Fix Windows console encoding for Unicode characters
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


class TC04VerifyPendingReceivables:
    """Test case TC_04: Verify Pending Receivables data."""

    def __init__(self, request_id: str = None, invoice_number: str = None, headless: bool = False):
        """
        Initialize TC_04 test.

        Args:
            request_id: Request ID from TC_03 invoice creation
            invoice_number: Invoice number to search for
            headless: Run browser in headless mode
        """
        self.base_url = "https://qaoneob.remit.in"
        self.headless = headless
        self.browser = None
        self.page = None
        self.context = None
        self.playwright = None

        # Test parameters
        self.request_id = request_id
        self.invoice_number = invoice_number

        # Test results
        self.test_result = None
        self.verification_results = []
        self.captured_data = {}
        self.expected_data = {}

        # Setup directories
        self.base_dir = Path(__file__).parent.parent
        self.reports_dir = self.base_dir / "Reports" / "Python_Automation"
        self.reports_dir.mkdir(parents=True, exist_ok=True)

        # Login credentials
        self.username = "visheshindindia@yopmail.com"
        self.password = "Password@2"

    def setup(self):
        """Setup browser and page."""
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=500,
            args=["--start-maximized", "--disable-infobars", "--no-first-run"]
        )
        self.context = self.browser.new_context(no_viewport=True)
        self.page = self.context.new_page()
        print(f"[SETUP] Browser initialized successfully")
        print(f"[SETUP] Reports will be saved to: {self.reports_dir}")

    def teardown(self):
        """Cleanup browser resources."""
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        print("[TEARDOWN] Browser closed successfully")

    def _take_screenshot(self, name: str, full_page: bool = True) -> str:
        """Take a screenshot and save it."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{name}_{timestamp}.png"
        filepath = self.reports_dir / filename
        self.page.screenshot(path=str(filepath), full_page=full_page)
        print(f"[SCREENSHOT] Saved: {filename}")
        return str(filepath)

    def _login(self) -> bool:
        """Login to the application."""
        print("\n[LOGIN] Logging into the application...")
        try:
            self.page.goto(f"{self.base_url}/login")
            self.page.wait_for_load_state("networkidle")

            # Enter credentials
            email_input = self.page.locator("input[type='email'], input[placeholder*='email']").first
            email_input.fill(self.username)

            password_input = self.page.locator("input[type='password']").first
            password_input.fill(self.password)

            # Submit login
            password_input.press("Enter")
            self.page.wait_for_timeout(2000)

            # Wait for dashboard
            self.page.wait_for_url("**/dashboard", timeout=30000)
            print("[LOGIN] Successfully logged in")
            return True

        except Exception as e:
            print(f"[LOGIN] Failed: {e}")
            return False

    def _find_invoice_in_receivables(self) -> bool:
        """Find the invoice in Pending Receivables section."""
        print("\n[SEARCH] Looking for invoice in Pending Receivables...")

        try:
            # Wait for dashboard to fully load
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            # Look for the invoice by invoice number
            if self.invoice_number:
                invoice_locator = self.page.locator(f"text={self.invoice_number}")
                if invoice_locator.is_visible(timeout=5000):
                    print(f"[SEARCH] Found invoice: {self.invoice_number}")
                    return True

            # If not found by invoice number, scroll and search in the table
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
            self.page.wait_for_timeout(1000)

            # Check if Pending Receivables section exists
            pending_section = self.page.locator("text=Pending Receivables")
            if pending_section.is_visible(timeout=5000):
                print("[SEARCH] Found Pending Receivables section")

                # Look for any invoice row
                if self.invoice_number:
                    row = self.page.locator(f"tr:has-text('{self.invoice_number}')")
                    if row.is_visible(timeout=3000):
                        return True

            return self.invoice_number is not None

        except Exception as e:
            print(f"[SEARCH] Error: {e}")
            return False

    def _click_eye_icon(self) -> bool:
        """Click the eye icon to view invoice details."""
        print("\n[ACTION] Clicking eye icon to view invoice details...")

        try:
            # Find the row with our invoice
            row_selector = f"tr:has-text('{self.invoice_number}')"
            row = self.page.locator(row_selector)

            if not row.is_visible(timeout=5000):
                print(f"[ERROR] Invoice row not found: {self.invoice_number}")
                return False

            # Find and click the eye icon (SVG with lucide-eye class)
            eye_icon = row.locator("svg.lucide-eye, svg[class*='eye']")

            if eye_icon.is_visible(timeout=3000):
                eye_icon.click()
                print("[ACTION] Clicked eye icon")
            else:
                # Try using JavaScript to click
                self.page.evaluate(f"""
                    const rows = document.querySelectorAll('tr');
                    for (const row of rows) {{
                        if (row.textContent.includes('{self.invoice_number}')) {{
                            const cells = row.querySelectorAll('td');
                            const lastCell = cells[cells.length - 1];
                            if (lastCell) {{
                                const eyeIcon = lastCell.querySelector('svg');
                                if (eyeIcon) {{
                                    eyeIcon.dispatchEvent(new MouseEvent('click', {{
                                        view: window,
                                        bubbles: true,
                                        cancelable: true
                                    }}));
                                }}
                            }}
                            break;
                        }}
                    }}
                """)
                print("[ACTION] Clicked eye icon via JavaScript")

            # Wait for navigation to details page
            self.page.wait_for_url("**/receivable-details", timeout=10000)
            self.page.wait_for_load_state("networkidle")
            print("[ACTION] Navigated to invoice details page")
            return True

        except Exception as e:
            print(f"[ACTION] Error clicking eye icon: {e}")
            return False

    def _capture_invoice_details(self) -> dict:
        """Capture all invoice details from the details page."""
        print("\n[CAPTURE] Extracting invoice details from page...")

        try:
            self.page.wait_for_timeout(2000)

            # Capture data using JavaScript
            captured_data = self.page.evaluate("""
                () => {
                    const data = {};
                    const pageText = document.body.innerText;

                    // Invoice Number
                    const invNumMatch = pageText.match(/Invoice Number:\\s*\\n?\\s*([A-Z0-9-]+)/i);
                    if (invNumMatch) data['Invoice Number'] = invNumMatch[1].trim();

                    // Date
                    const dateMatch = pageText.match(/Date:\\s*\\n?\\s*([A-Za-z]+ \\d{1,2}, \\d{4})/i);
                    if (dateMatch) data['Date'] = dateMatch[1].trim();

                    // Bill From
                    const billFromMatch = pageText.match(/Bill From:\\s*\\n?\\s*([A-Za-z\\s]+?)\\s*\\n?\\s*,\\s*([A-Za-z]+)/i);
                    if (billFromMatch) {
                        data['Bill From Name'] = billFromMatch[1].trim();
                        data['Bill From Country'] = billFromMatch[2].trim();
                    }

                    // Bank Name
                    const bankMatch = pageText.match(/Bank Name:\\s*\\n?\\s*([A-Za-z\\s]+?)\\s*(?:\\n|Account)/i);
                    if (bankMatch) data['Bank Name'] = bankMatch[1].trim();

                    // Account Number
                    const accMatch = pageText.match(/Account Number:\\s*\\n?\\s*([*\\d]+)/i);
                    if (accMatch) data['Account Number'] = accMatch[1].trim();

                    // Currency
                    const currMatch = pageText.match(/Currency:\\s*\\n?\\s*([A-Z]{3})?/i);
                    data['Currency'] = currMatch && currMatch[1] ? currMatch[1].trim() : '';

                    // Country (in bank details section)
                    const countryMatch = pageText.match(/Country:\\s*\\n?\\s*([A-Za-z\\s]+?)\\s*(?:\\n|Attached)/i);
                    if (countryMatch) data['Country'] = countryMatch[1].trim();

                    // Amount
                    const amountMatch = pageText.match(/Amount Due\\s*\\n?\\s*([\\d,.]+)/i);
                    if (amountMatch) data['Amount'] = amountMatch[1].trim();

                    // Alternative amount capture
                    if (!data['Amount']) {
                        const altAmountMatch = pageText.match(/Payment Request\\s*\\n?\\s*([\\d,.]+)/i);
                        if (altAmountMatch) data['Amount'] = altAmountMatch[1].trim();
                    }

                    // Attached Documents
                    const docMatch = pageText.match(/Attached Documents \\((\\d+)\\)/);
                    if (docMatch) data['Documents Count'] = docMatch[1];

                    // Document Name
                    const docNameMatch = pageText.match(/Document Name\\s*Action\\s*\\n?\\s*([A-Za-z0-9_.]+)/i);
                    if (docNameMatch) data['Document Name'] = docNameMatch[1].trim();

                    return data;
                }
            """)

            self.captured_data = captured_data
            print("[CAPTURE] Captured invoice details:")
            for key, value in captured_data.items():
                print(f"  {key}: {value}")

            return captured_data

        except Exception as e:
            print(f"[CAPTURE] Error: {e}")
            return {}

    def _verify_data(self, expected_data: dict) -> list:
        """Verify captured data against expected data."""
        print("\n[VERIFY] Verifying invoice data...")

        verification_results = []

        for field, expected_value in expected_data.items():
            actual_value = self.captured_data.get(field, '')

            # Normalize values for comparison
            expected_normalized = str(expected_value).strip().upper() if expected_value else ''
            actual_normalized = str(actual_value).strip().upper() if actual_value else ''

            # Check for match
            if actual_normalized == expected_normalized:
                status = "MATCH"
            elif not actual_value or actual_value == '':
                status = "DATA MISSING"
            else:
                status = "MISMATCH"

            result = {
                "field": field,
                "expected": expected_value,
                "actual": actual_value if actual_value else "(Blank)",
                "status": status
            }
            verification_results.append(result)

            status_icon = "✓" if status == "MATCH" else "✗"
            print(f"  {status_icon} {field}: Expected='{expected_value}' | Actual='{actual_value}' | {status}")

        self.verification_results = verification_results
        return verification_results

    def _generate_report(self, screenshots: list) -> str:
        """Generate HTML test report."""
        report_path = self.reports_dir / f"TC04_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        # Determine overall status
        failed_count = sum(1 for r in self.verification_results if r['status'] != 'MATCH')
        overall_status = "FAILED" if failed_count > 0 else "PASSED"
        status_class = "failed" if failed_count > 0 else "passed"

        # Generate verification table rows
        verification_rows = ""
        for r in self.verification_results:
            if r['status'] == 'MATCH':
                status_color = "green"
            elif r['status'] == 'DATA MISSING':
                status_color = "orange"
            else:
                status_color = "red"

            verification_rows += f"""
                <tr>
                    <td>{r['field']}</td>
                    <td>{r['expected']}</td>
                    <td>{r['actual']}</td>
                    <td style="color: {status_color}; font-weight: bold;">{r['status']}</td>
                </tr>"""

        # Generate screenshot gallery
        screenshot_html = ""
        for screenshot in screenshots:
            if screenshot:
                screenshot_name = Path(screenshot).name
                screenshot_html += f'''
                    <div class="screenshot-item">
                        <img src="{screenshot_name}" alt="{screenshot_name}" onclick="openModal(this)">
                        <p>{screenshot_name}</p>
                    </div>'''

        # Observations section
        observations = []
        for r in self.verification_results:
            if r['status'] == 'DATA MISSING':
                observations.append(f"<li><strong>{r['field']}:</strong> Field appears blank in the Invoice Details view page.</li>")
            elif r['status'] == 'MISMATCH':
                observations.append(f"<li><strong>{r['field']}:</strong> Shows \"{r['actual']}\" instead of expected \"{r['expected']}\".</li>")

        observations_html = ""
        if observations:
            observations_html = f"""
            <div style="margin-top: 20px; padding: 15px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                <strong>Observations:</strong>
                <ul style="margin-top: 10px; margin-left: 20px;">
                    {''.join(observations)}
                </ul>
            </div>"""

        report_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TC_04 Test Report - Verify Pending Receivables</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: white; padding: 30px 40px; text-align: center; }}
        .header h1 {{ font-size: 2rem; margin-bottom: 10px; }}
        .header p {{ opacity: 0.8; font-size: 1rem; }}
        .meta-info {{ display: flex; justify-content: center; gap: 40px; margin-top: 20px; flex-wrap: wrap; }}
        .meta-item {{ text-align: center; }}
        .meta-item label {{ display: block; font-size: 0.8rem; opacity: 0.7; text-transform: uppercase; }}
        .meta-item span {{ font-size: 1rem; font-weight: 600; }}
        .summary {{ display: flex; justify-content: center; padding: 30px; background: #f8f9fa; }}
        .summary-card {{ text-align: center; padding: 20px 60px; border-radius: 10px; background: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }}
        .summary-card.passed {{ border-top: 4px solid #28a745; }}
        .summary-card.failed {{ border-top: 4px solid #dc3545; }}
        .summary-card h2 {{ font-size: 2rem; margin-bottom: 5px; }}
        .summary-card.passed h2 {{ color: #28a745; }}
        .summary-card.failed h2 {{ color: #dc3545; }}
        .summary-card p {{ color: #6c757d; font-size: 0.9rem; text-transform: uppercase; }}
        .content {{ padding: 40px; }}
        .section-title {{ font-size: 1.5rem; color: #1a1a2e; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #667eea; }}
        .test-case {{ background: #f8f9fa; border-radius: 10px; margin-bottom: 30px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        .test-case-header {{ padding: 20px; display: flex; justify-content: space-between; align-items: center; }}
        .test-case-header.passed {{ background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; }}
        .test-case-header.failed {{ background: linear-gradient(135deg, #dc3545 0%, #e83e8c 100%); color: white; }}
        .test-case-id {{ font-size: 1.2rem; font-weight: 700; }}
        .test-case-scenario {{ font-size: 0.95rem; opacity: 0.9; }}
        .status-badge {{ padding: 8px 20px; border-radius: 20px; font-weight: 600; background: rgba(255,255,255,0.2); }}
        .test-case-body {{ padding: 25px; background: white; }}
        .data-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        .data-table th, .data-table td {{ padding: 12px 15px; text-align: left; border-bottom: 1px solid #e9ecef; }}
        .data-table th {{ background: #f8f9fa; font-weight: 600; color: #495057; font-size: 0.85rem; text-transform: uppercase; }}
        .data-table td {{ color: #1a1a2e; }}
        .screenshot-gallery {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; margin-top: 20px; }}
        .screenshot-item {{ background: #f8f9fa; border-radius: 8px; padding: 15px; text-align: center; }}
        .screenshot-item img {{ max-width: 100%; border-radius: 5px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); cursor: pointer; transition: transform 0.3s; }}
        .screenshot-item img:hover {{ transform: scale(1.02); }}
        .screenshot-item p {{ margin-top: 10px; font-size: 0.85rem; color: #6c757d; }}
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
            <h1>TC_04 - Verify Pending Receivables</h1>
            <p>Omney Business Automation Testing (Python Script)</p>
            <div class="meta-info">
                <div class="meta-item">
                    <label>Execution Date</label>
                    <span>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</span>
                </div>
                <div class="meta-item">
                    <label>Request ID</label>
                    <span>{self.request_id or 'N/A'}</span>
                </div>
                <div class="meta-item">
                    <label>Invoice Number</label>
                    <span>{self.invoice_number or 'N/A'}</span>
                </div>
            </div>
        </div>

        <div class="summary">
            <div class="summary-card {status_class}">
                <h2>{'✓' if overall_status == 'PASSED' else '✗'} {overall_status}</h2>
                <p>{len(self.verification_results) - failed_count}/{len(self.verification_results)} Fields Matched</p>
            </div>
        </div>

        <div class="content">
            <div class="test-case">
                <div class="test-case-header {status_class}">
                    <div>
                        <div class="test-case-id">TC_04</div>
                        <div class="test-case-scenario">To check if user can find Request ID in Pending Receivables and verify data</div>
                    </div>
                    <span class="status-badge">{'✓' if overall_status == 'PASSED' else '✗'} {overall_status}</span>
                </div>
                <div class="test-case-body">
                    <p><strong>Details:</strong> Invoice found in Pending Receivables. Data verification {'completed successfully' if overall_status == 'PASSED' else f'FAILED - {failed_count} fields have mismatched/missing data'}.</p>
                    <div class="screenshot-gallery">{screenshot_html}</div>
                </div>
            </div>

            <h2 class="section-title">Data Verification Results</h2>
            <table class="data-table">
                <tr>
                    <th>Field</th>
                    <th>Expected</th>
                    <th>Actual</th>
                    <th>Status</th>
                </tr>
                {verification_rows}
            </table>

            {observations_html}

            <h2 class="section-title" style="margin-top: 40px;">Environment Details</h2>
            <table class="data-table">
                <tr><th>Parameter</th><th>Value</th></tr>
                <tr><td>Automation Framework</td><td>Playwright Python</td></tr>
                <tr><td>Browser</td><td>Chromium</td></tr>
                <tr><td>Python Version</td><td>{sys.version.split()[0]}</td></tr>
                <tr><td>Application URL</td><td>{self.base_url}</td></tr>
            </table>
        </div>

        <div class="footer">
            <p>Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | TC_04 Verification Test</p>
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

        with open(report_path, "w", encoding="utf-8") as f:
            f.write(report_content)

        print(f"\n[REPORT] Generated: {report_path}")
        return str(report_path)

    def run_test(self, expected_data: dict = None) -> bool:
        """
        Run TC_04 test.

        Args:
            expected_data: Dictionary of expected field values for verification

        Returns:
            True if test passed, False otherwise
        """
        print("\n" + "=" * 70)
        print("TC_04: VERIFY PENDING RECEIVABLES")
        print("=" * 70)
        print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Request ID: {self.request_id}")
        print(f"Invoice Number: {self.invoice_number}")
        print("=" * 70)

        screenshots = []

        try:
            self.setup()

            # Step 1: Login
            if not self._login():
                raise Exception("Login failed")

            screenshots.append(self._take_screenshot("TC04_01_Dashboard"))

            # Step 2: Find invoice in Pending Receivables
            if not self._find_invoice_in_receivables():
                raise Exception(f"Invoice not found in Pending Receivables: {self.invoice_number}")

            screenshots.append(self._take_screenshot("TC04_02_Pending_Receivables"))

            # Step 3: Click eye icon to view details
            if not self._click_eye_icon():
                raise Exception("Failed to click eye icon")

            screenshots.append(self._take_screenshot("TC04_03_Invoice_Details"))

            # Step 4: Capture invoice details
            self._capture_invoice_details()

            # Step 5: Verify data
            if expected_data:
                self._verify_data(expected_data)
            else:
                # Use default expected data based on TC_03 input
                default_expected = {
                    "Invoice Number": self.invoice_number,
                    "Bank Name": "BANDHAN BANK",
                    "Account Number": "********5678",
                    "Currency": "INR",
                    "Country": "India",
                    "Amount": "9500.00"
                }
                self._verify_data(default_expected)

            # Step 6: Generate report
            report_path = self._generate_report(screenshots)

            # Determine overall result
            failed_count = sum(1 for r in self.verification_results if r['status'] != 'MATCH')
            self.test_result = "PASSED" if failed_count == 0 else "FAILED"

            print("\n" + "=" * 70)
            print(f"TC_04 RESULT: {self.test_result}")
            print(f"Fields Verified: {len(self.verification_results)}")
            print(f"Fields Matched: {len(self.verification_results) - failed_count}")
            print(f"Fields Mismatched/Missing: {failed_count}")
            print("=" * 70)

            return self.test_result == "PASSED"

        except Exception as e:
            print(f"\n[ERROR] Test failed: {e}")
            screenshots.append(self._take_screenshot("TC04_ERROR"))
            self._generate_report(screenshots)
            return False

        finally:
            self.teardown()


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description='TC_04 - Verify Pending Receivables')
    parser.add_argument('--request-id', type=str, help='Request ID from TC_03')
    parser.add_argument('--invoice-number', type=str, help='Invoice number to search')
    parser.add_argument('--headless', action='store_true', help='Run in headless mode')
    args = parser.parse_args()

    # Default values from the TC_03 execution
    request_id = args.request_id or "ITIwRDoaFU"
    invoice_number = args.invoice_number or "INV-MCP-TC04-20260108"

    # Expected data based on TC_03 input
    expected_data = {
        "Invoice Number": invoice_number,
        "Bill From Name": "SURAJ KUMAAR",
        "Bill From Country": "India",
        "Bank Name": "BANDHAN BANK",
        "Account Number": "********5678",
        "Currency": "INR",
        "Country": "India",
        "Amount": "9500.00",
        "Document Name": "Test.png"
    }

    # Run test
    test = TC04VerifyPendingReceivables(
        request_id=request_id,
        invoice_number=invoice_number,
        headless=args.headless
    )

    result = test.run_test(expected_data)
    sys.exit(0 if result else 1)


if __name__ == "__main__":
    main()
