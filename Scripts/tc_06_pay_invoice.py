"""
TC_06: Pay Invoice from View Page - Standalone Script
======================================================
This script automates TC_06 test case for Omney Business application.

Test Case: TC_06 - To Pay Invoice from View page
Prerequisites: Invoice must already exist in Pending Payables (created via TC_01-TC_03)

Steps:
    1. Login as Client_Business
    2. Find invoice in Pending Payables
    3. Click eye icon to view invoice details
    4. Click Approve button
    5. Click Pay Now to navigate to Pay Invoice form
    6. Capture and verify all form fields against TC_03 data
    7. Click Pay Now to complete payment
    8. Capture Transaction Success popup with Booking ID
    9. Close popup

Requirements:
    pip install playwright pandas openpyxl
    playwright install chromium

Usage:
    python tc_06_pay_invoice.py
    python tc_06_pay_invoice.py --invoice INV-202601094108
"""

import os
import sys
import io
import argparse
from datetime import datetime
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# Fix Windows console encoding for Unicode characters
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


class TC06PayInvoice:
    """TC_06: Pay Invoice from View Page automation class."""

    def __init__(self, headless: bool = False, invoice_number: str = None):
        """
        Initialize TC_06 automation.

        Args:
            headless: Run browser in headless mode (default: False)
            invoice_number: Specific invoice number to pay (optional)
        """
        self.base_url = "https://qaoneob.remit.in"
        self.headless = headless
        self.invoice_number = invoice_number
        self.browser = None
        self.page = None
        self.context = None
        self.playwright = None

        # Test results storage
        self.test_result = None
        self.tc03_invoice_data = {}  # Expected data from TC_03
        self.tc06_form_data = {}  # Captured data from Pay Invoice form
        self.tc06_verification_results = []  # Verification comparison
        self.transaction_data = {}  # Transaction success data
        self.screenshots = []

        # Setup directories
        self.base_dir = Path(__file__).parent.parent
        self.reports_dir = self.base_dir / "Reports" / "Python_Automation"
        self.testcase_file = self.base_dir / "Testcase" / "OB_Automation.xlsx"

        # Create reports directory
        self.reports_dir.mkdir(parents=True, exist_ok=True)

        # Credentials
        self.credentials_sheet = None

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

        # Load credentials from Excel
        self._load_credentials()

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

    def _load_credentials(self):
        """Load credentials from Excel file."""
        try:
            self.credentials_sheet = pd.read_excel(self.testcase_file, sheet_name="Credentials")
            print(f"[DATA] Loaded credentials from {self.testcase_file}")
        except Exception as e:
            print(f"[ERROR] Failed to load credentials: {e}")
            raise

    def _get_credentials(self, credential_type: str) -> tuple:
        """Get credentials by type."""
        try:
            cred_row = self.credentials_sheet[self.credentials_sheet['Type'] == credential_type]
            if cred_row.empty:
                raise ValueError(f"Credential type '{credential_type}' not found")
            email = cred_row['Email'].values[0]
            password = cred_row['Password'].values[0]
            print(f"[CREDENTIALS] Using {credential_type}: {email}")
            return email, password
        except Exception as e:
            print(f"[ERROR] Failed to get credentials: {e}")
            raise

    def _take_screenshot(self, name: str, full_page: bool = False) -> str:
        """Take a screenshot and save it."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{name}_{timestamp}.png"
        filepath = self.reports_dir / filename
        self.page.screenshot(path=str(filepath), full_page=full_page)
        self.screenshots.append(str(filepath))
        print(f"[SCREENSHOT] Saved: {filename}")
        return str(filepath)

    def set_tc03_data(self, invoice_data: dict):
        """Set expected invoice data from TC_03 for verification."""
        self.tc03_invoice_data = invoice_data
        if 'Invoice Number' in invoice_data:
            self.invoice_number = invoice_data['Invoice Number']

    def login_as_client(self) -> bool:
        """Login as Client_Business user."""
        print("\n[LOGIN] Logging in as Client_Business...")

        try:
            # Navigate to login page
            self.page.goto(f"{self.base_url}/login")
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(1000)

            # Get credentials
            email, password = self._get_credentials("Client_Business")

            # Fill login form
            email_input = self.page.locator("input[type='email'], input[type='text']").first
            email_input.fill(email)

            password_input = self.page.locator("input[type='password']").first
            password_input.fill(password)

            self._take_screenshot("TC_06_Login_Form")

            # Submit
            password_input.press("Enter")
            self.page.wait_for_timeout(2000)

            # Wait for dashboard
            try:
                self.page.wait_for_url("**/dashboard", timeout=30000)
            except:
                login_btn = self.page.locator("button:has-text('Log in')").first
                if login_btn.is_visible(timeout=2000):
                    login_btn.click()
                    self.page.wait_for_url("**/dashboard", timeout=30000)

            self.page.wait_for_load_state("networkidle")
            print("[LOGIN] Successfully logged in as Client_Business")
            self._take_screenshot("TC_06_Client_Dashboard")
            return True

        except Exception as e:
            print(f"[ERROR] Login failed: {e}")
            self._take_screenshot("TC_06_Login_Failed")
            return False

    def find_and_view_invoice(self) -> bool:
        """Find invoice in Pending Payables and click to view details."""
        print(f"\n[STEP] Finding invoice {self.invoice_number} in Pending Payables...")

        try:
            # Scroll to Pending Payables section
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            self.page.wait_for_timeout(2000)

            # Look for invoice
            invoice_visible = False
            try:
                invoice_element = self.page.locator(f"text={self.invoice_number}").first
                if invoice_element.is_visible(timeout=10000):
                    invoice_visible = True
            except:
                pass

            if not invoice_visible:
                # Try clicking "View all" for Pending Payables
                view_all_buttons = self.page.locator("button:has-text('View all')")
                if view_all_buttons.count() > 1:
                    view_all_buttons.nth(1).click()
                    self.page.wait_for_timeout(2000)

            self._take_screenshot("TC_06_Pending_Payables")

            # Click eye icon to view invoice details
            clicked = self.page.evaluate(f"""
                () => {{
                    const elements = document.querySelectorAll('tr, div');
                    for (const el of elements) {{
                        if (el.textContent.includes('{self.invoice_number}')) {{
                            const eyeIcon = el.querySelector('svg.lucide-eye, svg[class*="eye"]');
                            if (eyeIcon) {{
                                eyeIcon.dispatchEvent(new MouseEvent('click', {{
                                    view: window,
                                    bubbles: true,
                                    cancelable: true
                                }}));
                                return true;
                            }}
                        }}
                    }}
                    return false;
                }}
            """)

            if not clicked:
                raise Exception(f"Could not find eye icon for invoice {self.invoice_number}")

            # Wait for navigation to payable-details
            self.page.wait_for_url("**/payable-details", timeout=15000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            print(f"[STEP] Viewing invoice details for {self.invoice_number}")
            self._take_screenshot("TC_06_Invoice_Details")
            return True

        except Exception as e:
            print(f"[ERROR] Failed to find/view invoice: {e}")
            self._take_screenshot("TC_06_Find_Invoice_Failed")
            return False

    def approve_invoice(self) -> bool:
        """Click Approve button on invoice details page."""
        print("\n[STEP] Clicking Approve button...")

        try:
            # Find and click Approve button
            approve_btn = self.page.locator("button:has-text('Approve')").first
            if not approve_btn.is_visible(timeout=10000):
                raise Exception("Approve button not found")

            approve_btn.click()
            self.page.wait_for_timeout(3000)

            # Wait for Pay Now button to appear
            pay_now_btn = self.page.locator("button:has-text('Pay Now')").first
            if not pay_now_btn.is_visible(timeout=10000):
                raise Exception("Pay Now button not found after approval")

            print("[STEP] Invoice approved, Pay Now button visible")
            self._take_screenshot("TC_06_After_Approve")
            return True

        except Exception as e:
            print(f"[ERROR] Approval failed: {e}")
            self._take_screenshot("TC_06_Approve_Failed")
            return False

    def navigate_to_pay_invoice(self) -> bool:
        """Click Pay Now to navigate to Pay Invoice form."""
        print("\n[STEP] Clicking Pay Now to navigate to Pay Invoice form...")

        try:
            pay_now_btn = self.page.locator("button:has-text('Pay Now')").first
            pay_now_btn.click()
            self.page.wait_for_timeout(3000)

            # Wait for Pay Invoice page
            self.page.wait_for_url("**/pay**", timeout=15000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            print("[STEP] Navigated to Pay Invoice form")
            self._take_screenshot("TC_06_Pay_Invoice_Form", full_page=True)
            return True

        except Exception as e:
            print(f"[ERROR] Navigation to Pay Invoice failed: {e}")
            self._take_screenshot("TC_06_Navigate_Pay_Failed")
            return False

    def capture_and_verify_form_data(self) -> bool:
        """Capture all fields from Pay Invoice form and verify against TC_03 data."""
        print("\n[STEP] Capturing and verifying Pay Invoice form data...")

        try:
            # Capture all form data using JavaScript
            captured_data = self.page.evaluate("""
                () => {
                    const data = {};
                    const pageText = document.body.innerText;

                    // Choose Invoice - from page text
                    const chooseMatch = pageText.match(/Choose Invoice[\\s\\S]*?(INV-[\\d]+[^\\n]*)/);
                    if (chooseMatch) data['Choose Invoice'] = chooseMatch[1].trim();

                    // Vendor Details
                    const mobileMatch = pageText.match(/Mobile Number[\\s\\n]+\\+?(\\d+)[\\s\\n]+(\\d+)/);
                    if (mobileMatch) data['Mobile Number'] = '+' + mobileMatch[1] + ' ' + mobileMatch[2];

                    const countryMatch = pageText.match(/Country[\\s\\n]+([A-Z]{2})/);
                    if (countryMatch) data['Country'] = countryMatch[1];

                    // ========== Vendor Bank Account Details Section ==========
                    // Get all input elements on the page
                    const allInputs = document.querySelectorAll('input');

                    // Bank Name - Look for input field containing "BANK" in value
                    for (const input of allInputs) {
                        const val = input.value;
                        if (val && val.toUpperCase().includes('BANK')) {
                            data['Bank Name'] = val;
                            break;
                        }
                    }

                    // Account Number - Look for masked account number (********5678 pattern)
                    for (const input of allInputs) {
                        const val = input.value;
                        if (val && val.match(/^\\*+\\d+$/)) {
                            data['Account Number'] = val;
                            break;
                        }
                    }

                    // Invoice Number from input
                    const invoiceNumInput = document.querySelector('input[placeholder*="Invoice"], input[name*="invoice"]');
                    if (invoiceNumInput) data['Invoice Number'] = invoiceNumInput.value;
                    // Fallback: find input with INV- value
                    if (!data['Invoice Number']) {
                        for (const input of allInputs) {
                            if (input.value && input.value.startsWith('INV-')) {
                                data['Invoice Number'] = input.value;
                                break;
                            }
                        }
                    }

                    // Dates from inputs
                    const dateInputs = document.querySelectorAll('input[type="date"]');
                    if (dateInputs[0]) data['Invoice Date'] = dateInputs[0].value;
                    if (dateInputs[1]) data['Due Date'] = dateInputs[1].value;

                    // Currency - check select dropdowns and page text
                    const allSelects = document.querySelectorAll('select');
                    for (const sel of allSelects) {
                        const selectedOpt = sel.options[sel.selectedIndex];
                        if (selectedOpt) {
                            const txt = selectedOpt.text || selectedOpt.value;
                            if (txt && txt.match(/^[A-Z]{3}$/)) {
                                data['Currency'] = txt;
                                break;
                            }
                        }
                    }
                    if (!data['Currency']) {
                        const currencyMatch = pageText.match(/Currency \\*[\\s\\n]+([A-Z]{3})/);
                        if (currencyMatch) data['Currency'] = currencyMatch[1];
                    }

                    // Amount from input
                    const amountInput = document.querySelector('input[placeholder*="amount"], input[name*="amount"]');
                    if (amountInput) data['Amount'] = amountInput.value;
                    if (!data['Amount']) {
                        // Find numeric input that looks like an amount
                        for (const input of allInputs) {
                            if (input.value && input.value.match(/^[\\d.]+$/) && parseFloat(input.value) >= 100) {
                                data['Amount'] = input.value;
                                break;
                            }
                        }
                    }

                    // ========== Purpose and Source of Funds Section ==========
                    // Purpose - Look in select dropdowns for Purpose value
                    for (const sel of allSelects) {
                        const selectedOpt = sel.options[sel.selectedIndex];
                        if (selectedOpt) {
                            const txt = selectedOpt.text || selectedOpt.value;
                            if (txt && (txt.toLowerCase().includes('purpose') ||
                                       txt.toLowerCase().includes('demo') ||
                                       txt.toLowerCase().includes('payment'))) {
                                data['Purpose'] = txt;
                                break;
                            }
                        }
                    }
                    // Fallback: Look for "Demo Purpose" text pattern
                    if (!data['Purpose']) {
                        const purposeMatch = pageText.match(/(Demo Purpose|Payment Purpose|Business Purpose|Trade Purpose)/i);
                        if (purposeMatch) data['Purpose'] = purposeMatch[1];
                    }

                    // Invoice Document
                    const docMatch = pageText.match(/Invoice uploaded|([A-Za-z0-9._-]+\\.png)/i);
                    if (docMatch) data['Invoice Document'] = docMatch[0];

                    return data;
                }
            """)

            self.tc06_form_data = captured_data
            print("[DATA] Captured Pay Invoice Form Data:")
            for key, value in captured_data.items():
                print(f"  {key}: {value}")

            # Verify against TC_03 data if available
            if self.tc03_invoice_data:
                print("\n[VERIFY] Comparing with TC_03 data...")
                verification_results = []

                fields_to_verify = [
                    ('Invoice Number', 'Invoice Number'),
                    ('Bank Name', 'Bank Name'),
                    ('Account Number', 'Account Number'),
                    ('Currency', 'Currency'),
                    ('Amount', 'Amount'),
                    ('Purpose', 'Purpose'),
                ]

                for tc03_field, tc06_field in fields_to_verify:
                    expected = str(self.tc03_invoice_data.get(tc03_field, '')).strip().upper()
                    actual = str(captured_data.get(tc06_field, '')).strip().upper()

                    # Handle amount comparison
                    if 'amount' in tc03_field.lower():
                        try:
                            exp_num = float(str(self.tc03_invoice_data.get(tc03_field, 0)).replace(',', ''))
                            act_num = float(str(captured_data.get(tc06_field, 0)).replace(',', ''))
                            status = "MATCH" if abs(exp_num - act_num) < 0.01 else "MISMATCH"
                        except:
                            status = "MATCH" if expected == actual else "MISMATCH"
                    elif expected and actual and (expected in actual or actual in expected):
                        status = "MATCH"
                    elif not actual:
                        status = "DATA MISSING"
                    else:
                        status = "MISMATCH"

                    result = {
                        'field': tc03_field,
                        'expected': self.tc03_invoice_data.get(tc03_field, ''),
                        'actual': captured_data.get(tc06_field, '(Blank)'),
                        'status': status
                    }
                    verification_results.append(result)

                    icon = "+" if status == "MATCH" else "-"
                    print(f"  {icon} {tc03_field}: Expected='{result['expected']}' | Actual='{result['actual']}' | {status}")

                self.tc06_verification_results = verification_results

            return True

        except Exception as e:
            print(f"[ERROR] Failed to capture form data: {e}")
            self._take_screenshot("TC_06_Capture_Failed")
            return False

    def complete_payment(self) -> bool:
        """Click Pay Now to complete the payment and capture success popup."""
        print("\n[STEP] Completing payment...")

        try:
            # Find and click Pay Now button (the submit button)
            pay_now_btn = self.page.locator("button:has-text('Pay Now')").last
            if not pay_now_btn.is_visible(timeout=10000):
                raise Exception("Pay Now button not found on form")

            pay_now_btn.scroll_into_view_if_needed()
            self.page.wait_for_timeout(500)
            pay_now_btn.click()

            # Wait for success popup
            print("[WAIT] Waiting for transaction success popup...")
            self.page.wait_for_timeout(5000)

            # Check for success popup
            success_selectors = [
                "text=Transaction Successful",
                "text=Booking ID",
                "text=booked Successfully",
            ]

            popup_found = False
            for selector in success_selectors:
                try:
                    if self.page.locator(selector).first.is_visible(timeout=10000):
                        popup_found = True
                        break
                except:
                    continue

            if not popup_found:
                raise Exception("Transaction success popup not found")

            self._take_screenshot("TC_06_Transaction_Success")

            # Capture transaction details
            transaction_data = self.page.evaluate("""
                () => {
                    const data = {};
                    const pageText = document.body.innerText;

                    // Booking ID
                    const bookingMatch = pageText.match(/Booking ID[\\s\\n]+([A-Z0-9]+)/);
                    if (bookingMatch) data['Booking ID'] = bookingMatch[1];

                    // Bank
                    const bankMatch = pageText.match(/Send Money To[\\s\\S]*?([A-Za-z\\s]+Bank[A-Za-z\\s]*)/i);
                    if (bankMatch) data['Bank'] = bankMatch[1].trim();

                    // Account Holder
                    const holderMatch = pageText.match(/Bank[A-Za-z\\s]*[\\s\\n]+([A-Za-z]+)[\\s\\n]+\\d/);
                    if (holderMatch) data['Account Holder'] = holderMatch[1];

                    // Account Number
                    const accMatch = pageText.match(/([\\d]{10,})/);
                    if (accMatch) data['Account Number'] = accMatch[1];

                    // BIC Code
                    const bicMatch = pageText.match(/([A-Z]{8,11})/);
                    if (bicMatch) data['BIC Code'] = bicMatch[1];

                    return data;
                }
            """)

            self.transaction_data = transaction_data
            print("[SUCCESS] Transaction completed!")
            print(f"  Booking ID: {transaction_data.get('Booking ID', 'N/A')}")
            print(f"  Bank: {transaction_data.get('Bank', 'N/A')}")
            print(f"  Account: {transaction_data.get('Account Number', 'N/A')}")

            return True

        except Exception as e:
            print(f"[ERROR] Payment failed: {e}")
            self._take_screenshot("TC_06_Payment_Failed")
            return False

    def close_popup(self) -> bool:
        """Close the success popup."""
        print("\n[STEP] Closing success popup...")

        try:
            close_btn = self.page.locator("button:has-text('Close')").first
            if close_btn.is_visible(timeout=5000):
                close_btn.click()
                self.page.wait_for_timeout(2000)
                print("[STEP] Popup closed")
                self._take_screenshot("TC_06_After_Close")
                return True
            return False

        except Exception as e:
            print(f"[WARNING] Could not close popup: {e}")
            return False

    def generate_report(self) -> str:
        """Generate HTML report for TC_06."""
        report_path = self.reports_dir / f"TC_06_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        # Verification results HTML
        verification_html = ""
        if self.tc06_verification_results:
            rows = ""
            for r in self.tc06_verification_results:
                color = "green" if r['status'] == "MATCH" else ("orange" if r['status'] == "DATA MISSING" else "red")
                rows += f"""
                    <tr>
                        <td>{r['field']}</td>
                        <td>{r['expected']}</td>
                        <td>{r['actual']}</td>
                        <td style="color: {color}; font-weight: bold;">{r['status']}</td>
                    </tr>"""

            verification_html = f"""
            <h2>Data Verification: TC_03 vs TC_06 Pay Invoice Form</h2>
            <table class="data-table">
                <tr>
                    <th>Field</th>
                    <th>Expected (TC_03)</th>
                    <th>Actual (TC_06 Form)</th>
                    <th>Status</th>
                </tr>
                {rows}
            </table>"""

        # Screenshots HTML
        screenshots_html = ""
        for screenshot in self.screenshots:
            name = Path(screenshot).name
            screenshots_html += f"""
                <div class="screenshot-item">
                    <img src="{name}" alt="{name}" onclick="openModal(this)">
                    <p>{name}</p>
                </div>"""

        # Transaction data HTML
        transaction_html = ""
        if self.transaction_data:
            transaction_html = f"""
            <h2>Transaction Success Details</h2>
            <div class="success-box">
                <h3>Transaction Booked Successfully!</h3>
                <table class="data-table">
                    <tr><td>Booking ID</td><td><strong>{self.transaction_data.get('Booking ID', 'N/A')}</strong></td></tr>
                    <tr><td>Bank</td><td>{self.transaction_data.get('Bank', 'N/A')}</td></tr>
                    <tr><td>Account Holder</td><td>{self.transaction_data.get('Account Holder', 'N/A')}</td></tr>
                    <tr><td>Account Number</td><td>{self.transaction_data.get('Account Number', 'N/A')}</td></tr>
                    <tr><td>BIC Code</td><td>{self.transaction_data.get('BIC Code', 'N/A')}</td></tr>
                </table>
            </div>"""

        status = self.test_result if self.test_result else "UNKNOWN"
        status_class = "passed" if status == "PASSED" else "failed"

        report_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TC_06 Test Report - Pay Invoice</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', sans-serif; background: linear-gradient(135deg, #667eea, #764ba2); min-height: 100vh; padding: 20px; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #1a1a2e, #16213e); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ font-size: 2rem; margin-bottom: 10px; }}
        .status-badge {{ display: inline-block; padding: 10px 30px; border-radius: 25px; font-weight: bold; margin-top: 15px; }}
        .status-badge.passed {{ background: #28a745; color: white; }}
        .status-badge.failed {{ background: #dc3545; color: white; }}
        .content {{ padding: 30px; }}
        h2 {{ color: #1a1a2e; border-bottom: 2px solid #667eea; padding-bottom: 10px; margin: 30px 0 15px 0; }}
        .data-table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
        .data-table th, .data-table td {{ padding: 12px 15px; text-align: left; border-bottom: 1px solid #e9ecef; }}
        .data-table th {{ background: #f8f9fa; font-weight: 600; }}
        .success-box {{ background: linear-gradient(135deg, #28a745, #20c997); color: white; padding: 20px; border-radius: 10px; margin: 20px 0; }}
        .success-box h3 {{ text-align: center; margin-bottom: 15px; }}
        .success-box .data-table {{ background: rgba(255,255,255,0.1); }}
        .success-box .data-table td {{ color: white; border-color: rgba(255,255,255,0.2); }}
        .screenshot-gallery {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin-top: 20px; }}
        .screenshot-item {{ background: #f8f9fa; border-radius: 8px; padding: 10px; text-align: center; }}
        .screenshot-item img {{ max-width: 100%; border-radius: 5px; cursor: pointer; }}
        .screenshot-item p {{ margin-top: 8px; font-size: 0.85rem; color: #6c757d; }}
        .modal {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.9); z-index: 1000; }}
        .modal.active {{ display: flex; justify-content: center; align-items: center; }}
        .modal img {{ max-width: 90%; max-height: 90%; }}
        .modal-close {{ position: absolute; top: 20px; right: 30px; color: white; font-size: 2rem; cursor: pointer; }}
        .footer {{ background: #1a1a2e; color: white; padding: 20px; text-align: center; opacity: 0.8; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>TC_06: Pay Invoice from View Page</h1>
            <p>Omney Business Automation - Python Playwright</p>
            <div class="status-badge {status_class}">{status}</div>
        </div>

        <div class="content">
            <h2>Test Case Information</h2>
            <table class="data-table">
                <tr><td>Test Case ID</td><td>TC_06</td></tr>
                <tr><td>Test Scenario</td><td>To Pay Invoice from View page</td></tr>
                <tr><td>Invoice Number</td><td>{self.invoice_number or 'N/A'}</td></tr>
                <tr><td>Execution Time</td><td>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</td></tr>
            </table>

            {verification_html}

            {transaction_html}

            <h2>Screenshots</h2>
            <div class="screenshot-gallery">
                {screenshots_html}
            </div>
        </div>

        <div class="footer">
            <p>Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | TC_06 Pay Invoice Automation</p>
        </div>
    </div>

    <div class="modal" id="imageModal" onclick="closeModal()">
        <span class="modal-close">&times;</span>
        <img id="modalImage" src="">
    </div>

    <script>
        function openModal(img) {{
            document.getElementById('imageModal').classList.add('active');
            document.getElementById('modalImage').src = img.src;
        }}
        function closeModal() {{
            document.getElementById('imageModal').classList.remove('active');
        }}
    </script>
</body>
</html>"""

        with open(report_path, "w", encoding="utf-8") as f:
            f.write(report_content)

        print(f"\n[REPORT] Generated: {report_path}")
        return str(report_path)

    def run(self) -> bool:
        """Execute TC_06 test case."""
        print("\n" + "=" * 70)
        print("TC_06: PAY INVOICE FROM VIEW PAGE")
        print("=" * 70)
        print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Invoice: {self.invoice_number or 'Will be detected from Pending Payables'}")
        print("=" * 70)

        try:
            self.setup()

            # Step 1: Login as Client
            if not self.login_as_client():
                self.test_result = "FAILED"
                return False

            # Step 2: Find and view invoice
            if not self.find_and_view_invoice():
                self.test_result = "FAILED"
                return False

            # Step 3: Approve invoice
            if not self.approve_invoice():
                self.test_result = "FAILED"
                return False

            # Step 4: Navigate to Pay Invoice form
            if not self.navigate_to_pay_invoice():
                self.test_result = "FAILED"
                return False

            # Step 5: Capture and verify form data
            if not self.capture_and_verify_form_data():
                self.test_result = "FAILED"
                return False

            # Step 6: Complete payment
            if not self.complete_payment():
                self.test_result = "FAILED"
                return False

            # Step 7: Close popup
            self.close_popup()

            self.test_result = "PASSED"
            print("\n[RESULT] TC_06: PASSED")
            return True

        except Exception as e:
            print(f"\n[CRITICAL ERROR] {e}")
            self.test_result = "FAILED"
            return False

        finally:
            self.generate_report()
            self.teardown()

            print("\n" + "=" * 70)
            print(f"TC_06 Result: {self.test_result}")
            print(f"End Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("=" * 70)


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description='TC_06: Pay Invoice from View Page')
    parser.add_argument('--invoice', type=str, help='Invoice number to pay')
    parser.add_argument('--headless', action='store_true', help='Run in headless mode')
    args = parser.parse_args()

    tc06 = TC06PayInvoice(headless=args.headless, invoice_number=args.invoice)

    # If no invoice specified, use a default or prompt
    if not args.invoice:
        print("[INFO] No invoice number specified. Will use first available invoice in Pending Payables.")
        # You can set TC_03 data here if running standalone
        # tc06.set_tc03_data({'Invoice Number': 'INV-XXXXXXXX', ...})

    tc06.run()


if __name__ == "__main__":
    main()
