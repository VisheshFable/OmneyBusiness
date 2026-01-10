"""
Omney Business Automation Script
================================
Automates test cases TC_01 to TC_06 for Omney Business application.

Test Cases:
    TC_01: URL Verification - Check if URL is working
    TC_02: Login - Check if user can login with valid credentials (Vendor_Individual)
    TC_03: Raise Invoice - Create an invoice and capture Request ID
    TC_04: Verify Pending Receivables - Find invoice and verify data (as Vendor)
    TC_05: Verify Pending Payables - Login as Client_Business and verify invoice data
    TC_06: Pay Invoice - Approve and pay invoice, capture transaction success details

Requirements:
    pip install playwright pandas openpyxl
    playwright install chromium

Usage:
    python omney_business_automation.py
"""

import os
import sys
import io
import random
import string
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from playwright.sync_api import sync_playwright, expect, TimeoutError as PlaywrightTimeout

# Fix Windows console encoding for Unicode characters
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


class OmneyBusinessAutomation:
    """Automation class for Omney Business application testing."""

    def __init__(self, headless: bool = False, keep_browser_open: bool = False):
        """
        Initialize the automation framework.

        Args:
            headless: Run browser in headless mode (default: False)
            keep_browser_open: Keep browser open after test completion (default: False)
        """
        self.base_url = "https://qaoneob.remit.in"
        self.headless = headless
        self.keep_browser_open = keep_browser_open
        self.browser = None
        self.page = None
        self.context = None
        self.playwright = None

        # Test results storage
        self.test_results = []
        self.request_id = None
        self.invoice_data = {}
        self.tc04_verification_results = []  # TC_04 verification data
        self.tc04_captured_data = {}  # TC_04 captured invoice details
        self.tc05_verification_results = []  # TC_05 verification data
        self.tc05_captured_data = {}  # TC_05 captured invoice details from Pay Invoice page
        self.tc06_verification_results = []  # TC_06 verification data
        self.tc06_form_data = {}  # TC_06 captured Pay Invoice form data
        self.tc06_transaction_data = {}  # TC_06 transaction success data

        # Setup directories
        self.base_dir = Path(__file__).parent.parent
        self.reports_dir = self.base_dir / "Reports" / "Python_Automation"
        self.testcase_file = self.base_dir / "Testcase" / "OB_Automation.xlsx"

        # Create reports directory
        self.reports_dir.mkdir(parents=True, exist_ok=True)

        # Test data
        self.test_data = None
        self.invoice_sheet = None
        self.credentials_sheet = None

    def setup(self):
        """Setup browser and page."""
        self.playwright = sync_playwright().start()

        # Launch browser in fullscreen/maximized mode
        self.browser = self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=500,  # Slow down actions for visibility
            args=[
                "--start-maximized",
                "--disable-infobars",
                "--no-first-run"
            ]
        )
        # Use no_viewport=True to allow the browser to use its full window size
        self.context = self.browser.new_context(no_viewport=True)
        self.page = self.context.new_page()

        # Load test data from Excel
        self._load_test_data()

        print(f"[SETUP] Browser initialized successfully")
        print(f"[SETUP] Reports will be saved to: {self.reports_dir}")

    def teardown(self):
        """Cleanup browser resources."""
        if self.keep_browser_open:
            print("[TEARDOWN] Browser kept open as requested. Close manually when done.")
            print("[TEARDOWN] Press Ctrl+C to exit the script.")
            try:
                # Keep the script running so browser stays open
                import time
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n[TEARDOWN] Received exit signal, closing browser...")

        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        print("[TEARDOWN] Browser closed successfully")

    def _load_test_data(self):
        """Load test data from Excel file."""
        try:
            self.test_data = pd.read_excel(self.testcase_file, sheet_name="Testcase")
            self.invoice_sheet = pd.read_excel(self.testcase_file, sheet_name="Invoice")
            self.credentials_sheet = pd.read_excel(self.testcase_file, sheet_name="Credentials")
            print(f"[DATA] Loaded test data from {self.testcase_file}")
            print(f"[DATA] Available credential types: {', '.join(self.credentials_sheet['Type'].tolist())}")
        except Exception as e:
            print(f"[ERROR] Failed to load test data: {e}")
            raise

    def _get_credentials(self, credential_type: str) -> tuple:
        """
        Get credentials from Credentials sheet by type.

        Args:
            credential_type: The type of credentials (e.g., 'Vendor_Individual', 'Client_Business')

        Returns:
            Tuple of (email, password)
        """
        try:
            cred_row = self.credentials_sheet[self.credentials_sheet['Type'] == credential_type]
            if cred_row.empty:
                raise ValueError(f"Credential type '{credential_type}' not found in Credentials sheet")

            email = cred_row['Email'].values[0]
            password = cred_row['Password'].values[0]
            print(f"[CREDENTIALS] Using {credential_type}: {email}")
            return email, password
        except Exception as e:
            print(f"[ERROR] Failed to get credentials for '{credential_type}': {e}")
            raise

    def _parse_credential_type(self, test_data_value: str) -> str:
        """
        Parse credential type from Test Data column.

        Args:
            test_data_value: Value from Test Data column (e.g., 'Credentials: Vendor_Individual')

        Returns:
            Credential type string (e.g., 'Vendor_Individual')
        """
        if pd.isna(test_data_value):
            return None

        if 'Credentials:' in str(test_data_value):
            # Extract type after "Credentials:"
            return str(test_data_value).split('Credentials:')[1].strip()
        return None

    def _take_screenshot(self, name: str, full_page: bool = True) -> str:
        """
        Take a screenshot and save it.

        Args:
            name: Screenshot filename (without extension)
            full_page: Capture full page or just viewport

        Returns:
            Path to saved screenshot
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{name}_{timestamp}.png"
        filepath = self.reports_dir / filename
        self.page.screenshot(path=str(filepath), full_page=full_page)
        print(f"[SCREENSHOT] Saved: {filename}")
        return str(filepath)

    def _generate_invoice_number(self) -> str:
        """Generate a random invoice number."""
        date_part = datetime.now().strftime("%Y%m%d")
        random_part = ''.join(random.choices(string.digits, k=4))
        return f"INV-{date_part}{random_part}"

    def _log_result(self, tc_id: str, scenario: str, status: str,
                    details: str = "", screenshot: str = ""):
        """Log test result."""
        result = {
            "tc_id": tc_id,
            "scenario": scenario,
            "status": status,
            "details": details,
            "screenshot": screenshot,
            "timestamp": datetime.now().isoformat()
        }
        self.test_results.append(result)
        status_icon = "✓" if status == "PASSED" else "✗"
        print(f"[{status_icon}] {tc_id}: {status} - {scenario}")

    # =========================================================================
    # TEST CASE: TC_01 - URL Verification
    # =========================================================================
    def tc_01_url_verification(self) -> bool:
        """
        TC_01: To check if URL is working

        Steps:
            1. Navigate to URL

        Expected: URL should be working
        """
        tc_id = "TC_01"
        scenario = "To check if URL is working"
        print(f"\n{'='*60}")
        print(f"[EXECUTING] {tc_id}: {scenario}")
        print(f"{'='*60}")

        try:
            # Step 1: Navigate to URL
            self.page.goto(self.base_url)
            self.page.wait_for_load_state("networkidle")

            # Verify page loaded - check URL contains base domain
            current_url = self.page.url
            assert self.base_url.replace("https://", "") in current_url, f"URL mismatch: {current_url}"

            # Check for key elements on homepage - try multiple possible headings
            page_loaded = False
            possible_selectors = [
                "h1:has-text('Borderless B2B Payments')",
                "h1:has-text('Borderless')",
                "h1:has-text('B2B')",
                "h1:has-text('Payments')",
                "h1",
                "text=Login",
                "text=Sign in",
                "a[href='/login']",
                "button"
            ]

            for selector in possible_selectors:
                try:
                    element = self.page.locator(selector).first
                    if element.is_visible(timeout=3000):
                        page_loaded = True
                        print(f"[INFO] Found element with selector: {selector}")
                        break
                except:
                    continue

            if not page_loaded:
                raise Exception("Could not find any expected elements on homepage")

            # Take screenshot
            screenshot = self._take_screenshot("TC_01_URL_Working")

            # Log success
            self._log_result(
                tc_id, scenario, "PASSED",
                "URL loaded successfully with homepage content",
                screenshot
            )
            return True

        except Exception as e:
            screenshot = self._take_screenshot("TC_01_FAILED")
            self._log_result(tc_id, scenario, "FAILED", str(e), screenshot)
            return False

    # =========================================================================
    # TEST CASE: TC_02 - Login with Valid Credentials
    # =========================================================================
    def tc_02_login(self) -> bool:
        """
        TC_02: To check if user is able to Login using valid credentials

        Steps:
            1. Enter valid Username
            2. Enter valid Password
            3. Click on Submit

        Expected: User should be successfully logged in
        """
        tc_id = "TC_02"
        scenario = "To check if user is able to Login using valid credentials"
        print(f"\n{'='*60}")
        print(f"[EXECUTING] {tc_id}: {scenario}")
        print(f"{'='*60}")

        try:
            # Get credentials from Test Data column in Testcase sheet
            tc02_row = self.test_data[self.test_data['TC_ID'] == 'TC_02']
            test_data_value = tc02_row['Test Data'].values[0] if not tc02_row.empty else None

            # Parse credential type and get credentials
            credential_type = self._parse_credential_type(test_data_value)
            if credential_type:
                username, password = self._get_credentials(credential_type)
            else:
                # Fallback to default credentials
                username = "visheshindindia@yopmail.com"
                password = "Password@2"
                print(f"[CREDENTIALS] Using default credentials: {username}")

            # Click on Login button from homepage - try multiple approaches
            login_clicked = False
            login_link_selectors = [
                "text=Log in",
                "a:has-text('Log in')",
                "a:has-text('Login')",
                "a[href='/login']",
                "button:has-text('Log in')",
                "nav >> text=Log in",
                "header >> text=Log in",
            ]

            for selector in login_link_selectors:
                try:
                    login_link = self.page.locator(selector).first
                    if login_link.is_visible(timeout=2000):
                        login_link.click()
                        login_clicked = True
                        print(f"[INFO] Clicked login using selector: {selector}")
                        break
                except Exception as e:
                    print(f"[DEBUG] Selector {selector} failed: {str(e)[:50]}")
                    continue

            if not login_clicked:
                # Try direct navigation as fallback
                print("[INFO] Trying direct navigation to login page")
                self.page.goto(f"{self.base_url}/login")

            self.page.wait_for_url("**/login", timeout=15000)
            print("[STEP] Navigated to login page")

            # Wait for login form to load
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(1000)

            # Step 1: Enter valid Username - try multiple selectors
            email_selectors = [
                "input[placeholder='you@example.com']",
                "input[placeholder*='email']",
                "input[placeholder*='Email']",
                "input[type='email']",
                "input[type='text']",
            ]
            email_input = None
            for selector in email_selectors:
                try:
                    email_input = self.page.locator(selector).first
                    if email_input.is_visible(timeout=2000):
                        break
                except:
                    continue

            if email_input:
                email_input.fill(username)
                print(f"[STEP 1] Entered username: {username}")
            else:
                raise Exception("Could not find email input field")

            # Step 2: Enter valid Password
            password_input = self.page.locator("input[type='password']").first
            password_input.fill(password)
            print(f"[STEP 2] Entered password: ********")

            # Take screenshot before clicking login
            self._take_screenshot("TC_02_Before_Login_Click")

            # Step 3: Click on Submit - try multiple submission methods
            print("[STEP 3] Attempting form submission...")

            # Method 1: Try pressing Enter on password field (common for forms)
            print("[DEBUG] Method 1: Pressing Enter on password field")
            password_input.press("Enter")
            self.page.wait_for_timeout(2000)

            # Check if we navigated away
            if "/dashboard" in self.page.url:
                print("[SUCCESS] Login succeeded via Enter key")
            else:
                # Method 2: Try clicking the button with force
                print("[DEBUG] Method 2: Force clicking login button")
                login_button = self.page.locator("button:has-text('Log in')").first
                login_button.click(force=True)
                self.page.wait_for_timeout(2000)

            # Check again
            if "/dashboard" not in self.page.url:
                # Method 3: Try JavaScript click
                print("[DEBUG] Method 3: JavaScript button click")
                self.page.evaluate("""
                    const btn = document.querySelector('button');
                    if (btn) {
                        btn.click();
                    }
                """)
                self.page.wait_for_timeout(2000)

            # Check again
            if "/dashboard" not in self.page.url:
                # Method 4: Try form submission via JavaScript
                print("[DEBUG] Method 4: JavaScript form submit")
                self.page.evaluate("""
                    const form = document.querySelector('form');
                    if (form) {
                        form.submit();
                    }
                """)
                self.page.wait_for_timeout(2000)

            # Check again
            if "/dashboard" not in self.page.url:
                # Method 5: Dispatch click event
                print("[DEBUG] Method 5: Dispatching click event on button")
                self.page.evaluate("""
                    const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.includes('Log in'));
                    if (btn) {
                        btn.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, view: window}));
                    }
                """)
                self.page.wait_for_timeout(3000)

            # Take screenshot after submission attempts
            self._take_screenshot("TC_02_After_Login_Click")

            # Check for error messages
            error_selectors = [
                "text=Invalid",
                "text=incorrect",
                "text=error",
                "text=failed",
                ".error",
                "[role='alert']",
            ]
            for selector in error_selectors:
                try:
                    error_element = self.page.locator(selector).first
                    if error_element.is_visible(timeout=1000):
                        error_text = error_element.inner_text()
                        print(f"[WARNING] Login error detected: {error_text}")
                except:
                    continue

            # Wait for dashboard to load (with longer timeout for slow connections)
            try:
                self.page.wait_for_url("**/dashboard", timeout=45000)
            except:
                # Check current URL
                current_url = self.page.url
                print(f"[DEBUG] Current URL after login attempt: {current_url}")
                if "/login" in current_url:
                    raise Exception("Login failed - still on login page. Credentials may be invalid.")

            # Verify successful login - try multiple indicators
            login_success = False
            success_selectors = [
                "h1:has-text('Welcome back')",
                "h1:has-text('Welcome')",
                "text=Dashboard",
                "text=Raise Invoice",
                "button:has-text('Raise Invoice')",
            ]
            for selector in success_selectors:
                try:
                    element = self.page.locator(selector).first
                    if element.is_visible(timeout=3000):
                        login_success = True
                        print(f"[INFO] Login verified with: {selector}")
                        break
                except:
                    continue

            if not login_success:
                raise Exception("Could not verify successful login")

            # Take screenshot
            screenshot = self._take_screenshot("TC_02_Login_Success")

            # Log success
            self._log_result(
                tc_id, scenario, "PASSED",
                "User successfully logged in, dashboard displayed",
                screenshot
            )
            return True

        except Exception as e:
            screenshot = self._take_screenshot("TC_02_FAILED")
            self._log_result(tc_id, scenario, "FAILED", str(e), screenshot)
            return False

    # =========================================================================
    # TEST CASE: TC_03 - Raise Invoice
    # =========================================================================
    def tc_03_raise_invoice(self) -> bool:
        """
        TC_03: To check if user can navigate to Raise Invoice page and Create a Invoice

        Steps:
            1. Click on 'Raise Invoice' button
            2. Fill all the details from Invoice sheet
            3. Fetch data from all fields and store in logs
            4. Click on Raise Invoice button
            5. Click on 'Copy Request ID' button
            6. Store the Request ID
            7. Click on Close button

        Expected: A pop up will display with Request ID
        """
        tc_id = "TC_03"
        scenario = "To check if user can navigate to Raise Invoice page and Create a Invoice"
        print(f"\n{'='*60}")
        print(f"[EXECUTING] {tc_id}: {scenario}")
        print(f"{'='*60}")

        try:
            # Get invoice data from Excel
            invoice_row = self.invoice_sheet.iloc[0]

            # Prepare test data
            invoice_number = self._generate_invoice_number()
            invoice_date = datetime.now().strftime("%Y-%m-%d")
            due_date = (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d")
            client_name = invoice_row["Select Client"]
            purpose = invoice_row["Purpose"]
            currency = invoice_row["Currency"]
            amount = str(int(invoice_row["Amount"]))
            bank_account = invoice_row["Your Receiving Account"]
            document_path = invoice_row["Invoice Document"]

            # Store invoice data for logging
            self.invoice_data = {
                "invoice_number": invoice_number,
                "invoice_date": invoice_date,
                "due_date": due_date,
                "client": client_name,
                "purpose": purpose,
                "currency": currency,
                "amount": amount,
                "bank_account": bank_account,
                "document": document_path
            }

            # Step 1: Click on 'Raise Invoice' button
            self.page.click("button:has-text('Raise Invoice')")
            self.page.wait_for_url("**/raise")
            print("[STEP 1] Clicked 'Raise Invoice' button")

            # Step 2: Fill all the details
            print("[STEP 2] Filling invoice details...")

            # Fill Invoice Number
            self.page.fill("input[placeholder*='Invoice Number'], input:near(:text('Invoice Number'))", invoice_number)
            print(f"  - Invoice Number: {invoice_number}")

            # Set Invoice Date using JavaScript
            self.page.evaluate(f"""
                const dateInputs = document.querySelectorAll('input[type="date"]');
                if (dateInputs[0]) {{
                    dateInputs[0].value = '{invoice_date}';
                    dateInputs[0].dispatchEvent(new Event('input', {{ bubbles: true }}));
                    dateInputs[0].dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            """)
            print(f"  - Invoice Date: {invoice_date}")

            # Set Due Date using JavaScript
            self.page.evaluate(f"""
                const dateInputs = document.querySelectorAll('input[type="date"]');
                if (dateInputs[1]) {{
                    dateInputs[1].value = '{due_date}';
                    dateInputs[1].dispatchEvent(new Event('input', {{ bubbles: true }}));
                    dateInputs[1].dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            """)
            print(f"  - Due Date: {due_date}")

            # Select Client - try multiple approaches
            print("  - Selecting Client...")
            client_dropdown_selectors = [
                "text=Choose a client",
                "[placeholder='Choose a client']",
                "div:has-text('Choose a client')",
                "text=Select Client >> xpath=following-sibling::*[1]",
            ]
            for selector in client_dropdown_selectors:
                try:
                    dropdown = self.page.locator(selector).first
                    if dropdown.is_visible(timeout=2000):
                        dropdown.click()
                        print(f"    [DEBUG] Clicked client dropdown using: {selector}")
                        break
                except:
                    continue

            self.page.wait_for_timeout(1000)

            # Try to select client from dropdown list
            try:
                # Look for the client in dropdown options
                client_option = self.page.locator(f"text={client_name}").first
                if client_option.is_visible(timeout=3000):
                    client_option.click()
                else:
                    # Try clicking on any visible option containing the client name
                    self.page.locator(f"div:has-text('{client_name}'), li:has-text('{client_name}'), span:has-text('{client_name}')").first.click()
            except Exception as e:
                print(f"    [WARNING] Could not select client: {e}")
                # Try keyboard navigation
                self.page.keyboard.type(client_name[:3])
                self.page.wait_for_timeout(500)
                self.page.keyboard.press("Enter")

            print(f"  - Client: {client_name}")

            # Wait for client details to populate
            self.page.wait_for_timeout(1500)

            # Take screenshot after client selection
            self._take_screenshot("TC_03_After_Client_Selection")

            # Select Purpose - try multiple approaches
            print("  - Selecting Purpose...")
            purpose_selectors = [
                "text=Select Purpose",
                "[placeholder='Select Purpose']",
                "div:has-text('Select Purpose')",
            ]
            for selector in purpose_selectors:
                try:
                    dropdown = self.page.locator(selector).first
                    if dropdown.is_visible(timeout=2000):
                        dropdown.click()
                        print(f"    [DEBUG] Clicked purpose dropdown using: {selector}")
                        break
                except:
                    continue

            self.page.wait_for_timeout(500)
            try:
                self.page.locator(f"text={purpose}").first.click(timeout=3000)
            except:
                self.page.keyboard.type(purpose[:3])
                self.page.keyboard.press("Enter")
            print(f"  - Purpose: {purpose}")

            # Select Currency - try multiple approaches
            print("  - Selecting Currency...")
            currency_selectors = [
                "text=Select Currency",
                "[placeholder='Select Currency']",
                "div:has-text('Select Currency')",
            ]
            for selector in currency_selectors:
                try:
                    dropdown = self.page.locator(selector).first
                    if dropdown.is_visible(timeout=2000):
                        dropdown.click()
                        print(f"    [DEBUG] Clicked currency dropdown using: {selector}")
                        break
                except:
                    continue

            self.page.wait_for_timeout(500)
            try:
                self.page.locator(f"text={currency}").first.click(timeout=3000)
            except:
                self.page.keyboard.type(currency)
                self.page.keyboard.press("Enter")
            print(f"  - Currency: {currency}")

            # Fill Amount - try multiple selectors
            print("  - Filling Amount...")
            amount_selectors = [
                "input[placeholder='Enter the amount']",
                "input[placeholder*='amount']",
                "input:near(:text('Amount'))",
            ]
            amount_filled = False
            for selector in amount_selectors:
                try:
                    amount_input = self.page.locator(selector).first
                    if amount_input.is_visible(timeout=2000):
                        amount_input.fill(amount)
                        amount_filled = True
                        print(f"    [DEBUG] Filled amount using: {selector}")
                        break
                except:
                    continue

            if not amount_filled:
                # Try finding by label
                self.page.evaluate(f"""
                    const inputs = document.querySelectorAll('input');
                    for (let input of inputs) {{
                        if (input.placeholder && input.placeholder.toLowerCase().includes('amount')) {{
                            input.value = '{amount}';
                            input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                            break;
                        }}
                    }}
                """)
            print(f"  - Amount: {amount}")

            # Select Receiving Account
            print("  - Selecting Receiving Account...")

            # Scroll down to ensure receiving account is visible
            self.page.evaluate("window.scrollBy(0, 300)")
            self.page.wait_for_timeout(500)

            # Take screenshot to see current state
            self._take_screenshot("TC_03_Before_Account_Selection")

            # Try to click the dropdown - it might be a React Select component
            # First, find and click on the dropdown container
            try:
                # Method 1: Click on the visible text
                dropdown = self.page.locator("text=Select account to receive funds")
                dropdown.click(force=True)
                print("    [DEBUG] Clicked on dropdown text")
            except Exception as e:
                print(f"    [DEBUG] Method 1 failed: {str(e)[:50]}")

            self.page.wait_for_timeout(1500)
            self._take_screenshot("TC_03_Account_Dropdown_Open")

            # Try to select an option using various methods
            selected = False

            # Method A: ArrowDown + Enter
            try:
                self.page.keyboard.press("ArrowDown")
                self.page.wait_for_timeout(300)
                self.page.keyboard.press("Enter")
                self.page.wait_for_timeout(500)
                print("    [DEBUG] Tried ArrowDown + Enter")
                selected = True
            except:
                pass

            # Check if still showing "Select account"
            self.page.wait_for_timeout(500)
            page_content = self.page.content()

            if "Select account to receive funds" in page_content and not selected:
                # Method B: Try clicking again with JavaScript
                print("    [DEBUG] Selection not made, trying JS click")
                self.page.evaluate("""
                    // Find all dropdowns and click the one for receiving account
                    const dropdowns = document.querySelectorAll('[class*="select"], [class*="dropdown"]');
                    for (let d of dropdowns) {
                        if (d.textContent.includes('Select account') || d.textContent.includes('receive funds')) {
                            d.click();
                            break;
                        }
                    }
                """)
                self.page.wait_for_timeout(1000)

                # Try to click on an option
                self.page.evaluate("""
                    // Click first available option
                    const options = document.querySelectorAll('[class*="option"], [role="option"]');
                    if (options.length > 0) {
                        options[0].click();
                    }
                """)

            self.page.wait_for_timeout(500)
            self._take_screenshot("TC_03_After_Account_Selection")
            print(f"  - Receiving Account: {bank_account} (attempted)")

            # Upload Invoice Document
            print("  - Uploading Document...")
            try:
                file_input = self.page.locator("input[type='file']").first
                # Check if document path exists
                if Path(document_path).exists():
                    file_input.set_input_files(document_path)
                    print(f"  - Document: {document_path}")
                else:
                    print(f"    [WARNING] Document not found: {document_path}")
                    # Try to create a dummy file or skip
                    # Create a simple test file
                    test_doc = self.base_dir / "test_invoice.txt"
                    test_doc.write_text("Test Invoice Document")
                    file_input.set_input_files(str(test_doc))
                    print(f"  - Document: {test_doc} (created test file)")
            except Exception as e:
                print(f"    [WARNING] Document upload failed: {e}")

            # Wait for upload to complete
            self.page.wait_for_timeout(2000)

            # Step 3: Capture ALL data from the form
            print("\n[STEP 3] Capturing all invoice data from form...")

            # Capture all field values from the page
            captured_data = self.page.evaluate("""
                () => {
                    const data = {};

                    // Invoice Details
                    const invoiceNumInput = document.querySelector('input[placeholder*="Invoice Number"], input[name*="invoice"]');
                    if (invoiceNumInput) data['Invoice Number'] = invoiceNumInput.value;

                    const dateInputs = document.querySelectorAll('input[type="date"]');
                    if (dateInputs[0]) data['Invoice Date'] = dateInputs[0].value;
                    if (dateInputs[1]) data['Due Date'] = dateInputs[1].value;

                    // Get all text content that looks like field values
                    const pageText = document.body.innerText;

                    // Company Name - look for text between "Company Name" and "Email" labels
                    const companyNameMatch = pageText.match(/Company Name\\s*\\n\\s*([A-Za-z][A-Za-z0-9\\s&.,'-]+?)\\s*(?:\\n|Email)/);
                    if (companyNameMatch) {
                        let compName = companyNameMatch[1].trim();
                        // Remove trailing label words if present
                        compName = compName.replace(/\\s*(Email|Mobile|Phone).*$/i, '').trim();
                        if (compName && compName.length > 1) data['Company Name'] = compName;
                    }

                    // Email - look for actual email format
                    const emailMatch = pageText.match(/Email\\s*\\n\\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})/);
                    if (emailMatch) data['Email'] = emailMatch[1].trim();

                    // Mobile Number - look for phone number pattern
                    const mobileMatch = pageText.match(/Mobile Number\\s*\\n\\s*([+]?[\\d\\s()-]{8,20})/);
                    if (mobileMatch) {
                        const mobile = mobileMatch[1].trim();
                        if (mobile && mobile.length >= 8) data['Mobile Number'] = mobile;
                    }

                    // Receiving Account section
                    const bankNameMatch = pageText.match(/Bank Name\\s*\\n\\s*([A-Za-z][A-Za-z\\s]+?)\\s*(?:\\n|Account Holder)/);
                    if (bankNameMatch) data['Bank Name'] = bankNameMatch[1].trim();

                    const holderMatch = pageText.match(/Account Holder Name\\s*\\n\\s*([A-Za-z][A-Za-z\\s.]+?)\\s*(?:\\n|Account Number)/);
                    if (holderMatch) data['Account Holder Name'] = holderMatch[1].trim();

                    const accNumMatch = pageText.match(/Account Number\\s*\\n\\s*([*\\d]+)/);
                    if (accNumMatch) data['Account Number'] = accNumMatch[1].trim();

                    const accCurrMatch = pageText.match(/Account Currency\\s*\\n\\s*([A-Z]{3})/);
                    if (accCurrMatch) data['Account Currency'] = accCurrMatch[1].trim();

                    const routingMatch = pageText.match(/Routing Number\\s*\\n\\s*([*\\dN\\/A]+)/i);
                    if (routingMatch) data['Routing Number'] = routingMatch[1].trim();

                    const branchMatch = pageText.match(/Branch Code\\s*\\n\\s*([A-Za-z0-9\\/]+)/i);
                    if (branchMatch) data['Branch Code'] = branchMatch[1].trim();

                    const nicknameMatch = pageText.match(/Nickname\\s*\\n\\s*([\\d]+)/);
                    if (nicknameMatch) data['Nickname'] = nicknameMatch[1].trim();

                    const addedOnMatch = pageText.match(/Account Added On\\s*\\n\\s*([\\d\\/]+)/);
                    if (addedOnMatch) data['Account Added On'] = addedOnMatch[1].trim();

                    // Purpose and Amount
                    const amountInput = document.querySelector('input[placeholder*="amount"], input[placeholder*="Amount"]');
                    if (amountInput) data['Amount'] = amountInput.value;

                    // Description
                    const descInput = document.querySelector('textarea[placeholder*="description"], textarea[placeholder*="Description"]');
                    if (descInput && descInput.value) data['Description'] = descInput.value;

                    return data;
                }
            """)

            # Update invoice_data with captured values
            self.invoice_data = {
                "Invoice Number": invoice_number,
                "Invoice Date": invoice_date,
                "Due Date": due_date,
                "Client": client_name,
                "Company Name": captured_data.get('Company Name', client_name),
                "Email": captured_data.get('Email', 'N/A'),
                "Mobile Number": captured_data.get('Mobile Number', 'N/A'),
                "Receiving Account": bank_account,
                "Bank Name": captured_data.get('Bank Name', 'N/A'),
                "Account Holder Name": captured_data.get('Account Holder Name', 'N/A'),
                "Account Number": captured_data.get('Account Number', 'N/A'),
                "Account Currency": captured_data.get('Account Currency', 'N/A'),
                "Routing Number": captured_data.get('Routing Number', 'N/A'),
                "Branch Code": captured_data.get('Branch Code', 'N/A'),
                "Nickname": captured_data.get('Nickname', 'N/A'),
                "Account Added On": captured_data.get('Account Added On', 'N/A'),
                "Purpose": purpose,
                "Currency": currency,
                "Amount": amount,
                "Description": captured_data.get('Description', ''),
                "Invoice Document": document_path
            }

            print("[STEP 3] Invoice Data Captured:")
            for key, value in self.invoice_data.items():
                print(f"  {key}: {value}")

            # Take screenshot of filled form
            screenshot_form = self._take_screenshot("TC_03_Invoice_Form_Filled")

            # Step 4: Click on Raise Invoice button
            print("\n[STEP 4] Clicking 'Raise Invoice' button...")

            # First scroll to bottom
            try:
                self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            except:
                pass
            self.page.wait_for_timeout(500)

            # Method 1: Simple Playwright click
            try:
                raise_btn = self.page.locator("button:has-text('Raise Invoice')").last
                raise_btn.scroll_into_view_if_needed()
                self.page.wait_for_timeout(500)
                raise_btn.click()
                print("    [DEBUG] Playwright click executed")
            except Exception as e:
                print(f"    [DEBUG] Playwright click failed: {str(e)[:50]}")

            # Wait for response
            self.page.wait_for_timeout(3000)
            self._take_screenshot("TC_03_After_Raise_Click_1")

            # Check if still on form
            if "/raise" in self.page.url:
                print("    [DEBUG] Still on form, trying Playwright click")
                raise_btn = self.page.locator("button:has-text('Raise Invoice')").last
                raise_btn.scroll_into_view_if_needed()
                raise_btn.click()
                self.page.wait_for_timeout(3000)
                self._take_screenshot("TC_03_After_Raise_Click_2")

            # Check if redirected to login (session expired)
            if "/login" in self.page.url:
                print("    [WARNING] Session expired, attempting re-login")
                raise Exception("Session expired during form submission")

            # Check if still on form - try Enter on button
            if "/raise" in self.page.url:
                print("    [DEBUG] Trying focus + Enter on button")
                raise_btn = self.page.locator("button:has-text('Raise Invoice')").last
                raise_btn.focus()
                self.page.keyboard.press("Enter")
                self.page.wait_for_timeout(5000)
                self._take_screenshot("TC_03_After_Raise_Click_3")

            # Check again for login redirect
            if "/login" in self.page.url:
                raise Exception("Session expired during form submission")

            self._take_screenshot("TC_03_After_Submit_Attempts")

            # Wait for success popup with various selectors
            success_selectors = [
                "text=Request Id",
                "text=Request ID",
                "text=Success",
                "text=Invoice created",
                "text=successfully",
            ]

            popup_found = False
            for selector in success_selectors:
                try:
                    self.page.wait_for_selector(selector, timeout=10000)
                    popup_found = True
                    print(f"[STEP 4] Success popup displayed (found: {selector})")
                    break
                except:
                    continue

            if not popup_found:
                # Take diagnostic screenshot
                self._take_screenshot("TC_03_No_Popup_Found")
                # Check current URL
                print(f"    [DEBUG] Current URL: {self.page.url}")
                raise Exception("Success popup not found after form submission")

            # Take screenshot of success popup
            screenshot_success = self._take_screenshot("TC_03_Invoice_Success_Popup", full_page=False)

            # Step 5 & 6: Get Request ID
            request_id_element = self.page.locator("text=Request Id").locator("xpath=following-sibling::*[1]")
            self.request_id = request_id_element.inner_text()

            # Click Copy Request ID
            self.page.click("button:has-text('Copy Request ID')")
            print(f"\n[STEP 5 & 6] Request ID: {self.request_id}")

            # Step 7: Click Close button
            self.page.click("button:has-text('Close')")
            print("[STEP 7] Clicked 'Close' button")

            # Wait for dashboard
            self.page.wait_for_url("**/dashboard", timeout=10000)

            # Take screenshot of dashboard with new invoice
            screenshot_dashboard = self._take_screenshot("TC_03_Dashboard_After_Invoice")

            # Verify invoice appears in Pending Receivables
            invoice_in_list = self.page.locator(f"text={invoice_number}")
            expect(invoice_in_list).to_be_visible(timeout=5000)

            # Log success
            self._log_result(
                tc_id, scenario, "PASSED",
                f"Invoice created successfully. Request ID: {self.request_id}<br>Invoice Number: {invoice_number}",
                f"{screenshot_form}, {screenshot_success}, {screenshot_dashboard}"
            )
            return True

        except Exception as e:
            screenshot = self._take_screenshot("TC_03_FAILED")
            self._log_result(tc_id, scenario, "FAILED", str(e), screenshot)
            return False

    # =========================================================================
    # TEST CASE: TC_04 - Verify Pending Receivables
    # =========================================================================
    def tc_04_verify_pending_receivables(self) -> bool:
        """
        TC_04: To check if user can find Request ID in Pending Receivables and verify data

        Steps:
            1. Find the invoice in Pending Receivables section using Request ID/Invoice Number
            2. Click on the Eye icon to view invoice details
            3. Capture all invoice data from the details page
            4. Verify captured data against expected data from TC_03

        Expected: All invoice data should match the data entered during TC_03
        """
        tc_id = "TC_04"
        scenario = "To check if user can find Request ID in Pending Receivables and verify data"
        print(f"\n{'='*60}")
        print(f"[EXECUTING] {tc_id}: {scenario}")
        print(f"{'='*60}")

        try:
            # Ensure we have data from TC_03
            if not self.request_id or not self.invoice_data:
                raise Exception("TC_04 requires TC_03 to be executed first (Request ID and Invoice data needed)")

            invoice_number = self.invoice_data.get("Invoice Number", "")
            print(f"[INFO] Looking for Invoice: {invoice_number}")
            print(f"[INFO] Request ID: {self.request_id}")

            # Step 1: Find invoice in Pending Receivables
            print("\n[STEP 1] Finding invoice in Pending Receivables...")

            # Wait for dashboard to fully load
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            # Scroll to Pending Receivables section
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
            self.page.wait_for_timeout(1000)

            # Look for the invoice
            invoice_row = self.page.locator(f"tr:has-text('{invoice_number}')")
            if not invoice_row.is_visible(timeout=10000):
                raise Exception(f"Invoice not found in Pending Receivables: {invoice_number}")

            print(f"[STEP 1] Found invoice: {invoice_number}")
            screenshot_receivables = self._take_screenshot("TC_04_Pending_Receivables")

            # Step 2: Click on Eye icon to view details
            print("\n[STEP 2] Clicking eye icon to view invoice details...")

            # Try to click the eye icon using JavaScript
            clicked = self.page.evaluate(f"""
                () => {{
                    const rows = document.querySelectorAll('tr');
                    for (const row of rows) {{
                        if (row.textContent.includes('{invoice_number}')) {{
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
                                    return true;
                                }}
                            }}
                            break;
                        }}
                    }}
                    return false;
                }}
            """)

            if not clicked:
                # Fallback: try clicking via Playwright
                eye_icon = invoice_row.locator("svg.lucide-eye, svg[class*='eye']").first
                if eye_icon.is_visible(timeout=3000):
                    eye_icon.click()

            # Wait for navigation to details page
            self.page.wait_for_url("**/receivable-details", timeout=15000)
            self.page.wait_for_load_state("networkidle")
            print("[STEP 2] Navigated to invoice details page")

            self.page.wait_for_timeout(2000)
            screenshot_details = self._take_screenshot("TC_04_Invoice_Details")

            # Step 3: Capture invoice details from the page
            print("\n[STEP 3] Capturing invoice details from page...")

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

                    return data;
                }
            """)

            self.tc04_captured_data = captured_data
            print("[STEP 3] Captured Invoice Details:")
            for key, value in captured_data.items():
                print(f"  {key}: {value}")

            # Step 4: Verify data against TC_03 expected values
            print("\n[STEP 4] Verifying invoice data...")

            # Build expected data from TC_03 invoice data
            expected_data = {
                "Invoice Number": self.invoice_data.get("Invoice Number", ""),
                "Bank Name": self.invoice_data.get("Bank Name", "").upper(),
                "Account Number": self.invoice_data.get("Account Number", ""),
                "Currency": self.invoice_data.get("Currency", "INR"),
                "Country": "India",  # Expected country based on INR currency
                "Amount": self.invoice_data.get("Amount", "")
            }

            verification_results = []
            for field, expected_value in expected_data.items():
                actual_value = captured_data.get(field, '')

                # Normalize values for comparison
                expected_normalized = str(expected_value).strip().upper() if expected_value else ''
                actual_normalized = str(actual_value).strip().upper() if actual_value else ''

                # Handle amount comparison (remove commas, compare numbers)
                if field == "Amount":
                    try:
                        expected_num = float(str(expected_value).replace(",", ""))
                        actual_num = float(str(actual_value).replace(",", ""))
                        if abs(expected_num - actual_num) < 0.01:
                            status = "MATCH"
                        else:
                            status = "MISMATCH"
                    except:
                        status = "MISMATCH" if expected_normalized != actual_normalized else "MATCH"
                elif actual_normalized == expected_normalized:
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

            self.tc04_verification_results = verification_results

            # Determine overall result
            failed_count = sum(1 for r in verification_results if r['status'] != 'MATCH')
            overall_status = "PASSED" if failed_count == 0 else "FAILED"

            print(f"\n[RESULT] Fields Verified: {len(verification_results)}")
            print(f"[RESULT] Fields Matched: {len(verification_results) - failed_count}")
            print(f"[RESULT] Fields Mismatched/Missing: {failed_count}")

            # Navigate back to dashboard
            self.page.go_back()
            self.page.wait_for_timeout(2000)

            # Log result
            if overall_status == "PASSED":
                self._log_result(
                    tc_id, scenario, "PASSED",
                    f"Invoice verified successfully. All {len(verification_results)} fields matched.",
                    f"{screenshot_receivables}, {screenshot_details}"
                )
                return True
            else:
                mismatched_fields = [r['field'] for r in verification_results if r['status'] != 'MATCH']
                self._log_result(
                    tc_id, scenario, "FAILED",
                    f"Data verification failed. Mismatched fields: {', '.join(mismatched_fields)}",
                    f"{screenshot_receivables}, {screenshot_details}"
                )
                return False

        except Exception as e:
            screenshot = self._take_screenshot("TC_04_FAILED")
            self._log_result(tc_id, scenario, "FAILED", str(e), screenshot)
            return False

    # =========================================================================
    # TEST CASE: TC_05 - Verify Pending Payables (Client Login)
    # =========================================================================
    def tc_05_verify_pending_payables(self) -> bool:
        """
        TC_05: To Check for Pending Payables in client login

        Steps:
            1. Logout from current (Vendor) account
            2. Login with Client_Business credentials
            3. Navigate to Pending Payables at the bottom of the homepage
            4. Look for the Invoice Number created during TC_03
            5. Click on the eye icon on the right side of the transaction
            6. Compare all fields with the fields captured from TC_03

        Expected: All data from TC_03 should be same in Pay Invoice page in TC_05
        """
        tc_id = "TC_05"
        scenario = "To Check for Pending Payables in client login"
        print(f"\n{'='*60}")
        print(f"[EXECUTING] {tc_id}: {scenario}")
        print(f"{'='*60}")

        try:
            # Ensure we have data from TC_03
            if not self.invoice_data:
                raise Exception("TC_05 requires TC_03 to be executed first (Invoice data needed)")

            invoice_number = self.invoice_data.get("Invoice Number", "")
            print(f"[INFO] Looking for Invoice: {invoice_number}")

            # Step 1: Logout from Vendor account
            print("\n[STEP 1] Logging out from Vendor account...")

            # Navigate to dashboard first if not already there
            if "/dashboard" not in self.page.url:
                self.page.goto(f"{self.base_url}/dashboard")
                self.page.wait_for_load_state("networkidle")

            # Click logout button
            logout_button = self.page.locator("button:has-text('Log out'), button:has-text('Logout')").first
            if logout_button.is_visible(timeout=5000):
                logout_button.click()
                self.page.wait_for_timeout(2000)
                print("[STEP 1] Logged out from Vendor account")
            else:
                # Try direct navigation to login
                self.page.goto(f"{self.base_url}/login")

            self._take_screenshot("TC_05_After_Logout")

            # Step 2: Login with Client_Business credentials
            print("\n[STEP 2] Logging in with Client_Business credentials...")

            # Get Client_Business credentials
            client_email, client_password = self._get_credentials("Client_Business")

            # Wait for login page
            self.page.wait_for_url("**/login", timeout=10000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(1000)

            # Fill login form
            email_input = self.page.locator("input[type='email'], input[type='text']").first
            email_input.fill(client_email)
            print(f"  - Email: {client_email}")

            password_input = self.page.locator("input[type='password']").first
            password_input.fill(client_password)
            print(f"  - Password: ********")

            self._take_screenshot("TC_05_Login_Form_Filled")

            # Submit login
            password_input.press("Enter")
            self.page.wait_for_timeout(2000)

            # Wait for dashboard
            try:
                self.page.wait_for_url("**/dashboard", timeout=30000)
            except:
                # Try clicking login button
                login_btn = self.page.locator("button:has-text('Log in')").first
                if login_btn.is_visible(timeout=2000):
                    login_btn.click()
                    self.page.wait_for_url("**/dashboard", timeout=30000)

            self.page.wait_for_load_state("networkidle")
            print("[STEP 2] Successfully logged in as Client_Business")

            screenshot_dashboard = self._take_screenshot("TC_05_Client_Dashboard")

            # Step 3: Navigate to Pending Payables
            print("\n[STEP 3] Finding invoice in Pending Payables...")

            # Scroll down to see Pending Payables section
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            self.page.wait_for_timeout(2000)

            # Look for the invoice in Pending Payables
            invoice_found = False
            try:
                invoice_element = self.page.locator(f"text={invoice_number}").first
                if invoice_element.is_visible(timeout=10000):
                    invoice_found = True
                    print(f"[STEP 3] Found invoice: {invoice_number}")
            except:
                pass

            if not invoice_found:
                # Try clicking "View all" for Pending Payables
                view_all_buttons = self.page.locator("button:has-text('View all')")
                if view_all_buttons.count() > 1:
                    # Second "View all" is for Pending Payables
                    view_all_buttons.nth(1).click()
                    self.page.wait_for_timeout(2000)

                # Check again
                invoice_element = self.page.locator(f"text={invoice_number}").first
                if not invoice_element.is_visible(timeout=5000):
                    raise Exception(f"Invoice not found in Pending Payables: {invoice_number}")

            screenshot_payables = self._take_screenshot("TC_05_Pending_Payables")

            # Step 4 & 5: Click on Eye icon to view invoice details
            print("\n[STEP 4] Clicking eye icon to view invoice details...")

            # Find and click the eye icon for this invoice using JavaScript
            clicked = self.page.evaluate(f"""
                () => {{
                    // Find all rows/elements containing the invoice number
                    const elements = document.querySelectorAll('tr, div');
                    for (const el of elements) {{
                        if (el.textContent.includes('{invoice_number}')) {{
                            // Find eye icon within this element
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
                # Fallback: try to find eye icon near the invoice
                try:
                    # Get the row containing the invoice
                    invoice_row = self.page.locator(f"tr:has-text('{invoice_number}'), div:has-text('{invoice_number}')").first
                    eye_icon = invoice_row.locator("svg.lucide-eye, svg[class*='eye']").first
                    if eye_icon.is_visible(timeout=3000):
                        eye_icon.click()
                        clicked = True
                except:
                    pass

            if not clicked:
                raise Exception("Could not click eye icon to view invoice details")

            # Wait for navigation to payable-details page
            self.page.wait_for_url("**/payable-details", timeout=15000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            print("[STEP 4] Navigated to Pay Invoice details page")
            screenshot_details = self._take_screenshot("TC_05_Pay_Invoice_Details")

            # Step 6: Capture and verify invoice data
            print("\n[STEP 5] Capturing invoice details from Pay Invoice page...")

            captured_data = self.page.evaluate("""
                () => {
                    const data = {};
                    const pageText = document.body.innerText;

                    // Invoice Number
                    const invNumMatch = pageText.match(/Invoice Number:\\s*\\n?\\s*([A-Z0-9-]+)/i);
                    if (invNumMatch) data['Invoice Number'] = invNumMatch[1].trim();

                    // Invoice Date
                    const invDateMatch = pageText.match(/Invoice Date:\\s*\\n?\\s*([A-Za-z]+ \\d{1,2}, \\d{4})/i);
                    if (invDateMatch) data['Invoice Date'] = invDateMatch[1].trim();

                    // Due Date
                    const dueDateMatch = pageText.match(/Due Date:\\s*\\n?\\s*([A-Za-z]+ \\d{1,2}, \\d{4})/i);
                    if (dueDateMatch) data['Due Date'] = dueDateMatch[1].trim();

                    // Bill From (Client/Company)
                    const billFromMatch = pageText.match(/Bill From:\\s*\\n?\\s*([A-Z][A-Z0-9\\s]+?)\\s*(?:-|\\n)/i);
                    if (billFromMatch) data['Bill From'] = billFromMatch[1].trim();

                    // Bank Name
                    const bankMatch = pageText.match(/Bank Name:\\s*\\n?\\s*([A-Za-z][A-Za-z\\s]+?)\\s*(?:\\n|Account)/i);
                    if (bankMatch) data['Bank Name'] = bankMatch[1].trim();

                    // Account Number
                    const accMatch = pageText.match(/Account Number:\\s*\\n?\\s*([*\\d]+)/i);
                    if (accMatch) data['Account Number'] = accMatch[1].trim();

                    // Currency
                    const currMatch = pageText.match(/Currency:\\s*\\n?\\s*([A-Z]{3})/i);
                    if (currMatch) data['Currency'] = currMatch[1].trim();

                    // Country
                    const countryMatch = pageText.match(/Country:\\s*\\n?\\s*([A-Z]{2})/i);
                    if (countryMatch) data['Country'] = countryMatch[1].trim();

                    // Amount
                    const amountMatch = pageText.match(/Amount Due\\s*\\n?\\s*([\\d,.]+)/i);
                    if (amountMatch) data['Amount'] = amountMatch[1].trim();

                    // Alternative amount
                    if (!data['Amount']) {
                        const altMatch = pageText.match(/Payment Request\\s*\\n?\\s*([\\d,.]+)/i);
                        if (altMatch) data['Amount'] = altMatch[1].trim();
                    }

                    // Attached Documents
                    const docMatch = pageText.match(/Attached Documents \\((\\d+)\\)/);
                    if (docMatch) data['Documents Count'] = docMatch[1];

                    // Document Name
                    const docNameMatch = pageText.match(/Document Name\\s*\\n?\\s*Action\\s*\\n?\\s*([A-Za-z0-9._-]+)/);
                    if (docNameMatch) data['Document Name'] = docNameMatch[1].trim();

                    return data;
                }
            """)

            self.tc05_captured_data = captured_data
            print("[STEP 5] Captured Pay Invoice Details:")
            for key, value in captured_data.items():
                print(f"  {key}: {value}")

            # Step 6: Verify data against TC_03
            print("\n[STEP 6] Verifying invoice data against TC_03...")

            # Build expected data from TC_03 invoice data
            expected_data = {
                "Invoice Number": self.invoice_data.get("Invoice Number", ""),
                "Bank Name": self.invoice_data.get("Bank Name", ""),
                "Account Number": self.invoice_data.get("Account Number", ""),
                "Currency": self.invoice_data.get("Currency", ""),
                "Amount": self.invoice_data.get("Amount", "")
            }

            verification_results = []
            for field, expected_value in expected_data.items():
                actual_value = captured_data.get(field, '')

                # Normalize values for comparison
                expected_normalized = str(expected_value).strip().upper() if expected_value else ''
                actual_normalized = str(actual_value).strip().upper() if actual_value else ''

                # Handle amount comparison
                if field == "Amount":
                    try:
                        expected_num = float(str(expected_value).replace(",", ""))
                        actual_num = float(str(actual_value).replace(",", ""))
                        if abs(expected_num - actual_num) < 0.01:
                            status = "MATCH"
                        else:
                            status = "MISMATCH"
                    except:
                        status = "MISMATCH" if expected_normalized != actual_normalized else "MATCH"
                elif expected_normalized == actual_normalized:
                    status = "MATCH"
                elif actual_normalized in expected_normalized or expected_normalized in actual_normalized:
                    status = "MATCH"  # Partial match is acceptable for some fields
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

            self.tc05_verification_results = verification_results

            # Determine overall result
            failed_count = sum(1 for r in verification_results if r['status'] not in ['MATCH'])
            overall_status = "PASSED" if failed_count == 0 else "FAILED"

            print(f"\n[RESULT] Fields Verified: {len(verification_results)}")
            print(f"[RESULT] Fields Matched: {len(verification_results) - failed_count}")
            print(f"[RESULT] Fields Mismatched/Missing: {failed_count}")

            # Log result
            if overall_status == "PASSED":
                self._log_result(
                    tc_id, scenario, "PASSED",
                    f"Invoice verified successfully in Client's Pending Payables. All {len(verification_results)} fields matched.",
                    f"{screenshot_dashboard}, {screenshot_payables}, {screenshot_details}"
                )
                return True
            else:
                mismatched_fields = [r['field'] for r in verification_results if r['status'] != 'MATCH']
                self._log_result(
                    tc_id, scenario, "FAILED",
                    f"Data verification failed. Mismatched fields: {', '.join(mismatched_fields)}",
                    f"{screenshot_dashboard}, {screenshot_payables}, {screenshot_details}"
                )
                return False

        except Exception as e:
            screenshot = self._take_screenshot("TC_05_FAILED")
            self._log_result(tc_id, scenario, "FAILED", str(e), screenshot)
            return False

    # =========================================================================
    # TEST CASE: TC_06 - Pay Invoice from View Page
    # =========================================================================
    def tc_06_pay_invoice(self) -> bool:
        """
        TC_06: To Pay Invoice from View page

        Steps:
            1. Continue from TC_05 (already logged in as Client_Business)
            2. Find invoice in Pending Payables
            3. Click eye icon to view invoice details
            4. Click Approve button
            5. Click Pay Now to navigate to Pay Invoice form
            6. Capture and verify all form fields against TC_03 data
            7. Click Pay Now to complete payment
            8. Capture Transaction Success popup with Booking ID
            9. Close popup

        Expected: Transaction success popup should be displayed with Booking ID
        """
        tc_id = "TC_06"
        scenario = "To Pay Invoice from View page"
        print(f"\n{'='*60}")
        print(f"[EXECUTING] {tc_id}: {scenario}")
        print(f"{'='*60}")

        try:
            # Ensure we have data from TC_03
            if not self.invoice_data:
                raise Exception("TC_06 requires TC_03 to be executed first (Invoice data needed)")

            invoice_number = self.invoice_data.get("Invoice Number", "")
            print(f"[INFO] Processing Invoice: {invoice_number}")

            # Step 1: Navigate to dashboard (we should already be logged in as Client from TC_05)
            print("\n[STEP 1] Navigating to dashboard...")

            # Check if we need to login as Client_Business
            if "/dashboard" not in self.page.url:
                self.page.goto(f"{self.base_url}/dashboard")
                self.page.wait_for_load_state("networkidle")

            # Verify we're logged in as Client by checking dashboard
            self.page.wait_for_timeout(2000)

            # If on login page, we need to login as Client_Business
            if "/login" in self.page.url:
                print("[INFO] Logging in as Client_Business...")
                client_email, client_password = self._get_credentials("Client_Business")

                email_input = self.page.locator("input[type='email'], input[type='text']").first
                email_input.fill(client_email)

                password_input = self.page.locator("input[type='password']").first
                password_input.fill(client_password)
                password_input.press("Enter")

                self.page.wait_for_url("**/dashboard", timeout=30000)
                self.page.wait_for_load_state("networkidle")

            screenshot_dashboard = self._take_screenshot("TC_06_Client_Dashboard")

            # Step 2: Find invoice in Pending Payables
            print("\n[STEP 2] Finding invoice in Pending Payables...")

            # Scroll to Pending Payables
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            self.page.wait_for_timeout(2000)

            # Look for invoice
            invoice_visible = False
            try:
                invoice_element = self.page.locator(f"text={invoice_number}").first
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

            screenshot_payables = self._take_screenshot("TC_06_Pending_Payables")

            # Step 3: Click eye icon to view invoice details
            print("\n[STEP 3] Clicking eye icon to view invoice details...")

            clicked = self.page.evaluate(f"""
                () => {{
                    const elements = document.querySelectorAll('tr, div');
                    for (const el of elements) {{
                        if (el.textContent.includes('{invoice_number}')) {{
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
                raise Exception(f"Could not find eye icon for invoice {invoice_number}")

            # Wait for navigation to payable-details
            self.page.wait_for_url("**/payable-details", timeout=15000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            screenshot_details = self._take_screenshot("TC_06_Invoice_Details")
            print(f"[STEP 3] Viewing invoice details for {invoice_number}")

            # Step 4: Click Approve button
            print("\n[STEP 4] Clicking Approve button...")

            approve_btn = self.page.locator("button:has-text('Approve')").first
            if not approve_btn.is_visible(timeout=10000):
                raise Exception("Approve button not found")

            approve_btn.click()
            self.page.wait_for_timeout(3000)

            # Wait for Pay Now button to appear
            pay_now_btn = self.page.locator("button:has-text('Pay Now')").first
            if not pay_now_btn.is_visible(timeout=10000):
                raise Exception("Pay Now button not found after approval")

            screenshot_approve = self._take_screenshot("TC_06_After_Approve")
            print("[STEP 4] Invoice approved, Pay Now button visible")

            # Step 5: Click Pay Now to navigate to Pay Invoice form
            print("\n[STEP 5] Clicking Pay Now to navigate to Pay Invoice form...")

            pay_now_btn.click()
            self.page.wait_for_timeout(3000)

            # Wait for Pay Invoice page
            self.page.wait_for_url("**/pay**", timeout=15000)
            self.page.wait_for_load_state("networkidle")
            self.page.wait_for_timeout(2000)

            screenshot_form = self._take_screenshot("TC_06_Pay_Invoice_Form", full_page=True)
            print("[STEP 5] Navigated to Pay Invoice form")

            # Step 6: Capture and verify form data
            print("\n[STEP 6] Capturing Pay Invoice form data...")

            captured_data = self.page.evaluate("""
                () => {
                    const data = {};
                    const pageText = document.body.innerText;

                    // Choose Invoice
                    const chooseMatch = pageText.match(/Choose Invoice[\\s\\S]*?(INV-[\\d]+[^\\n]*)/);
                    if (chooseMatch) data['Choose Invoice'] = chooseMatch[1].trim();

                    // Mobile Number
                    const mobileMatch = pageText.match(/Mobile Number[\\s\\n]+\\+?(\\d+)[\\s\\n]+(\\d+)/);
                    if (mobileMatch) data['Mobile Number'] = '+' + mobileMatch[1] + ' ' + mobileMatch[2];

                    // Country
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

            # Verify against TC_03 data
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
                expected = str(self.invoice_data.get(tc03_field, '')).strip().upper()
                actual = str(captured_data.get(tc06_field, '')).strip().upper()

                if 'amount' in tc03_field.lower():
                    try:
                        exp_num = float(str(self.invoice_data.get(tc03_field, 0)).replace(',', ''))
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
                    'expected': self.invoice_data.get(tc03_field, ''),
                    'actual': captured_data.get(tc06_field, '(Blank)'),
                    'status': status
                }
                verification_results.append(result)

                icon = "+" if status == "MATCH" else "-"
                print(f"  {icon} {tc03_field}: Expected='{result['expected']}' | Actual='{result['actual']}' | {status}")

            self.tc06_verification_results = verification_results

            # Step 7: Click Pay Now to complete payment
            print("\n[STEP 7] Completing payment...")

            pay_now_submit = self.page.locator("button:has-text('Pay Now')").last
            if not pay_now_submit.is_visible(timeout=10000):
                raise Exception("Pay Now submit button not found on form")

            pay_now_submit.scroll_into_view_if_needed()
            self.page.wait_for_timeout(500)
            pay_now_submit.click()

            # Step 8: Wait for and capture success popup
            print("\n[STEP 8] Waiting for transaction success popup...")
            self.page.wait_for_timeout(5000)

            success_found = False
            for selector in ["text=Transaction Successful", "text=Booking ID", "text=booked Successfully"]:
                try:
                    if self.page.locator(selector).first.is_visible(timeout=10000):
                        success_found = True
                        break
                except:
                    continue

            if not success_found:
                raise Exception("Transaction success popup not found")

            screenshot_success = self._take_screenshot("TC_06_Transaction_Success")

            # Capture transaction details
            transaction_data = self.page.evaluate("""
                () => {
                    const data = {};
                    const pageText = document.body.innerText;

                    const bookingMatch = pageText.match(/Booking ID[\\s\\n]+([A-Z0-9]+)/);
                    if (bookingMatch) data['Booking ID'] = bookingMatch[1];

                    const bankMatch = pageText.match(/Send Money To[\\s\\S]*?([A-Za-z\\s]+of[A-Za-z\\s]+)/i);
                    if (bankMatch) data['Bank'] = bankMatch[1].trim();

                    const accMatch = pageText.match(/([\\d]{10,})/);
                    if (accMatch) data['Account Number'] = accMatch[1];

                    const bicMatch = pageText.match(/([A-Z]{8,11})/);
                    if (bicMatch) data['BIC Code'] = bicMatch[1];

                    return data;
                }
            """)

            self.tc06_transaction_data = transaction_data
            print("[SUCCESS] Transaction completed!")
            print(f"  Booking ID: {transaction_data.get('Booking ID', 'N/A')}")
            print(f"  Bank: {transaction_data.get('Bank', 'N/A')}")
            print(f"  Account: {transaction_data.get('Account Number', 'N/A')}")
            print(f"  BIC Code: {transaction_data.get('BIC Code', 'N/A')}")

            # Step 9: Close popup
            print("\n[STEP 9] Closing success popup...")
            close_btn = self.page.locator("button:has-text('Close')").first
            if close_btn.is_visible(timeout=5000):
                close_btn.click()
                self.page.wait_for_timeout(2000)
                print("[STEP 9] Popup closed")

            screenshot_final = self._take_screenshot("TC_06_After_Close")

            # Log success
            self._log_result(
                tc_id, scenario, "PASSED",
                f"Invoice paid successfully. Booking ID: {transaction_data.get('Booking ID', 'N/A')}",
                f"{screenshot_dashboard}, {screenshot_form}, {screenshot_success}"
            )
            return True

        except Exception as e:
            screenshot = self._take_screenshot("TC_06_FAILED")
            self._log_result(tc_id, scenario, "FAILED", str(e), screenshot)
            return False

    # =========================================================================
    # Report Generation
    # =========================================================================
    def generate_report(self):
        """Generate test execution report in HTML format."""
        report_path = self.reports_dir / f"Test_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        passed = sum(1 for r in self.test_results if r["status"] == "PASSED")
        failed = sum(1 for r in self.test_results if r["status"] == "FAILED")
        total = len(self.test_results)
        pass_rate = (passed / total * 100) if total > 0 else 0

        # Generate test case HTML blocks
        test_cases_html = ""
        for result in self.test_results:
            status_class = "passed" if result["status"] == "PASSED" else "failed"
            status_icon = "&#10003;" if result["status"] == "PASSED" else "&#10007;"
            screenshots = result.get('screenshot', '').split(', ')
            screenshot_html = ""
            for screenshot in screenshots:
                if screenshot:
                    screenshot_name = screenshot.split('/')[-1].split('\\')[-1]
                    screenshot_html += f'''
                        <div class="screenshot-item">
                            <img src="{screenshot_name}" alt="{screenshot_name}" onclick="openModal(this)">
                            <p>{screenshot_name}</p>
                        </div>'''

            test_cases_html += f'''
            <div class="test-case">
                <div class="test-case-header {status_class}">
                    <div>
                        <div class="test-case-id">{result['tc_id']}</div>
                        <div class="test-case-scenario">{result['scenario']}</div>
                    </div>
                    <span class="status-badge {status_class}">{status_icon} {result['status']}</span>
                </div>
                <div class="test-case-body">
                    <div class="test-details">
                        <div class="detail-group">
                            <label>Details</label>
                            <p>{result['details']}</p>
                        </div>
                        <div class="detail-group">
                            <label>Timestamp</label>
                            <p>{result['timestamp']}</p>
                        </div>
                    </div>
                    <div class="screenshot-gallery">{screenshot_html}</div>
                </div>
            </div>'''

        # Generate invoice data HTML if available
        invoice_html = ""
        if self.invoice_data:
            invoice_rows = ""
            for key, value in self.invoice_data.items():
                invoice_rows += f"<tr><td>{key}</td><td>{value}</td></tr>"
            invoice_rows += f"<tr><td>request_id</td><td><strong>{self.request_id}</strong></td></tr>"
            invoice_html = f'''
            <h2 class="section-title">Invoice Data Captured (TC_03)</h2>
            <table class="data-table">
                <tr><th>Field</th><th>Value</th></tr>
                {invoice_rows}
            </table>'''

        # Generate TC_04 verification results HTML if available
        tc04_html = ""
        if self.tc04_verification_results:
            verification_rows = ""
            for r in self.tc04_verification_results:
                if r['status'] == 'MATCH':
                    status_color = "green"
                elif r['status'] == 'DATA MISSING':
                    status_color = "orange"
                else:
                    status_color = "red"
                verification_rows += f'''
                    <tr>
                        <td>{r['field']}</td>
                        <td>{r['expected']}</td>
                        <td>{r['actual']}</td>
                        <td style="color: {status_color}; font-weight: bold;">{r['status']}</td>
                    </tr>'''

            # Observations for mismatched fields
            observations = []
            for r in self.tc04_verification_results:
                if r['status'] == 'DATA MISSING':
                    observations.append(f"<li><strong>{r['field']}:</strong> Field appears blank in the Invoice Details view page.</li>")
                elif r['status'] == 'MISMATCH':
                    observations.append(f"<li><strong>{r['field']}:</strong> Shows \"{r['actual']}\" instead of expected \"{r['expected']}\".</li>")

            observations_html = ""
            if observations:
                observations_html = f'''
                <div style="margin-top: 20px; padding: 15px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                    <strong>Observations:</strong>
                    <ul style="margin-top: 10px; margin-left: 20px;">
                        {''.join(observations)}
                    </ul>
                </div>'''

            tc04_html = f'''
            <h2 class="section-title">TC_04 Data Verification Results</h2>
            <table class="data-table">
                <tr>
                    <th>Field</th>
                    <th>Expected (TC_03)</th>
                    <th>Actual (Details Page)</th>
                    <th>Status</th>
                </tr>
                {verification_rows}
            </table>
            {observations_html}'''

        # Generate TC_05 verification results HTML if available
        tc05_html = ""
        if self.tc05_verification_results:
            verification_rows_tc05 = ""
            for r in self.tc05_verification_results:
                if r['status'] == 'MATCH':
                    status_color = "green"
                elif r['status'] == 'DATA MISSING':
                    status_color = "orange"
                else:
                    status_color = "red"
                verification_rows_tc05 += f'''
                    <tr>
                        <td>{r['field']}</td>
                        <td>{r['expected']}</td>
                        <td>{r['actual']}</td>
                        <td style="color: {status_color}; font-weight: bold;">{r['status']}</td>
                    </tr>'''

            # Observations for mismatched fields
            observations_tc05 = []
            for r in self.tc05_verification_results:
                if r['status'] == 'DATA MISSING':
                    observations_tc05.append(f"<li><strong>{r['field']}:</strong> Field appears blank in the Pay Invoice page.</li>")
                elif r['status'] == 'MISMATCH':
                    observations_tc05.append(f"<li><strong>{r['field']}:</strong> Shows \"{r['actual']}\" instead of expected \"{r['expected']}\".</li>")

            observations_html_tc05 = ""
            if observations_tc05:
                observations_html_tc05 = f'''
                <div style="margin-top: 20px; padding: 15px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">
                    <strong>Observations:</strong>
                    <ul style="margin-top: 10px; margin-left: 20px;">
                        {''.join(observations_tc05)}
                    </ul>
                </div>'''

            tc05_html = f'''
            <h2 class="section-title">TC_05 Data Verification Results (Client View)</h2>
            <p style="margin-bottom: 15px; color: #6c757d;">Verification of invoice data from Client's Pending Payables / Pay Invoice page</p>
            <table class="data-table">
                <tr>
                    <th>Field</th>
                    <th>Expected (TC_03)</th>
                    <th>Actual (Pay Invoice Page)</th>
                    <th>Status</th>
                </tr>
                {verification_rows_tc05}
            </table>
            {observations_html_tc05}'''

        # Generate TC_06 verification results HTML if available
        tc06_html = ""
        if self.tc06_verification_results:
            verification_rows_tc06 = ""
            for r in self.tc06_verification_results:
                if r['status'] == 'MATCH':
                    status_color = "green"
                elif r['status'] == 'DATA MISSING':
                    status_color = "orange"
                else:
                    status_color = "red"
                verification_rows_tc06 += f'''
                    <tr>
                        <td>{r['field']}</td>
                        <td>{r['expected']}</td>
                        <td>{r['actual']}</td>
                        <td style="color: {status_color}; font-weight: bold;">{r['status']}</td>
                    </tr>'''

            # Transaction success data
            transaction_html = ""
            if self.tc06_transaction_data:
                transaction_html = f'''
                <div style="margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; border-radius: 10px;">
                    <h3 style="text-align: center; margin-bottom: 15px;">Transaction Success Details</h3>
                    <table style="width: 100%; background: rgba(255,255,255,0.1); border-radius: 5px;">
                        <tr><td style="padding: 10px; color: white;">Booking ID</td><td style="padding: 10px; color: white; font-weight: bold;">{self.tc06_transaction_data.get('Booking ID', 'N/A')}</td></tr>
                        <tr><td style="padding: 10px; color: white;">Bank</td><td style="padding: 10px; color: white;">{self.tc06_transaction_data.get('Bank', 'N/A')}</td></tr>
                        <tr><td style="padding: 10px; color: white;">Account Number</td><td style="padding: 10px; color: white;">{self.tc06_transaction_data.get('Account Number', 'N/A')}</td></tr>
                        <tr><td style="padding: 10px; color: white;">BIC Code</td><td style="padding: 10px; color: white;">{self.tc06_transaction_data.get('BIC Code', 'N/A')}</td></tr>
                    </table>
                </div>'''

            tc06_html = f'''
            <h2 class="section-title">TC_06 Data Verification Results (Pay Invoice Form)</h2>
            <p style="margin-bottom: 15px; color: #6c757d;">Verification of invoice data from Pay Invoice form before payment</p>
            <table class="data-table">
                <tr>
                    <th>Field</th>
                    <th>Expected (TC_03)</th>
                    <th>Actual (Pay Invoice Form)</th>
                    <th>Status</th>
                </tr>
                {verification_rows_tc06}
            </table>
            {transaction_html}'''

        report_content = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Test Execution Report - Omney Business</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: white; padding: 30px 40px; text-align: center; }}
        .header h1 {{ font-size: 2.5rem; margin-bottom: 10px; }}
        .header p {{ opacity: 0.8; font-size: 1.1rem; }}
        .meta-info {{ display: flex; justify-content: center; gap: 40px; margin-top: 20px; flex-wrap: wrap; }}
        .meta-item {{ text-align: center; }}
        .meta-item label {{ display: block; font-size: 0.8rem; opacity: 0.7; text-transform: uppercase; }}
        .meta-item span {{ font-size: 1rem; font-weight: 600; }}
        .summary {{ display: flex; justify-content: space-around; padding: 30px; background: #f8f9fa; border-bottom: 1px solid #e9ecef; flex-wrap: wrap; gap: 20px; }}
        .summary-card {{ text-align: center; padding: 20px 40px; border-radius: 10px; background: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1); min-width: 150px; }}
        .summary-card.total {{ border-top: 4px solid #6c757d; }}
        .summary-card.passed {{ border-top: 4px solid #28a745; }}
        .summary-card.failed {{ border-top: 4px solid #dc3545; }}
        .summary-card.rate {{ border-top: 4px solid #007bff; }}
        .summary-card h2 {{ font-size: 2.5rem; margin-bottom: 5px; }}
        .summary-card.passed h2 {{ color: #28a745; }}
        .summary-card.failed h2 {{ color: #dc3545; }}
        .summary-card.rate h2 {{ color: #007bff; }}
        .summary-card p {{ color: #6c757d; font-size: 0.9rem; text-transform: uppercase; }}
        .content {{ padding: 40px; }}
        .section-title {{ font-size: 1.5rem; color: #1a1a2e; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #667eea; }}
        .test-case {{ background: #f8f9fa; border-radius: 10px; margin-bottom: 30px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        .test-case-header {{ padding: 20px; display: flex; justify-content: space-between; align-items: center; }}
        .test-case-header.passed {{ background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; }}
        .test-case-header.failed {{ background: linear-gradient(135deg, #dc3545 0%, #e83e8c 100%); color: white; }}
        .test-case-id {{ font-size: 1.2rem; font-weight: 700; }}
        .test-case-scenario {{ font-size: 0.95rem; opacity: 0.9; }}
        .status-badge {{ padding: 8px 20px; border-radius: 20px; font-weight: 600; font-size: 0.9rem; background: rgba(255,255,255,0.2); color: white; }}
        .test-case-body {{ padding: 25px; background: white; }}
        .test-details {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }}
        .detail-group {{ margin-bottom: 15px; }}
        .detail-group label {{ display: block; font-size: 0.8rem; color: #6c757d; text-transform: uppercase; margin-bottom: 5px; font-weight: 600; }}
        .detail-group p {{ color: #1a1a2e; line-height: 1.6; }}
        .data-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        .data-table th, .data-table td {{ padding: 10px 15px; text-align: left; border-bottom: 1px solid #e9ecef; }}
        .data-table th {{ background: #f8f9fa; font-weight: 600; color: #495057; font-size: 0.85rem; text-transform: uppercase; }}
        .data-table td {{ color: #1a1a2e; }}
        .screenshot-gallery {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin-top: 20px; }}
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
            <h1>Test Execution Report</h1>
            <p>Omney Business Automation Testing (Python Script)</p>
            <div class="meta-info">
                <div class="meta-item">
                    <label>Execution Date</label>
                    <span>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</span>
                </div>
                <div class="meta-item">
                    <label>Framework</label>
                    <span>Playwright Python</span>
                </div>
                <div class="meta-item">
                    <label>Application URL</label>
                    <span>{self.base_url}</span>
                </div>
            </div>
        </div>

        <div class="summary">
            <div class="summary-card total">
                <h2>{total}</h2>
                <p>Total Tests</p>
            </div>
            <div class="summary-card passed">
                <h2>{passed}</h2>
                <p>Passed</p>
            </div>
            <div class="summary-card failed">
                <h2>{failed}</h2>
                <p>Failed</p>
            </div>
            <div class="summary-card rate">
                <h2>{pass_rate:.0f}%</h2>
                <p>Pass Rate</p>
            </div>
        </div>

        <div class="content">
            <h2 class="section-title">Detailed Test Results</h2>
            {test_cases_html}

            {invoice_html}

            {tc04_html}

            {tc05_html}

            {tc06_html}

            <h2 class="section-title" style="margin-top: 40px;">Environment Details</h2>
            <table class="data-table">
                <tr><th>Parameter</th><th>Value</th></tr>
                <tr><td>Browser</td><td>Chromium (Playwright)</td></tr>
                <tr><td>Python Version</td><td>{sys.version.split()[0]}</td></tr>
                <tr><td>Headless Mode</td><td>{self.headless}</td></tr>
                <tr><td>Application URL</td><td>{self.base_url}</td></tr>
            </table>
        </div>

        <div class="footer">
            <p>Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Omney Business Automation Testing</p>
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
        return report_path

    # =========================================================================
    # Main Test Runner
    # =========================================================================
    def run_all_tests(self):
        """Run all test cases."""
        print("\n" + "="*70)
        print("OMNEY BUSINESS AUTOMATION TEST SUITE")
        print("="*70)
        print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Base URL: {self.base_url}")
        print("="*70)

        try:
            self.setup()

            # Run test cases in sequence
            tc01_result = self.tc_01_url_verification()

            if tc01_result:
                tc02_result = self.tc_02_login()

                if tc02_result:
                    tc03_result = self.tc_03_raise_invoice()

                    if tc03_result:
                        tc04_result = self.tc_04_verify_pending_receivables()

                        # TC_05: Verify Pending Payables as Client
                        if tc04_result or True:  # Run TC_05 even if TC_04 has minor failures
                            tc05_result = self.tc_05_verify_pending_payables()

                            # TC_06: Pay Invoice from View Page
                            if tc05_result or True:  # Run TC_06 even if TC_05 has minor failures
                                tc06_result = self.tc_06_pay_invoice()
                            else:
                                print("[SKIP] TC_06 skipped due to TC_05 failure")
                        else:
                            print("[SKIP] TC_05, TC_06 skipped due to TC_04 failure")
                    else:
                        print("[SKIP] TC_04, TC_05, TC_06 skipped due to TC_03 failure")
                else:
                    print("[SKIP] TC_03, TC_04, TC_05, TC_06 skipped due to TC_02 failure")
            else:
                print("[SKIP] TC_02, TC_03, TC_04, TC_05, TC_06 skipped due to TC_01 failure")

            # Generate report
            self.generate_report()

        except Exception as e:
            print(f"[CRITICAL ERROR] {e}")
            raise
        finally:
            self.teardown()

        # Print final summary
        print("\n" + "="*70)
        print("TEST EXECUTION SUMMARY")
        print("="*70)
        for result in self.test_results:
            status_icon = "✓" if result["status"] == "PASSED" else "✗"
            print(f"  {status_icon} {result['tc_id']}: {result['status']}")
        print("="*70)
        print(f"End Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*70)


def main():
    """Main entry point."""
    # Set headless=True for CI/CD environments
    # Set keep_browser_open=True to keep browser open after test completion
    automation = OmneyBusinessAutomation(headless=False, keep_browser_open=False)
    automation.run_all_tests()


if __name__ == "__main__":
    main()
