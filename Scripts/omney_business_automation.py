"""
Omney Business Automation Script
================================
Automates test cases TC_01, TC_02, TC_03 for Omney Business application.

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

    def __init__(self, headless: bool = False):
        """
        Initialize the automation framework.

        Args:
            headless: Run browser in headless mode (default: False)
        """
        self.base_url = "https://qaoneob.remit.in"
        self.headless = headless
        self.browser = None
        self.page = None
        self.context = None
        self.playwright = None

        # Test results storage
        self.test_results = []
        self.request_id = None
        self.invoice_data = {}

        # Setup directories
        self.base_dir = Path(__file__).parent.parent
        self.reports_dir = self.base_dir / "Reports" / "Python_Automation"
        self.testcase_file = self.base_dir / "Testcase" / "OB_Automation.xlsx"

        # Create reports directory
        self.reports_dir.mkdir(parents=True, exist_ok=True)

        # Test data
        self.test_data = None
        self.invoice_sheet = None

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
            print(f"[DATA] Loaded test data from {self.testcase_file}")
        except Exception as e:
            print(f"[ERROR] Failed to load test data: {e}")
            raise

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

        # Test data from Excel
        username = "visheshindindia@yopmail.com"
        password = "Password@2"

        try:
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

            # Step 3: Log all fetched data
            print("\n[STEP 3] Invoice Data Logged:")
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
                f"Invoice created successfully. Request ID: {self.request_id}",
                f"{screenshot_form}, {screenshot_success}, {screenshot_dashboard}"
            )
            return True

        except Exception as e:
            screenshot = self._take_screenshot("TC_03_FAILED")
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
            <h2 class="section-title">Invoice Data Captured</h2>
            <table class="data-table">
                <tr><th>Field</th><th>Value</th></tr>
                {invoice_rows}
            </table>'''

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
                else:
                    print("[SKIP] TC_03 skipped due to TC_02 failure")
            else:
                print("[SKIP] TC_02, TC_03 skipped due to TC_01 failure")

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
    automation = OmneyBusinessAutomation(headless=False)
    automation.run_all_tests()


if __name__ == "__main__":
    main()
