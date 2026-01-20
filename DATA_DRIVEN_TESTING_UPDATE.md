# Data-Driven Testing Framework Update

**Date**: 2026-01-20
**Status**: ✅ COMPLETED

---

## Overview

The Omney Business automation framework has been successfully updated to support data-driven testing with multi-sheet Excel structure. This enables repeating test flows (TC_03-08) with different credential and invoice combinations.

---

## Changes Made

### 1. New Helper Methods (Lines 194-295)

#### `_parse_credential_type(test_data_value: str) -> str`
- Parses credential type from Test Data column
- Supports two formats:
  - Format 1: `Credentials: Vendor_Individual`
  - Format 2: `Credentials TC_03, TC_04: Vendor_Individual`
- Returns credential type (e.g., `Vendor_Individual`)

#### `_parse_invoice_reference(test_data_value: str) -> str`
- Parses invoice reference from Test Data column
- Format: `Invoice: Vendor_Individual + Client_Business`
- Returns invoice reference string

#### `_get_invoice_data(invoice_reference: str) -> dict`
- Retrieves invoice data from Invoice sheet by Vendor Type
- Returns dictionary with invoice fields:
  - Select Client
  - Purpose
  - Currency
  - Amount
  - Your Receiving Account
  - Invoice Document

#### `_get_invoice_data_for_tc(tc_id: str) -> dict`
- Convenience method to get invoice data for a specific test case
- Combines Test Data parsing with invoice lookup
- Used by TC_03 and TC_09

### 2. Updated Test Methods

#### `tc_03_raise_invoice(context_tc_id: str = "TC_03") -> bool`
**Changes:**
- Added `context_tc_id` parameter (default: "TC_03")
- Now uses `_get_invoice_data_for_tc(context_tc_id)` instead of hardcoded `iloc[0]`
- Supports TC_09 by accepting different context

**Usage:**
```python
# Original TC_03
automation.tc_03_raise_invoice()  # Uses TC_03 data

# TC_09 reusing TC_03 logic
automation.tc_03_raise_invoice(context_tc_id="TC_09")  # Uses TC_09 data
```

#### `tc_05_verify_pending_payables(context_tc_id: str = "TC_05") -> bool`
**Changes:**
- Added `context_tc_id` parameter (default: "TC_05")
- Now uses `_get_credentials_for_tc(context_tc_id)` for dynamic credential lookup
- Supports TC_09 with Client_Individual credentials

**Usage:**
```python
# Original TC_05
automation.tc_05_verify_pending_payables()  # Uses TC_05 data (Client_Business)

# TC_09 reusing TC_05 logic
automation.tc_05_verify_pending_payables(context_tc_id="TC_09")  # Uses Client_Individual
```

### 3. New Test Case: TC_09 (Lines 3440-3652)

#### `tc_09_raise_and_pay_invoice_individual() -> bool`

**Purpose:**
Repeats the complete TC_03-08 flow with different credentials and invoice data.

**Test Flow:**
1. **TC_01-02**: URL verification and login as Vendor_Individual
2. **TC_03**: Create invoice using TC_09 invoice data (Client: Yash Kumaar, Bandhan Bank)
3. **TC_04**: Verify invoice in Pending Receivables (as Vendor_Individual)
4. **TC_05**: Logout → Login as Client_Individual → Verify in Pending Payables
5. **TC_06**: Pay invoice from View Page
6. **TC_07**: Create new invoice → Pay from Homepage
7. **TC_08**: Create new invoice → Pay from Pay Invoice Page

**Key Differences from TC_03-08:**
| Aspect | TC_03-08 | TC_09 |
|--------|----------|-------|
| Vendor Credentials | Vendor_Individual | Vendor_Individual (same) |
| Client Credentials | Client_Business | Client_Individual |
| Client Name | Haier Electionics | Yash Kumaar |
| Bank Account | IDFC Bank | Bandhan Bank |
| Invoice Reference | Vendor_Individual + Client_Business | Vendor_Individual + Client_Individual |

### 4. Command Line Options

#### Updated `main()` function (Lines 4387-4435)

**Option 1: `--tc03-only`**
```bash
python Scripts/omney_business_automation.py --tc03-only
```
- Runs TC_01, TC_02, TC_03 only
- Creates invoice and saves data to `tc07_invoice_data.json`
- Useful for preparing test data

**Option 2: `--tc09`** (NEW)
```bash
python Scripts/omney_business_automation.py --tc09
```
- Runs only TC_09 (complete flow with Individual client)
- Executes ~10 test operations (3 invoice creations + 7 payment verifications)
- Generates HTML report

**Option 3: Default (no arguments)**
```bash
python Scripts/omney_business_automation.py
```
- Runs all test cases (TC_01 through TC_08)
- Does NOT include TC_09 by default (run separately with --tc09)

---

## Excel Structure

### Sheet 1: Testcase
Contains test case definitions with Test Data column.

**TC_03 Test Data:**
```
Credentials: Vendor_Individual,
Invoice sheet: Vendor_Individual + Client_Business
```

**TC_09 Test Data:**
```
Credentials TC_03, TC_04: Vendor_Individual
Credentials TC_05, TC_06, TC_07, TC_08: Client_Individual
Invoice: Vendor_Individual + Client_Individual
```

### Sheet 2: Credentials
| Type | Email | Password | Name |
|------|-------|----------|------|
| Vendor_Individual | visheshindindia@yopmail.com | Password@2 | Suraj |
| Client_Business | ganeshthakurpm@yopmail.com | Password@2 | Haier Electionics |
| Vendor_Business | qwerty1@yopmail.com | Password@2 | Ahmed |
| Client_Individual | yashkumaarob@yopmail.com | Password@2 | (Individual) |

### Sheet 3: Invoice
| Sr No | Vendor Type | Invoice Number | Select Client | Purpose | Currency | Amount | Your Receiving Account | Invoice Document |
|-------|-------------|----------------|---------------|---------|----------|---------|----------------------|------------------|
| 1 | Vendor_Individual + Client_Business | {Random} | Haier Electionics | Demo Purpose | USD | {Random} | IDFC Bank | D:\Vishesh\FAB\Test.png |
| 2 | Vendor_Individual + Client_Individual | {Random} | Yash Kumaar | Demo Purpose | USD | {Random} | Bandhan Bank | D:\Vishesh\FAB\Test.png |

**Placeholders:**
- `{Random}`: Automatically generated during test execution
- `{Current Date}`: Uses current date
- `{Current Date +2}`: Uses current date + 2 days

---

## Data Flow Architecture

### TC_03-08 Flow (Original)
```
┌─────────────────────────────────────────────────────┐
│ TC_03: Read Test Data → Parse "Vendor_Individual +  │
│        Client_Business" → Look up in Invoice sheet  │
│        → Use row 1 data (Haier Electionics, IDFC)   │
└─────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────┐
│ TC_05: Read Test Data → Parse "Client_Business"     │
│        → Look up in Credentials sheet → Login       │
└─────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────┐
│ TC_06-08: Continue with same Client_Business user   │
└─────────────────────────────────────────────────────┘
```

### TC_09 Flow (New)
```
┌─────────────────────────────────────────────────────┐
│ TC_09 Step 3: Read TC_09 Test Data → Parse          │
│               "Vendor_Individual + Client_Individual"│
│               → Look up in Invoice sheet → Use row 2 │
│               data (Yash Kumaar, Bandhan Bank)       │
└─────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────┐
│ TC_09 Step 5: Read TC_09 Test Data → Parse          │
│               "Client_Individual" → Look up in       │
│               Credentials sheet → Login              │
└─────────────────────────────────────────────────────┘
                          ↓
┌─────────────────────────────────────────────────────┐
│ TC_09 Steps 6-8: Continue with Client_Individual    │
└─────────────────────────────────────────────────────┘
```

---

## Benefits

### 1. **Reusability**
- Same test logic (TC_03-08) can be executed with different data
- No code duplication
- Easy to add TC_10, TC_11, etc. with new data combinations

### 2. **Maintainability**
- Test data centralized in Excel sheets
- Easy to update credentials and invoice data without code changes
- Clear separation of test logic and test data

### 3. **Scalability**
- Add new credential types: `Vendor_Business + Client_Business`
- Add new invoice combinations: `Vendor_Business + Client_Individual`
- Framework automatically adapts to new data

### 4. **Coverage**
- Tests both Business and Individual client types
- Verifies different bank accounts and client configurations
- Comprehensive validation of invoice workflows

---

## Usage Examples

### Example 1: Run TC_09 Only
```bash
cd D:\Vishesh\OmneyBusiness
python Scripts/omney_business_automation.py --tc09
```

**Expected Output:**
```
======================================================================
RUNNING TC_09 ONLY
Complete invoice creation and payment flow with Individual client
======================================================================

[TC_09] Step 1-2: Executing TC_01 and TC_02...
[TC_09] Logging in as Vendor_Individual...
[DATA] TC_09 Test Data: Credentials TC_03, TC_04: Vendor_Individual...
[CREDENTIALS] Using Vendor_Individual: visheshindindia@yopmail.com
[TC_09] Login successful - Vendor_Individual logged in

[TC_09] Step 3: Executing TC_03 with TC_09 invoice data...
[EXECUTING] TC_03: To check if user can navigate to Raise Invoice...
[CONTEXT] Using data from TC_09
[INVOICE DATA] Using invoice: Vendor_Individual + Client_Individual
[INVOICE DATA] Client: Yash Kumaar, Currency: USD, Bank: Bandhan Bank
...
```

### Example 2: Create Test Data for TC_09
```bash
# Create invoice for manual testing of TC_09 steps
python Scripts/omney_business_automation.py --tc03-only

# Edit tc07_invoice_data.json to use TC_09 data
# Then run standalone TC_08 with TC_09 invoice
```

### Example 3: Add New Test Case (TC_10)
1. **Add to Credentials sheet:**
   - Type: `Vendor_Business`
   - Email: `qwerty1@yopmail.com`

2. **Add to Invoice sheet:**
   - Sr No: 3
   - Vendor Type: `Vendor_Business + Client_Business`
   - Select Client: `Haier Electionics`

3. **Add to Testcase sheet:**
   - TC_ID: `TC_10`
   - Test Data: `Credentials TC_03, TC_04: Vendor_Business\nCredentials TC_05-08: Client_Business\nInvoice: Vendor_Business + Client_Business`

4. **Create TC_10 method:**
```python
def tc_10_raise_and_pay_invoice_business(self) -> bool:
    # Similar to TC_09, just calls existing methods with context_tc_id="TC_10"
    tc03_result = self.tc_03_raise_invoice(context_tc_id="TC_10")
    tc05_result = self.tc_05_verify_pending_payables(context_tc_id="TC_10")
    # ...
```

---

## Testing Checklist

- [x] TC_03 accepts context_tc_id parameter
- [x] TC_05 accepts context_tc_id parameter
- [x] TC_09 method created and integrated
- [x] _parse_credential_type handles both formats
- [x] _parse_invoice_reference extracts invoice reference
- [x] _get_invoice_data looks up by Vendor Type
- [x] _get_invoice_data_for_tc combines parsing and lookup
- [x] --tc09 command line option added
- [x] Excel structure supports multi-line Test Data

**Ready to test:** ✅

---

## Files Modified

1. **Scripts/omney_business_automation.py**
   - Lines 170-199: Updated `_parse_credential_type`
   - Lines 201-295: Added new helper methods
   - Lines 657-704: Updated `tc_03_raise_invoice`
   - Lines 1483-1537: Updated `tc_05_verify_pending_payables`
   - Lines 3440-3652: Added `tc_09_raise_and_pay_invoice_individual`
   - Lines 4387-4435: Updated `main()` with --tc09 option

2. **Testcase/OB_Automation.xlsx**
   - Sheet 1 (Testcase): TC_09 added with Test Data
   - Sheet 2 (Credentials): All 4 credential types defined
   - Sheet 3 (Invoice): 2 invoice combinations defined

---

## Next Steps

### Immediate Testing
```bash
# Test TC_09 execution
python Scripts/omney_business_automation.py --tc09
```

### Future Enhancements
1. Add TC_10, TC_11 with more data combinations
2. Update TestAutomationAgent to support TC_09 generation
3. Create Excel validation script to check data integrity
4. Add data-driven reporting (show which data combination was used)
5. Implement parameterized test execution (run TC_03-08 with multiple datasets)

---

**Framework Version**: 2.0 (Data-Driven Testing)
**Last Updated**: 2026-01-20
**Author**: Claude Code Assistant
