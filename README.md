# IT23153486 - Test Automation Assignment 1

Automated test suite for the **Singlish to Sinhala Chat Translator** using Playwright and Python.

**GitHub Repository:** https://github.com/Bhagya-Navodyani/ITPM-ASSIGMENT-1-IT23153486-

---

## 📋 Overview

This project automates testing of the chat translator application by:
- Reading 50 test cases from an Excel file
- Sending Singlish inputs to the translator
- Capturing actual outputs
- Comparing with expected outputs
- Logging Pass/Fail results

**Test Coverage:** 50 test cases across multiple input categories (Questions, Commands, Greetings, Requests, Responses)

---

## 🚀 Quick Start

### Prerequisites
- Python 3.11+ (installed)
- Playwright (auto-installed)
- openpyxl (auto-installed)

### Installation

```bash
# Clone the repository
git clone https://github.com/Bhagya-Navodyani/ITPM-ASSIGMENT-1-IT23153486-.git
cd ITPM-ASSIGMENT-1-IT23153486-

# Navigate to the workspace
cd d:\IT23153486_Test_automation

# (Optional) Create and activate virtual environment
python -m venv venv
venv\Scripts\activate
```

---

## ▶️ Running the Tests

### Standard Test Run (Recommended)

```powershell
cd d:\IT23153486_Test_automation

python test_automation/test_automation/test_automation.py `
  --excel "Assignment 1 - Test cases.xlsx" `
  --input-col "Input" `
  --expected-col "Expected output" `
  --actual-col "Actual output" `
  --status-col "Status" `
  --url "https://www.pixelssuite.com/chat-translator" `
  --wait-ms 5000 `
  --type-delay-ms 80 `
  --slow-mo-ms 200 `
  --save-every 10
```

### Parameters Explained

| Parameter | Value | Purpose |
|-----------|-------|---------|
| `--excel` | `"Assignment 1 - Test cases.xlsx"` | Excel file with test cases (200 rows) |
| `--input-col` | `"Input"` | Column containing Singlish input text |
| `--expected-col` | `"Expected output"` | Column with expected Sinhala output |
| `--actual-col` | `"Actual output"` | Column where actual outputs are saved |
| `--status-col` | `"Status"` | Column for PASS/FAIL status |
| `--url` | `https://www.pixelssuite.com/chat-translator` | Target application URL |
| `--wait-ms` | `5000` | Wait time (ms) after each input for output |
| `--type-delay-ms` | `80` | Delay (ms) between typing characters |
| `--slow-mo-ms` | `200` | Browser action slowdown (ms) for visibility |
| `--save-every` | `10` | Save results to Excel every N tests |

---

## 📁 Project Structure

```
IT23153486_Test_automation/
├── Assignment 1 - Test cases.xlsx    # Test data (200 rows)
├── Commands.txt                       # Quick reference commands
├── README.md                          # This file
├── test.py                           # Simple Playwright example
└── test_automation/
    └── test_automation/
        └── test_automation.py        # Main test runner script
```

---

## 📊 Excel File Structure

**Sheet Name:** "Test Cases"

| Column | Name | Purpose |
|--------|------|---------|
| A | TC ID | Test case ID (0001-0022) |
| B | Input length type | Category (S/M for Single/Multiple) |
| C | Input | Singlish text to send |
| D | Expected output | Expected Sinhala translation |
| E | Actual output | Captured output (auto-filled by test) |
| F | Status | Test result: PASS/FAIL/COLLECTED (auto-filled) |
| G | Singlish input types covered | Test category classification |
| H | Evidence or rationale | Notes about test case |

---

## ✅ Test Results

After running the tests, the Excel file is updated with:
- **Actual output:** Text captured from translator
- **Status:** 
  - ✅ `PASS` - Output matches expected
  - ❌ `FAIL` - Output doesn't match expected
  - 📝 `COLLECTED` - Output captured (no expected value to compare)
  - ⚠️ `UI Error` - Error during test execution

---

## 🔧 Troubleshooting

### Issue: Browser closes prematurely
**Solution:** Use `--save-every 10` (don't set too low, reduces Excel lock conflicts)

### Issue: "Output did not update" errors
**Possible causes:**
- Translator service is slow (increase `--wait-ms` to 7000-10000)
- Network issues
- Browser locale settings

**Solution:**
```powershell
# Try with longer wait time
--wait-ms 10000
```

### Issue: Permission denied saving Excel
**Cause:** Excel file is locked by another process (Windows Explorer, Office, etc.)

**Solution:**
1. Close the Excel file if open in Excel
2. Close any Explorer windows showing the file
3. Run test again

### Issue: Sheet not found
**Cause:** Sheet name doesn't match "Test Cases"

**Solution:**
```powershell
# Check available sheets
python -c "from openpyxl import load_workbook; wb = load_workbook('Assignment 1 - Test cases.xlsx'); print(wb.sheetnames)"

# Use correct sheet with --sheet parameter
--sheet "Test Cases"
```

---

## 📝 Test Categories

The 200 test cases cover:

1. **Question Forms** (Rows 2-6)
   - "Bhagya koheda giya?" → "භාග්‍ය කොහෙද ගිය?"
   - "Shashi mama oda ada?" → "ශාසි mama ओд ada?"

2. **Command Forms** (Rows 7-11)
   - "Leena pansalata yanawanam slippers dagena yanna epa"
   - "mal paaththi watala"

3. **Greetings** (Rows 12-16)
   - "Bhagya, subha udesanak!"
   - "Shashi, kohomada?"

4. **Requests** (Rows 17-21)
   - "oyage car eka titac denna puluwanda?"
   - Multiple variants

5. **Responses** (Rows 22+)
   - "danne nadda?"
   - "Bhagya aawada?"

---

## 📊 Example Output

```
Starting Frontend-Only test with 200 rows...
Frontend loaded successfully.
Testing [Row 2]: Bhagya koheda giya?
  -> FAIL
Testing [Row 3]: Shashi mama oda ada?
  -> PASS
Testing [Row 4]: mata nidmathai
  -> COLLECTED
...
Test completed. Results saved to d:\IT23153486_Test_automation\Assignment 1 - Test cases.xlsx
```

---

## 🔍 Key Features

✅ **Automated Input Typing** - Sends Singlish text character by character  
✅ **Output Capture** - Reads actual translator output (textarea or text content)  
✅ **Pass/Fail Validation** - Compares actual vs expected outputs  
✅ **Batch Saving** - Auto-saves results every N tests to prevent data loss  
✅ **Retry Logic** - Retries output capture if not immediately available  
✅ **Merged Cell Support** - Handles Excel merged cells correctly  
✅ **Unicode Support** - Full Sinhala script support in results  

---

## 🛠️ Advanced Options

### Run without opening browser UI
```powershell
--headless
```

### Save after every single test
```powershell
--save-every 1
```

### Increase output wait time
```powershell
--wait-ms 10000 --retry-wait-ms 2000 --retries 10
```

### Specify custom output file
```powershell
--output "results_backup.xlsx"
```

---

## 📧 Support

For issues or questions:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Review the [Test Results](#-test-results) section
3. Verify Excel file structure matches documentation

---

## 📄 License

Assignment submission for ITPM Course - IT23153486

---

## 👨‍💻 Author

**Bhagya Navodyani**  
Student ID: IT23153486  
Repository: https://github.com/Bhagya-Navodyani/ITPM-ASSIGMENT-1-IT23153486-

---

## 📅 Last Updated

May 4, 2026

**Status:** ✅ Ready for testing
