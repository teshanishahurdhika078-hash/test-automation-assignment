# PixelSuite Chat Translator - Test Automation

This repository contains the Playwright automated test script for the PixelSuite Chat Translator. It reads test cases from an Excel sheet, automatically runs them against the frontend using Playwright Chromium, and updates the Excel sheet with the results.

## Prerequisites
Before running the script, ensure you have the following installed:
1. **Python 3.11 or newer** (Make sure to check "Add Python to PATH" during installation)
2. **Git** (if you are cloning this repository)

## Installation

1. Open a Command Prompt or Terminal in this folder.
2. Install the required Python libraries (`playwright` and `openpyxl`):
   ```bash
   pip install playwright openpyxl
   ```
   *(If `pip` is not recognized, use `py -m pip install playwright openpyxl` instead)*

3. Install the Playwright browsers:
   ```bash
   playwright install
   ```
   *(Or `py -m playwright install` if `playwright` is not recognized)*

## Running the Tests

To run the automated tests, simply execute the python script in your terminal:

```bash
python test_automation.py
```
*(Or `py test_automation.py`)*

### What the script does:
1. Reads the `Assignment 1 - Test cases.xlsx` file.
2. Opens Chromium and navigates to the Chat Translator URL.
3. Automatically inputs the "Singlish" test cases into the frontend.
4. Compares the generated translation with the "Expected output".
5. Updates the Excel sheet's "Actual output" and "Status" (PASS/FAIL) columns.

## Notes
- To see the browser typing live (not headless), simply run the script normally.
- You can leave the browser open after tests finish by appending the `--keep-open` flag.
