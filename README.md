# URL Checker

Opens a list of websites in Google Chrome and records what a visitor sees, e.g. `200 - OK` or `522 - Connection timed out`.

## Files

- `url_checker.py` - the script
- `Run URL Checker.bat` - starts the script

## Setup (once)

1. Install Python from [python.org](https://www.python.org/downloads/) and tick **"Add python.exe to PATH"** during install.
2. Make sure Google Chrome is installed.

## How to use

1. Put your URLs in column A of an Excel file (a header in row 1 is fine).
2. In File Explorer, click and hold your Excel file, drop it onto the `Run URL Checker.bat` icon and let go. A black window opens and the check starts.
   - Or double-click `Run URL Checker.bat`, drag your Excel file into the window that opens, and press Enter.
3. Wait while it runs. Each URL appears on screen as it is checked.
4. Open `<your file name> - results.xlsx`, saved next to your Excel file. Column B has the result for each URL.

You can also run it from a terminal: `python url_checker.py "your file.xlsx"`

If you see `No module named ...`, run this once: `python -m pip install --user openpyxl playwright dnspython`
