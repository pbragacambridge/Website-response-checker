# Website-response-checker

Checks a list of websites and records what each one shows, e.g. `200 - OK` or `522 - Connection timed out`.

## Files

- `url_checker.py` - the script
- `Run URL Checker.bat` - starts the script

## Setup (once)

Install Python from [python.org](https://www.python.org/downloads/) and tick **"Add python.exe to PATH"** during install.

## How to use

1. Put your URLs in column A of an Excel file (a header in row 1 is fine).
2. Drag the Excel file onto `Run URL Checker.bat`.
3. Wait while it runs. Each URL appears on screen as it is checked.
4. Open `<your file name> - results.xlsx`, saved next to your Excel file. Column B has the result for each URL.

You can also run it from a terminal: `python url_checker.py "your file.xlsx"`

If you see `No module named requests`, run this once: `python -m pip install --user openpyxl requests`
