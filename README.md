This repository contains a Python script that automates the calculation of annualised one-month rolling returns for various mutual fund schemes across categories and AMCs using data from the public API https://api.mfapi.in/.

Features
- Fetches latest NAV data for each scheme.
- Calculates rolling returns using:
((NAV_t1 - NAV_t0) / NAV_t0) * (365 / days)

where t1 = last day of previous month, t0 = last day of two months ago.
- Exports category-wise Excel reports with formatted tables.
- Can be converted into an `.exe` file that runs on systems without Python.
- Creates log files and output reports automatically.

Repository Structure
rolling_returns.py # Main executable script
scheme_codes.json # AMC and scheme code mapping
outputs/ # Generated Excel files (auto-created)
app_log.txt # Log file (auto-created when run)


Requirements
- Python 3.9+
- Packages: `pandas`, `numpy`, `requests`, `xlsxwriter`

Usage
1. Place `scheme_codes.json` in the same folder as the script.
2. Run:
python rolling_returns.py

3. The generated Excel file will appear inside `/outputs`.

License
This project is licensed under the MIT License.  

If you'd like to run it without python, I have the EXE file here. Make sure the scheme code file and the exe are in the same folder.
https://drive.google.com/drive/folders/1rHs5JvrQD62QqX4Vd0TUycxwkDpQQtRc?usp=drive_link 
