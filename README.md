# OSBA Policy Tool

The OSBA Policy Tool is a Python application that processes Excel files to generate policy recommendations based on specified criteria. It is designed for users interested in analyzing policy data efficiently.

## Prerequisites

- Python 3.12.3 or above
- Libraries: tkinter, pandas, BeautifulSoup (bs4), os

## Features

- **Data Input:** Enables selection of input Excel files containing policy data.
- **Chart Integration:** Allows users to include an OPSS chart Excel file for detailed analysis.
- **Data Processing:** Adds recommendations to the input data based on criteria defined within the OPSS chart.
- **Results Output:** Outputs the processed data with recommendations into a new Excel file.

## Installation

1. **Install Python:** Ensure Python 3.12.3 or newer is installed. [Download Python](https://www.python.org/downloads/)

2. **Install Libraries:** Use pip to install the required libraries:
```bash
pip3 install -r requirements.txt
```
3. Get the Code: Clone the repository or download the source code:
```bash
git clone https://github.com/Yobo123o/policyTool.git
cd policyTool
```
## Usage
1. Run the Program: Execute the script to start the application:
```bash
python process_spreadsheet.py
```
2. File Selection: Use the "Browse" buttons to select both the input Excel file and the OPSS chart file. 

3. Process Data: Click "Process Data" and follow on-screen instructions to specify the output file details.

4. View Results: The new Excel file with recommendations will be saved to your specified location. Open it to view the results.

## Building the App
These instructions will guide you through building a standalone executable for different architectures:

ARM Build:
```bash
pyinstaller process_spreadsheet.py --name PolicyTool_ARM -F --argv-emulation
```
Intel Build:
```bash
pyinstaller process_spreadsheet.py --name PolicyTool_INTEL -F --argv-emulation
```
Windows Build:
```bash
pyinstaller process_spreadsheet.py --name PolicyTool_WIN -F --argv-emulation
```
## Support
For assistance, please open an issue in the GitHub repository or contact Brendan Swartz at bswartz@ohioschoolboards.org.

This README is organized to provide clarity and ease of use, improving the onboarding process for new users and enhancing support accessibility. If you need any further modifications or additional sections, feel free to ask!
