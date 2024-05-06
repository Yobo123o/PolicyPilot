# OSBA Policy Tool

The OSBA Policy Tool is a Python application that processes Excel files to generate policy recommendations based on specified criteria. It requires Python 3.x and the following libraries: tkinter, pandas, bs4 (BeautifulSoup), and os.

## Features

- Allows users to select an input Excel file containing policy data.
- Allows users to select an OPSS chart Excel file.
- Processes the input data and adds recommendations based on criteria defined in the OPSS chart.
- Outputs the processed data with recommendations to a new Excel file.

## How to Use

1. **Install Python:** Make sure you have Python 3.x installed on your system. If not, you can download it from [python.org](https://www.python.org/downloads/).

2. **Install Required Libraries:** Open a terminal or command prompt and install the required libraries using pip: `pip install pandas beautifulsoup4`

3. **Download the Repository:** Download or clone the repository to your local machine.

4. **Run the Program:** Navigate to the directory containing the script files and run the main script: `process_spreadsheet.py`

5. **Select Input Files:**
- Click on the "Browse" button next to "Select the Input Excel File" to choose your input Excel file containing policy data.
- Click on the "Browse" button next to "Select the OPSS Chart File" to choose the OPSS chart Excel file.

6. **Process Data:**
- Once both files are selected, click on the "Process Data" button to start processing.
- Follow any on-screen prompts to enter the output file name.

7. **View Output:**
- After processing, the program will generate a new Excel file with the processed data and recommendations.
- The file will be saved in your Documents folder by default.

8. **Explore Recommendations:**
- Open the generated Excel file to view the processed data and recommendations.

## Building the App
For ARM Build run `pyinstaller process_spreadsheet.py -N PolicyTool_ARM -F --argv-emulation ` in the Terminal to build the app.

For Intel Build run `python3 -m pyinstaller process_spreadsheet.py -N PolicyTool_INTEL -F --argv-emulation`
* Arguments:
  * `--argv-emulation`: Enable argv emulation for macOS app bundles. If enabled, the initial open document/URL event is processed by the bootloader and the passed file paths or URLs are appended to sys.argv.
  * `-n NAME, --name NAME`: Name to assign to the bundled app and spec file (default: first script’s basename)
  * `-F, --onefile` : Create a one-file bundled executable.

## Support

For any issues or questions, please open an issue in the GitHub repository or contact `Brendan Swartz` at `bswartz@ohioschoolboards.org`.
