# Project: Run Rates Tool

## Overview/How-To-Use

1. **Running the Program**: Double Click to .exe file
2. **Using the Interface**: You can either manually edit the cell run rates or paste the run rates from the live sheet into the box at the bottom.
3. **Run**: First - Press update the cell rate then press svae, the program will then run in the background
4. **Completion**: When the program is done an excel sheet with the output will be created.

## TECHNICAL STUFF 

### Logging Configuration
Configures logging to record events and errors in `app.log` and stream them to the console.

### Directory Management
Ensures a clean downloads directory by deleting existing files at the start of each run.

### Cell Run Rates Data
Initializes a dictionary of cell run rates and converts it into a Pandas DataFrame for further analysis.

### Web Scrapers
#### Item List Scraper (`itemlistscraper`)
- Scrapes the item list from a specific URL.
- Downloads the data as an Excel file.
- Loads the downloaded file into a Pandas DataFrame for analysis.

#### CoDate Scraper (`codedatescraper`)
- Scrapes the CoDate report from a specific URL.
- Downloads the data as an Excel file.
- Loads the downloaded file into a Pandas DataFrame for analysis.

### Data Analysis
#### `file_analysis`
- Analyzes the data from the CoDate and item list scrapers.
- Merges the data with cell run rates.
- Performs calculations to identify critical entries.
- Exports the analyzed data to an Excel file.

### GUI
Provides a Tkinter-based GUI for user interaction:
- Allows users to view and edit cell run rates.
- Provides a text area for pasting data, which is then processed and integrated into the tool.

## Usage

### Prerequisites
- Python 3.x
- Required libraries: `os`, `time`, `logging`, `pandas`, `datetime`, `plyer`, `selenium`, `tkinter`, `xlsxwriter`, `office365`, `requests_ntlm`

### Setup
1. Install the required libraries:
   ```bash
   pip install pandas plyer selenium tk office365-connector requests_ntlm xlsxwriter
   ```
