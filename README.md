# Unstructured Data Extraction from Excel

This repository hosts a Python-based GUI application tailored for extracting specific data from Excel sheets containing unstructured information. Designed to streamline and automate the data extraction process, this tool is both robust and user-friendly, making it a valuable asset for those dealing with disorganized Excel datasets.

![Screenshot of the Application](Screenshot%202023-10-26%20134520.png)

## Features

- **Directory Selection**: The application allows users to browse and select directories that house multiple Excel files.
  
- **Dynamic Keyword Downloading**: Once a directory is chosen, users can load unique sheets from the Excel files. The application then offers the functionality to dynamically download and display keywords from the selected sheets, giving users a clear view of available data points.
  
- **Keyword-Based Extraction with Value Association**: Users can opt to extract data based on specific keywords. Moreover, the tool lets users specify where the associated value is in relation to the keyword, effectively handling key-pair associations within the sheets.
  
- **Automated Summary Generation**: Post-extraction, the application generates a structured summary CSV file. This summary is saved in the initially selected directory with the name `0 - summary`, providing users with an organized overview of the extracted data at a glance.
  
- **Excel Test File Generation**: A script, `generate_test_files.py`, is included to create 50 Excel files in a folder on your desktop. These files contain unstructured data, making them ideal for testing the extraction capabilities of the GUI application.

## Installation

1. **Clone the Repository**:
   Begin by cloning this repository to your local machine.
   ```bash
   git clone https://github.com/dawsonporter/Unstructured-Data-Extraction-from-Excel.git

