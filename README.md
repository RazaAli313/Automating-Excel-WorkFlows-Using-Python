# Automating-Excel-WorkFlows-Using-Python

Excel Workflow Automation with OpenPyXL
This Python script automates Excel spreadsheet processing using the openpyxl library. It modifies numerical values, adds a bar chart, and saves an updated version of the file.

Features
Reads an Excel file and processes data in Sheet1.

Updates values in the third column (reducing by 10%) and writes them to the fourth column.

Adds a bar chart to visualize the updated values.

Saves a new Excel file with modifications.

Requirements
Python 3.x

openpyxl library

Installation
bash
Copy
Edit
pip install openpyxl
Usage
Ensure your Excel file (e.g., sales.xlsx) exists in the script's directory.

Run the script with:

python
Copy
Edit
process_spreadsheet("sales.xlsx")
The modified file is saved as sales_updated.xlsx.
