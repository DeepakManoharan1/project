Document Search Tool
Overview
This project is a Document Search Tool built with Python and Tkinter. It allows users to search for specific requirement numbers across multiple DOCX files and extract relevant information, including:

Requirements from an SRD (System Requirements Document)
Traces from an SDD (System Design Document)
Technical specifications
Test case tables
The tool also provides options to save the extracted data into an Excel file.

Features
Search and display requirement details from multiple DOCX files.
Extract trace information and technical specifications.
Navigate between requirements using "Previous" and "Next" buttons.
Save results to an Excel file with formatted output.
How to Run the Project
Install the necessary libraries:

Copy code
pip install tkinter python-docx openpyxl
Update the paths for your DOCX files in the script:

python
Copy code
SRD_MFD_PATH = r'path\to\SRD_MFD.docx'
SDD_MFD_PATH = r'path\to\SDD_MFD.docx'
TC_PATH = r'path\to\TC_MFD_GROUND_SPEED.docx'
TEC_SPEC_PATH = r'path\to\Tech Spec MFD.docx'
Run the script using Python:

Copy code
python document_search_tool.py
How to Use
Select a requirement number from the dropdown menu.
Click "Search" to view the details.
Use the "Previous" and "Next" buttons to navigate between requirements.
Click "Save" to export the data to an Excel file.
