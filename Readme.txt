Marks Entry Automation Script

This Python script allows users to enter marks for students based on their registration numbers from an Excel file. It provides a smooth and efficient way to update marks without manually opening the file.
Features

    3-digit Registration Number Entry:
        The user enters the last three digits of a registration number.
        The script automatically detects the full registration number.
    Handles Multiple Matches:
        If multiple students have the same last three digits, the script asks for the full registration number.
    One-Key Marks Entry:
        Once a student is identified, the script immediately prompts for marks (without pressing Enter).
        Only a single-digit mark (0-9) is accepted.
    Automatic Data Saving:
        Changes are saved only when the user exits the program (q).
    Case-Insensitive Matching:
        Works regardless of uppercase/lowercase differences in registration numbers.
    Error Handling:
        Notifies the user if a registration number is not found.
        Prevents incorrect or duplicate entries.

Usage
Run the Script

python marks.py "OOP SP25 Marks List.xlsx"

Replace "OOP SP25 Marks List.xlsx" with the actual Excel file name containing student records.
User Input Workflow

    Enter a 3-digit registration number (e.g., 044).
    If a single match is found → Enter marks (0-9).
    If multiple matches exist → Enter the full registration number (e.g., SP24-BSI-044).
    Marks are updated, and the next entry begins.
    Enter q at any point to exit and save changes.

Excel File Format

The Excel file must contain two columns:
Reg Number	Marks
SP24-BSI-044	5
FA23-ELC-011	8
SP21-BSI-002	7
Dependencies
And these column names must be in the first row of the file i.e no blank rows at start

Install required libraries before running the script:

pip install pandas openpyxl keyboard

License

This script is open-source and can be modified as needed.