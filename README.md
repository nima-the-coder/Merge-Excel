# Merge-Excel

Excel File Merger & Cleaner
A Python tool with GUI to merge multiple Excel files, clean the data, and export results in Excel and CSV formats.

    ✨ Features
-Graphical User Interface (GUI) using Tkinter

-Easy folder selection with standard dialog

-Automatic merging of all Excel (.xlsx) files in selected folder

-Data cleaning: removal of empty rows/columns, duplicate columns

-Preview of files to be processed

-Export to both Excel (.xlsx) and CSV formats

-Comprehensive error handling and user feedback

    🛠️ Technologies Used
-Python 3.8+

-Pandas – Data manipulation and Excel operations

-Tkinter – Graphical user interface

-OS Module – File system operations

    📋 requirements
-Python 3.8 or higher

-Install Required Libraries

-bash
pip install pandas openpyxl

    🚀 Usage

Tool automatically:

-Lists all Excel files in folder

-Merges them into single DataFrame

-Cleans data (removes empty rows/columns)

-Saves results to output/ folder as:

-clean_merged.xlsx

-clean_merged.csv

    ⚠️ Notes
-Only processes .xlsx files (not .xls)

-Creates output/ folder if it doesn't exist

-Preserves all data types during merging

-CSV file uses UTF-8-SIG encoding for Persian/Arabic compatibility

    📄 License
This project is released under the MIT License.

    Developer
Name: [Nima]

Email: [nima.hashemi.dev@gmail.com]
