📂 ExcelReader
============================
This Python script processes an Excel workbook by reading specified cell ranges from a given worksheet, calculates the total and average of the numeric values (ignoring zeros and empty cells) within those ranges, and saves the summarized results into a new Excel file.

----------------------------
✨ Features
----------------------------
- Loads a specified sheet from an Excel workbook
- Processes multiple predefined cell ranges
- Calculates the total sum and average of non-zero values for each range
- Saves the summary results to a new Excel file- 

----------------------------
📋 Requirements
----------------------------
- Python 3.x
- openpyxl
- pandas 

Install dependencies:

    pip install openpyxl pandas

----------------------------
⚙️ Configuration
----------------------------
- Modify file_path variable to point to your input Excel file (.xlsx)
- Set sheet_name to the worksheet containing your data
- Adjust ranges dictionary to specify desired cell ranges for processing
- Change output_path to define where the summary Excel file will be saved

----------------------------
🕹️ Running the Script
----------------------------
To start the , run: ExcelAv.py
