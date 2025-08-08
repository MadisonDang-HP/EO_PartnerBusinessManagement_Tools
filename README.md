# EO Partner Business Management Tools

A comprehensive suite of Python tools designed to streamline partner business management Documentation Kit processes including cost analysis, specification comparison, and price tracking.

## Overview

This toolkit consists of three powerful applications that help manage and analyze partner business data:

1. **Cost Upload Tool** - Automates the generation of cost data files by looking up part prices from supplier/ODM databases and populating cost upload templates
2. **Spec Comparator** - Matches parts against specification databases to find identical or similar specifications across different suppliers
3. **Historical Cost Delta Analyzer** - Detects and analyzes price increases and variances between different time periods

## Prerequisites

### Required Software
- Python 3.7 or higher
- Microsoft Excel (for viewing/editing output files)

### Required Python Packages
Install the following packages using pip:

```bash
pip install pandas openpyxl tkinter xlrd pyxlsb
```

### Optional Dependencies
- `pyxlsb` - Required for reading Excel Binary Worksheet (.xlsb) files in the Historical Cost Delta Analyzer

```bash
pip install pyxlsb
```

## Installation

1. Clone or download this repository to your local machine
2. Install the required Python packages (see Prerequisites)
3. Ensure all Python files are in the same directory
4. Your directory structure should look like:
   ```
   EO_PartnerBusinessManagement_Tools/
   ├── Cost_Upload_Tool.py
   ├── Historical_Cost_Delta_Analyzer.py
   ├── Spec_Comparator.py
   └── README.md
   ```

## Creating Executable Applications

To convert these Python scripts into standalone executable applications that don't require Python to be installed:

### Install PyInstaller
```bash
pip install pyinstaller
```

### Create Executables

**For Cost Upload Tool:**
```bash
pyinstaller --onefile --windowed --name "Cost_Upload_Tool" Cost_Upload_Tool.py
```

**For Historical Cost Delta Analyzer:**
```bash
pyinstaller --onefile --windowed --name "Historical_Cost_Delta_Analyzer" Historical_Cost_Delta_Analyzer.py
```

**For Spec Comparator:**
```bash
pyinstaller --onefile --windowed --name "Spec_Comparator" Spec_Comparator.py
```

The executable files will be created in the `dist/` folder and can be distributed to users without Python installed.

## Tool Usage Instructions

### 1. Cost Upload Tool

**Purpose:** Automatically populates cost upload templates by looking up part prices from supplier/ODM databases based on site codes and requested dates.

**Input Requirements:**
- Media Tracker Excel file with part numbers, site codes, and requested dates
- Site Info file mapping site codes to suppliers, ODMs, and MS4 vendor codes
- Root folder containing supplier/ODM price databases organized by date
- Cost Upload Template (Excel file)

**How to Use:**

1. **Launch the application:**
   - Run `python Cost_Upload_Tool.py` or use the executable
   
2. **Select Media Tracker File:**
   - Click "Browse" next to "Media Tracker File"
   - Select your Excel file containing the requested parts list
   - Choose the appropriate sheet from the dropdown

3. **Select Site Info File:**
   - Click "Browse" next to "Site Info File"
   - Select the Excel file containing site code mappings

4. **Select Root Folder:**
   - Click "Browse" next to "Root Folder"
   - Select the folder containing your supplier/ODM price databases
   - Folder structure should be: `Root/Supplier/ODM/Date/PriceFiles.xlsx`

5. **Select Template:**
   - Click "Submit" and you'll be prompted to select the Cost Upload Template
   - The tool will create a timestamped copy for output

6. **Review Results:**
   - The tool will process all parts and generate price lookups
   - Progress will be shown via progress bar
   - Output file will be saved with timestamp
   - Comments will be updated in the original Media Tracker

**Expected Output:**
- Populated cost upload template with prices, suppliers, vendor codes
- Updated Media Tracker with "Cost Uploaded" comments for processed parts

### 2. Spec Comparator

**Purpose:** Matches parts from a quote against specification databases to find identical or similar specifications and compare pricing.

**Input Requirements:**
- Quote Items Excel file with part specifications
- Specs folder containing specification database files (.xls, .xlsx, .xlsb)

**How to Use:**

1. **Launch the application:**
   - Run `python Spec_Comparator.py` or use the executable

2. **Select Quote Items File:**
   - Click "Browse" next to "Quote Items Excel File"
   - Select your Excel file containing the parts to be matched

3. **Select Specs Folder:**
   - Click "Browse" next to "Specs Folder"
   - Select the folder containing your specification database files

4. **Set Output File:**
   - Enter desired output filename or click "Browse" to select location
   - Default: "Quote_Spec_Comparison.xlsx"

5. **Run Comparison:**
   - Click "Run Comparison"
   - Progress will be displayed during processing

**Expected Output:**
- Excel file with two sheets:
  - **Matched Parts:** Parts with exact specification matches, including pricing from spec databases
  - **Unmatched Parts:** Parts without exact matches, with closest specifications and confidence scores
- Color-coded pricing (green=lowest, red=highest, yellow=tied prices)
- Specification difference analysis for unmatched parts

### 3. Historical Cost Delta Analyzer

**Purpose:** Analyzes historical cost data to identify price increases and variances between different time periods.

**Input Requirements:**
- Excel file (.xlsx, .xls, or .xlsb) containing historical pricing data with variance information

**How to Use:**

1. **Launch the application:**
   - Run `python Historical_Cost_Delta_Analyzer.py` or use the executable

2. **Select Excel File:**
   - Click "Browse" to select your historical cost data file
   - Supports .xlsx, .xls, and .xlsb formats (install pyxlsb for .xlsb support)

3. **Configure Analysis Options:**
   - **Include BOM Variances:** Check to analyze Bill of Materials related variances
   - **Include Spec Variances:** Check to analyze specification-related variances

4. **Set Output Options:**
   - **Auto-open result file:** Check to automatically open the output file when complete

5. **Run Analysis:**
   - Click "Analyze File"
   - Progress will be shown via progress bar
   - Results will be displayed in the text area

**Expected Output:**
- Excel file named "Historical Cost Delta Analyzer [DATE].xlsx"
- Separate sheets for BOM variances and Spec variances
- Analysis summary showing total variances found
- Detailed variance information with part numbers, prices, and delta amounts

## File Structure Requirements

### Cost Upload Tool Directory Structure:
```
Root_Folder/
├── Supplier1/
│   ├── ODM1/
│   │   ├── JAN'25/
│   │   │   ├── Final_PriceList.xlsx
│   │   │   └── New_PriceList.xlsx
│   │   └── FEB'25/
│   │       └── Initial_PriceList.xlsx
│   └── ODM2/
└── Supplier2/
```

### Spec Comparator Directory Structure:
```
Specs_Folder/
├── SpecDatabase1.xlsx
├── SpecDatabase2.xlsb
└── SpecDatabase3.xls
```

## Troubleshooting

### Common Issues:

1. **Missing Dependencies Error:**
   - Ensure all required packages are installed: `pip install pandas openpyxl tkinter xlrd pyxlsb`

2. **File Access Errors:**
   - Close Excel files before processing
   - Ensure you have write permissions to output directories

3. **XLSB Support Error:**
   - Install pyxlsb: `pip install pyxlsb`

4. **Memory Issues with Large Files:**
   - Close other applications
   - Process smaller batches if necessary

5. **Path Issues on Windows:**
   - Use absolute paths when possible
   - Ensure folder paths don't contain special characters
