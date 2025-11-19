# Tableau Calculation and Lineage Extractor

A tool to extract and analyze calculations from Tableau workbooks with visual dependency diagrams.

## 🙏 Acknowledgments

This project is inspired by [tableauCalculationExport](https://github.com/scinana/tableauCalculationExport) by Ana Milana. The goal was to quickly develop and enhance it for Tableau users with additional features like interactive lineage diagrams and GUI interface. Special thanks to GitHub Copilot for accelerating the development process!

---

## 📋 Overview

This application parses XML metadata from Tableau workbooks (`.twb` and `.twbx` files) to extract all calculations, fields, and parameters, then generates:
- **Excel file** with detailed field metadata (formulas, data types, worksheets, datasources, usage indicators)
- **Interactive HTML diagram** showing calculation lineage and dependencies between fields

---

## ⚙️ System Requirements

- **Operating System**: Windows only
- **Python**: 3.7+ (if running Python scripts)
- **Tableau Files**: `.twb` or `.twbx` format

### Required Python Packages
```bash
# Core dependencies
pandas
tableaudocumentapi
xlsxwriter

# For GUI application
tkinter  # Usually included with Python on Windows

# For notebook version
jupyter

# For building executable
pyinstaller

```

## 🚀 Quick Start

### **Option 1: Use the GUI (Recommended for non-technical users)**

1. Run `tableau_extractor_gui.exe` (located in `dist/` folder)
2. Click "Browse" to select your Tableau workbook (`.twb` or `.twbx`)
3. Choose output directory (defaults to `outputs/` folder)
4. Check desired output options:
   - ☑ Generate Excel
   - ☑ Generate Mermaid Diagram
5. Click "Process Workbook"
6. Find results in the output folder

### **Option 2: Run Python Script**

```bash
# Place your .twb/.twbx file in the "inputs" folder
# NOTE: The script processes the first .twb/.twbx file it finds
# If adding a new workbook, empty the inputs folder first or remove old files

# Run the extractor (updated filename)
python "Tableau calculation and lineage extractor.py"

# Check "outputs" folder for results
```

## 🔧 How It Works

### Data Extraction Process
1. **Opens Tableau workbook** (.twb or .twbx)
   - For `.twbx` files: Extracts the embedded `.twb` XML file from the packaged workbook
   - For `.twb` files: Reads XML directly
2. **Parses XML structure** using Tableau Document API
3. **Extracts metadata from XML**:
   - Field names and internal IDs
   - Calculation formulas (with field ID references)
   - Data types
   - Worksheet usage information
   - Associated datasource information
4. **Processes calculations**:
   - Replaces field IDs with friendly field names in formulas
   - Identifies which fields are used in worksheets
   - Categorizes by field type (Parameters, Calculated, Default)
5. **Generates outputs**:
   - Excel file with 8 columns including usage indicators
   - Interactive HTML diagram showing only used fields and their dependencies

---

## 🔒 Data Security & Privacy

### What This Tool Accesses
✅ **XML metadata only**: Reads workbook structure from XML files  
✅ **Field definitions**: Extracts calculation formulas and field names  
✅ **Workbook structure**: Datasource names, worksheet names, field relationships  

### What This Tool Does NOT Access
❌ **No data connections**: Does not connect to databases or data sources  
❌ **No actual data**: Does not extract, view, or store any row-level data  
❌ **No credentials**: Does not access database passwords or connection strings  
❌ **No external calls**: Operates entirely offline on local files  


## 📝 Files Explained

| File | Purpose |
|------|------|
| `tableau_extractor_gui.exe` | **GUI executable** - Windows application (no Python needed) |
| `tableau_extractor_gui.py` | GUI source code (Tkinter interface) |
| `Tableau calculation and lineage extractor.py` | Standalone command-line Python script (v3.1) |
| `Tableau calculation and lineage extractor.ipynb` | Jupyter Notebook version with markdown documentation |
| `Excelcreator.py` | Helper module for Excel formatting with xlsxwriter |
| `tableau_extractor_gui.spec` | PyInstaller build configuration for GUI executable |
| `convert_md_to_docx.py` | Utility to convert markdown documentation to Word format |




 


