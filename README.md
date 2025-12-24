# DataLens 📊

**Explore 300 columns of data in just 3 clicks.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

## The Problem

Six years ago, I faced a daunting challenge: analyzing Excel files with up to 300 columns. Each column required manual exploration—scrolling, filtering, calculating statistics, creating charts. It was tedious, time-consuming, and error-prone. Hours turned into days just to understand the data landscape.

## The Solution

**DataLens** automates the entire process. What used to take hours now takes **3 clicks**:

1. **Click 1**: Browse and select your Excel file
2. **Click 2**: Run Analysis (with optional grouping)
3. **Click 3**: Export to Excel with full visualizations

That's it. Instant comprehensive analysis with statistical summaries, frequency distributions, and interactive charts for every single column.

## Key Benefits

- ✅ **Analyze 300+ columns automatically** - No manual work required
- ✅ **Statistical summaries** - Min, max, median, percentiles, averages instantly
- ✅ **Visual insights** - Histograms and distribution density charts for all numeric data
- ✅ **Smart detection** - Automatically identifies numeric vs categorical columns
- ✅ **Professional reports** - Export to Excel with navigation, charts, and formatting
- ✅ **Group-by analysis** - Compare data across categories effortlessly
- ✅ **In-app visualization** - View results with charts and tables before exporting

## Features

- Load Excel files (.xlsx, .xls)
- Choose a column to group analysis by (optional)
- Automatic detection of quantitative and qualitative columns
- **Quantitative Analysis** (numeric columns):
  - Minimum value
  - Maximum value
  - Average
  - Percentiles (25th, 50th/Median, 75th)
  - Sum
  - Percentage of total
  - Count of values
  - Frequency distribution (value, count, percentage, value sum, % of total)
- **Qualitative Analysis** (categorical columns):
  - Label (unique values)
  - Frequency (count of each value)
  - Percentage of column with this value
- **Excel Export with Visualizations**:
  - Index sheet with hyperlinks to all columns
  - Visualizations overview sheet
  - Individual sheet for each column with charts
  - Bar charts for quantitative columns
  - Pie/bar charts for qualitative columns

## Installation

1. Install Python 3.8 or higher
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the application:

```bash
python main.py
```

### Steps:
1. Click "Browse Excel File" to select your Excel file
2. (Optional) Select a column to group analysis by from the dropdown
3. Click "Run Analysis" to perform the analysis
4. View results in the text area
5. Click "Export Results" to save the analysis to Excel with visualizations

## Requirements

- Python 3.8+
- pandas
- openpyxl
- numpy
- tkinter (included with Python)

## Testing

A test Excel file with dummy employee data is included:

```bash
python create_test_data.py  # Regenerate test data if needed
```

Then load `test_employee_data.xlsx` in the application to try it out!

See `USAGE_GUIDE.md` for detailed examples and tutorials.

## 📁 Project Structure

```
DataLens/
├── src/
│   ├── main.py              # Main GUI application
│   ├── analyzer.py          # Core analysis engine
│   └── excel_exporter.py    # Excel export with visualizations
├── tests/
│   └── test_export.py       # Export functionality tests
├── examples/
│   ├── test_employee_data.xlsx  # Sample data file
│   └── create_test_data.py      # Generate test data
├── docs/
│   ├── USAGE_GUIDE.md       # Detailed usage examples
│   ├── FEATURES.md          # Feature documentation
│   └── LINKEDIN_POST.md     # Marketing content
├── dist/
│   └── DataLens.exe         # Windows executable
├── requirements.txt         # Python dependencies
├── LICENSE                  # MIT License
└── README.md               # This file
```

## 🚀 Quick Start

### Option 1: Windows Executable (No Installation)

1. Download `DataLens.exe` from the [releases page](../../releases)
2. Double-click to run
3. Start analyzing your Excel files!

### Option 2: Run from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/datalens.git
cd datalens

# Install dependencies
pip install -r requirements.txt

# Run the application
python src/main.py
```

## Example Output

The exported Excel file contains:
- **Index Sheet**: Clickable links to navigate to any column
- **Visualizations Sheet**: Overview and instructions
- **Column Sheets**: One per column with:
  - Statistical summary or category breakdown
  - Frequency distribution tables
  - Embedded charts (bar charts for numeric, pie/bar for categorical)
