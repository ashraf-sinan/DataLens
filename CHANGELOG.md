# Changelog

All notable changes to DataLens will be documented in this file.

## [1.0.0] - 2024-12-24

### Added
- **New Name**: Rebranded from "Excel Analysis System" to "DataLens"
- **Sheet Selection**: Select which sheet to analyze in multi-sheet Excel files
- **In-App Visualizations**: Visual statistics cards with color-coded metrics and horizontal bar charts
- **Modern UI Design**: Complete redesign with bigger fonts, better spacing, and professional color scheme
- **Tabbed Results**: Separate tabs for text summary and visualizations
- **GitHub-Ready Structure**: Organized project structure with src/, tests/, examples/, and docs/ directories
- **MIT License**: Added open-source license
- **.gitignore**: Comprehensive gitignore for Python projects
- **Project Documentation**: README with badges, quick start guide, and project structure

### Changed
- **UI Improvements**:
  - Increased font sizes (Title: 18pt, Headings: 13pt, Normal: 11pt)
  - Brighter color scheme with professional blues and greens
  - Card-based layout for better visual hierarchy
  - Step-by-step workflow (STEP 1, STEP 2, STEP 3)
  - Emoji icons on buttons for better UX
  - Modern status bar with white text on blue background

- **Visualizations**:
  - Replaced box plots with distribution density diagrams (line charts)
  - Added in-app visualization with statistics cards
  - Color-coded stat boxes for quantitative data
  - Horizontal bar charts showing frequency distributions
  - Type-specific visualizations (numeric vs categorical)

- **Analysis Engine**:
  - Added sheet_name parameter to ExcelAnalyzer
  - Support for multi-sheet Excel file analysis
  - Dynamic sheet selection with live updates

### Technical
- **File Structure**:
  - Moved source files to `src/` directory
  - Moved tests to `tests/` directory
  - Moved examples to `examples/` directory
  - Moved documentation to `docs/` directory

- **Build**:
  - Updated PyInstaller spec file for new structure
  - Renamed executable from ExcelAnalyzer.exe to DataLens.exe
  - Build output: ~540 MB standalone Windows executable

### Features
- Analyze 300+ columns automatically
- Statistical summaries (min, max, median, quartiles, average, sum)
- Frequency distributions with percentages
- Histogram charts for numeric data
- Distribution density diagrams for data spread
- Pie/bar charts for categorical data
- Group-by analysis for comparative insights
- Professional Excel export with navigation
- In-app visualization preview
- Multi-sheet support with sheet selector

## [0.1.0] - Initial Release

### Added
- Basic Excel file analysis
- Group-by functionality
- Excel export with charts
- Statistical analysis for numeric columns
- Frequency analysis for categorical columns
