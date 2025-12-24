"""Test script to verify Excel export functionality."""

from analyzer import ExcelAnalyzer
from excel_exporter import ExcelExporter

# Load test data
print("Loading test data...")
analyzer = ExcelAnalyzer('test_employee_data.xlsx')

# Test ungrouped analysis
print("Running ungrouped analysis...")
results = analyzer.analyze_all_columns()

print(f"Analyzed {len(results)} columns")

# Export ungrouped
print("Exporting ungrouped results to Excel...")
exporter = ExcelExporter(results, None)
exporter.export_ungrouped('test_output_ungrouped.xlsx')
print("[OK] Ungrouped export complete: test_output_ungrouped.xlsx")

# Test grouped analysis
print("\nRunning grouped analysis by Department...")
results_grouped = analyzer.analyze_by_group('Department')

print(f"Analyzed {len(results_grouped)} groups")

# Export grouped
print("Exporting grouped results to Excel...")
exporter_grouped = ExcelExporter(results_grouped, 'Department')
exporter_grouped.export_grouped('test_output_grouped.xlsx')
print("[OK] Grouped export complete: test_output_grouped.xlsx")

print("\n" + "=" * 60)
print("SUCCESS! Test exports completed.")
print("=" * 60)
print("\nGenerated files:")
print("  1. test_output_ungrouped.xlsx - Analysis without grouping")
print("  2. test_output_grouped.xlsx - Analysis grouped by Department")
print("\nOpen these files to see:")
print("  - Index sheet with hyperlinks")
print("  - Visualizations overview sheet")
print("  - Individual sheets for each column with charts")
