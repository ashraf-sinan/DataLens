import pandas as pd
import numpy as np
from typing import Dict, List, Any


class ExcelAnalyzer:
    """Handles Excel file analysis with grouping capabilities."""

    def __init__(self, file_path: str, sheet_name: str = 0):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df = None
        self.load_file()

    def load_file(self):
        """Load Excel file into pandas DataFrame."""
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
        except Exception as e:
            raise Exception(f"Error loading file: {str(e)}")

    def get_columns(self) -> List[str]:
        """Get list of all columns in the dataset."""
        return list(self.df.columns)

    def analyze_quantitative(self, column: str, grouped_data=None) -> Dict[str, Any]:
        """
        Analyze quantitative (numeric) column.
        Returns: min, max, average, percentiles (25, 50, 75), sum, % of total, frequency
        """
        data = grouped_data[column] if grouped_data is not None else self.df[column]

        # Filter out non-numeric values
        numeric_data = pd.to_numeric(data, errors='coerce').dropna()

        if len(numeric_data) == 0:
            return None

        total_sum = numeric_data.sum()
        grand_total = pd.to_numeric(self.df[column], errors='coerce').sum()

        # Calculate frequency distribution
        value_counts = numeric_data.value_counts().sort_index()
        frequency_data = []
        total_count = len(numeric_data)

        for value, count in value_counts.items():
            value_sum = value * count
            percent_of_total_column = (value_sum / total_sum * 100) if total_sum != 0 else 0
            frequency_data.append({
                'value': value,
                'frequency': int(count),
                'percentage': (count / total_count * 100) if total_count > 0 else 0,
                'value_sum': value_sum,
                'percent_of_total_column': percent_of_total_column
            })

        result = {
            'min': numeric_data.min(),
            'max': numeric_data.max(),
            'average': numeric_data.mean(),
            'percentile_25': numeric_data.quantile(0.25),
            'percentile_50': numeric_data.quantile(0.50),
            'percentile_75': numeric_data.quantile(0.75),
            'sum': total_sum,
            'percent_of_total': (total_sum / grand_total * 100) if grand_total != 0 else 0,
            'count': len(numeric_data),
            'frequency': frequency_data
        }

        return result

    def analyze_qualitative(self, column: str, grouped_data=None) -> List[Dict[str, Any]]:
        """
        Analyze qualitative (categorical) column.
        Returns: label, frequency, % of column that has this value
        """
        data = grouped_data[column] if grouped_data is not None else self.df[column]

        # Get value counts
        value_counts = data.value_counts()
        total_count = len(data.dropna())

        results = []
        for label, frequency in value_counts.items():
            results.append({
                'label': str(label),
                'frequency': int(frequency),
                'percentage': (frequency / total_count * 100) if total_count > 0 else 0
            })

        return results

    def is_numeric_column(self, column: str) -> bool:
        """Check if a column is numeric."""
        return pd.api.types.is_numeric_dtype(self.df[column])

    def analyze_by_group(self, group_column: str) -> Dict[str, Dict]:
        """
        Perform analysis grouped by a specific column.
        Returns results for each group.
        """
        if group_column not in self.df.columns:
            raise ValueError(f"Column '{group_column}' not found in dataset")

        results = {}
        grouped = self.df.groupby(group_column)

        for group_name, group_data in grouped:
            group_results = {
                'group_name': str(group_name),
                'row_count': len(group_data),
                'columns': {}
            }

            # Analyze each column in the group
            for column in self.df.columns:
                if column == group_column:
                    continue

                if self.is_numeric_column(column):
                    analysis = self.analyze_quantitative(column, group_data)
                    if analysis:
                        group_results['columns'][column] = {
                            'type': 'quantitative',
                            'data': analysis
                        }
                else:
                    analysis = self.analyze_qualitative(column, group_data)
                    if analysis:
                        group_results['columns'][column] = {
                            'type': 'qualitative',
                            'data': analysis
                        }

            results[str(group_name)] = group_results

        return results

    def analyze_all_columns(self) -> Dict[str, Dict]:
        """
        Perform analysis on all columns without grouping.
        """
        results = {}

        for column in self.df.columns:
            if self.is_numeric_column(column):
                analysis = self.analyze_quantitative(column)
                if analysis:
                    results[column] = {
                        'type': 'quantitative',
                        'data': analysis
                    }
            else:
                analysis = self.analyze_qualitative(column)
                if analysis:
                    results[column] = {
                        'type': 'qualitative',
                        'data': analysis
                    }

        return results
