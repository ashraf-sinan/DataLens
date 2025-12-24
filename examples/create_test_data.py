import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Set random seed for reproducibility
np.random.seed(42)

# Generate sample data
n_rows = 100

# Create dummy data
data = {
    'Department': np.random.choice(['Sales', 'Marketing', 'Engineering', 'HR', 'Finance'], n_rows),
    'Employee_Name': [f'Employee_{i}' for i in range(1, n_rows + 1)],
    'Age': np.random.randint(22, 65, n_rows),
    'Salary': np.random.randint(30000, 150000, n_rows),
    'Years_Experience': np.random.randint(0, 25, n_rows),
    'Performance_Score': np.random.choice([1, 2, 3, 4, 5], n_rows, p=[0.05, 0.15, 0.40, 0.30, 0.10]),
    'Location': np.random.choice(['New York', 'San Francisco', 'Chicago', 'Austin', 'Seattle'], n_rows),
    'Education': np.random.choice(['High School', 'Bachelor', 'Master', 'PhD'], n_rows, p=[0.10, 0.50, 0.30, 0.10]),
    'Bonus': np.random.randint(0, 20000, n_rows),
    'Projects_Completed': np.random.randint(0, 30, n_rows),
}

# Create DataFrame
df = pd.DataFrame(data)

# Add some calculated fields
df['Total_Compensation'] = df['Salary'] + df['Bonus']

# Save to Excel
output_file = 'test_employee_data.xlsx'
df.to_excel(output_file, index=False)

print(f"Test Excel file created successfully: {output_file}")
print(f"\nFile contains {len(df)} rows and {len(df.columns)} columns")
print(f"\nColumns:")
for col in df.columns:
    print(f"  - {col}")

print("\n\nSample data (first 5 rows):")
print(df.head())

print("\n\nColumn types:")
print(df.dtypes)

print("\n\nSuggested grouping columns: Department, Location, Education, Performance_Score")
