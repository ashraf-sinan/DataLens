# Excel Analysis System - Usage Guide

## Installation

1. Make sure Python 3.8+ is installed
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

```bash
python main.py
```

## Using the Test Data

A test Excel file `test_employee_data.xlsx` has been created with 100 rows of employee data including:

**Quantitative Columns:**
- Age (22-64 years)
- Salary ($30,000-$150,000)
- Years_Experience (0-25 years)
- Performance_Score (1-5)
- Bonus ($0-$20,000)
- Projects_Completed (0-30)
- Total_Compensation (Salary + Bonus)

**Qualitative Columns:**
- Department (Sales, Marketing, Engineering, HR, Finance)
- Employee_Name (Employee_1 to Employee_100)
- Location (New York, San Francisco, Chicago, Austin, Seattle)
- Education (High School, Bachelor, Master, PhD)

## Step-by-Step Tutorial

### Example 1: Analyze without grouping
1. Click "Browse Excel File" and select `test_employee_data.xlsx`
2. Leave "Group By Column" set to "None"
3. Click "Run Analysis"
4. View the results showing statistics for all columns

### Example 2: Analyze by Department
1. Load the test file
2. Select "Department" from the "Group By Column" dropdown
3. Click "Run Analysis"
4. See separate statistics for Sales, Marketing, Engineering, HR, and Finance departments

### Example 3: Analyze by Location
1. Load the test file
2. Select "Location" from dropdown
3. Click "Run Analysis"
4. Compare statistics across different office locations

### Example 4: Export Results
1. After running any analysis
2. Click "Export Results"
3. Choose where to save the text file
4. Open in any text editor

## Understanding the Results

### Quantitative (Numeric) Columns Show:
- **Count**: Number of values
- **Minimum**: Smallest value
- **25th Percentile**: 25% of values are below this
- **Median (50th)**: Middle value
- **75th Percentile**: 75% of values are below this
- **Maximum**: Largest value
- **Average**: Mean of all values
- **Sum**: Total of all values
- **% of Total**: Percentage this group represents of the overall total
- **Frequency Distribution**: Table showing each unique value, how often it appears, and its percentage

### Qualitative (Categorical) Columns Show:
- **Label**: Each unique category/value
- **Frequency**: How many times this value appears
- **Percentage**: What % of the column has this value

## Tips

- Use "Department" or "Location" for meaningful group comparisons
- Try different grouping columns to discover insights
- Export results for reporting or further analysis
- The frequency distribution in numeric columns helps identify common values

## Example Insights from Test Data

By grouping by Department, you can answer:
- Which department has the highest average salary?
- Which department has the most experienced employees?
- How does performance score distribution vary by department?

By grouping by Education, you can see:
- How education level correlates with salary
- Which education levels have the most employees
- Performance differences across education levels
