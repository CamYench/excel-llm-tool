#!/usr/bin/env python3
"""
Script to generate a sample Excel file for testing the Excel-to-LLM Context Feeder Tool.
"""

import pandas as pd
import numpy as np
from pathlib import Path

def create_sample_data():
    """Create sample data for testing."""
    
    # Sample employee dataOpe now
    np.random.seed(42)  # For reproducible data
    
    # Generate sample data
    n_employees = 100
    
    departments = ['Engineering', 'Marketing', 'Sales', 'HR', 'Finance', 'Operations']
    cities = ['New York', 'Los Angeles', 'Chicago', 'Boston', 'San Francisco', 'Seattle', 'Austin', 'Denver']
    
    data = {
        'Employee_ID': range(1001, 1001 + n_employees),
        'First_Name': [f'Employee_{i}' for i in range(1, n_employees + 1)],
        'Last_Name': [f'Last_{i}' for i in range(1, n_employees + 1)],
        'Department': np.random.choice(departments, n_employees),
        'Position': np.random.choice(['Manager', 'Senior', 'Junior', 'Lead'], n_employees),
        'City': np.random.choice(cities, n_employees),
        'Age': np.random.randint(22, 65, n_employees),
        'Salary': np.random.randint(40000, 150000, n_employees),
        'Years_Experience': np.random.randint(0, 20, n_employees),
        'Performance_Rating': np.random.uniform(3.0, 5.0, n_employees).round(1),
        'Start_Date': pd.date_range('2020-01-01', periods=n_employees, freq='D').strftime('%Y-%m-%d'),
        'Is_Remote': np.random.choice([True, False], n_employees),
        'Projects_Completed': np.random.randint(0, 25, n_employees)
    }
    
    return pd.DataFrame(data)

def create_financial_data():
    """Create sample financial data."""
    
    np.random.seed(123)
    
    # Generate monthly financial data for 2 years
    dates = pd.date_range('2022-01-01', '2023-12-31', freq='M')
    
    data = {
        'Month': dates.strftime('%Y-%m'),
        'Revenue': np.random.uniform(50000, 200000, len(dates)).round(2),
        'Expenses': np.random.uniform(30000, 150000, len(dates)).round(2),
        'Profit': np.random.uniform(10000, 80000, len(dates)).round(2),
        'Marketing_Spend': np.random.uniform(5000, 25000, len(dates)).round(2),
        'R_D_Spend': np.random.uniform(8000, 35000, len(dates)).round(2),
        'Employee_Count': np.random.randint(80, 120, len(dates)),
        'Customer_Satisfaction': np.random.uniform(3.5, 5.0, len(dates)).round(2)
    }
    
    # Calculate derived metrics
    df = pd.DataFrame(data)
    df['Profit_Margin'] = (df['Profit'] / df['Revenue'] * 100).round(2)
    df['ROI'] = (df['Profit'] / df['Marketing_Spend'] * 100).round(2)
    
    return df

def main():
    """Generate sample Excel file with multiple sheets."""
    
    print("📊 Generating sample Excel file for testing...")
    
    # Create output directory if it doesn't exist
    output_dir = Path("sample_data")
    output_dir.mkdir(exist_ok=True)
    
    output_file = output_dir / "sample_company_data.xlsx"
    
    # Create Excel writer
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # Sheet 1: Employee Data
        print("  📋 Creating Employee Data sheet...")
        employee_df = create_sample_data()
        employee_df.to_excel(writer, sheet_name='Employee_Data', index=False)
        
        # Sheet 2: Financial Data
        print("  📋 Creating Financial Data sheet...")
        financial_df = create_financial_data()
        financial_df.to_excel(writer, sheet_name='Financial_Data', index=False)
        
        # Sheet 3: Department Summary
        print("  📋 Creating Department Summary sheet...")
        dept_summary = employee_df.groupby('Department').agg({
            'Employee_ID': 'count',
            'Salary': ['mean', 'min', 'max'],
            'Age': 'mean',
            'Performance_Rating': 'mean'
        }).round(2)
        dept_summary.columns = ['Employee_Count', 'Avg_Salary', 'Min_Salary', 'Max_Salary', 'Avg_Age', 'Avg_Performance']
        dept_summary.to_excel(writer, sheet_name='Department_Summary')
        
        # Sheet 4: City Analysis
        print("  📋 Creating City Analysis sheet...")
        city_analysis = employee_df.groupby('City').agg({
            'Employee_ID': 'count',
            'Salary': 'mean',
            'Performance_Rating': 'mean'
        }).round(2)
        city_analysis.columns = ['Employee_Count', 'Avg_Salary', 'Avg_Performance']
        city_analysis.to_excel(writer, sheet_name='City_Analysis')
    
    print(f"✅ Sample Excel file created: {output_file}")
    print(f"📁 File size: {output_file.stat().st_size / 1024:.2f} KB")
    
    # Display file info
    print("\n📊 File Contents:")
    xls = pd.ExcelFile(output_file)
    for i, sheet in enumerate(xls.sheet_names, 1):
        df = pd.read_excel(output_file, sheet_name=sheet)
        print(f"  {i}. {sheet}: {len(df)} rows × {len(df.columns)} columns")
    
    print(f"\n🎯 You can now test the tool with:")
    print(f"   GUI: streamlit run src/app.py")
    print(f"   CLI: python src/cli.py {output_file} -f csv --summary")
    print(f"   CLI: python src/cli.py {output_file} -f json -s Employee_Data,Financial_Data")

if __name__ == "__main__":
    main()
