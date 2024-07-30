import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use the 'Agg' backend for non-interactive environments
import matplotlib.pyplot as plt

# Load the data
data_path = 'Copy of BizTech2 Case Comp - Supporting Data 2024.xlsx'
spend_data = pd.read_excel(data_path, sheet_name='Spend Data')
business_divisions = pd.read_excel(data_path, sheet_name='Business Divisions')

# Convert relevant columns to appropriate data types
spend_data['Invoice Date'] = pd.to_datetime(spend_data['Invoice Date'], errors='coerce')
spend_data['Line Amount'] = pd.to_numeric(spend_data['Line Amount'], errors='coerce')

# Drop rows with missing values in critical columns for analysis
spend_data_cleaned = spend_data.dropna(subset=['Invoice Date', 'Line Amount', 'Business Unit'])

# Merge the spend data with the business divisions data
merged_data = pd.merge(spend_data_cleaned, business_divisions, on='Business Unit', how='left')

# Check if the merge was successful
print(merged_data.head())

# Visualization: Spending by Business Division
spend_by_division = merged_data.groupby('BU Dvision')['Line Amount'].sum().sort_values(ascending=False)
plt.figure(figsize=(10, 8))
spend_by_division.plot(kind='pie', autopct='%1.1f%%', startangle=140)
plt.title('Spending by Business Division')
plt.ylabel('')  # Hide the y-label for a cleaner look
plt.tight_layout()
plt.savefig('spending_by_business_division.png')
plt.close()
