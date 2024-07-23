import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use the 'Agg' backend for non-interactive environments
import matplotlib.pyplot as plt

# Load the data
data_path = 'Copy of BizTech2 Case Comp - Supporting Data 2024.xlsx'
data = pd.read_excel(data_path, sheet_name='Spend Data')

# Convert relevant columns to appropriate data types
data['Invoice Date'] = pd.to_datetime(data['Invoice Date'], errors='coerce')
data['Line Amount'] = pd.to_numeric(data['Line Amount'], errors='coerce')

# Drop rows with missing values in critical columns for analysis
data_cleaned = data.dropna(subset=['Invoice Date', 'Line Amount', 'Business Unit'])

# Visualization 1: Spending by Category
spend_by_category = data_cleaned.groupby('Line Remark')['Line Amount'].sum().sort_values(ascending=False)
plt.figure(figsize=(10, 8))
spend_by_category.plot(kind='pie', autopct='%1.1f%%', startangle=140)
plt.title('Spending by Category')
plt.ylabel('')  # Hide the y-label for a cleaner look
plt.tight_layout()
plt.savefig('spending_by_category.png')
plt.close()

# Visualization 2: Spending by Supplier
spend_by_supplier = data_cleaned.groupby('Supplier No')['Line Amount'].sum().sort_values(ascending=False)

# Top 10 suppliers
top_supplier_spending = spend_by_supplier.head(10)
plt.figure(figsize=(10, 8))
top_supplier_spending.plot(kind='pie', autopct='%1.1f%%', startangle=140)
plt.title('Spending by Top 10 Suppliers')
plt.ylabel('')  # Hide the y-label for a cleaner look
plt.tight_layout()
plt.savefig('spending_by_top_10_suppliers.png')
plt.close()

# Top 15 suppliers with the rest combined
top_15_suppliers = spend_by_supplier.head(15)
other_suppliers_spend = spend_by_supplier.iloc[15:].sum()

# Add the "Other" category
top_15_suppliers_combined = top_15_suppliers.append(pd.Series(other_suppliers_spend, index=["Other"]))

plt.figure(figsize=(10, 8))
top_15_suppliers_combined.plot(kind='pie', autopct='%1.1f%%', startangle=140)
plt.title('Spending by Top 15 Suppliers and Others Combined')
plt.ylabel('')  # Hide the y-label for a cleaner look
plt.tight_layout()
plt.savefig('spending_by_top_15_suppliers_combined.png')
plt.close()

# Visualization 3: Invoice Amounts Over Time
plt.figure(figsize=(10, 6))
data_cleaned.groupby('Invoice Date')['Line Amount'].sum().plot()
plt.title('Invoice Amounts Over Time')
plt.ylabel('Total Spend')
plt.xlabel('Invoice Date')
plt.tight_layout()
plt.savefig('invoice_amounts_over_time.png')
plt.close()

# Visualization 4: Transactions by Business Unit
transactions_by_bu = data_cleaned['Business Unit'].value_counts().sort_values(ascending=False)
plt.figure(figsize=(10, 6))
transactions_by_bu.plot(kind='bar')
plt.title('Transactions by Business Unit')
plt.ylabel('Number of Transactions')
plt.xlabel('Business Unit')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('transactions_by_business_unit.png')
plt.close()
