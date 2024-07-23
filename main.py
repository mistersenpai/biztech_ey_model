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
#
# # Visualization 1: Spending by Category
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
top_15_suppliers_combined = pd.concat([top_15_suppliers, pd.Series(other_suppliers_spend, index=["Other"])])

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

# Visualization 5: Top Business Units by Spend
spend_by_bu = data_cleaned.groupby('Business Unit')['Line Amount'].sum().sort_values(ascending=False)
plt.figure(figsize=(10, 6))
spend_by_bu.plot(kind='bar')
plt.title('Top Business Units by Spend')
plt.ylabel('Total Spend')
plt.xlabel('Business Unit')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('top_business_units_by_spend.png')
plt.close()

# -------
# Extract month and year from the Invoice Date
data_cleaned['Month'] = data_cleaned['Invoice Date'].dt.to_period('M')

# Spend by Month
spend_by_month = data_cleaned.groupby('Month')['Line Amount'].sum()
plt.figure(figsize=(10, 6))
spend_by_month.plot(kind='line', marker='o')
plt.title('Spend by Month')
plt.ylabel('Total Spend')
plt.xlabel('Month')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('spend_by_month.png')
plt.close()

# Calculate the number of transactions for each supplier
transactions_by_supplier = data_cleaned['Supplier No'].value_counts().sort_values(ascending=False).head(15)
plt.figure(figsize=(10, 6))
transactions_by_supplier.plot(kind='bar')
plt.title('Number of Transactions by Top 15 Suppliers')
plt.ylabel('Number of Transactions')
plt.xlabel('Supplier No')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('transactions_by_supplier.png')
plt.close()


# Calculate average spend per transaction for each supplier
avg_spend_per_transaction = data_cleaned.groupby('Supplier No')['Line Amount'].mean().sort_values(ascending=False).head(15)
plt.figure(figsize=(10, 6))
avg_spend_per_transaction.plot(kind='bar')
plt.title('Average Spend per Transaction by Top 15 Suppliers')
plt.ylabel('Average Spend')
plt.xlabel('Supplier No')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.savefig('avg_spend_per_transaction_by_supplier.png')
plt.close()