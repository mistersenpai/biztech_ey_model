import pandas as pd
import matplotlib.pyplot as plt

# Set the backend for matplotlib
plt.switch_backend('agg')

# Load the Excel file
file_path = 'Copy of BizTech2 Case Comp - Supporting Data 2024.xlsx'
xls = pd.ExcelFile(file_path)

# Load each sheet into a DataFrame
spend_data = pd.read_excel(xls, 'Spend Data')
object_account_mapping = pd.read_excel(xls, 'Object Account Mapping')
category_hierarchy = pd.read_excel(xls, 'Category Hierarchy')

# Extract the relevant date information from the Spend Data sheet
spend_data['Invoice Date'] = pd.to_datetime(spend_data['Invoice Date'], errors='coerce')

# Filter data for the last 5 years
end_date = pd.Timestamp.now()
start_date = end_date - pd.DateOffset(years=5)
spend_data_filtered = spend_data[(spend_data['Invoice Date'] >= start_date) & (spend_data['Invoice Date'] <= end_date)]

# Clean and filter necessary data
spend_data_filtered = spend_data_filtered[['Line Amount', 'Object Account', 'Invoice Date']]
object_account_mapping = object_account_mapping[['Object Account', 'Mapped Category']]
category_hierarchy = category_hierarchy[['Mega_Category', 'Code']]

# Drop any rows with missing critical data
spend_data_filtered.dropna(subset=['Line Amount', 'Object Account', 'Invoice Date'], inplace=True)
object_account_mapping.dropna(subset=['Object Account', 'Mapped Category'], inplace=True)
category_hierarchy.dropna(subset=['Mega_Category', 'Code'], inplace=True)

# Merge spend_data with object_account_mapping on 'Object Account'
merged_data_filtered = spend_data_filtered.merge(object_account_mapping, how='left', left_on='Object Account', right_on='Object Account')

# Merge the result with category_hierarchy on 'Mapped Category' and 'Code'
final_data_filtered = merged_data_filtered.merge(category_hierarchy, how='left', left_on='Mapped Category', right_on='Code')

# Drop rows with missing Mega_Category
final_data_filtered.dropna(subset=['Mega_Category'], inplace=True)

# Group by date and Mega_Category and sum the Line Amount
final_data_filtered['YearMonth'] = final_data_filtered['Invoice Date'].dt.to_period('M')
grouped_time_series = final_data_filtered.groupby(['YearMonth', 'Mega_Category'])['Line Amount'].sum().reset_index()

# Pivot the data to get Mega_Category as columns
time_series_pivot = grouped_time_series.pivot(index='YearMonth', columns='Mega_Category', values='Line Amount').fillna(0)

# Plot the time series
plt.figure(figsize=(14, 8))
time_series_pivot.plot(kind='line', figsize=(14, 8))
plt.title('Expenditure of Line Amount Over the Last 5 Years by Mega Category')
plt.xlabel('Date')
plt.ylabel('Line Amount')
plt.legend(title='Mega Category', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.grid(True)
plt.tight_layout()
plt.savefig('line_amount_distribution_time_series.png')  # Save the figure as a file
plt.show()

print(time_series_pivot)
