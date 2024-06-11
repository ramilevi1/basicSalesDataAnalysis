import pandas as pd
import matplotlib.pyplot as plt



# Read in the Excel file
df = pd.read_excel("data.xlsx", sheet_name="Data")

# Calculate the total sales and expenses for each month
df['Month'] = pd.DatetimeIndex(df['Date']).month
monthly_sales = df.groupby('Month')['Sales'].sum()
monthly_expenses = df.groupby('Month')['Expenses'].sum()

# Calculate the profit for each month
monthly_profit = monthly_sales - monthly_expenses

# Print to check if the sums are computed correctly
print("Monthly Sales:")
print(monthly_sales)
print("Monthly Expenses:")
print(monthly_expenses)
print("Monthly Profit:")
print(monthly_profit)


# Plot the monthly sales, expenses, and profit
plt.figure(figsize=(10, 5))
plt.plot(monthly_sales, label='Sales')
plt.plot(monthly_expenses, label='Expenses')
plt.plot(monthly_profit, label='Profit')
plt.xlabel('Month')
plt.ylabel('Amount')
plt.title('Monthly Sales, Expenses, and Profit')
plt.legend()
plt.show()

# Calculate the top-selling products
product_sales = df.groupby('Product')['Sales'].sum().sort_values(ascending=False)

# Plot a bar chart of the top-selling products
plt.bar(product_sales.index, product_sales.values)
plt.xlabel('Product')
plt.ylabel('Sales')
plt.title('Top-Selling Products')
plt.show()
