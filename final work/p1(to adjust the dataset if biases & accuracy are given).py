import pandas as pd
import numpy as np
from statsmodels.tsa.arima.model import ARIMA

# Load the historical sales data from the Excel file
sales_file_path = 'dataraw.xlsx'
dataraw_df = pd.read_excel(sales_file_path)

# Extract the relevant data for modeling
product_names = dataraw_df['Product Name '].values
sales_data = dataraw_df.iloc[:, 3:].astype(float)

# Load the demand forecast data from the Excel file
demand_file_path = 'demandforecast.xlsx'
demand_df = pd.read_excel(demand_file_path)

demand_df['Product Name '] = demand_df['Product Name '].astype(str).str.strip()

sales_data = sales_data.ffill(axis=1)

order = (1, 1, 1)  # You may need to tune these parameters based on the data

for i, selected_product in enumerate(product_names):
    selected_sales = sales_data.loc[dataraw_df['Product Name '] == selected_product].values.flatten()

    # Check if demand forecast data is available for the selected product
    demand_data = demand_df[demand_df['Product Name '] == selected_product]
    if not demand_data.empty:
        for quarter in ['2023q1', '2023q2', '2023q3']:
            accuracy_col = f"{quarter}Accuracy"
            bias_col = f"{quarter}BIAS"

            accuracy = float(str(demand_data[accuracy_col].values[0]).rstrip('%')) / 100.0
            bias = float(str(demand_data[bias_col].values[0]).rstrip('%')) / 100.0

            # Apply adjustments to the sales data for each quarter
            sales_data.loc[dataraw_df['Product Name '] == selected_product] *= (1 + bias)
            sales_data.loc[dataraw_df['Product Name '] == selected_product] *= (1 + accuracy)

# Save the adjusted dataraw DataFrame to a new Excel file
adjusted_sales_file_path = 'adjusted_dataraw.xlsx'
sales_data.to_excel(adjusted_sales_file_path, index=False)
print(f"Adjusted sales data saved to {adjusted_sales_file_path}")
