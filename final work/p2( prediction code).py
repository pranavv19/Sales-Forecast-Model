import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
import warnings

# Load the Excel file into a pandas DataFrame
file_path = 'book.xlsx'
df = pd.read_excel(file_path)

# Extract the relevant data for modeling
product_names = df['Product Name '].values  # Note the extra space at the end
sales_data = df.iloc[:, 3:].astype(float)

# Specify ARIMA order
order = (1, 1, 1)  # You may need to tune these parameters based on the data

# Forecast sales for each product in the list
for selected_product in [
    'ROUTER 1', 'TRANSCIEVER', 'SWITCH 1', 'ACCESS POINT 1', 'ACCESS POINT 2',
    'SWITCH 2', 'SWITCH 3', 'POWER SUPPLY 1', 'SWITCH 4', 'SWITCH 5',
    'SWITCH 6', 'ACCESS POINT 3', 'SUPERVISOR ENGINE', 'SWITCH 7',
    'WIRELESS CONTROLLER', 'SWITCH 8', 'SWITCH 9', 'ACCESS POINT 4',
    'SWITCH 10', 'POWER SUPPLY 2'
]:
    selected_sales = sales_data.loc[df['Product Name '] == selected_product].values.flatten()

    # Suppress the specific warning
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")

        # Perform ARIMA forecasting
        model = ARIMA(selected_sales, order=order)
        fit_model = model.fit()

    # Forecast sales for the 4th quarter of 2023 (assuming your data is quarterly)
    forecast_steps = 1
    forecast = fit_model.forecast(steps=forecast_steps)

    # Print the forecasted values
    print(f"Forecasted sales for {selected_product} in 4th quarter of 2023: {forecast}")
