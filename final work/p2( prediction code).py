import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
import warnings

file_path = 'book.xlsx'
df = pd.read_excel(file_path)

product_names = df['Product Name '].values 
sales_data = df.iloc[:, 3:].astype(float)

order = (1, 1, 1) 
for selected_product in [
    'ROUTER 1', 'TRANSCIEVER', 'SWITCH 1', 'ACCESS POINT 1', 'ACCESS POINT 2',
    'SWITCH 2', 'SWITCH 3', 'POWER SUPPLY 1', 'SWITCH 4', 'SWITCH 5',
    'SWITCH 6', 'ACCESS POINT 3', 'SUPERVISOR ENGINE', 'SWITCH 7',
    'WIRELESS CONTROLLER', 'SWITCH 8', 'SWITCH 9', 'ACCESS POINT 4',
    'SWITCH 10', 'POWER SUPPLY 2'
]:
    selected_sales = sales_data.loc[df['Product Name '] == selected_product].values.flatten()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")

        model = ARIMA(selected_sales, order=order)
        fit_model = model.fit()

    forecast_steps = 1
    forecast = fit_model.forecast(steps=forecast_steps)

    print(f"Forecasted sales for {selected_product} in 4th quarter of 2023: {forecast}")
