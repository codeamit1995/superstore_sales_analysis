import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.arima.model import ARIMA

sns.set(style="whitegrid")
file = "Superstore.xlsx"
df = pd.read_excel(file)

#checking the dataset
print("Data Shape:", df.shape)
print(df.head())
#coverting to Dates
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])

#Profit Margin
df['Profit Margin'] = (df['Profit'] / df['Sales'])*100

# Sales Trends (monthly)
sales_trend = df.groupby(df['Order Date'].dt.to_period("M"))['Sales'].sum()
sales_trend.plot(kind='line', figsize=(20,10), title="Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Total Sales")
plt.show()


#Profit by Region
plt.figure(figsize=(8,5))
sns.barplot(x="Region", y="Profit", data=df, estimator=sum)
plt.title("Total Profit by Region")
plt.show()

# Top10 largest Products by Sales
top_products = df.groupby("Product Name")['Sales'].sum().nlargest(10)
top_products.plot(kind='barh', figsize=(10,8), title="Top 10 Products by Sales")
plt.xlabel("Total Sales")
plt.show()

# Discount vs Profit graph
plt.figure(figsize=(8,5))
sns.scatterplot(x="Discount", y="Profit", data=df, alpha=.5)
plt.title("Discount vs Profit")
plt.show()


# Prepare monthly sales
monthly_sales = df.groupby(df['Order Date'].dt.to_period("M"))['Sales'].sum().reset_index()
monthly_sales['Order Date'] = monthly_sales['Order Date'].dt.to_timestamp()
monthly_sales = monthly_sales.set_index('Order Date')
monthly_sales = monthly_sales.asfreq('MS')
#ARIMA
model = ARIMA(monthly_sales['Sales'], order=(1,1,1))  # simple config
model_fit = model.fit()

# Forecast for next 12 months
forecast = model_fit.get_forecast(steps=12)
forecast_index = pd.date_range(start=monthly_sales.index[-1] + pd.offsets.MonthBegin(),
                               periods=12, freq='MS')
forecast_df = pd.DataFrame({
    "Date": forecast_index,
    "Predicted": forecast.predicted_mean,
    "Lower Bound": forecast.conf_int()['lower Sales'].values,
    "Upper Bound": forecast.conf_int()['upper Sales'].values
})

# Plot forecast
plt.figure(figsize=(10,5))
plt.plot(monthly_sales.index, monthly_sales['Sales'], label="Actual Sales")
plt.plot(forecast_index, forecast_df['Predicted'], label="Forecast", color="red")
plt.fill_between(forecast_index,forecast_df['Lower Bound'],forecast_df['Upper Bound'],color="green",alpha=0.5)
#plot
plt.title("Sales Forecast (Next 12 Months)")
plt.xlabel("Date")
plt.ylabel("Sales")
plt.legend()
plt.show()

excel_n = "Superstore_Analysis_Dashboard.xlsx"

with pd.ExcelWriter(excel_n, engine="openpyxl") as writer:
    sales_trend.to_timestamp().to_frame("Sales").to_excel(writer,sheet_name="Sales Trend")
    df.groupby("Region")[['Sales','Profit']].sum().to_excel(writer,sheet_name="Region Summary")
    df.groupby("Category")[['Sales','Profit']].sum().to_excel(writer,sheet_name="Category Summary")
    df.groupby("Segment")[['Sales','Profit']].sum().to_excel(writer,sheet_name="Segment Summary")
    top_products.to_frame("Sales").to_excel(writer,sheet_name="Top Products")
    df[['Discount','Profit']].to_excel(writer,sheet_name="Discount vs Profit",index=False)
    forecast_df.to_excel(writer,sheet_name="Sales Forecast",index=False)
    df.to_excel(writer,sheet_name="Cleaned Data",index=False)

print("Analysis Complete!")

