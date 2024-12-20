import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
df=pd.read_excel("C:\\data analytics\\sales_data.xlsx")
# print(df.to_string())
#
# print(df.isnull().sum())
df['ProductName']= df['ProductName'].fillna(value='Not Available')
df['OrderDate']= df['OrderDate'].ffill()
df=df.fillna({"Quantity" : "0" })
# df.drop_duplicates()
# df = df.dropna(axis=1, how='all')
# # print(df.info())

df['Year'] = df['OrderDate'].dt.year
df['Month'] = df['OrderDate'].dt.month
df["Total Sales"] =df["Quantity"]*df["PricePerUnit"]
df['MonthYear'] = df['Month'].astype(str) + '-' + df['Year'].astype(str)
print(df.info())
print(df.to_string())

x = df.groupby("Country")["Total Sales"].min()
print(x)
y = df.groupby("Country")["Total Sales"].max()
print(y)
z = df.groupby("Country")["Total Sales"].mean()
print(z)
df["maxium sales"]=df.groupby("ProductName")["Quantity"].sum()
print(df)

# Create a line graph
plt.figure(figsize=(10, 6))
x.plot(kind='line',marker=0)
plt.title('Maximum sales')
plt.xlabel('Country')
plt.ylabel('Total Sales')
plt.xticks(rotation=45)  # Rotate x labels for better readability
plt.grid(True)  # Add grid for better readability
plt.show()

plt.figure(figsize=(10, 6))
y.plot(kind='line',marker=0)
plt.title('Maximum sales')
plt.xlabel('Country')
plt.ylabel('Total Sales')
plt.xticks(rotation=45)  # Rotate x labels for better readability
plt.grid(True)  # Add grid for better readability
plt.show()

plt.figure(figsize=(10, 6))
z.plot(kind='line',marker=0)
plt.title('Maximum sales')
plt.xlabel('Country')
plt.ylabel('Total Sales')
plt.xticks(rotation=45)  # Rotate x labels for better readability
plt.grid(True)  # Add grid for better readability
plt.show()