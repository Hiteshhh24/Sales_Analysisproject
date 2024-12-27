import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
df = pd.read_excel("C:\\excel file\\sales_data.xlsx")
# x = df.isnull().sum()
#missing values
df['ProductName']= df['ProductName'].fillna(value='Not Available')
df['OrderDate']= df['OrderDate'].ffill()

#remove the duplicates
df.drop_duplicates()

#convert data types
df['Year'] = df['OrderDate'].dt.year
df['Month'] = df['OrderDate'].dt.month

#Create a new column
df["Total Sales"] =df["Quantity"]*df["PricePerUnit"]

#Extract the month and year
df['MonthYear'] = df['Month'].astype(str) + '-' + df['Year'].astype(str)
# df= df.describe()
# print(df.to_string())

#Perform any additional data transformations
x = df.groupby("Country")["Total Sales"].min()
# print(x)
y = df.groupby("Country")["Total Sales"].max()
# print(y)
z = df.groupby("Country")["Total Sales"].mean()
# print(z)


salesdf = df.groupby("Country")["Total Sales"].sum()
zx=(salesdf.nlargest(10))
print(zx)

#Create visualizations (e.g., bar plots, line charts) to explore trends in sales over time and by country.

zx.plot(kind='bar', figsize=(10, 6), color='skyblue')
plt.title('Total Sales by Country')
plt.xlabel('Country')
plt.ylabel('Total Sales')
plt.xticks(rotation=90)
plt.grid(axis='y')
plt.show()


zx.plot(kind= 'line')
plt.title('largest sales country')
plt.xlabel('Country')
plt.ylabel('Total Sales')
plt.xticks(rotation=90)# Rotate x labels for better readability
plt.yticks(rotation=90)
plt.grid(True)  # Add grid for better readability
plt.show()

#Save the cleaned and transformed dataset to a new Excel file
df.to_excel('C:\\excel file\\sales_data_transformed.xlsx', index=False)