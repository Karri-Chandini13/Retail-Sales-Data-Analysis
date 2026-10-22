
import pandas as pd
import matplotlib.pyplot as plt

# Load the sales dataset
df = pd.read_csv("sales_data.csv")

# Display the first 5 rows
print(df.head())


# Display all column names
print("\nColumn Names:")
print(df.columns)

# Display number of rows and columns
print("\nDataset Shape:")
print(df.shape)

# Display dataset information
print("\nDataset Information:")
df.info()

# Display statistical summary
print("\nStatistical Summary:")
print(df.describe())

# Step 10: Check missing values

print("\nMissing Values in Each Column:")
print(df.isnull().sum())

print("\nTotal Missing Values:")
print(df.isnull().sum().sum())

# Step 11: Check duplicate rows

print("\nDuplicate Rows:")
print(df.duplicated().sum())

# Step 12: Remove duplicate rows

df = df.drop_duplicates()

# Remove extra spaces from text columns
text_columns = df.select_dtypes(include="object").columns

for column in text_columns:
    df[column] = df[column].str.strip()

print("\nDataset Shape After Cleaning:")
print(df.shape)

# Step 14: Calculate Sales

df["Sales"] = df["Quantity"] * df["UnitPrice"] * (1 - df["Discount"])

print("\nSales Column Created:")
print(df[["Product", "Quantity", "UnitPrice", "Discount", "Sales"]].head())

# Step 15: Calculate Total Sales

total_sales = df["Sales"].sum()

print("\nTotal Sales Amount:")
print(round(total_sales, 2))

# Step 16: Calculate Average Sales

average_sales = df["Sales"].mean()

print("\nAverage Sales Per Order:")
print(round(average_sales, 2))

# Step 17: Product-wise Sales Analysis

product_sales = df.groupby("Product")["Sales"].sum()

product_sales = product_sales.sort_values(ascending=False)

print("\nProduct-wise Sales:")
print(product_sales)

# Step 18: Region-wise Sales Analysis

region_sales = df.groupby("Region")["Sales"].sum()

region_sales = region_sales.sort_values(ascending=False)

print("\nRegion-wise Sales:")
print(region_sales)

# Step 19: Find Top Product and Region

top_product = product_sales.idxmax()
top_region = region_sales.idxmax()

print("\nBest-Selling Product:", top_product)
print("Highest-Sales Region:", top_region)

# Step 21: Category-wise Sales Analysis

category_sales = df.groupby("Category")["Sales"].sum()

category_sales = category_sales.sort_values(ascending=False)

print("\nCategory-wise Sales:")
print(category_sales)

# Step 22: Category-wise Quantity Analysis

category_quantity = df.groupby("Category")["Quantity"].sum()

category_quantity = category_quantity.sort_values(ascending=False)

print("\nCategory-wise Quantity Sold:")
print(category_quantity)

# Step 23: Top 5 Products by Sales

top_5_products = product_sales.head(5)

print("\nTop 5 Products by Sales:")
print(top_5_products)

# Step 24: Save Analysis Results

product_sales.to_csv("product_sales.csv")
region_sales.to_csv("region_sales.csv")
category_sales.to_csv("category_sales.csv")
category_quantity.to_csv("category_quantity.csv")
top_5_products.to_csv("top_5_products.csv")

print("\nAnalysis results saved successfully!")


# Step 27: Product-wise Sales Graph

plt.figure(figsize=(10, 6))

product_sales.plot(kind="bar")

plt.title("Product-wise Sales Analysis")
plt.xlabel("Product")
plt.ylabel("Total Sales")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig("product_sales_chart.png")
plt.show()


# Step 28: Region-wise Sales Graph

plt.figure(figsize=(8, 5))

region_sales.plot(kind="bar")

plt.title("Region-wise Sales Analysis")
plt.xlabel("Region")
plt.ylabel("Total Sales")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig("region_sales_chart.png")
plt.show()


# Step 29: Category-wise Sales Pie Chart

plt.figure(figsize=(8, 8))

category_sales.plot(
    kind="pie",
    autopct="%1.1f%%",
    startangle=90
)

plt.title("Category-wise Sales Distribution")
plt.ylabel("")
plt.tight_layout()

plt.savefig("category_sales_chart.png")
plt.show()

# Step 31: Export Analysis Results to Excel

with pd.ExcelWriter("Sales_Analysis_Report.xlsx", engine="openpyxl") as writer:

    # Sheet 1: Cleaned Dataset
    df.to_excel(writer, sheet_name="Sales Data", index=False)

    # Sheet 2: Product-wise Sales
    product_sales.to_frame("Total Sales").to_excel(
        writer, sheet_name="Product Sales"
    )

    # Sheet 3: Region-wise Sales
    region_sales.to_frame("Total Sales").to_excel(
        writer, sheet_name="Region Sales"
    )

    # Sheet 4: Category-wise Sales
    category_sales.to_frame("Total Sales").to_excel(
        writer, sheet_name="Category Sales"
    )

    # Sheet 5: Category-wise Quantity
    category_quantity.to_frame("Total Quantity").to_excel(
        writer, sheet_name="Category Quantity"
    )

    # Sheet 6: Top 5 Products
    top_5_products.to_frame("Total Sales").to_excel(
        writer, sheet_name="Top 5 Products"
    )

print("\nExcel Report Created Successfully!")




# Step 34: Project Summary Dashboard

total_orders = df["OrderID"].nunique()
total_quantity = df["Quantity"].sum()
total_sales = df["Sales"].sum()
average_sales = df["Sales"].mean()

top_product = product_sales.idxmax()
top_region = region_sales.idxmax()
top_category = category_sales.idxmax()

summary = pd.DataFrame({
    "Metric": [
        "Total Orders",
        "Total Quantity Sold",
        "Total Sales",
        "Average Sales Per Row",
        "Top Product",
        "Highest-Sales Region",
        "Top Category by Sales"
    ],
    "Value": [
        total_orders,
        total_quantity,
        round(total_sales, 2),
        round(average_sales, 2),
        top_product,
        top_region,
        top_category
    ]
})

print("\n===== RETAIL SALES SUMMARY DASHBOARD =====")
print(summary.to_string(index=False))

# Add Summary sheet to the existing Excel workbook
with pd.ExcelWriter(
    "Sales_Analysis_Report.xlsx",
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)

print("\nSummary sheet added to Excel report!")

