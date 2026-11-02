import pandas as pd

# =================================
# SALES DATA CLEANING AND ANALYSIS
# File Name: sales_analysis.py
# =================================

# 1. Load Excel File
df = pd.read_excel("Sales Analysis.xlsx")

print("ORIGINAL SALES DATA")
print(df)

# 2. Check Dataset Information
print("\nDATA INFORMATION")
print(df.info())

# 3. Check Missing Values
print("\nMISSING VALUES")
print(df.isnull().sum())

# 4. Remove Duplicate Records
df = df.drop_duplicates()

# 5. Remove Extra Spaces
for col in df.select_dtypes(include="object").columns:
    df[col] = df[col].str.strip()

# 6. Handle Missing Values
df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
df["Sales"] = df["Sales"].fillna(0)

# 7. Remove Rows with Missing Customer Names
df = df.dropna(subset=["Profit"])

# 8. Filter Sales Above 1000
filtered_data = df[df["Sales"] > 1000]

# 9. Sort Sales in Descending Order
sorted_data = df.sort_values(
    by="Sales",
    ascending=False
)

# 10. Find Top 5 Sales Records
top_5_sales = df.nlargest(5, "Sales")

# 11. Calculate Total Sales
total_sales = df["Sales"].sum()

# 12. Calculate Average Sales
average_sales = df["Sales"].mean()

# 13. Find Maximum Sales
maximum_sales = df["Sales"].max()

# 14. Find Minimum Sales
minimum_sales = df["Sales"].min()

# 15. Count Total Records
total_records = len(df)

# 16. Sales by City
sales_by_country = df.groupby("Country")["Sales"].sum()

# 17. Export Results to Excel
with pd.ExcelWriter("sales_analysis_output.xlsx") as writer:

    df.to_excel(
        writer,
        sheet_name="Cleaned Data",
        index=False
    )

    filtered_data.to_excel(
        writer,
        sheet_name="Filtered Data",
        index=False
    )

    top_5_sales.to_excel(
        writer,
        sheet_name="Top 5 Sales",
        index=False
    )

    sales_by_country.to_excel(
        writer,
        sheet_name="Country Analysis"
    )

# 18. Display Final Results
print("\n========== SALES ANALYSIS ==========")

print("Total Records:", total_records)
print("Total Sales:", total_sales)
print("Average Sales:", average_sales)
print("Maximum Sales:", maximum_sales)
print("Minimum Sales:", minimum_sales)

print("\nTOP 5 SALES")
print(top_5_sales)

print("\nSALES BY COUNTRY")
print(sales_by_country)

print("\nAnalysis Completed Successfully!")
print("Output saved as sales_analysis_output.xlsx")