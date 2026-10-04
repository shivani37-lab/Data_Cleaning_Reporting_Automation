# ============================================================
# DATA CLEANING & REPORTING AUTOMATION
# ============================================================

import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ------------------------------------------------------------
# 1. SETUP
# ------------------------------------------------------------

INPUT_FILE = "data/sales_data.csv"
OUTPUT_FOLDER = "output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

print("=" * 60)
print("DATA CLEANING & REPORTING AUTOMATION")
print("=" * 60)

# ------------------------------------------------------------
# 2. LOAD DATA
# ------------------------------------------------------------

df = pd.read_csv(INPUT_FILE)

print("\n[1] DATA LOADED")
print("-" * 60)
print(f"Rows    : {df.shape[0]}")
print(f"Columns : {df.shape[1]}")

# ------------------------------------------------------------
# 3. INITIAL DATA QUALITY CHECK
# ------------------------------------------------------------

print("\n[2] INITIAL DATA QUALITY CHECK")
print("-" * 60)

print("\nMissing values:")
print(df.isnull().sum())

print("\nDuplicate rows:", df.duplicated().sum())

# ------------------------------------------------------------
# 4. STANDARDIZE TEXT DATA
# ------------------------------------------------------------

print("\n[3] STANDARDIZING TEXT DATA")
print("-" * 60)

df["Region"] = df["Region"].astype(str).str.strip().str.title()
df["Category"] = df["Category"].astype(str).str.strip().str.title()
df["Customer"] = df["Customer"].astype(str).str.strip()
df["Product"] = df["Product"].astype(str).str.strip()

print("Text formatting standardized.")

# ------------------------------------------------------------
# 5. CONVERT DATA TYPES
# ------------------------------------------------------------

df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

numeric_columns = ["Quantity", "Unit_Price", "Sales"]

for column in numeric_columns:
    df[column] = pd.to_numeric(df[column], errors="coerce")

# ------------------------------------------------------------
# 6. HANDLE MISSING VALUES
# ------------------------------------------------------------

print("\n[4] HANDLING MISSING VALUES")
print("-" * 60)

# Fill Quantity using median
if df["Quantity"].isnull().sum() > 0:
    df["Quantity"] = df["Quantity"].fillna(df["Quantity"].median())

# Fill Unit Price using median
if df["Unit_Price"].isnull().sum() > 0:
    df["Unit_Price"] = df["Unit_Price"].fillna(df["Unit_Price"].median())

# Fill missing Sales using Quantity × Unit Price
df["Sales"] = df["Sales"].fillna(
    df["Quantity"] * df["Unit_Price"]
)

# Fill missing text values
df["Customer"] = df["Customer"].replace("nan", "Unknown")
df["Product"] = df["Product"].replace("nan", "Unknown")
df["Region"] = df["Region"].replace("Nan", "Unknown")
df["Category"] = df["Category"].replace("Nan", "Unknown")

print("Missing values handled.")

# ------------------------------------------------------------
# 7. REMOVE DUPLICATES
# ------------------------------------------------------------

print("\n[5] REMOVING DUPLICATES")
print("-" * 60)

duplicates_before = df.duplicated().sum()

df = df.drop_duplicates()

duplicates_after = df.duplicated().sum()

print("Duplicates removed:", duplicates_before)
print("Remaining duplicates:", duplicates_after)

# ------------------------------------------------------------
# 8. FIX INVALID VALUES
# ------------------------------------------------------------

print("\n[6] CHECKING INVALID VALUES")
print("-" * 60)

# Remove records with invalid quantity or price
df = df[df["Quantity"] > 0]
df = df[df["Unit_Price"] > 0]

# Recalculate Sales to ensure consistency
df["Sales"] = df["Quantity"] * df["Unit_Price"]

print("Invalid values checked and corrected.")

# ------------------------------------------------------------
# 9. ADD USEFUL COLUMNS
# ------------------------------------------------------------

df["Month"] = df["Date"].dt.month_name()
df["Year"] = df["Date"].dt.year

df["Profit_Estimate"] = df["Sales"] * 0.20

# ------------------------------------------------------------
# 10. SAVE CLEANED DATA
# ------------------------------------------------------------

cleaned_file = os.path.join(
    OUTPUT_FOLDER,
    "cleaned_sales_data.xlsx"
)

df.to_excel(cleaned_file, index=False)

print("\n[7] CLEANED DATA SAVED")
print("-" * 60)
print(cleaned_file)

# ------------------------------------------------------------
# 11. SALES SUMMARY
# ------------------------------------------------------------

total_sales = df["Sales"].sum()
total_orders = df["Order_ID"].nunique()
total_quantity = df["Quantity"].sum()
average_order_value = total_sales / total_orders
estimated_profit = df["Profit_Estimate"].sum()

summary = pd.DataFrame({
    "Metric": [
        "Total Sales",
        "Total Orders",
        "Total Quantity Sold",
        "Average Order Value",
        "Estimated Profit"
    ],
    "Value": [
        total_sales,
        total_orders,
        total_quantity,
        average_order_value,
        estimated_profit
    ]
})

# ------------------------------------------------------------
# 12. REGION SUMMARY
# ------------------------------------------------------------

region_summary = (
    df.groupby("Region")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Orders=("Order_ID", "nunique"),
        Quantity_Sold=("Quantity", "sum")
    )
    .reset_index()
    .sort_values("Total_Sales", ascending=False)
)

# ------------------------------------------------------------
# 13. CATEGORY SUMMARY
# ------------------------------------------------------------

category_summary = (
    df.groupby("Category")
    .agg(
        Total_Sales=("Sales", "sum"),
        Total_Orders=("Order_ID", "nunique"),
        Quantity_Sold=("Quantity", "sum")
    )
    .reset_index()
    .sort_values("Total_Sales", ascending=False)
)

# ------------------------------------------------------------
# 14. PRODUCT SUMMARY
# ------------------------------------------------------------

product_summary = (
    df.groupby("Product")
    .agg(
        Total_Sales=("Sales", "sum"),
        Quantity_Sold=("Quantity", "sum")
    )
    .reset_index()
    .sort_values("Total_Sales", ascending=False)
)

# ------------------------------------------------------------
# 15. CREATE EXCEL REPORT
# ------------------------------------------------------------

report_file = os.path.join(
    OUTPUT_FOLDER,
    "sales_report.xlsx"
)

with pd.ExcelWriter(report_file, engine="openpyxl") as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    region_summary.to_excel(
        writer,
        sheet_name="Region Analysis",
        index=False
    )

    category_summary.to_excel(
        writer,
        sheet_name="Category Analysis",
        index=False
    )

    product_summary.to_excel(
        writer,
        sheet_name="Product Analysis",
        index=False
    )

    df.to_excel(
        writer,
        sheet_name="Cleaned Data",
        index=False
    )

print("\n[8] EXCEL REPORT CREATED")
print("-" * 60)
print(report_file)

# ------------------------------------------------------------
# 16. CREATE DASHBOARD CHARTS
# ------------------------------------------------------------

sns.set_theme(style="whitegrid")

# -------- Region Sales Chart --------

plt.figure(figsize=(10, 6))

sns.barplot(
    data=region_summary,
    x="Region",
    y="Total_Sales"
)

plt.title("Sales by Region")
plt.xlabel("Region")
plt.ylabel("Total Sales")

plt.tight_layout()

region_chart = os.path.join(
    OUTPUT_FOLDER,
    "sales_by_region.png"
)

plt.savefig(region_chart, dpi=300)
plt.close()

# -------- Category Sales Chart --------

plt.figure(figsize=(10, 6))

sns.barplot(
    data=category_summary,
    x="Category",
    y="Total_Sales"
)

plt.title("Sales by Category")
plt.xlabel("Category")
plt.ylabel("Total Sales")

plt.tight_layout()

category_chart = os.path.join(
    OUTPUT_FOLDER,
    "sales_by_category.png"
)

plt.savefig(category_chart, dpi=300)
plt.close()

# -------- Product Sales Chart --------

plt.figure(figsize=(10, 6))

sns.barplot(
    data=product_summary,
    x="Product",
    y="Total_Sales"
)

plt.title("Sales by Product")
plt.xlabel("Product")
plt.ylabel("Total Sales")

plt.xticks(rotation=30)

plt.tight_layout()

product_chart = os.path.join(
    OUTPUT_FOLDER,
    "sales_by_product.png"
)

plt.savefig(product_chart, dpi=300)
plt.close()

# ------------------------------------------------------------
# 17. MONTHLY SALES TREND
# ------------------------------------------------------------

monthly_sales = (
    df.groupby("Date")["Sales"]
    .sum()
    .reset_index()
    .sort_values("Date")
)

plt.figure(figsize=(12, 6))

plt.plot(
    monthly_sales["Date"],
    monthly_sales["Sales"],
    marker="o"
)

plt.title("Sales Trend Over Time")
plt.xlabel("Date")
plt.ylabel("Sales")

plt.xticks(rotation=45)

plt.tight_layout()

trend_chart = os.path.join(
    OUTPUT_FOLDER,
    "sales_trend.png"
)

plt.savefig(trend_chart, dpi=300)
plt.close()

# ------------------------------------------------------------
# 18. FINAL REPORT
# ------------------------------------------------------------

print("\n[9] FINAL REPORT")
print("=" * 60)

print(f"Total Sales       : ₹{total_sales:,.2f}")
print(f"Total Orders      : {total_orders}")
print(f"Quantity Sold     : {total_quantity:,.0f}")
print(f"Average Order     : ₹{average_order_value:,.2f}")
print(f"Estimated Profit  : ₹{estimated_profit:,.2f}")

print("\nTop Region:")
print(region_summary.iloc[0]["Region"])

print("\nTop Category:")
print(category_summary.iloc[0]["Category"])

print("\nTop Product:")
print(product_summary.iloc[0]["Product"])

# ------------------------------------------------------------
# 19. COMPLETION MESSAGE
# ------------------------------------------------------------

print("\n" + "=" * 60)
print("🎉 AUTOMATION COMPLETED SUCCESSFULLY!")
print("=" * 60)

print("\nGenerated files:")

print("✔ cleaned_sales_data.xlsx")
print("✔ sales_report.xlsx")
print("✔ sales_by_region.png")
print("✔ sales_by_category.png")
print("✔ sales_by_product.png")
print("✔ sales_trend.png")

print("\nAll files are available inside the 'output' folder.")