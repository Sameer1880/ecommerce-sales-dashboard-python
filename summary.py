def print_project_summary(data, monthly_sales, category_sales):
    total_sales = data['SalesAmount'].sum()
    total_profit = data['Profit'].sum()
    avg_order_value = data['SalesAmount'].mean()
    total_customers = data['CustomerID'].nunique()
    print("\n📋 PROJECT SUMMARY")
    print(f"Total Sales: ${total_sales:,.2f}")
    print(f"Total Profit: ${total_profit:,.2f}")
    print(f"Average Order Value: ${avg_order_value:.2f}")
    print(f"Number of Orders: {len(data)}")
    print(f"Unique Customers: {total_customers}")
    print(f"Top Category: {category_sales.iloc[0]['ProductCategory']} with ${category_sales.iloc[0]['SalesAmount']:,.2f}")

# -------------------------------
# 6. MAIN FUNCTION
# -------------------------------
def main():
    data = generate_sample_data()
    wb, monthly_sales, category_sales = create_excel_dashboard(data)
    add_charts_to_excel(wb, wb['Analysis'], wb['Dashboard Summary'], monthly_sales, category_sales)
    create_portfolio_visualizations(monthly_sales, category_sales)
    print_project_summary(data, monthly_sales, category_sales)

if __name__ == "__main__":
    main()