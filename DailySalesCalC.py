def calculate_sales_for_store(store_name):
    print(f"\nEnter sales details for store: {store_name}")
    
    # Input Global and UK sales amounts
    global_sales = float(input("Enter Global Sales amount (£): "))
    uk_sales = float(input("Enter UK Sales amount (£): "))

    # Step 1: Calculate Euro Sales Grand
    euro_sales_grand = global_sales - uk_sales

    # Step 2: Calculate Euro NAT Sale (Euro Sales Grand / 120%)
    euro_nat_sale = euro_sales_grand / 1.2  # Divide by 1.2 to get the NAT sale (100%)

    # Step 3: Calculate Euro VAT (Euro Sales Grand - Euro NAT Sale)
    euro_vat = euro_sales_grand - euro_nat_sale

    # Step 4: Calculate Total Sales Net (UK Sales + Euro NAT Sale)
    total_sales_net = uk_sales + euro_nat_sale

    # Create result string
    results = (
        f"--- Sales Breakdown for {store_name} ---\n"
        f"Global Sales        : £{global_sales:,.2f}\n"
        f"UK Sales            : £{uk_sales:,.2f}\n"
        f"Euro Sales Grand    : £{euro_sales_grand:,.2f}\n"
        f"Euro NAT Sale       : £{euro_nat_sale:,.2f}\n"
        f"Euro VAT            : £{euro_vat:,.2f}\n"
        f"Total Sales Net     : £{total_sales_net:,.2f}\n"
        "-------------------------------\n"
    )
    print("\n" + results)

    # Return values for aggregation
    return global_sales, euro_vat, total_sales_net, results

def main():
    # List of store names
    stores = ["NW", "SP", "JM", "KC"]

    # Initialize totals
    total_gross_sales = 0
    total_vat = 0
    total_net_sales = 0

    # Initialize the file
    file_name = "C:/All Sales/Sales/sales_breakdown_all_stores.txt"
    with open(file_name, "w") as file:
        file.write("Sales Breakdown for All Stores\n")
        file.write("===============================\n")

    # Loop through each store and calculate sales
    for store in stores:
        global_sales, euro_vat, total_sales_net, store_results = calculate_sales_for_store(store)
        
        # Update totals
        total_gross_sales += global_sales
        total_vat += euro_vat
        total_net_sales += total_sales_net
        
        # Append results to the file
        with open(file_name, "a") as file:
            file.write(store_results)

    # Append total calculations to the file
    total_results = (
        "\n--- Totals for All Stores ---\n"
        f"Total Gross Sales   : £{total_gross_sales:,.2f}\n"
        f"Total VAT           : £{total_vat:,.2f}\n"
        f"Total Net Sales     : £{total_net_sales:,.2f}\n"
        "===============================\n"
    )
    print(total_results)
    with open(file_name, "a") as file:
        file.write(total_results)

    print(f"Results for all stores and totals have been saved to {file_name}")

# Run the script
main()
