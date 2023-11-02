from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import csv
import pandas as pd


#Define the soup function
def get_html(driver, symbol, parameters):
    url = f"https://www.reuters.com/markets/companies/{symbol}.N/key-metrics/{parameters}"
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    return soup

#Make the data convertable
def get_data_as_float(soup, query):
    th = soup.find("th", string=query)
    if th is not None:
        span = th.next_sibling.span
        if span is not None:
            value_text = span.text
            try:
                value = float(value_text.replace(",", ""))
                return value
            except ValueError:
                return None
    return None  # Return None if the query or value is not found


def extract_data(driver, symbol, data_type, search_queries):
    soup = get_html(driver, symbol, data_type)
    symbol_data = {}

    for query in search_queries:
        value = get_data_as_float(soup, query)
        symbol_data[query] = value if value is not None else "Not found"

    return symbol_data

def main():
    # Getting stock symbols
    def get_valid_symbol(value):
        while True:
            symbol = input(value).upper()  # Convert to uppercase
            if any(char.isdigit() for char in symbol):  # Check if any character in the symbol is a digit
                print("Error: Symbol must not contain numbers.")
            elif len(symbol) > 5:
                print("Error: Symbol must be 5 characters or less.")
            else:
                return symbol


    # Getting stock symbols with validation
    first_stock = get_valid_symbol('Enter the first symbol:')
    second_stock = get_valid_symbol('Enter the second symbol:')
    third_stock = get_valid_symbol('Enter the third symbol:')
    fourth_stock = get_valid_symbol('Enter the fourth symbol:')
    fifth_stock = get_valid_symbol('Enter the index fund:')

    symbols = [first_stock, second_stock, third_stock, fourth_stock, fifth_stock]

    search_queries = [
        'Gross Margin (5Y)',
        'Operating Margin (5Y)',
        'Net Profit Margin (5Y)',
        'Free Operating Cash Flow/Revenue (5Y)'
    ]

    search_queries2 = [
        'P/E Normalized (Annual)',
        'Price To Sales (Annual)',
        'Price To Tangible Book (Annual)',
        'Price To Free Cash Flow (Per Share Annual)',
        'Price To Book (Annual)',
        'Dividend Yield (5Y)'
    ]

    search_queries3 = [
        'Return On Assets (Annual)',
        'Return On Equity (TTM)',
        'Return On Investment (Annual)',
        'Asset Turnover (Annual)',
        'Inventory Turnover (Annual)',
        'Receivables Turnover (Annual)'
    ]

    search_queries4 = [
        'Revenue Growth Rate (5Y)',
        'EPS Growth Rate (5Y)',
        'Dividend Growth Rate (5Y)',
        'Book Value Growth Rate (Per Share 5Y)',
        'Net Profit Margin Growth Rate (5Y)'
    ]

    search_queries5 = [
        'Free Cash Flow (Annual)',
        'Current Ratio (Annual)',
        'Quick Ratio (Annual)',
        'Net Interest Coverage (Annual)',
        'Total Debt/Total Equity (Annual)',
        'Payout Ratio (Annual)'
    ]
    

    # Create a single WebDriver instance and reuse it
    options = Options()
    options.page_load_strategy = 'eager'
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    output_metrics_data = {}
    output_valuation_data = {}
    output_management_data = {}
    output_growth_data = {}
    output_financial_data = {}

    for symbol in symbols:
        print(f"{symbol}:")

        # Extract margins data
        symbol_metrics_data = extract_data(driver, symbol, 'margins', search_queries)
        for query, value in symbol_metrics_data.items():
            print(f"\t{query}: {value}")

        output_metrics_data[symbol] = symbol_metrics_data

        # Extract valuation data
        symbol_valuation_data = extract_data(driver, symbol, 'valuation', search_queries2)
        for query, value in symbol_valuation_data.items():
            print(f"\t{query}: {value}")
        output_valuation_data[symbol] = symbol_valuation_data

        # Extract management effectiveness data
        symbol_management_data = extract_data(driver, symbol, 'management-effectiveness', search_queries3)
        for query, value in symbol_management_data.items():
            print(f"\t{query}: {value}")
        output_management_data[symbol] = symbol_management_data

        #Extract growth data
        symbol_growth_data = extract_data(driver, symbol, 'growth', search_queries4)
        for query, value in symbol_growth_data.items():
            print(f"\t{query}: {value}")
        output_growth_data[symbol] = symbol_growth_data

        #Extract financial-strenght data
        symbol_financial_data = extract_data(driver, symbol, 'financial-strength', search_queries5)
        for query, value in symbol_financial_data.items():
            print(f"\t{query}: {value}")
        output_financial_data[symbol] = symbol_financial_data


    # Close the WebDriver instance
    driver.quit()

    valuation_df = pd.DataFrame.from_dict(output_valuation_data, orient='index', columns=search_queries2)
    metrics_df = pd.DataFrame.from_dict(output_metrics_data, orient='index', columns=search_queries)
    management_df = pd.DataFrame.from_dict(output_management_data, orient='index', columns=search_queries3)
    growth_df = pd.DataFrame.from_dict(output_growth_data, orient='index', columns=search_queries4)
    financial_df = pd.DataFrame.from_dict(output_financial_data, orient='index', columns=search_queries5)

    # Create a Pandas ExcelWriter object to write to a single Excel file
    with pd.ExcelWriter("output_data.xlsx", engine="xlsxwriter") as writer:
        # Write each DataFrame to a separate sheet in the same Excel file
        valuation_df.to_excel(writer, sheet_name= "Valuation Data", index_label="Symbol")
        metrics_df.to_excel(writer, sheet_name= "Margin Data", index_label="Symbol")
        management_df.to_excel(writer, sheet_name= "Management Data", index_label="Symbol")
        growth_df.to_excel(writer, sheet_name = "Growth Data", index_label = "Symbol" )
        financial_df.to_excel(writer, sheet_name = "Financial Strength Data", index_label = "Symbol")
        
    
    

if __name__ == "__main__":
    main()