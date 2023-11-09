import tkinter as tk
from tkinter import messagebox
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import threading

# Define the soup function
def get_html(driver, symbol, parameters):
    url = f"https://www.reuters.com/markets/companies/{symbol}.N/key-metrics/{parameters}"
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    return soup

# Make the data convertible
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

# Extract data
def extract_data(driver, symbol, data_type, search_queries):
    soup = get_html(driver, symbol, data_type)
    symbol_data = {}

    for query in search_queries:
        value = get_data_as_float(soup, query)
        symbol_data[query] = value if value is not None else "Not found"

    return symbol_data

# The function to be run by the scraping thread
def run_scraping(symbols, progress_label, callback=None):
    # Define all your search queries
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

    # Initialize data dictionaries for different categories
    output_valuation_data = {}
    output_metrics_data = {}
    output_management_data = {}
    output_growth_data = {}
    output_financial_data = {}

    try:
        for symbol in symbols:
            progress_label.config(text=f"Processing {symbol}")
            
            # Extract and store data for each category
            output_metrics_data[symbol] = extract_data(driver, symbol, 'margins', search_queries)
            output_valuation_data[symbol] = extract_data(driver, symbol, 'valuation', search_queries2)
            output_management_data[symbol] = extract_data(driver, symbol, 'management-effectiveness', search_queries3)
            output_growth_data[symbol] = extract_data(driver, symbol, 'growth', search_queries4)
            output_financial_data[symbol] = extract_data(driver, symbol, 'financial-strength', search_queries5)

        # After extraction, save the data to Excel
        with pd.ExcelWriter("output_data.xlsx", engine="xlsxwriter") as writer:
            pd.DataFrame.from_dict(output_valuation_data, orient='index', columns=search_queries2).to_excel(writer, sheet_name="Valuation Data", index_label="Symbol")
            pd.DataFrame.from_dict(output_metrics_data, orient='index', columns=search_queries).to_excel(writer, sheet_name="Margin Data", index_label="Symbol")
            pd.DataFrame.from_dict(output_management_data, orient='index', columns=search_queries3).to_excel(writer, sheet_name="Management Data", index_label="Symbol")
            pd.DataFrame.from_dict(output_growth_data, orient='index', columns=search_queries4).to_excel(writer, sheet_name="Growth Data", index_label="Symbol")
            pd.DataFrame.from_dict(output_financial_data, orient='index', columns=search_queries5).to_excel(writer, sheet_name="Financial Strength Data", index_label="Symbol")

        progress_label.config(text="Scraping finished. Check the Excel file.")

        # If a callback is provided, call it after the scraping is complete
        if callback:
            callback()

    except Exception as e:
        messagebox.showerror("Error", str(e))
    
    finally:
        driver.quit()

# Function to start the scraping in a thread and handle GUI elements
def start_scraping(entries, progress_label):
    symbols = [entry.get().upper() for entry in entries if entry.get()]
    if not symbols:
        messagebox.showwarning("Warning", "Please enter at least one symbol.")
        return

    # Run scraping in a separate thread to prevent freezing of GUI
    scraping_thread = threading.Thread(target=run_scraping, args=(symbols, progress_label))
    scraping_thread.daemon = True
    scraping_thread.start()

# GUI setup
def setup_gui():
    root = tk.Tk()
    root.title("Stock Data Scraper")

    instructions = tk.Label(root, text="Enter up to five stock symbols:")
    instructions.grid(row=0, columnspan=2)

    entries = []
    for i in range(5):
        entry = tk.Entry(root)
        entry.grid(row=i+1, column=1, padx=10, pady=5)
        entries.append(entry)

    progress_label = tk.Label(root, text="")
    progress_label.grid(row=7, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

    scrape_button = tk.Button(root, text="Scrape Data", command=lambda: start_scraping(entries, progress_label))
    scrape_button.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

    root.mainloop()

# Start the GUI application
if __name__ == "__main__":
    setup_gui()
