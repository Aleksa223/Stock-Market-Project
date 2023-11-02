# Stock-Market-Project

## Description

The Stock-Market-Project is a Python-based project that scrapes stock information from Reuters and visualizes the data through charts. It performs data extraction using BeautifulSoup and Selenium, stores the scraped data in an Excel file (`output_data.xlsx`), and then creates insightful charts for each sheet within the Excel file to help with data analysis.

## Components

- `scraping.py`: This is the main scraping script that utilizes BeautifulSoup and Selenium to extract stock data from Reuters.
- `charts.py`: This script takes the data stored in `output_data.xlsx` and generates charts for each sheet. It's designed to run after `scraping.py` has populated the Excel file with data.
- `runner.py`: This module is a utility script that runs `scraping.py` and `charts.py` sequentially using the `subprocess` module to streamline the process.

## Getting Started

To get this project up and running on your local machine, follow these instructions.

### Prerequisites

Requirements:

- Python 3
- pip (Python package installer)
- Selenium WebDriver
- BeautifulSoup4
- Openpyxl (for working with Excel files)
- Matplotlib (for creating charts)
- Pandas
- NumPy

You can install the necessary Python packages by running:

```bash
pip install selenium beautifulsoup4 openpyxl matplotlib pandas numpy


