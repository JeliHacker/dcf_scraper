from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import pandas as pd


def get_tickers() -> list:
    df = pd.read_excel('Russell 2000 Stocks List.xlsx', sheet_name='Characteristics')

    tickers = df['Ticker']
    tickers = tickers.to_list()

    return tickers


def scrape_margin_of_safety(ticker: str) -> float:
    service = Service(ChromeDriverManager().install())

    # Setup Selenium WebDriver
    driver = webdriver.Chrome(service=service)

    # URL you want to scrape
    url = "https://www.gurufocus.com/"

    # Use Selenium to get the page
    driver.get(f"https://www.gurufocus.com/stock/{ticker}/dcf")

    # Wait a few seconds for JavaScript to load content
    time.sleep(5)

    # Now you can use BeautifulSoup to parse the rendered HTML
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    result_span = soup.find('span', class_='dcf-result')
    print(result_span, bool(result_span))

    percentage_text = result_span.text.strip()

    if percentage_text == 'N/A':
        return -69

    percentage_value = float(percentage_text.replace('%', '').replace(',', ''))

    # Don't forget to close the browser
    driver.quit()

    return percentage_value


def main():
    df = pd.read_excel('Russell 2000 Stocks List.xlsx', sheet_name='Characteristics')
    df['Margin of Safety'] = None

    for index, row in df.head(3).iterrows():
        # Calculate the margin of safety for each ticker
        mos = scrape_margin_of_safety(row['Ticker'])
        # Assign the result to the 'Margin of Safety' column
        df.at[index, 'Margin of Safety'] = mos

    # Save the updated DataFrame to a new Excel file
    df.to_excel('Russell 2000 DCFs.xlsx', sheet_name='Stocks', index=False)


main()