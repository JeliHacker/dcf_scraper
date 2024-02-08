from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time


def scrape_margin_of_safety(driver, ticker: str):
    driver.get(f"https://www.gurufocus.com/stock/{ticker}/dcf")

    # Wait for the page to load sufficiently
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'dcf-result')))

    # Implement a retry mechanism for BeautifulSoup
    attempts = 0
    while attempts < 3:
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        result_span = soup.find('span', class_='dcf-result')
        if result_span:
            percentage_text = result_span.text.strip()
            break
        else:
            print(f"Retry {attempts + 1} for {ticker}")
            time.sleep(2)  # Wait a bit before retrying
            attempts += 1

    if not result_span:
        print(f"Failed to find margin of safety for {ticker} after {attempts} attempts.")
        return None  # or a default value, or raise an exception

    if percentage_text == 'N/A':
        return 'N/A'

    percentage_value = float(percentage_text.replace('%', '').replace(',', ''))
    return percentage_value


def main():
    df = pd.read_excel('Russell 2000 Stocks List.xlsx', sheet_name='Characteristics')
    df['Margin of Safety'] = None
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        for index, row in df.head(3).iterrows():
            mos = scrape_margin_of_safety(driver, row['Ticker'])
            df.at[index, 'Margin of Safety'] = mos
    finally:
        driver.quit()

    df.to_excel('Russell 2000 DCFs.xlsx', sheet_name='Stocks', index=False)


if __name__ == "__main__":
    main()
