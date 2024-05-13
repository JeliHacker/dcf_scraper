from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, WebDriverException
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import time


def create_driver():
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    return driver


def scrape_data(driver, ticker: str):
    """
    :param driver: the selenium webdriver
    :param ticker: the company's ticker symbol
    :return: a list [margin of safety, fair value, predictability stars]
    """
    driver.get(f"https://www.gurufocus.com/stock/{ticker}/dcf")

    # Wait for the page to load sufficiently
    try:
        print(f"trying... {ticker}")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'dcf-result')))
        print("done waiting, got past webdriverwait")
    except TimeoutException:
        print(f"Timed out waiting for the page for {ticker} to load.")
        # Optionally, take a screenshot or dump the page source to help with debugging.
        # driver.save_screenshot('page_timeout.png')
        # with open('page_source.html', 'w') as f:
        #    f.write(driver.page_source)
        # Handle the exception or re-raise it if you want the script to stop here.
        return [-69420, -69420, -69420]

    percentage_text = 'Blank'
    predictability_stars = 0
    fair_value = 0

    # Implement a retry mechanism for BeautifulSoup
    attempts = 0
    timed_out = True
    while attempts < 3:
        attempts += 1
        try:
            soup = BeautifulSoup(driver.page_source, 'html.parser')

        except WebDriverException as e:
            print("WebDriverException encountered:", e)
            return
            driver.quit()  # Quit the crashed driver
            time.sleep(5)  # Wait a bit before restarting
            driver = create_driver()  # Create a new driver instance
            continue

        result_span = soup.find('span', class_='dcf-result')

        # Find the outermost div that contains the stars.
        outer_div = soup.find('div', class_='el-rate')

        if not result_span or not outer_div:
            continue

        # Get all elements with the class 'el-icon-star-on' within the outer div.
        stars = outer_div.find_all('i', class_='el-icon-star-on')

        # Filter full stars with the exact color style.
        full_stars = [star for star in stars if
                      'color:#F7BA2A;' in star.get('style', '') and not ('width:50%' in star.get('style', ''))]

        # Filter half stars with both the color and the width style.
        half_stars = [star for star in stars if
                      'color:#F7BA2A;' in star.get('style', '') and 'width:50%' in star.get('style', '')]

        predictability_stars = len(full_stars) + (0.5 * len(half_stars))

        price_div = soup.find('div', class_='fs-x-large fw-bolder text-right el-col el-col-12')
        fair_value = price_div.get_text(strip=True)

        if result_span:
            percentage_text = result_span.text.strip()
            timed_out = False
            break
        else:
            print(f"Retry {attempts + 1} for {ticker}")
            time.sleep(2)  # Wait a bit before retrying

    if timed_out:
        print(f"Failed to find margin of safety for {ticker} after {attempts} attempts.")
        return [percentage_text, predictability_stars, fair_value]

    if percentage_text == 'N/A':
        percentage_value = 'N/A'
    else:
        percentage_value = float(percentage_text.replace('%', '').replace(',', ''))

    return [percentage_value, predictability_stars, fair_value]


def scrape_all_stocks(input_file):
    """
    Writes to the stocks list (which is the input file), adding 3 columns relating to fair value,
    based on data scraped from Gurufocus.
    :param input_file: What file to read and write. This should be the relevant stocks list.
    :return:
    """

    # Load the DataFrame
    df = pd.read_excel(input_file, sheet_name='Stocks')
    # Ensure the DataFrame has the necessary columns
    if 'Margin of Safety' not in df.columns:
        df['Margin of Safety'] = None
    if 'Business Predictability' not in df.columns:
        df['Business Predictability'] = None
    if 'Fair Value' not in df.columns:
        df['Fair Value'] = None

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    processed_count = 0
    try:
        df['Symbol'] = df['Symbol'].str.replace('/', '.', regex=False)
        df['Symbol'] = df['Symbol'].str.replace('^', '.', regex=False)

        for index, row in df.iterrows():
            # Skip rows that have already been processed
            if pd.notna(row['Fair Value']):
                continue

            data = scrape_data(driver, row['Symbol'])
            print(data)
            if data:
                mos = data[0]
                if mos == 'N/A':
                    mos = np.nan

                stars = data[1]
                if stars == 'N/A':
                    stars = np.nan

                fair_value = data[2]
                if fair_value == 'N/A':
                    fair_value = np.nan

                df.at[index, 'Margin of Safety'] = mos
                df.at[index, 'Business Predictability'] = stars
                df.at[index, 'Fair Value'] = fair_value

            processed_count += 1

            if processed_count >= 50:
                break
    finally:
        driver.quit()

        df.to_excel(input_file, sheet_name='Stocks', index=False)


if __name__ == "__main__":
    scrape_all_stocks()
