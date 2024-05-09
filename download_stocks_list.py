from selenium.webdriver.common.by import By
import pandas as pd
import time
from main import create_driver
import shutil
import os
from glob import glob


def download_nasdaq_list():
    driver = create_driver()
    driver.get("https://www.nasdaq.com/market-activity/stocks/screener")
    time.sleep(5)
    download_button = driver.find_element(By.CLASS_NAME, 'nasdaq-screener__form-button--download')
    download_button.click()
    time.sleep(5)

    # Path to your Downloads folder
    downloads_folder = '/Users/eligooch/Downloads'

    # Path to your project directory
    project_directory = '/Users/eligooch/PycharmProjects/dcf_scraper'

    # List all files in the Downloads folder
    files = glob(os.path.join(downloads_folder, '*'))

    # Find the most recently downloaded file (based on modification time)
    latest_file = max(files, key=os.path.getmtime)

    # Move the most recently downloaded file to your project directory
    shutil.move(latest_file, os.path.join(project_directory, "stocks_list.csv"))

    df = pd.read_csv("stocks_list.csv")

    excel_file_path = "stocks_list.xlsx"
    df.to_excel(excel_file_path, index=False, engine="openpyxl", sheet_name="Stocks")


download_nasdaq_list()