from download_stocks_list import download_nasdaq_list
from scrape_all_stocks import scrape_all_stocks
import datetime


def main():
    print("main")
    current_date = datetime.datetime.now().strftime("%m_%d_%Y")  # Formats date as MM_DD_YYYY
    master_filename = f"stocks_list_{current_date}" # e.g. stocks_list_05_12_2024

    download_nasdaq_list(master_filename)
    scrape_all_stocks(f"{master_filename}.xlsx")


if __name__ == "__main__":
    main()
