import logging
import re
import time
from datetime import datetime, date
import os
import shutil

import openpyxl
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler = logging.StreamHandler()
handler.setFormatter(formatter)
logger.addHandler(handler)

class APNewsScraper:
    """
    A class to scrape news articles from AP News website within a specified date range.
    """

    def __init__(self, search_phrase, delta, chromedriver_path, base_url, file_name):
        """
        Initializes the APNewsScraper with provided parameters.

        Args:
            search_phrase (str): The phrase to search for in AP News.
            delta (int): Number of months prior to current month as the start date.
            chromedriver_path (str): Path to the ChromeDriver executable.
            base_url (str): Base URL of the AP News website.
            file_name (str): Name of the output Excel file.
        """
        self.search_phrase = search_phrase
        self.base_url = base_url
        self.chromedriver_path = chromedriver_path
        self.file_name = file_name
        self.driver = None
        self.start_date, self.end_date = self.calc_dates(delta)
        self.output_path = self.get_output_path()

    def __enter__(self):
        """Initializes the ChromeDriver on entering the context."""
        self.driver = webdriver.Chrome(service=webdriver.chrome.service.Service(executable_path=self.chromedriver_path))
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Quits the ChromeDriver on exiting the context."""
        if self.driver:
            self.driver.quit()

    def get_output_path(self):
        """Creates and returns the output path for the Excel file."""
        output_folder = os.path.join(os.getcwd(), "OUTPUT")
        os.makedirs(output_folder, exist_ok=True)
        return os.path.join(output_folder, self.file_name)

    @staticmethod
    def calc_dates(delta):
        """Calculates the start and end dates based on the given delta."""
        now = datetime.now()
        current_year = int(now.year)
        current_month = int(now.month)

        if delta <= 1:
            year = current_year
            month = current_month
        else:
            delta -= 1
            difference = current_month - delta
            if difference > 0:
                month = difference
                year = current_year
            else:
                residual = delta % 12
                quotient = delta // 12
                month = 12 + difference  # Correcting negative month calculation
                year = current_year - quotient

        start_date = date(year, month, 1).strftime("%m/%d/%Y")
        end_date = now.strftime("%m/%d/%Y")
        return start_date, end_date

    def run(self):
        """Main method to execute the scraping process."""
        try:
            save_folder = self.create_images_folder()
            self.load_website()
            self.close_popup()
            self.search_news()
            self.order_by_newest()
            news_data = self.scrape_news_articles(save_folder)
            self.save_to_excel(news_data)
            logger.info("Scraping completed successfully.")
        except Exception as e:
            logger.error(f"An error occurred: {e}")

    def load_website(self):
        """Loads the AP News website."""
        try:
            self.driver.get(self.base_url)
            logger.info(f"Website loaded successfully: {self.base_url}")
        except Exception as e:
            logger.error(f"Failed to load website: {e}")
            raise

    def close_popup(self):
        """Closes the pop-up window if present."""
        try:
            close_button = WebDriverWait(self.driver, 1).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "fancybox-close"))
            )
            close_button.click()
            time.sleep(1)
            logger.debug("Pop-up closed.")  # Using debug for less verbose logging
        except Exception:
            logger.debug("No pop-up found.")

    def search_news(self):
        """Enters the search phrase and waits for the results to load."""
        try:
            search_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "SearchOverlay-search-button"))
            )
            self.close_popup()
            search_button.click()
            self.close_popup()

            search_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "SearchOverlay-search-input"))
            )
            self.close_popup()
            search_input.click()
            self.close_popup()
            search_input.send_keys(self.search_phrase)
            self.close_popup()
            search_input.submit()

            logger.info(f"Search initiated for phrase: {self.search_phrase}")

            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'SearchResultsModule-count-desktop'))
            )
            if not self.driver.find_elements(By.CLASS_NAME, 'SearchResultsModule-count-desktop'):
                raise Exception('Search returned no results.') 

        except Exception as e:
            logger.error(f"Failed to search for news: {e}")
            raise

    def order_by_newest(self):
        """Orders the search results by newest articles."""
        try:
            current_url = self.driver.current_url
            newest_url = current_url.replace("#nt=navsearch", "&s=3")
            self.driver.get(newest_url)
            logger.debug("Ordered by Newest Articles.")
        except Exception as e:
            logger.error(f"Failed to order by Newest Articles: {e}")
            raise

    @staticmethod
    def convert_to_date(timestamp_ms):
        """Converts a timestamp in milliseconds to a formatted date string."""
        timestamp_s = timestamp_ms / 1000
        dt_object = datetime.fromtimestamp(timestamp_s)
        return dt_object.strftime("%m/%d/%Y")

    def scrape_news_articles(self, save_folder):
        """Scrapes data from each news article within the specified date range."""
        news_data = []

        while True:
            article_elements = self.driver.find_elements(
                By.XPATH, '//div[@class="SearchResultsModule-results"]//div[@class="PageList-items-item"]/div[@class="PagePromo"]'
            )
            for article_element in article_elements:
                self.close_popup()
                try:
                    timestamp_element = article_element.find_element(By.CSS_SELECTOR, 'bsp-timestamp[data-timestamp]')
                    date_str = self.convert_to_date(int(timestamp_element.get_attribute('data-timestamp')))

                    # Date filtering should be done here to avoid unnecessary processing
                    if not (self.start_date <= date_str <= self.end_date):
                        continue

                    title = article_element.find_element(By.CLASS_NAME, 'PagePromo-title').text
                    description = article_element.find_element(By.CLASS_NAME, 'PagePromo-description').text
                    image_url = article_element.find_element(By.CLASS_NAME, 'Image').get_attribute("src")
                    image_path = self.download_image(image_url, save_folder)

                    news_data.append({
                        'title': title,
                        'date': date_str,
                        'description': description,
                        'picture_filename': image_path,
                        'search_phrase_count': self.count_search_phrase(title, description),
                        'money_mention': self.detect_money(title, description)
                    })

                except Exception as e:
                    logger.warning(f"Error scraping article: {e}")  # Using warning for non-critical errors
                    continue

            next_page_element = self.driver.find_element(By.CLASS_NAME, 'Pagination-nextPage')
            if next_page_element.is_enabled():
                next_page_element.click()
                time.sleep(10)
            else:
                break

        return news_data

    @staticmethod
    def create_images_folder():
        """Creates the IMAGES folder if it doesn't exist."""
        save_folder = os.path.join(os.getcwd(), "IMAGES")
        os.makedirs(save_folder, exist_ok=True)
        return save_folder

    def download_image(self, image_url, save_folder):
        """Downloads the image and saves it to the specified folder."""
        try:
            response = requests.get(image_url)
            response.raise_for_status()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"image_{timestamp}.jpg"
            save_path = os.path.join(save_folder, filename)

            with open(save_path, 'wb') as f:
                f.write(response.content)
            logger.debug(f"Image downloaded successfully: {save_path}")  # Using debug for less verbose logging
            return save_path

        except Exception as e:
            logger.warning(f"Failed to download image: {e}")  # Using warning for non-critical errors
            return "N/A"

    def count_search_phrase(self, title, description):
        """Counts the occurrences of the search phrase in the title and description."""
        return title.lower().count(self.search_phrase.lower()) + description.lower().count(
            self.search_phrase.lower()
        )

    @staticmethod
    def detect_money(title, description):
        """Detects the presence of monetary values in the text."""
        text = f"{title} {description}"
        money_pattern = r"\$\d+[\.,]?\d*|\d+[\.,]?\d*\s*(?:dollars|USD)"
        return bool(re.search(money_pattern, text))

    def save_to_excel(self, news_data):
        """Saves the scraped data to an Excel file."""
        try:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(['Title', 'Date', 'Description', 'Picture Filename', 'Search Phrase Count', 'Money Mention'])
            for data in news_data:
                sheet.append(
                    [
                        data['title'],
                        data['date'],
                        data['description'],
                        data['picture_filename'],
                        data['search_phrase_count'],
                        data['money_mention'],
                    ]
                )
            wb.save(self.output_path)
            logger.info(f"Data saved to Excel file: {self.output_path}")
        except Exception as e:
            logger.error(f"Failed to save data to Excel: {e}")


if __name__ == "__main__":
    # Configuration parameters 
    delta_months = 1
    search_term = "OpenAI"
    chromedriver_exe = "chromedriver.exe"  # Update if necessary
    ap_news_url = "https://apnews.com"
    output_file = "ap_news_data.xlsx"

    with APNewsScraper(
        search_term, delta_months, chromedriver_exe, ap_news_url, output_file
    ) as scraper:
        scraper.run()