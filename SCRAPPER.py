import logging
import re
import time
from datetime import datetime, date
import openpyxl
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import os
import shutil

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class APNewsScraper:
    """
    A class to scrape news articles from AP News website within a specified date range.
    """
    def __init__(self, search_phrase, delta,chromedriver_name,base_url,fileName):
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
        self.chromedriver_name=chromedriver_name
        self.fileName=fileName
        self.driver = None
        self.start_date, self.end_date =self.calcDates(delta)
        self.output_path=self.createfileOutput() 


    def __enter__(self):
        """Initializes the ChromeDriver on entering the context."""
        self.driver = webdriver.Chrome(service=webdriver.chrome.service.Service(executable_path=self.setDriverpath()))
        return self


    def __exit__(self, exc_type, exc_val, exc_tb):
        """Quits the ChromeDriver on exiting the context."""
        if self.driver:
            self.driver.quit()

    def setDriverpath(self):
        """Creates and returns the dynamic ChromeDriver Path"""
        script_dir = os.path.dirname(os.path.abspath(__file__)) 
        driverfolder = os.path.join(script_dir, "CONFIG")
        driverPath=os.path.join(driverfolder, self.chromedriver_name)
        if not os.path.exists(driverPath):
            raise FileNotFoundError(f"No such file or directory: '{driverPath}'")
        return driverPath

    def createFolderImages(self):
        """Creates and returns the output path for the downloaded images"""
        script_dir = os.path.dirname(os.path.abspath(__file__)) 
        save_folder = os.path.join(script_dir, "IMAGES")
        if os.path.exists(save_folder):
            shutil.rmtree(save_folder)
            os.makedirs(save_folder)
        else:
            os.makedirs(save_folder)
        return save_folder
    
    def createfileOutput(self):
        """Creates and returns the output path for the Excel file."""
        script_dir = os.path.dirname(os.path.abspath(__file__)) 
        output_folder = os.path.join(script_dir, "OUTPUT")
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)
            os.makedirs(output_folder)
        else:
            os.makedirs(output_folder)
        output_path = os.path.join(output_folder, self.fileName)
        return output_path
    
    @staticmethod
    def calcDates(delta):
        """Calculates the start and end dates based on the given delta."""
        delta=int(delta)
        now = datetime.now()
        current_year = int(now.year)
        current_month = int(now.month)
        if delta==0 or delta==1:
            year=current_year
            month=current_month
        else:
            delta=delta-1
            difference=current_month-delta
            if difference>0:
                month=difference
                year=current_year
            else:
                residual=delta%12
                quotient=delta//12
                month=current_month-residual
                if month<0:
                    month=12-month
                year=current_year-quotient
        d=date(year, month, 1)
        start_date = d.strftime("%m/%d/%Y")
        today = datetime.today()
        end_date = today.strftime("%m/%d/%Y")
        return start_date, end_date
        

    def run(self):
        """Main method to execute the scraping process."""
        try:
            save_folder=self.createFolderImages()
            self.load_website()
            self.close_popup()
            self.search_news()
            self.close_popup()
            self.orderPageFromNewest()
            self.close_popup()
            news_data = self.scrape_news_articles(save_folder)
            self.close_popup()
            self.save_to_excel(news_data)
            logging.info("Scrapper Robot ran successfully")

        except Exception as e:
            logging.error(f"An error occurred: {e}")

    def load_website(self):
        """Loads the AP News website."""
        try:
            self.driver.get(self.base_url)
            logging.info(f"Website loaded successfully: {self.base_url}")
        except Exception as e:
            logging.error(f"Failed to load website: {e}")
            raise


    def close_popup(self):
        """Closes the pop-up window if present."""
        try:
            close_button = WebDriverWait(self.driver, 1).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "fancybox-close"))
            )
            close_button.click()
            time.sleep(1)
            logging.info("Pop-up closed.")
        except:
            logging.info("No pop-up found.")


    def search_news(self):
        """Enters the search phrase and waits for the results to load."""
        try:
            search_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "SearchOverlay-search-button"))
            )
            self.close_popup()
            search_field.click()
            self.close_popup()
            search_tab = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "SearchOverlay-search-input"))
            )
            self.close_popup()
            search_tab.click()
            self.close_popup()
            search_tab.send_keys(self.search_phrase)
            self.close_popup()
            search_tab.submit()
            logging.info(f"Search initiated for phrase: {self.search_phrase}")
            element_present = EC.presence_of_element_located((By.CLASS_NAME, 'SearchResultsModule-count-desktop'))
            WebDriverWait(self.driver, 10).until(element_present)
            elements = self.driver.find_elements(By.CLASS_NAME, 'SearchResultsModule-count-desktop')
            if not elements:
                raise Exception('Search Returned empty')    

        except Exception as e:
            logging.error(f"Failed to search for news: {e}")
            raise
    

    def orderPageFromNewest(self):
        """Orders the search results by newest articles."""
        try:
            currentURL=self.driver.current_url
            newestURL=currentURL.replace("#nt=navsearch","&s=3")
            self.driver.get(newestURL)
            logging.info("Ordered by Newest Articles.")
        except Exception as e:
            logging.error(f"Failed to load website by Newest Articles: {e}")
            raise

    @staticmethod
    def convertToDate(timestamp_ms):
        """Converts a timestamp in milliseconds to a formatted date string."""
        timestamp_s = timestamp_ms / 1000
        dt_object = datetime.fromtimestamp(timestamp_s)
        formatted_date = dt_object.strftime("%m/%d/%Y")
        return formatted_date


    def scrape_news_articles(self,save_folder):
        """Scrapes data from each news article within the specified date range."""
        news_data = []
        while True:
            titles=[]
            descriptions=[]
            dates=[]
            save_paths=[]
            article_elements = self.driver.find_elements(By.XPATH, '//div[@class="SearchResultsModule-results"]//div[@class="PageList-items-item"]/div[@class="PagePromo"]')
            for article_element in article_elements:
                self.close_popup()
                try:
                    timestamp_element = article_element.find_element(By.CSS_SELECTOR, 'bsp-timestamp[data-timestamp]')
                    date = self.convertToDate(int(timestamp_element.get_attribute('data-timestamp')))
                    if self.start_date<=date<=self.end_date:
                        dates.append(date)
                    else:
                        continue
                except:
                    raise Exception('The date Reference could not be found. Please check the HTML anchors as webpage might have changed') 

                try:
                    title_element = article_element.find_element(By.CLASS_NAME, 'PagePromo-title')
                    title_element=title_element.text
                except:
                    title_element='N/A'
                finally:
                    titles.append(title_element)
                
                try:
                    description_element = article_element.find_element(By.CLASS_NAME, 'PagePromo-description')
                    description_element=description_element.text
                except:
                    description_element='N/A'
                finally:
                    descriptions.append(description_element)

                try:
                    image_element=article_element.find_element(By.CLASS_NAME, 'Image')
                    imagePath = self.download_image(str(image_element.get_attribute("src")),save_folder)
                except:
                    imagePath='N/A'
                finally:
                    save_paths.append(imagePath)
            
            for title, description, date, save_path in zip(titles, descriptions, dates,save_paths):
                news_data.append({
                    'title': title,
                    'date': date,
                    'description': description,
                    'picture_filename': save_path,
                    'search_phrase_count': self.count_search_phrase(title, description),
                    'money_mention': self.detect_money(title, description)
                })

            next_page = self.driver.find_element(By.CLASS_NAME, 'Pagination-nextPage')
            if len(dates)==0:
                break
            else:
                if next_page:
                    next_page.click()
                    time.sleep(10)
                else:
                    break 
        return news_data

    @staticmethod
    def download_image(image_url,save_folder):
        """Downloads the image and saves it to an IMAGES folder in the script's directory."""
        try:
            response = requests.get(image_url)
            response.raise_for_status()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"image_{timestamp}.jpg"
            save_path = os.path.join(save_folder, filename)

            with open(save_path, 'wb') as f:
                f.write(response.content)
            logging.info(f"Image downloaded successfully: {save_path}")
            return save_path 

        except Exception as e:
            logging.error(f"Failed to download image: {e}")
            return "N/A"

    def count_search_phrase(self, title, description):
        """Counts the occurrences of the search phrase in the title and description."""
        return title.lower().count(self.search_phrase.lower()) + description.lower().count(self.search_phrase.lower())

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
                sheet.append([data['title'], data['date'], data['description'], data['picture_filename'],
                             data['search_phrase_count'], data['money_mention']])
            wb.save(self.output_path)
            logging.info(f"Data saved to Excel file: {self.output_path}")
        except Exception as e:
            logging.error(f"Failed to save data to Excel: {e}")

if __name__ == "__main__":

    DELTA = 1
    SEARCH_PHRASE = "OpenAI"
    CHROMEDRIVER_NAME = "chromedriver.exe"
    BASE_URL = "https://apnews.com"
    FILE_NAME="ap_news_data.xlsx"
    with APNewsScraper(SEARCH_PHRASE, DELTA, CHROMEDRIVER_NAME,BASE_URL,FILE_NAME) as scraper:
        scraper.run()