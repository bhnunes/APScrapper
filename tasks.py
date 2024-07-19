import logging
import re
import time
from datetime import datetime, date
import openpyxl
import requests
from RPA.Browser.Selenium import Selenium
import os
import shutil
from robocorp.tasks import task
import platform
from selenium.webdriver.chrome.options import Options

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class APNewsScraper:
    """
    A class to scrape news articles from AP News website within a specified date range.
    """
    def __init__(self, search_phrase, delta,base_url,fileName):
        """
        Initializes the APNewsScraper with provided parameters.

        Args:
            search_phrase (str): The phrase to search for in AP News.
            delta (int): Number of months prior to current month as the start date.
            base_url (str): Base URL of the AP News website.
            file_name (str): Name of the output Excel file.
        """
        self.search_phrase = search_phrase
        self.base_url = base_url
        self.fileName=fileName
        self.driver = None
        self.start_date, self.end_date =self.calcDates(delta)
        self.output_path=self.createfileOutput()
        self.sleepTime=int(10) 

    def __enter__(self):
        """Initializes the ChromeDriver on entering the context."""
        os_name = platform.system().lower()
        script_dir = os.path.dirname(os.path.abspath(__file__)) 
        chromeDriver_folder = os.path.join(script_dir, "SETUP")
        if os_name == 'windows':
            chrome_win = os.path.join(chromeDriver_folder, "WIN")
            chrome_driver_path=os.path.join(chrome_win, "chromedriver.exe")
            self.driver = Selenium()
            self.driver.open_browser(browser="headlesschrome",executable_path=chrome_driver_path)
            #self.driver.open_browser(browser="chrome",executable_path=chrome_driver_path)
        else:
            chrome_linux = os.path.join(chromeDriver_folder, "LINUX")
            chrome_driver_path = os.path.join(chrome_linux, "chromedriver")
            options = Options()
            options.binary_location = "/usr/bin/chromium"
            options.add_argument("--headless")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--start-maximized")
            self.driver = Selenium()
            self.driver.open_browser(browser="headlesschrome",executable_path=chrome_driver_path, options=options)
        return self


    def __exit__(self, exc_type, exc_val, exc_tb):
        """Quits the ChromeDriver on exiting the context."""
        if self.driver:
            self.driver.close_browser()

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
        output_folder = os.path.join(script_dir, "output")
        output_path = os.path.join(output_folder, self.fileName)
        if os.path.exists(output_path):
            os.remove(output_path)
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
        retry=0
        while retry<3:
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
                return self
            except Exception as e:
                logging.info(f"An error occurred: {e}")
                self.close_popup()
                retry=retry+1
        logging.error(f"An Irreversible error occurred!")
        raise


    def load_website(self):
        """Loads the AP News website."""
        try:
            self.driver.go_to(self.base_url)
            logging.info(f"Website loaded successfully: {self.base_url}")
        except Exception as e:
            logging.error(f"Failed to load website: {e}")
            raise


    def close_popup(self):
        """Closes the pop-up window if present."""
        try:
            self.driver.wait_until_element_is_visible("class:fancybox-close", timeout=1)
            self.driver.click_element("class:fancybox-close")
            time.sleep(1)
            logging.info("Pop-up closed.")
        except:
            logging.info("No pop-up found.")

        try:
            self.driver.wait_until_element_is_visible("xpath://div[@id='onetrust-close-btn-container']/button", timeout=1)
            self.driver.click_element("xpath://div[@id='onetrust-close-btn-container']/button")
            time.sleep(1)
            logging.info("Button closed.")
        except:
            logging.info("Button not found.")

    
    def search_news(self):
        """Enters the search phrase and waits for the results to load."""
        try:
            # self.driver.wait_until_element_is_visible("class:SearchOverlay-search-button", timeout=10)
            # self.close_popup()
            # self.driver.click_element("class:SearchOverlay-search-button")
            # self.close_popup()
            # self.driver.wait_until_element_is_visible("class:SearchOverlay-search-input", timeout=10)
            # self.close_popup()
            # self.driver.click_element("class:SearchOverlay-search-input")
            # self.close_popup()
            # self.driver.input_text("class:SearchOverlay-search-input", self.search_phrase)
            # self.close_popup()
            # self.driver.press_keys("class:SearchOverlay-search-input", "ENTER")
            adjusted_search=str(self.search_phrase)
            new_url=str(self.base_url)+"/search?q="+adjusted_search.replace(" ","+")
            self.driver.go_to(new_url)
            logging.info(f"Search initiated for phrase: {self.search_phrase}")
            # self.driver.wait_until_element_is_visible("class:SearchResultsModule-count-desktop", timeout=self.sleepTime)
            # if self.driver.get_element_count("class:SearchResultsModule-count-desktop") == 0:
            #     raise Exception('Search Returned empty')    

        except Exception as e:
            logging.error(f"Failed to search for news: {e}")
            raise

    def orderPageFromNewest(self):
        """Orders the search results by newest articles."""
        try:
            currentURL=self.driver.get_location()
            newestURL=currentURL+"&s=3"
            #newestURL=currentURL.replace("#nt=navsearch","&s=3")
            self.driver.go_to(newestURL)
            logging.info("Ordered by Newest Articles.")
        except Exception as e:
            logging.error(f"Failed to load website by Newest Articles: {e}")
            raise

    @staticmethod
    def convertToDate(timestamp_ms):
        """Converts a timestamp in milliseconds to a formatted date string."""
        timestamp_ms=int(timestamp_ms)
        timestamp_s = timestamp_ms / 1000
        dt_object = datetime.fromtimestamp(timestamp_s)
        formatted_date = dt_object.strftime("%m/%d/%Y")
        return formatted_date


    def scrape_news_articles(self,save_folder):
        """Scrapes data from each news article within the specified date range."""
        news_data = []
        while True:
            titles, descriptions, dates, save_paths = [], [], [], []
            self.driver.wait_until_element_is_visible("xpath://div[@class='SearchResultsModule-results']", timeout=self.sleepTime)
            article_elements = self.driver.get_webelements(
            "xpath://div[@class='SearchResultsModule-results']//div[@class='PageList-items-item']/div[@class='PagePromo']")
            for article_element in article_elements:
                self.close_popup()
                try:
                    timestamp_element = self.driver.find_element('tag:bsp-timestamp',article_element)
                    date = self.convertToDate(self.driver.get_element_attribute(timestamp_element,'data-timestamp'))
                    if self.start_date <= date <= self.end_date:
                        dates.append(date)
                    else:
                        continue
                except Exception as e:
                    print(f"Error processing date: {e}")
                    raise
                
                try:
                    title_element = self.driver.find_element('class:PagePromo-title',article_element)
                    title=self.driver.get_text(title_element)
                except:
                    title='N/A'
                finally:
                    titles.append(title)
                
                try:
                    description_element = self.driver.find_element('class:PagePromo-description',article_element)
                    description=self.driver.get_text(description_element)
                except:
                    description='N/A'
                finally:
                    descriptions.append(description)

                try:
                    image_element = self.driver.find_element('class:Image',article_element)
                    image_URL=self.driver.get_element_attribute(image_element,'src')
                    imagePath = self.download_image(image_URL,save_folder)
                except:
                    imagePath='N/A'
                finally:
                    save_paths.append(imagePath)
            
            if not dates:
                break

            news_data.extend([{
            'title': title,
            'date': date,
            'description': description,
            'picture_filename': save_path,
            'search_phrase_count': self.count_search_phrase(title, description),
            'money_mention': self.detect_money(title, description)} for title, description, date, save_path in zip(titles, descriptions, dates, save_paths)])

            next_page = self.driver.find_element('class:Pagination-nextPage')
            if next_page:
                self.driver.click_element(next_page)
                time.sleep(self.sleepTime)
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

@task
def runBot():
    DELTA = 1
    SEARCH_PHRASE = "OpenAI"
    BASE_URL = "https://www.apnews.com"
    FILE_NAME="ap_news_data.xlsx"
    with APNewsScraper(SEARCH_PHRASE, DELTA, BASE_URL, FILE_NAME) as scraper:
        scraper.run()