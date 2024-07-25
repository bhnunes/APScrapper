import logging
import time
from RPA.Browser.Selenium import Selenium
import os
from robocorp.tasks import task
from robocorp import workitems
import platform
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv

from utils import (
    Create_Folder_Images,
    Create_File_Output,
    Calculate_Dates,
    Convert_Timestamp_To_Date,
    Download_Image, 
    Detect_Money, 
    Count_Search_Phrase,
    Create_Zip_File_With_Images, 
    Save_Search_To_Excel
)

load_dotenv()

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class APNewsScraper:
    """
    A class to scrape news articles from AP News website within a specified date range.
    """
    def __init__(self, search_phrase, delta,base_url,fileName, LinuxChromiumPath):
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
        self.start_date, self.end_date =Calculate_Dates(delta)
        self.output_path=Create_File_Output(self.fileName)
        self.sleepTime=int(10)
        self.LinuxChromiumPath=LinuxChromiumPath 

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
        else:
            chrome_linux = os.path.join(chromeDriver_folder, "LINUX")
            chrome_driver_path = os.path.join(chrome_linux, "chromedriver")
            options = Options()
            options.binary_location = self.LinuxChromiumPath
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


    def run(self):
        """Main method to execute the scraping process."""
        retry=0
        error=""
        while retry<3:
            try:
                save_folder_images, zipFilePath = Create_Folder_Images()
                self.load_website()
                self.close_popup()
                self.search_news()
                self.close_popup()
                self.orderPageFromNewest()
                self.close_popup()
                news_data = self.scrape_news_articles(save_folder_images)
                self.close_popup()
                Save_Search_To_Excel(news_data,self.output_path)
                Create_Zip_File_With_Images(save_folder_images,zipFilePath)
                logging.info("Scrapper Robot ran successfully")
                return self
            except Exception as e:
                logging.info(f"An error occurred: {e}")
                self.close_popup()
                retry=retry+1
                error=str(e)
        logging.error(f"An Irreversible error occurred!")
        raise Exception(f"An Irreversible error occurred! - {error}")


    def load_website(self):
        """Loads the AP News website."""
        try:
            self.driver.go_to(self.base_url)
            logging.info(f"Website loaded successfully: {self.base_url}")
        except Exception as e:
            logging.error(f"Failed to load website: {e}")
            raise Exception(f"Failed to load website. Error: {e}")


    def close_popup(self):
        """Closes the pop-up window if present."""
        try:
            self.driver.wait_until_element_is_visible("class:fancybox-close", timeout=1)
            self.driver.click_element("class:fancybox-close")
            self.driver.wait_until_element_is_not_visible("class:fancybox-close", timeout=1)
            logging.info("ADD Pop-up closed.")
        except:
            pass

        try:
            self.driver.wait_until_element_is_visible("xpath://*[@id='onetrust-accept-btn-handler']", timeout=1)
            self.driver.click_element("xpath://*[@id='onetrust-accept-btn-handler']")
            self.driver.wait_until_element_is_not_visible("xpath://*[@id='onetrust-accept-btn-handler']", timeout=1)
            logging.info("GDPR Button closed.")
        except:
            pass

    
    def search_news(self):
        """Enters the search phrase and waits for the results to load."""
        try:
            adjusted_search=str(self.search_phrase)
            new_url=str(self.base_url)+"/search?q="+adjusted_search.replace(" ","+")
            self.driver.go_to(new_url)
            logging.info(f"Search initiated for phrase: {self.search_phrase}")
        except Exception as e:
            logging.error(f"Failed to search for news: {e}")
            raise Exception(f"Failed to search for news. Error : {e}")

    def orderPageFromNewest(self):
        """Orders the search results by newest articles."""
        try:
            currentURL=self.driver.get_location()
            newestURL=currentURL+"&s=3"
            self.driver.go_to(newestURL)
            logging.info("Ordered by Newest Articles.")
        except Exception as e:
            logging.error(f"Failed to load website by Newest Articles: {e}")
            raise Exception(f"Failed to load website by Newest Articles. Error: {e}")


    def scrape_news_articles(self,save_folder_images):
        """Scrapes data from each news article within the specified date range."""
        news_data = []
        while True:
            titles, descriptions, dates, save_paths = [], [], [], []
            timeToSleep=self.sleepTime
            retry=True
            countRetries=1
            while retry==True and countRetries<3:
                try:
                    self.driver.wait_until_element_is_visible("xpath://div[@class='SearchResultsModule-results']", timeout=timeToSleep)
                    retry=False
                except:
                    timeToSleep=self.sleepTime+10
                    countRetries=countRetries+1
            article_elements = self.driver.get_webelements(
            "xpath://div[@class='SearchResultsModule-results']//div[@class='PageList-items-item']/div[@class='PagePromo']")
            for article_element in article_elements:
                self.close_popup()
                try:
                    timestamp_element = self.driver.find_element('tag:bsp-timestamp',article_element)
                    date = Convert_Timestamp_To_Date(self.driver.get_element_attribute(timestamp_element,'data-timestamp'))
                    if self.start_date <= date <= self.end_date:
                        dates.append(date)
                    else:
                        continue
                except Exception as e:
                    continue
                
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
                    imagePath = Download_Image(image_URL,save_folder_images)
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
            'search_phrase_count': Count_Search_Phrase(self.search_phrase,title, description),
            'money_mention': Detect_Money(title, description)} for title, description, date, save_path in zip(titles, descriptions, dates, save_paths)])

            next_page = self.driver.find_element('class:Pagination-nextPage')
            if next_page:
                self.driver.click_element(next_page)
                self.driver.wait_for_condition('return document.readyState == "complete"',self.sleepTime)
            else:
                break

        return news_data


@task
def runBot():

    LinuxChromiumPath=str(os.getenv("LINUX_CHROMIUM_PATH"))
    outputFileName=str(os.getenv("OUTPUT_FILE_NAME"))
    baseUrl = str(os.getenv("BASE_URL"))

    if int(os.getenv("IS_PROD"))!=1:
        delta = 1
        search_phrase = "openai"
        try:
            print(f"Processing Workitem: {delta}, {search_phrase}")
            with APNewsScraper(search_phrase, delta, baseUrl, outputFileName, LinuxChromiumPath) as scraper:
                scraper.run()
        except Exception as err:
            raise Exception(f"Robot Failed - {err}")

    else:
        for item in workitems.inputs:
            try:
                delta = item.payload["DELTA"]
                search_phrase = item.payload["SEARCH_PHRASE"]
                print(f"Processing Workitem: {delta}, {search_phrase}")
                with APNewsScraper(search_phrase, delta, baseUrl, outputFileName, LinuxChromiumPath) as scraper:
                    scraper.run()
                item.done()
            except Exception as err:
                item.fail("BUSINESS", code="INVALID_STEP", message=str(err))
