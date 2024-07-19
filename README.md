# AP News Scraper 📰 💰

This Python-based web scraper, built using the Robocorp RPA Framework, efficiently extracts news articles from the AP News website (https://apnews.com/) based on your search criteria. 

## Features ✨

* **Targeted Search:**  Find articles containing specific keywords or phrases.
* **Date Range Filtering:** Specify the desired time frame for your news search.
* **Data Extraction:** Retrieves article titles, publication dates, descriptions, and associated images.
* **Search Phrase Analysis:** Counts the occurrences of your search phrase within each article.
* **Financial Insights:**  Detects and flags articles mentioning monetary values.
* **Structured Output:**  Organizes all scraped data into a user-friendly Excel spreadsheet and the images into a zip file.
* **Robocorp Integration:**  Designed for seamless integration with Robocorp's automation platform for scheduled or triggered execution.

## How it Works ⚙️

1. **Initialization:** Provide your search phrase, desired date range (in months prior to the current date), and the output file name.

2. **Website Navigation:** The scraper automatically launches a headless Chrome browser, loads the AP News website, and enters your search query.

3. **Article Filtering:** It filters and selects articles published within your chosen date range.

4. **Data Extraction:**  For each relevant article, the scraper extracts the title, publication date, description, and downloads the associated image.

5. **Data Enrichment:**  The scraper analyzes the title and description of each article, counting occurrences of your search phrase and identifying mentions of monetary values.

6. **Output Generation:** All extracted and enriched data is neatly compiled into an Excel spreadsheet.

## Getting Started 🚀

### Prerequisites

* **Python 3.7 or higher**
* **Robocorp RPA Framework:** Install using `pip install robocorp`
* **Required Python packages:** Install from `requirements.txt` using `pip install -r requirements.txt`
* **Chrome Browser:** Ensure you have Chrome installed on your system.
* **ChromeDriver:** Download the appropriate ChromeDriver for your Chrome version from [https://chromedriver.chromium.org/downloads](https://chromedriver.chromium.org/downloads). Place it in your project directory or in your system's PATH.
* **Environment Variables (.env):**
    * `LINUX_CHROMIUM_PATH`:  (Linux Only) Set this to the path of your Chromium executable if it's not in a standard location. 
    * `OUTPUT_FILE_NAME`: Specify the desired name for your output Excel file (e.g., "ap_news_data.xlsx").
    * `BASE_URL`: The base URL of the AP News website (https://apnews.com/). 
    * `IS_PROD`: Set to `1` for production mode (using Robocorp work items), otherwise, it will run a sample task.

### Running the Scraper

1. **Clone the repository:** `git clone https://github.com/your-username/ap-news-scraper.git`
2. **Navigate to the project directory:**  `cd ap-news-scraper`
3. **Run the Robocorp Task:** `robocorp run`

## Contributing 🤝

Feel free to fork this repository and submit pull requests to enhance the scraper's functionality or address any issues. 

## License 📄

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. 
