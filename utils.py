from datetime import datetime, date
import os
import shutil
import requests
import re
import zipfile
import openpyxl


def Create_Folder_Images():
    """Creates and returns the output path for the downloaded images"""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        save_folder = os.path.join(script_dir, "output")
        zipFilePath=os.path.join(save_folder, "output_images.zip")
        save_folder_images = os.path.join(save_folder, "IMAGES")
        if os.path.exists(save_folder_images):
            shutil.rmtree(save_folder_images)
            os.makedirs(save_folder_images)
        else:
            os.makedirs(save_folder_images)
        return save_folder_images, zipFilePath
    except Exception as e:
        raise Exception(f'Could not create Folder Images - Error: {str(e)}')


def Create_File_Output(fileName):
    """Creates and returns the output path for the Excel file."""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__)) 
        output_folder = os.path.join(script_dir, "output")
        output_path = os.path.join(output_folder,fileName)
        if os.path.exists(output_path):
            os.remove(output_path)
        return output_path
    except Exception as e:
        raise Exception(f'Could not Obtain the Path for the Output file - Error: {str(e)}')


def Calculate_Dates(delta):
    """Calculates the start and end dates based on the given delta."""
    try:
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
    except Exception as e:
        raise Exception(f'Could not Obtain Start and End Dates for the Scrapper - Error: {str(e)}')


def Convert_Timestamp_To_Date(timestamp_ms):
    """Converts a timestamp in milliseconds to a formatted date string."""
    try:
        timestamp_ms=int(timestamp_ms)
        timestamp_s = timestamp_ms / 1000
        dt_object = datetime.fromtimestamp(timestamp_s)
        formatted_date = dt_object.strftime("%m/%d/%Y")
        return formatted_date
    except Exception as e:
        raise Exception(f'Could not Obtain the timestamp of the Article - Error: {str(e)}')


def Download_Image(image_url,save_folder):
    """Downloads the image and saves it to an IMAGES folder in the script's directory."""
    response = requests.get(image_url)
    response.raise_for_status()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"image_{timestamp}.jpg"
    save_path = os.path.join(save_folder, filename)
    with open(save_path, 'wb') as f:
        f.write(response.content)
    return save_path 



def Detect_Money(title, description):
    """Detects the presence of monetary values in the text."""
    text = f"{title} {description}"
    money_pattern = r"\$\d+[\.,]?\d*|\d+[\.,]?\d*\s*(?:dollars|USD)"
    return bool(re.search(money_pattern, text))


def Count_Search_Phrase(search_phrase, title, description):
    """Counts the occurrences of the search phrase in the title and description."""
    return title.lower().count(search_phrase.lower()) + description.lower().count(search_phrase.lower())


def Create_Zip_File_With_Images(save_folder,zipFilePath):
    try:
        zipf = zipfile.ZipFile(zipFilePath, 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk(save_folder):
            for filename in files:
                actual_file_path = os.path.join(root, filename)
                zipped_file_path = os.path.relpath(actual_file_path, save_folder)
                zipf.write(actual_file_path, zipped_file_path)
        zipf.close()
        shutil.rmtree(save_folder)
    except Exception as e:
        raise Exception(f'Could not Create Zip File for the Images - Error: {str(e)}')


def Save_Search_To_Excel(news_data,output_path):
    """Saves the scraped data to an Excel file."""
    try:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(['Title', 'Date', 'Description', 'Picture Filename', 'Search Phrase Count', 'Money Mention'])
        for data in news_data:
            sheet.append([data['title'], data['date'], data['description'], data['picture_filename'],
                            data['search_phrase_count'], data['money_mention']])
        wb.save(output_path)
    except Exception as e:
        raise Exception(f'Could not Save Scrapped data to Excel - Error: {str(e)}')