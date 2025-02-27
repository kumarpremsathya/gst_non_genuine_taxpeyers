
import os
import time
import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from datetime import datetime
from config import non_gst_config
from functions import combined_all_excels, log, send_mail
import sys
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException


def parse_filename_date(filename):
    
    """Extract date from filename and return year and month."""
    # Look for date pattern like "dd-mm-yyyy" or similar
    date_pattern = r"(\d{2})[-.](\d{2})[-.](\d{4})"
    match = re.search(date_pattern, filename)
    
    if match:
        day, month, year = match.groups()
        return year, month
    return None, None


def get_month_name(month_number):

    """Convert month number to month name."""
    month_names = {
        '01': 'jan', '02': 'feb', '03': 'mar', '04': 'apr',
        '05': 'may', '06': 'jun', '07': 'jul', '08': 'aug',
        '09': 'sep', '10': 'oct', '11': 'nov', '12': 'dec'
    }
    return month_names.get(month_number, month_number)

def get_relative_path_while_download(year, month, filename):

    """Create relative path structure: excel_downloads/non_gst/year/month_name/filename"""
    if year and month:
        month_name = get_month_name(month)
        return os.path.join('excel_downloads', 'non_gst', year, month_name, filename)
    return os.path.join('excel_downloads', filename)


def create_folder_structure(base_path, year, month):

    """Create nested folder structure and return the final path."""
    if not year or not month:
        return base_path
        
    # Convert month number to name and create the folder structure
    month_name = get_month_name(month)
    folder_path = os.path.join(base_path, 'non_gst', year, month_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def extract_all_data_in_website():
    """
    Extracts data from a website containing non-GST taxpayer information and downloads related Excel files.
    This function performs the following operations:
    1. Sets up Chrome WebDriver with custom download options
    2. Navigates to the configured URL and extracts data from an HTML table
    3. For each row in the table:
        - Extracts file information and downloads Excel files
        - Organizes downloads into year/month folder structure
        - Stores metadata about each file
    4. Creates a summary Excel sheet with all extracted metadata
    5. Triggers combination of all downloaded Excel files
    
    Required Configuration:
        - base_download_path: Base directory for downloaded files
        - url: Target website URL
        - table_locator: HTML table tag name
        - table_row: Row element identifier
        - table_column: Column element identifier
        - link_element: Link element identifier
        - link_attribute: Link attribute to extract
        - first_excel_sheet_path: Path for summary Excel file
    Returns:
        None
    Raises:
        Exception: If website access fails or data extraction encounters errors
        TimeoutException: If page elements don't load within timeout
        WebDriverException: If browser automation fails
        NoSuchElementException: If required elements are not found
 
    Example:
        extract_all_data_in_website()
    """
   
      
    print("extract_all_data_in_website function is called")
    try:
        # Set up Chrome options for downloads
        chrome_options = Options()
        
        # Base download path
        base_download_path = non_gst_config.base_download_path
       
        
        # Initially set download directory to base path
        os.makedirs(base_download_path, exist_ok=True)
        
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": base_download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        
        try:
            browser = webdriver.Chrome(options=chrome_options)
            
            url = non_gst_config.url
            
            browser.get(url)
            browser.maximize_window()
            
            # Wait for the table to be present
            wait = WebDriverWait(browser, 20)
            table = wait.until(EC.presence_of_element_located((By.TAG_NAME, non_gst_config.table_locator)))
            
            # Find all rows in the table body (excluding the header row)
            rows = table.find_elements(By.TAG_NAME, non_gst_config.table_row)[1:]

        except (TimeoutException, WebDriverException, NoSuchElementException) as e:
            raise Exception("Website not opened correctly") from e
                
            
        all_data = []
        
        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, non_gst_config.table_column)
                if len(cells) >= 4:
                    # Get the file name from the link text
                    anchor_elements = cells[2].find_elements(By.TAG_NAME, non_gst_config.link_element)
                    if anchor_elements:
                        href = anchor_elements[0].get_attribute(non_gst_config.link_attribute)
                        file_name = anchor_elements[0].text.strip()
                        
                        # Parse year and month from filename
                        year, month = parse_filename_date(file_name)
                        
                        if year and month:
                            # Create appropriate folder structure for actual download
                            download_dir = create_folder_structure(base_download_path, year, month)
                            
                            # Update Chrome preferences for new download directory
                            browser.execute_cdp_cmd('Page.setDownloadBehavior', {
                                'behavior': 'allow',
                                'downloadPath': download_dir
                            })
                        
                        # Click the link to download
                        anchor_elements[0].click()
                        # Wait for download to complete (adjust sleep time if needed)
                        time.sleep(3)
                        
                        # Store relative path instead of full path
                        relative_path = get_relative_path_while_download(year, month, file_name)
                        
                        row_data = {
                            'sr_no': cells[0].text.strip(),
                            'financial_year': cells[1].text.strip(),
                            'filers_subject_excel_links': href,
                            'file_type_size': cells[3].text.strip(),
                            'file_path': relative_path
                        }
                        all_data.append(row_data)
            except Exception as e:
                print(f"Error processing row: {e}")
                continue
        
        if all_data:
            df = pd.DataFrame(all_data)
            
            first_excel_sheet_name = f"first_excel_sheet_{non_gst_config.current_date}.xlsx"
        
            # Build full path using os.path.join
            first_exceL_sheet_path = os.path.join(non_gst_config.first_excel_sheet_path, first_excel_sheet_name)
            
            df.to_excel(first_exceL_sheet_path, index=False)

            print(f"Successfully scraped {len(all_data)} records. Data has been saved to {first_excel_sheet_name}")

            combined_all_excels.combined_all_excels()
        
        else:
            print("No data was collected during the scraping process")
            return None
    
    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()
    
    except Exception as e:
        non_gst_config.log_list[1] = "Failure"
        if str(e) == "Website not opened correctly":
            non_gst_config.log_list[3] = "Website is not opened"
        else:
            non_gst_config.log_list[3] = "Error in data extraction part" 
        print("error in data extraction part======", non_gst_config.log_list)
        log.insert_log_into_table(non_gst_config.log_list)
        non_gst_config.log_list = [None] * 8

        traceback.print_exc()
        send_mail.send_email("non gst taxpayers extract data in website error", e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")
        sys.exit("script error")
    
    finally:
        # Close the browser
        try:
            browser.quit()
        except:
            pass
