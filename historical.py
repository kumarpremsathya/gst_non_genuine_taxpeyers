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
import traceback
import mysql.connector
from mysql.connector import Error
import sys
import time
import warnings
from openpyxl.utils import get_column_letter
from config import non_gst_config


"""
This script scrapes data from a specified website, processes the downloaded excel files,
and inserts the data into a MySQL database. The main functions include:

- extract_all_data_in_website: Scrapes data from the website and downloads excel files.
- process_files: Processes the downloaded excel files and combines the data.
- insert_into_database: Inserts the combined data into the MySQL database.
- create_db_connection_with_retry: Establishes a connection to the MySQL database with retries.
- parse_filename_date: Extracts the date from a filename.
- get_month_name: Converts a month number to its name.
- get_relative_path_while_download: Creates a relative path for the downloaded files.
- create_folder_structure: Creates a nested folder structure for the downloads.
- get_relative_path_after_download: Converts a full path to a relative path from the base directory.
"""



def extract_all_data_in_website():
    """
    Scrapes data from a website using Selenium, downloads excel files, and creates a summary Excel sheet.

    This function:
    1. Sets up Chrome browser with custom download settings
    2. Navigates to the configured URL and scrapes table data
    3. Downloads files to organized folder structure by year/month
    4. Creates a summary Excel sheet with metadata of downloaded files

    Returns:
        None

    Raises:
        Exception: For any errors during scraping, downloading or file operations
    """
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
        
        browser = webdriver.Chrome(options=chrome_options)
        
        url = non_gst_config.url
        
        browser.get(url)
        browser.maximize_window()
        
        # Wait for the table to be present
        wait = WebDriverWait(browser, 20)
        table = wait.until(EC.presence_of_element_located((By.TAG_NAME, non_gst_config.table_locator)))
        
        # Find all rows in the table body (excluding the header row)
        rows = table.find_elements(By.TAG_NAME, non_gst_config.table_row)[1:]
        
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
        
        # Convert to DataFrame and save
        if all_data:
            df = pd.DataFrame(all_data)
            current_date = datetime.now().strftime("%Y%m%d")
            first_excel_sheet_name = f"first_excel_sheet_{current_date}.xlsx"
            first_excel_sheet_path = os.path.join(non_gst_config.first_excel_sheet_path, first_excel_sheet_name)
            
            
            # Create directory for the summary Excel file
            os.makedirs(os.path.dirname(first_excel_sheet_path), exist_ok=True)
            
            df.to_excel(first_excel_sheet_path, index=False)
            print(f"Successfully scraped {len(all_data)} records. Data has been saved to {first_excel_sheet_name}")
        
        else:
            print("No data was collected during the scraping process")
            return None
    
    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()
    
    finally:
        # Close the browser
        try:
            browser.quit()
        except:
            pass



def get_relative_path_after_download(base_path, full_path):

    """
    Convert a full file path to a relative path from the base directory.

    Args:
        base_path (str): The base directory path
        full_path (str): The complete file path

    Returns:
        str: The relative path from base_path to full_path
    """
    return os.path.relpath(full_path, base_path)


def process_files():
    """
    Process downloaded Excel files and combine them into a single dataset.

    This function:
    1. Walks through the directory structure to find Excel files
    2. Standardizes column names across different files
    3. Cleans and formats data (e.g., phone numbers)
    4. Adds metadata columns (file year, month, date)
    5. Combines all data into a single DataFrame

    Returns:
        pandas.DataFrame: Combined data from all processed files
                        Returns empty DataFrame if no data or error

    Raises:
        Exception: For errors during file processing or data combination
    """
    # Configuration
    base_path = non_gst_config.base_download_path

    output_file = "combined_output.xlsx"
    date_pattern = r"(\d{2})[-.](\d{2})[-.](\d{4})"

    # Define the standard column order
    standard_columns = [
        'GSTN', 'NGTP_ID', 'TRADE_NAME', 'ASSIGNED_TO', 'RC_DATE', 
        'RCC_DATE', 'KPI_NGTP_STATUS', 'EVIDENCE', 'DIVISION', 'AUTHORITY', 
        'PAN', 'MOBILE_NO', 'EMAIL_ID', 'LETTER_NUMBER', 'LETTER_REC_DATE', 'EVIDENCE_FOLDER_REC_DATE'
    ]

    # Define column name variations mapping
    column_variations = {
        'ASSIGNED_TO': ['ASSIGNED_TO', 'ASSIGNED_T', 'ASSIGNED TO'],
        'MOBILE_NO': ['MOBILE_NO', 'MOBILE', 'CONTACT'],
        'EMAIL_ID': ['EMAIL_ID', 'EMAIL', 'MAIL'], 
        'NGTP_ID': ['NGTP ID', 'UNQ NGTP CODE', 'NGTP_ID', 'NGTPID'],
        'KPI_NGTP_STATUS': ['KPI NGTP STATUS', 'KPI_NGTP_STATUS', 'NGTP_STATUS', 'KPI STATUS'],
        'GSTN': ['GSTN','NEW NGTP LIST'],
        'TRADE_NAME': ['TRADE_NAME','Trade Name', 'TRADE NAME'],
        'RC_DATE': ['RC_DATE','RC_EFFECT_DATE','RC DATE'],
        'RCC_DATE': ['RCC_DATE','RCC_EFFECT_DATE','RCC DATE'],
        'AUTHORITY': ['AUTHORITY','Authority'],
        'PAN': ['PAN','PAN_NO'],
        'LETTER_NUMBER': ['LETTER_NUMBER','Letter number'],
        'LETTER_REC_DATE': ['LETTER_REC_DATE','Letter receive date'],
        'EVIDENCE_FOLDER_REC_DATE': ['EVIDENCE_FOLDER_REC_DATE','Evidence folder receive date']
    }

    # Suppress specific openpyxl warnings
    warnings.filterwarnings('ignore', category=UserWarning, 
                          module='openpyxl.worksheet._reader')
    
    combined_data = []

    def clean_phone_number(phone):
        """Clean phone number strings"""
        if pd.isna(phone) or phone == '':
            return ''
        # Convert to string and remove non-numeric characters
        return str(phone).replace(' ', '').replace('-', '').replace('+', '')

    try:
        # Walk through the directory structure
        for root, dirs, files in os.walk(base_path):
            # Skip the first_excel_sheet directory
            if "first_excel_sheet" in root:
                continue

            for file in files:
                if not file.lower().endswith(('.xlsx', '.xls')):
                    continue

                full_file_path = os.path.join(root, file)

                relative_file_path = get_relative_path_after_download(os.path.dirname(base_path), full_file_path)

                
                print(f"Processing file: {relative_file_path}")
                
                # Extract date from filename
                match = re.search(date_pattern, file)
                if not match:
                    print(f"Skipping {file} (date not found)")
                    continue

                try:
                    # Create a dictionary for column data types
                    dtype_dict = {
                        'MOBILE_NO': str,
                        'MOBILE': str,
                        'PHONE': str,
                        'CONTACT': str,
                        'NGTP_ID': str,
                        'K': str,
                        'LETTER_NUMBER': str,
                        'LETTER_REC_DATE': str
                    }

                    # Read Excel file
                    df = pd.read_excel(
                        full_file_path,
                        dtype=dtype_dict,
                        na_values=['', 'NA', 'N/A'],
                        keep_default_na=False,
                        skiprows=1
                    )
                    
                    df = df.iloc[:2]  # Keep only first two rows for testing

                    # Create column mapping
                    column_mapping = {}
                    for col in df.columns:
                        col_upper = str(col).upper().strip()
                        
                        mapped = False
                        for standard_name, variations in column_variations.items():
                            if any(variation.upper() in col_upper for variation in variations):
                                column_mapping[col] = standard_name
                                mapped = True
                                break
                        
                        if not mapped:
                            standard_name = re.sub(r'[^A-Z0-9_]', '', col_upper)
                            column_mapping[col] = standard_name

                    # Rename columns
                    df = df.rename(columns=column_mapping)
                    
                    # Add missing standard columns
                    for col in standard_columns:
                        if col not in df.columns:
                            df[col] = ''
                    
                    # Clean phone numbers
                    if 'MOBILE_NO' in df.columns:
                        df['MOBILE_NO'] = df['MOBILE_NO'].apply(clean_phone_number)
                    
                    # Add metadata columns
                    day, month, year = match.groups()
                    
                   
                    
                    df["file_year"] = int(year)
                    df["file_month"] = int(month)
                    df["file_date"] = int(day)
                    df["excel_file"] = file
                    df["excel_file_path"] = relative_file_path
            
                    
                    combined_data.append(df)
                    print(f"Successfully processed {relative_file_path}")

                except Exception as e:
                    print(f"Error processing {relative_file_path}: {str(e)}")
                    continue

        if combined_data:
            try:
                # Combine all dataframes
                final_df = pd.concat(combined_data, ignore_index=True)

                # Convert standard columns to lowercase
                standard_columns = [col.lower() for col in standard_columns]

                # Add metadata columns to keep
                metadata_columns = [
                    "file_year", "file_month", "file_date", "excel_file",
                    "excel_file_path"
                ]

                # Convert DataFrame columns to lowercase
                final_df.columns = final_df.columns.str.lower()

                # Select columns
                final_df = final_df[standard_columns + metadata_columns]
                  # Print final_df columns
                print("Columns in final_df:", final_df.columns.tolist())

                
                # Save to Excel
                final_df.to_excel(output_file, index=False)
                print(f"Successfully created {output_file} with {len(final_df)} records")
                return final_df

            except Exception as e:
                print(f"Error saving combined file: {str(e)}")
                return pd.DataFrame()

        else:
            print("No valid data found")
            return pd.DataFrame()

    except Exception as e:
        print(f"General error in process_files: {str(e)}")
        return pd.DataFrame()


def parse_filename_date(filename):
    """
    Extract date components from a filename.

    Args:
        filename (str): Name of the file containing a date pattern

    Returns:
        tuple: (year, month) if date found, (None, None) otherwise
        
    Example:
        >>> parse_filename_date("data-15-06-2023.xlsx")
        ('2023', '06')
    """
    # Look for date pattern like "dd-mm-yyyy" or similar
    date_pattern = r"(\d{2})[-.](\d{2})[-.](\d{4})"
    match = re.search(date_pattern, filename)
    
    if match:
        day, month, year = match.groups()
        return year, month
    return None, None


def get_month_name(month_number):
    """
    Convert a two-digit month number to its three-letter abbreviation.

    Args:
        month_number (str): Two-digit month number (01-12)

    Returns:
        str: Three-letter month abbreviation (e.g., 'jan', 'feb')

    Example:
        >>> get_month_name('01')
        'jan'
    """
    month_names = {
        '01': 'jan', '02': 'feb', '03': 'mar', '04': 'apr',
        '05': 'may', '06': 'jun', '07': 'jul', '08': 'aug',
        '09': 'sep', '10': 'oct', '11': 'nov', '12': 'dec'
    }
    return month_names.get(month_number, month_number)


def get_relative_path_while_download(year, month, filename):
    """
    Create a relative path structure for downloaded files.

    Args:
        year (str): Year component of the date
        month (str): Month component of the date
        filename (str): Name of the file

    Returns:
        str: Relative path in format 'excel_downloads/non_gst/year/month_name/filename'

    Example:
        >>> get_relative_path_while_download('2023', '01', 'data.xlsx')
        'excel_downloads/non_gst/2023/jan/data.xlsx'
    """
    if year and month:
        month_name = get_month_name(month)
        return os.path.join('excel_downloads', 'non_gst', year, month_name, filename)
    return os.path.join('excel_downloads', filename)


def create_folder_structure(base_path, year, month):
    """
    Create nested folder structure for organizing downloads.

    Args:
        base_path (str): Root directory path
        year (str): Year for folder creation
        month (str): Month for folder creation

    Returns:
        str: Complete path to the created folder structure

    Example:
        >>> create_folder_structure('/downloads', '2023', '01')
        '/downloads/non_gst/2023/jan'
    """
    if not year or not month:
        return base_path
        
    # Convert month number to name and create the folder structure
    month_name = get_month_name(month)
    folder_path = os.path.join(base_path, 'non_gst', year, month_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def create_db_connection_with_retry(max_retries=3, retry_delay=5):
    """
    Establish a MySQL database connection with retry mechanism.

    Args:
        max_retries (int, optional): Maximum number of connection attempts. Defaults to 3.
        retry_delay (int, optional): Delay in seconds between retries. Defaults to 5.

    Returns:
        mysql.connector.connection.MySQLConnection: Database connection object if successful
        None: If connection fails after all retries

    Raises:
        mysql.connector.Error: If connection fails
    """
    for attempt in range(max_retries):
        try:
            connection = non_gst_config.db_connection()
            if connection.is_connected():
                print(f"Successfully connected to MySQL on attempt {attempt + 1}")
                return connection
        except Error as e:
            print(f"Error connecting to MySQL (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                print(f"Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
    print("Failed to connect to MySQL after multiple attempts")
    return None


def insert_into_database(connection, data_row):
    """
    Insert processed data into MySQL database.

    Args:
        connection (mysql.connector.connection.MySQLConnection): Active database connection
        data_row (pandas.DataFrame): DataFrame containing rows to insert

    Raises:
        mysql.connector.Error: If database insertion fails

    Note:
        - Handles NULL value conversions from pandas to MySQL
        - Processes timestamp conversions automatically
        - Performs row-by-row insertion with transaction support
    """
    try:

        cursor = connection.cursor()
        
        # SQL Insert Query
        insert_query = """
        INSERT INTO non_gst_config.table_name
        (gstn, ngtp_id, trade_name, assigned_to, rc_date, rcc_date, kpi_ngtp_status, 
        evidence, division, authority, pan, mobile_no, email_id, letter_number, 
        letter_rec_date, evidence_folder_rec_date, file_year, file_month, file_date, 
        excel_file, excel_file_path)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        # Process each row separately
        for index, row in data_row.iterrows():

            # Convert pandas NaN/empty values to None for MySQL
            processed_values = []
            for column in ['gstn', 'ngtp_id', 'trade_name', 'assigned_to', 'rc_date', 
                          'rcc_date', 'kpi_ngtp_status', 'evidence', 'division', 
                          'authority', 'pan', 'mobile_no', 'email_id', 'letter_number',
                          'letter_rec_date', 'evidence_folder_rec_date', 'file_year', 
                          'file_month', 'file_date', 'excel_file', 'excel_file_path']:
                
                value = row.get(column)
                
                # Handle different types of null/empty values
                if pd.isna(value) or value == '' or value == 'nan' or value == 'NaN':
                    processed_values.append(None)
                else:
                    # Convert timestamps to datetime if needed
                    if isinstance(value, pd.Timestamp):
                        value = value.to_pydatetime()
                    processed_values.append(value)

            # Insert into MySQL with processed values
            cursor.execute(insert_query, tuple(processed_values))

        # Commit the transaction
        connection.commit()
        print("Data inserted successfully!")

    except Error as e:
        print(f"Error inserting data: {e}")
        connection.rollback()
    finally:
        if cursor:
            cursor.close()


def main():
    """
    Main execution function that orchestrates the entire process.

    This function:
    1. Establishes database connection
    2. Extracts data from website
    3. Processes downloaded files
    4. Inserts data into database

    Note:
        Exits with status code 1 if database connection fails
        
    """
    try:
        # Connect to MySQL database
        connection = create_db_connection_with_retry()
        if not connection:
            sys.exit(1)

        extract_all_data_in_website()

        final_df_results = process_files()

        # # Insert the combined data into the database
        insert_into_database(connection, final_df_results)
    
    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
    
    print("Done")

