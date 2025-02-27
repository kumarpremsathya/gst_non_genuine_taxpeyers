import os
import pandas as pd
import re
import warnings
from openpyxl.utils import get_column_letter
from config import non_gst_config
from functions import check_increment_data,log,send_mail
import traceback
import sys



def get_relative_path_after_download(base_path, full_path):
    """Convert full path to relative path from base directory"""
    return os.path.relpath(full_path, base_path)


def combined_all_excels():
    """
    Process downloaded Excel files and combine them into a single dataset.

    This function:
    1. Walks through the directory structure to find Excel files
    2. Standardizes column names across different files
    3. Cleans and formats data (e.g., phone numbers)
    4. Adds metadata columns (file year, month, date)
    5. Combines all data into a single DataFrame
    6. converts DataFrame in to Single Excel.

    Returns:
        pandas.DataFrame: Combined data from all processed files
                        Returns empty DataFrame if no data or error

    Raises:
        Exception: For errors during file processing or data combination
    """

    print("combined_all_excels function is called")

    # Configuration
    base_path = non_gst_config.base_download_path

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
        'EVIDENCE_FOLDER_REC_DATE': ['EVIDENCE_FOLDER_REC_DATE','Evidence folder receive date'],
    
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
                        'LETTER_REC_DATE': str,
                        'EVIDENCE_FOLDER_REC_DATE' :str
                    }

                    # Read Excel file
                    df = pd.read_excel(
                        full_file_path,
                        dtype=dtype_dict,
                        na_values=['', 'NA', 'N/A'],
                        keep_default_na=False,
                        skiprows=1
                    )
                    
                    # df = df.iloc[:]  # Keep only first two rows for testing

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
                    
                    
                    # Add metadata columns
                    day, month, year = match.groups()
                    
                    
                    df["file_year"] = str(year)
                    # Ensure month is two digits
                    df["file_month"] = f"{int(month):02d}"
                    df["file_date"] = str(day)
                    df["excel_file"] = file
                    df["excel_file_path"] = relative_file_path
                    

                    # Get financial year data from first_excel_sheet
                    try:
                        first_excel_sheet_name = f"first_excel_sheet_{non_gst_config.current_date}.xlsx"
                
                        # Build full path using os.path.join
                        first_exceL_sheet_path = os.path.join(non_gst_config.first_excel_sheet_path, first_excel_sheet_name)

                        first_excel_df = pd.read_excel(first_exceL_sheet_path)
                        
                        
                        # Extract date from filename
                        date_match = re.search(r"(\d{2})[-.](\d{2})[-.](\d{4})", file)
                   
                        if date_match:
                            file_date = date_match.group(0)  # Get the full matched date
                            # Compare with 'filers_subject_excel_links' in first_excel_df
                            matching_date_row = first_excel_df[first_excel_df['filers_subject_excel_links'].str.contains(file_date, na=False)]
                            if not matching_date_row.empty:
                                financial_year = matching_date_row['financial_year'].iloc[0]
                                df['financial_year'] = financial_year
                                combined_data.append(df)
                            else:
                                df['financial_year'] = ''
                                print(f"No matching financial year found for file: {file} and date: {file_date}")
                                combined_data.append(df)
                        else:
                            df['financial_year'] = ''
                            print(f"No matching financial year found for file: {file}")
                            combined_data.append(df)

                    except Exception as e:
                        print(f"Error getting financial year: {str(e)}")
                        df['financial_year'] = ''
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
                    "financial_year","file_year", "file_month", "file_date", "excel_file",
                    "excel_file_path", 
                ]

                # Convert DataFrame columns to lowercase
                final_df.columns = final_df.columns.str.lower()

                # Select columns
                final_df = final_df[standard_columns + metadata_columns]
                
                # Convert date columns to date format without time
                date_columns = ['rc_date', 'rcc_date', 'letter_rec_date','evidence_folder_rec_date']
                for col in date_columns:
                    # Convert to datetime first
                    final_df[col] = pd.to_datetime(final_df[col], errors='coerce')
                    # Extract only the date part
                    final_df[col] = final_df[col].dt.strftime('%Y-%m-%d')
                    # Replace 'NaT' strings with empty string
                    final_df[col] = final_df[col].replace('NaT', '')
                
                # Remove the existing file_month conversion line and add:
                final_df['file_month'] = final_df['file_month'].apply(
                    lambda x: f"{int(str(x).strip()):02d}" if pd.notna(x) and str(x).strip() and str(x).strip().isdigit() else None
                    )
               
                print("final_df\n", final_df)
                # Print final_df columns and their datatypes
                print("Columns in final_df:", final_df.columns.tolist())
                print("\nColumn datatypes:")
                for col, dtype in final_df.dtypes.items():
                    print(f"{col}: {dtype}")
                
                # Print final_df columns
                print("Columns in final_df:", final_df.columns.tolist())


                combined_excel_sheet_name = f"combined_all_excels_{non_gst_config.current_date}.xlsx"
        
                # Build full path using os.path.join
                combined_exceL_sheet_path = os.path.join(non_gst_config.combined_exceL_sheet_path, combined_excel_sheet_name)

                # Save to Excel
                final_df.to_excel(combined_exceL_sheet_path, index=False)
                print(f"Successfully created {combined_exceL_sheet_path} with {len(final_df)} records")
                # return final_df

                check_increment_data.check_increment_data(combined_exceL_sheet_path)

            except Exception as e:
                print(f"Error saving combined file: {str(e)}")
                return pd.DataFrame()

        else:
            print("No valid data found")
            return pd.DataFrame()

    except Exception as e:
        print(f"General error in process_files: {str(e)}")
    
        non_gst_config.log_list[1] = "Failure" 
        non_gst_config.log_list[3] = "Error in combining all excels part" 
        print("error in data extraction part======", non_gst_config.log_list)
        log.insert_log_into_table(non_gst_config.log_list)
        non_gst_config.log_list = [None] * 8

        traceback.print_exc()
        # send_mail.send_email("non gst taxpayers Error in combining all excels part", e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")

        return pd.DataFrame()
        # sys.exit("script error")


