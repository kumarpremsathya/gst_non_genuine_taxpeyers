
import pandas as pd
from mysql.connector import Error
from config import non_gst_config
import sys 
from functions import log, send_mail
import traceback

def insert_final_data_to_mysql(data_row):
    """
    Inserts processed non-GST taxpayer data into MySQL database table.
    This function takes processed taxpayer data and inserts it into the configured MySQL database table. 
    It handles data cleaning, formatting and validation before insertion.
    Args:
        data_row (list/dict): List or dictionary containing taxpayer records to be inserted
                             Must contain columns matching the target table schema
    Returns:
        None
            - Exits with sys.exit() on successful completion
            - Exits with sys.exit("script error") on failure
    Side Effects:
        - Updates non_gst_config.log_list with execution status
        - Updates non_gst_config tracking variables (no_data_scraped, newly_added_count)
        - Inserts logs into logging table
        - Sends email notification on errors
        - Commits database transaction on success
        - Rolls back transaction on failure
    
    Date Handling:
        - Dates are stored in YYYY-MM-DD format
        - Timestamps are converted to date-only format
        - File month is stored as 2-digit string (e.g. "01" for January)
    Null Handling:
        - Empty values, 'nan', 'NaN', pd.NA are converted to NULL in database
        - Proper data type conversion is handled for each field
    """

    print("insert_final_data_to_mysql function is called")

    connection = non_gst_config.db_connection()
    cursor = connection.cursor()

    df = pd.DataFrame(data_row)


    try:
        # Add source_name to each row
        
        # SQL Insert Query
        insert_query = f"""
        INSERT INTO {non_gst_config.table_name}
        (gstn, ngtp_id, trade_name, assigned_to, rc_date, rcc_date, kpi_ngtp_status, 
        evidence, division, authority, pan, mobile_no, email_id, letter_number, 
        letter_rec_date, evidence_folder_rec_date, financial_year, file_year, file_month, file_date, 
        excel_file, excel_file_path, source_name)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        # df = df.iloc[:]
        
        count = 0
        # Process each row separately
        for index, row in df.iterrows():
            # Add source_name to the current row
            # row['source_name'] = non_gst_config.source_name  # Convert pandas NaN/empty values to None for MySQL
            processed_values = []
            for column in ['gstn', 'ngtp_id', 'trade_name', 'assigned_to', 'rc_date', 
                          'rcc_date', 'kpi_ngtp_status', 'evidence', 'division', 
                          'authority', 'pan', 'mobile_no', 'email_id', 'letter_number',
                          'letter_rec_date', 'evidence_folder_rec_date', 'financial_year', 'file_year', 
                          'file_month', 'file_date', 'excel_file', 'excel_file_path', 'source_name']:
                if column == 'source_name':
                    value = non_gst_config.source_name
                else:  
                    value = row.get(column)

                # Handle different types of null/empty values
                if pd.isna(value) or value == '' or value == 'nan' or value == 'NaN':
                    processed_values.append(None)
                else:
                    # Convert timestamps to only date (YYYY-MM-DD) for specific columns
                    if column in ['rc_date', 'rcc_date', 'letter_rec_date', 'evidence_folder_rec_date']:
                        if isinstance(value, pd.Timestamp):
                            value = value.strftime('%Y-%m-%d')  # Convert to string format YYYY-MM-DD
                        elif isinstance(value, str) and " " in value:
                            value = value.split(" ")[0]  # Remove time if present

                    # Ensure file_month has leading zero and is stored as a string
                    elif column == 'file_month':
                        if isinstance(value, (int, float)):
                            value = f"{int(value):02d}"  # Convert number to string with two digits
                        elif isinstance(value, str) and value.isdigit():
                            value = f"{int(value):02d}"  # Convert string number to two-digit format
                        value = str(value)  # Ensure it is stored as a string

                    processed_values.append(value)

            # Execute and commit for each row after all values are processed
            cursor.execute(insert_query, tuple(processed_values))
            connection.commit()
            
            count += 1
            # Insert into MySQL with processed values
            # cursor.execute(insert_query, tuple(processed_values))

        # Commit the transaction
        # connection.commit()
        file_month = df['file_month'].iloc[0] if 'file_month' in df.columns else 'Unknown'
        file_year = df['file_year'].iloc[0] if 'file_year' in df.columns else 'Unknown'
        excel_file_name = df['excel_file'].iloc[0] if 'excel_file' in df.columns else 'Unknown'

        print(f"Total rows inserted: {count}")

        non_gst_config.log_list[1] = "Success"
        non_gst_config.no_data_scraped = count
        non_gst_config.newly_added_count = count
        non_gst_config.log_list[3] = f"{non_gst_config.newly_added_count} new data for {excel_file_name}"
        non_gst_config.log_list[4] = file_month
        non_gst_config.log_list[5] = file_year
        non_gst_config.log_list[6] = excel_file_name
        
        
        print("log table====", non_gst_config.log_list)
        log.insert_log_into_table(non_gst_config.log_list)
        non_gst_config.log_list = [None] * 8
        print("Data has been successfully inserted into the MySQL database.")
        sys.exit()

    except Exception as e:
        connection.rollback()
        non_gst_config.log_list[1] = "Failure"
        non_gst_config.log_list[2] = "error in insert part"
        print("log table====", non_gst_config.log_list)
        log.insert_log_into_table(non_gst_config.log_list)
        
        non_gst_config.log_list = [None] * 8
        traceback.print_exc()
        send_mail.send_email("non gst taxpayers  inserting data into the database error", e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")
         
        sys.exit("script error")

       
    finally:
        if cursor:
            cursor.close()

