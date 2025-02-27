import os
import sys
import traceback
import pandas as pd
from datetime import datetime
from config import non_gst_config
from functions import insert_final_data_to_mysql, log, send_mail

def check_increment_data(excel_path):
    """
    Check for incremental data between the database and the provided Excel file,
    and insert the new data into the database.
    """
    print("check_increment_data function is called")

    try:
       
        with non_gst_config.db_connection() as connection:
            query = f"SELECT * FROM  {non_gst_config.table_name}"
            db_df = pd.read_sql(query, con=connection)
 
        excel_df = pd.read_excel(excel_path)
        # excel_df = xl_df.drop(columns=['excel_file_path'])

        # Clean mobile numbers by removing .0
        if 'mobile_no' in excel_df.columns:
            excel_df['mobile_no'] = excel_df['mobile_no'].apply(
            lambda x: str(x).split('.')[0] if pd.notna(x) else None
            
            )

        database_df = db_df.drop(columns=['sr_no', 'source_name', 'updated_date','removal_date', 'date_scraped'])
        print("database columns", db_df.columns)
        print("excel columns", excel_df.columns)


         # Convert evidence_folder_rec_date to string and handle NaN values
        excel_df['evidence_folder_rec_date'] = excel_df['evidence_folder_rec_date'].apply(
            lambda x: str(x) if pd.notna(x) else None
        )
        database_df['evidence_folder_rec_date'] = database_df['evidence_folder_rec_date'].apply(
            lambda x: str(x) if pd.notna(x) else None
        )

        # Ensure 'file_month' is two digits in both DataFrames
        if 'file_month' in excel_df.columns:
            excel_df['file_month'] = excel_df['file_month'].apply(
            lambda x: f"{int(x):02d}" if pd.notna(x) else None
            )
        if 'file_month' in database_df.columns:
            database_df['file_month'] = database_df['file_month'].apply(
            lambda x: f"{int(x):02d}" if pd.notna(x) else None
            )

        # Convert numeric columns to strings in both DataFrames, skipping empty/NULL values
        numeric_columns = ['gstn', 'ngtp_id', 'trade_name', 'assigned_to', 'rc_date', 'rcc_date', 'kpi_ngtp_status', 'evidence', 'division', 'authority', 'pan', 'mobile_no', 'email_id', 'letter_number', 'letter_rec_date', 'evidence_folder_rec_date', 'financial_year', 'file_year', 'file_month', 'file_date', 'excel_file', 'excel_file_path']
        for col in numeric_columns:
            excel_df[col] = excel_df[col].apply(lambda x: str(x) if pd.notna(x) and str(x).strip() != '' else x)
            database_df[col] = database_df[col].apply(lambda x: str(x) if pd.notna(x) and str(x).strip() != '' else x)# Convert numeric columns to strings in both DataFrames
        

        # Compare 'gstn' and 'ngtp_id' columns to find new data
        if not database_df.empty:
            new_data = excel_df.merge(database_df[['gstn', 'file_year','file_date', 'file_month','excel_file']], on=['gstn', 'file_year',  'file_date', 'file_month', 'excel_file'], how='left', indicator=True).query('_merge == "left_only"').drop(columns=['_merge'])
        else:
            new_data = excel_df

        # Add diagnostic prints
        print("\nafter dropping columns:")
        print("Database DataFrame shape:", database_df.shape)
        print("Excel DataFrame shape:", excel_df.shape)
        
        # Safely print samples
        print("\nSample from database:")
        if not database_df.empty:
            print(database_df.head(1).to_string())
        else:
            print("Database is empty")
            
        print("\nSample from Excel:")
        if not excel_df.empty:
            print(excel_df.head(1).to_string())
        else:
            print("Excel file is empty")
        
        # Check data types of columns
        print("\nDatabase column types:")
        print(database_df.dtypes)
        print("\nExcel column types:")
        print(excel_df.dtypes)
        
        # Safely compare values
        print("\nSample values comparison:")
        for col in database_df.columns:
            print(f"\nColumn: {col}")
            db_value = repr(database_df[col].iloc[0]) if not database_df.empty else "No data"
            excel_value = repr(excel_df[col].iloc[0]) if not excel_df.empty else "No data"
            print("Database first value:", db_value)
            print("Excel first value:", excel_value)

        database_df.columns = database_df.columns.str.strip().str.lower()
        excel_df.columns = excel_df.columns.str.strip().str.lower()

    
        new_data.to_excel("non_gst_new_data.xlsx", index=False)

        # Print the missing rows in database and Excel
        print("Rows in Excel but not in database (New Data):")
        print(new_data)

        non_gst_config.no_data_avaliable = len(new_data)
        non_gst_config.no_data_scraped = len(new_data)
    
        print( "missing rows in database", len(new_data))
        
        if len(new_data) == 0:
            non_gst_config.log_list[1] = "Success"
           
            non_gst_config.log_list[3] = "no new data"
            log.insert_log_into_table(non_gst_config.log_list)
            print("log table====", non_gst_config.log_list)
            non_gst_config.log_list = [None] * 4
            sys.exit()

        current_date = datetime.now().strftime("%Y-%m-%d")
        increment_file_name = f"incremental_excel_sheet_{current_date}.xlsx"
        increment_data_excel_path = os.path.join(non_gst_config.increment_data_excel_path, increment_file_name)
        
     
        # Save to Excel file
        pd.DataFrame(new_data).to_excel(increment_data_excel_path, index=False)
       
        # Pass DataFrame directly instead of file path
        insert_final_data_to_mysql.insert_final_data_to_mysql(new_data)
 

    except Exception as e:
            traceback.print_exc()
            non_gst_config.log_list[1] = "Failure"
    
            non_gst_config.log_list[2] = "error in checking in incremental part"
            log.insert_log_into_table(non_gst_config.log_list)
            print("checking incremental part error:", non_gst_config.log_list)
            send_mail.send_email("non_gst taxpayers checking incremental part error", e)
            non_gst_config.log_list = [None] * 4
            exc_type, exc_obj, exc_tb = sys.exc_info()
            print(f"Error occurred at line {exc_tb.tb_lineno}:")
            print(f"Exception Type: {exc_type}")
            print(f"Exception Object: {exc_obj}")
            print(f"Traceback: {exc_tb}")
            sys.exit()
