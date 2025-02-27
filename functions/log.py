from config import non_gst_config
from functions import get_data_count_database
import sys
import json
from datetime import datetime


def insert_log_into_table(log_list):
    """
    Inserts logging information into the specified database table.
    Args:
        log_list (list): A list containing logging information with following indices:
            - log_list[1]: Script execution status
            - log_list[2]: Failure reason (if any)
            - log_list[3]: Additional comments
    Dependencies:
        - non_gst_config: Module containing database configuration and global variables
        - get_data_count_database: Module containing function to get total record count
        - datetime: For timestamp generation
        - json: For serializing deleted source information
   
    Raises:
        Exception: If there's any error during database operation, prints detailed error information
                  including line number, exception type and traceback
    Note:
        - Uses parameterized queries to prevent SQL injection
        - Automatically closes database connection after operation
        - Handles NULL values for optional fields
    """
    
    print("insert_log_into_table function is called")
    
    removal_date = datetime.now().strftime('%Y-%m-%d')
    connection = non_gst_config.db_connection()
    cursor = connection.cursor()

    try:
        query = f"""
            INSERT INTO {non_gst_config.log_table_name} (source_name, script_status, data_available, data_scraped, total_record_count, month, year, excel_file_name,failure_reason, comments, source_status, newly_added_count, deleted_source, deleted_source_count, removal_date)
            VALUES (%(source_name)s, %(script_status)s, %(data_available)s, %(data_scraped)s, %(total_record_count)s, %(month)s, %(year)s, %(excel_file_name)s, %(failure_reason)s, %(comments)s, %(source_status)s, %(newly_added_count)s, %(deleted_source)s, %(deleted_source_count)s, %(removal_date)s)
        """
        values = {
            'source_name': non_gst_config.source_name,
            'script_status': log_list[1] if log_list[1] else None,
            'data_available': non_gst_config.no_data_avaliable if non_gst_config.no_data_avaliable else None,
            'data_scraped': non_gst_config.no_data_scraped if non_gst_config.no_data_scraped else None,
            'total_record_count': get_data_count_database.get_data_count_database(),
            'month':log_list[4] if log_list[4] else None,
            'year':log_list[5] if log_list[5] else None,
            'excel_file_name':log_list[6] if log_list[6] else None,
            'failure_reason': log_list[2] if log_list[2] else None,
            'comments': log_list[3] if log_list[3] else None,
            'source_status': non_gst_config.source_status,
            'newly_added_count': non_gst_config.newly_added_count,
            # 'updated_source_count': non_gst_config.updated_count,
            'deleted_source':json.dumps(non_gst_config.deleted_source) if non_gst_config.deleted_source else None,
            'deleted_source_count': non_gst_config.deleted_source_count,
            'removal_date':removal_date if non_gst_config.deleted_source else None,


        }

        cursor.execute(query, values)
        print("log list ====", values)
        connection.commit()
        connection.close()

    except Exception as e:
        print("Error in insert_log_into_table function:", e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")
           