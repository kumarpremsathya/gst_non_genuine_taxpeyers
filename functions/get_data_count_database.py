import sys
from config import non_gst_config


def get_data_count_database():
    """
    Retrieves the total count of records from the specified database table.
    This function establishes a database connection using configuration from non_gst_config
    and executes a COUNT query on the specified table.
    Returns:
        int: The total number of records in the table.
    Raises:
        ValueError: If the query returns no results.
        Exception: If there are any database connection or query execution errors.
            Provides detailed error information including:
            - Exception type
            - Exception message
            - Line number where error occurred
            - Full traceback

    """
    
    print("get_data_count_database function is called")
    connection = non_gst_config.db_connection()
    cursor = connection.cursor()
 
    try:
        print("get_data_count_database function is called")
        cursor.execute(f"SELECT COUNT(*) FROM {non_gst_config.table_name};")
        result = cursor.fetchone()
        print("Result from database query:", result)
        if result:
            return result[0]
        else:
            raise ValueError("Query did not return any results")
    except Exception as e:
        print("Error in get_data_count_database :", e)
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")