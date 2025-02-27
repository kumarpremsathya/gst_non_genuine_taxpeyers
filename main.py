
"""
Main script for Non-GST data processing workflow.
This script orchestrates the end-to-end workflow for processing Non-GST data including:

- Web scraping of excel files from specified source website
- Processing and combining downloaded excel files 
- Checking for incremental data changes
- Inserting data into MySQL database
- Logging execution status and handling errors

The workflow is controlled by a source_status configuration that can be:
- Active: Executes full data processing workflow
- Hibernated: Logs non-execution status without processing
- Inactive: Logs non-execution and exits script

Functions:
    main(): Controls workflow execution based on source status configuration

Dependencies:
    - config.non_gst_config: Configuration settings and logging variables
    - functions.extract_all_data_in_website: Web scraping functionality
    - functions.combined_all_excels: Excel processing functionality  
    - functions.log: Logging utilities
    - functions.check_increment_data: Data comparison utilities
"""

import sys
import traceback
from config import non_gst_config
from functions import extract_all_data_in_website, combined_all_excels,log,check_increment_data


def main():

    """Main function to control the workflow based on source status.
    This function manages the execution flow of the application based on the source_status
    configuration setting in non_gst_config. It handles three possible states:
    - Active: Executes the main data processing workflow by combining Excel files
    - Hibernated: Logs the non-execution status and prints error information 
    - Inactive: Logs the non-execution status, prints error details and exits the script
    The function interacts with several modules:
    - combined_all_excels: Handles Excel file combination
    - non_gst_config: Contains configuration and logging variables
    - log: Manages logging functionality
    Returns:
        None
    Raises:
        SystemExit: If source_status is "Inactive", exits with "script error"
    """
    print("main function is called")

    if non_gst_config.source_status == "Active":

        # extract_all_data_in_website.extract_all_data_in_website()
        combined_all_excels.combined_all_excels()  
    
        print("finsihed")

    elif non_gst_config.source_status == "Hibernated":
        non_gst_config.log_list[1] = "not run"
        print(non_gst_config.log_list)
        log.insert_log_into_table(non_gst_config.log_list)

        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")
            
    elif non_gst_config.source_status == "Inactive":
        non_gst_config.log_list[1] = "not run"
       
        traceback.print_exc()
        print(non_gst_config.log_list)
        log.insert_log_into_table(non_gst_config.log_list)
        
        print(non_gst_config.log_list)
        non_gst_config.log_list = [None] * 4
        traceback.print_exc()
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"Error occurred at line {exc_tb.tb_lineno}:")
        print(f"Exception Type: {exc_type}")
        print(f"Exception Object: {exc_obj}")
        print(f"Traceback: {exc_tb}")
        sys.exit("script error")


if __name__ == "__main__":
    main()
