
from datetime import datetime
import mysql.connector
from selenium import webdriver

source_status = "Active"
source_name = "gst_non_genuine_taxpayers"
table_name = "gst_non_genuine_taxpayers"

current_date = datetime.now().strftime("%Y-%m-%d")

log_table_name = "gst_log"
log_list = [None] * 8
no_data_avaliable = 0
no_data_scraped = 0
newly_added_count = 0
deleted_source = []
deleted_source_count = 0
update_new_deleted_count = 0

url = "https://mahagst.gov.in/en/list/filers/143?field_filers_finacial_year_value%5Bvalue%5D%5Byear%5D=&title_field_value="



host = "localhost"
user = "root"
password = "root"
database = "gst"
auth_plugin = "mysql_native_password"


# host = "4.213.77.165"
# user = "root1"
# password = "Mysql1234$"
# database = "gst"
# auth_plugin = "mysql_native_password"


def db_connection():
    connection = mysql.connector.connect(
        host = host,
        user = user,
        password = password,
        database = database,
        auth_plugin = auth_plugin

    )
    return connection



# file and download folder paths

base_download_path = rf"C:\Users\Premkumar.8265\Desktop\non_gst\excel_downloads"
first_excel_sheet_path = rf"C:\Users\Premkumar.8265\Desktop\non_gst\data\first_excel_sheet"
combined_exceL_sheet_path = rf"C:\Users\Premkumar.8265\Desktop\non_gst\data\combined_all_excels"
increment_data_excel_path = rf"C:\Users\Premkumar.8265\Desktop\non_gst\data\increment_excel_sheet"


# Web element locators

table_locator = "table"
table_row = "tr"
table_column = "td"
link_element = "a"
link_attribute = "href"