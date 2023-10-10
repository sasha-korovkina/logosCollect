import os
import pyodbc
import pandas as pd
import re
import datetime
import time
import pyautogui
import subprocess

# Replace 'firefox' with the correct path to your Firefox executable if needed
firefox_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'  # Replace with the actual path
subprocess.Popen([firefox_path])

# Wait for Firefox to open (you may need to adjust the delay)
time.sleep(2)

# OPEN FIREFOX #

firefox_windows = pyautogui.getWindowsWithTitle("Mozilla Firefox")
if firefox_windows:
    firefox_window = firefox_windows[0]
    firefox_window.activate()
    pyautogui.sleep(1)
    website_url = "https://www.netzeroassetmanagers.org/"
    pyautogui.write(website_url)
    pyautogui.press('enter')
    pyautogui.sleep(1)
    pyautogui.hotkey("ctrl", "s")
    pyautogui.sleep(2)
    web_page_name = "website"  # Replace with your desired name
    pyautogui.write(web_page_name)
    pyautogui.press("enter")  # Save the web page
else:
    print("Firefox window not found. Please make sure Firefox is open.")

# Navigate to the desired URL
pyautogui.write('https://www.netzeroassetmanagers.org/signatories/')
pyautogui.press('enter')

folder_path = r"C:\Users\sasha\OneDrive - CMi2i\Desktop\Signatories – The Net Zero Asset Managers initiative_files"

file_list = os.listdir(folder_path)
png_files = [file for file in file_list if file.endswith('.png.webp')]

cn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server}; Server=CMDS-SQL02.CMDS.local; Database=CDB; '
                    'Trusted_Connection=yes;')
cr = cn.cursor()
sql_query = """
SELECT [entity_proper_name]
FROM ent.entity
WHERE entity_proper_name like 'A%'
"""
cr.execute(sql_query)
results = cr.fetchall()
df = pd.DataFrame(results, columns = ['Entity'])

filtered_png_files = []

for png_file in png_files:
    file_name = png_file.replace('.png.webp', '')
    if 'logo' not in file_name.lower():
        # Split the file_name by '-' or '_'
        parts = file_name.split('-')  # You can also use '_' instead of '-'
        # Take the first part of the split result
        first_word = parts[0]
        filtered_png_files.append(first_word)

print(filtered_png_files)

# # Check if the table exists in the schema
# cr.execute('''
#     IF NOT EXISTS (
#         SELECT 1
#         FROM INFORMATION_SCHEMA.TABLES
#         WHERE TABLE_SCHEMA = 'etl'
#         AND TABLE_NAME = 'websiteNames'
#     )
#     BEGIN
#         -- Create the table in the 'etl' schema
#         CREATE TABLE [etl].[websiteNames] (
#             Name NVARCHAR(MAX),
#             uploadTime DATETIME
#         )
#     END
# ''')
#
#
# # Insert data into the table
# for name in filtered_png_files:
#     upload_time = datetime.datetime.now()  # Corrected datetime import
#     cr.execute('INSERT INTO cdb.etl.websiteNames (Name, uploadTime) VALUES (?, ?)', (name, upload_time))
#
# # Commit the changes and close the connection
# cr.commit()
# cn.close()

# sql_query_template = "SELECT [entity_proper_name] FROM ent.entity WHERE entity_proper_name LIKE '?'"
#
# results = []
#
# for value in filtered_png_files:
#     sql_query = sql_query_template.replace('?', f'%{value}%')
#     print(sql_query)
#     cr.execute(sql_query)
#     results.extend(cr.fetchall())
#
# print(results)
#
# final_results_df = pd.DataFrame(results, columns=['Entity'])
#
# # Get the current date and time
# upload_time = datetime.datetime.now()
#
# # Insert data into the 'websiteNamesFound' table
# for value in filtered_png_files:
#     sql_query = sql_query_template.replace('?', f'%{value}%')
#     cr.execute(sql_query)
#     search_results = cr.fetchall()
#
#     # Insert the data into the table for each value
#     for result in search_results:
#         original_name = value
#         results_data = ', '.join(result)  # Assuming your result is a list of strings
#
#         # Insert the data into the 'websiteNamesFound' table
#         insert_sql = '''
#         INSERT INTO [etl].[websiteNamesFound] (originalName, Results, uploadTime)
#         VALUES (?, ?, ?)
#         '''
#         cr.execute(insert_sql, (original_name, results_data, upload_time))
#         cn.commit()  # Commit the transaction
#
# # Close the database connection
# cn.close()

# excel_file_path = r"M:\CDB\Analyst\Sasha\Pycharm\Matches Logo Finder.xlsx"
# final_results_df.to_excel(excel_file_path, index=False)
# print(f"Results have been saved to {excel_file_path}")

