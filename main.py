import os
import pyodbc
import pandas as pd
import re
import datetime

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

sql_query_template = "SELECT [entity_proper_name] FROM ent.entity WHERE entity_proper_name LIKE '?'"

results = []

for value in search:
    sql_query = sql_query_template.replace('?', f'%{value}%')
    print(sql_query)
    cr.execute(sql_query)
    results.extend(cr.fetchall())

cn.close()
print(results)

final_results_df = pd.DataFrame(results, columns=['Entity'])
# excel_file_path = r"M:\CDB\Analyst\Sasha\Pycharm\Matches Logo Finder.xlsx"
# final_results_df.to_excel(excel_file_path, index=False)
# print(f"Results have been saved to {excel_file_path}")

