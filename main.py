import os
import pyodbc
import pandas as pd
import re

folder_path = r'C:\Users\sasha\Desktop\logos_files'

file_list = os.listdir(folder_path)
png_files = [file for file in file_list if file.endswith('.png.webp')]

for png_file in png_files:
    print(png_file)

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
print(df.head())

filtered_png_files = []

for png_file in png_files:
    file_name = png_file.replace('.png.webp', '')

    if 'logo' not in file_name.lower():
        filtered_png_files.append(file_name)

print(filtered_png_files)

search = []

for file_name in filtered_png_files:
    parts = re.split(r'[-_]', file_name)
    words_before_symbols = [part for part in parts if part]
    search.extend(words_before_symbols)

print(search)

sql_query_template = "SELECT [entity_proper_name] FROM ent.entity WHERE entity_proper_name LIKE '?'"

results = []

for value in search:
    sql_query = sql_query_template.replace('?', f'%{value}%')
    print(sql_query)
    cr.execute(sql_query)
    results.extend(cr.fetchall())

cn.close()
print(results)

# ... your existing code ...

final_results_df = pd.DataFrame(results, columns=['Entity'])
excel_file_path = r"M:\CDB\Analyst\Sasha\Pycharm\Matches Logo Finder.xlsx"
final_results_df.to_excel(excel_file_path, index=False)
print(f"Results have been saved to {excel_file_path}")

############################################################################################
#
# from fuzzywuzzy import fuzz
#
# matching_results = []
#
# threshold_score = 80  # Adjust this value based on your requirements
#
# for png_file in png_files:
#     best_match = None
#     highest_score = 0
#
#     for index, row in df.iterrows():
#         string_to_match = row['Entity']
#         match_score = fuzz.ratio(png_file, string_to_match)
#         if match_score > threshold_score and match_score > highest_score:
#             best_match = (png_file, string_to_match, row['Entity'])
#             highest_score = match_score
#
#     if best_match is not None:
#         matching_results.append(best_match)
#         print(matching_results)
#
# matching_df = pd.DataFrame(matching_results, columns=['PNG File', 'Matched String', 'Matched ID'])
