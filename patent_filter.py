import os
import datetime
import pandas as pd
import requests
import re

# Set the log file name and path
log_filename = 'log_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + '.txt'
log_path = ''

# Select the file to be read
while True:
    file_path = input('Please enter the file path to be read：')
    if os.path.isfile(file_path) and file_path.endswith('.xlsx'):
        break
    else:
        print('The file path is wrong or the file format is not .xlsx, please re-enter.')

# Read the Excel file
df = pd.read_excel(file_path)

# Get the keywords to be matched
keywords_input = input('Please enter the keywords that need to be matched, separated by spaces, and the phrases must be enclosed in quotation marks:')
keywords_list = []
temp_keyword = ''
in_quote = False
for char in keywords_input:
    if char == ' ' and not in_quote:
        if temp_keyword != '':
            keywords_list.append(temp_keyword)
            temp_keyword = ''
    elif char == '"' or char == "'":
        in_quote = not in_quote
    else:
        temp_keyword += char
if temp_keyword != '':
    keywords_list.append(temp_keyword)

# Match the keywords
start_time = datetime.datetime.now()
for index, row in df.iterrows():
    result_link = row['result link']
    if pd.isna(result_link):
        continue
    print('Processing the link ', index + 1, '：', result_link)
    try:
        response = requests.get(result_link)
        text = response.text.lower()
        text = re.sub(r'<\/?\w+[^>]*>', '', text)
        matches = []     
        for keyword in keywords_list:
            reg = re.compile(keyword, re.IGNORECASE)
            matched_keyword = reg.findall(text)
            if matched_keyword:
                if 'Match' in df.columns:
                    matches.extend(matched_keyword)
                else:
                    df['Match'] = ''
                    matches.extend(matched_keyword)
        matches = list(set(matches))
        matches.sort()
        df.loc[index, 'Match'] = '|'.join(matches)
    except BaseException as exc:
        print('An error occurred: ', result_link)
        print(exc)
        df.loc[index, 'Match'] = exc
print('All links done!')
end_time = datetime.datetime.now()

# Output an Excel file containing a 'Match' column
output_path = os.path.dirname(file_path)
output_filename = os.path.splitext(os.path.basename(file_path))[0] + '_Match.xlsx'
df.to_excel(os.path.join(output_path, output_filename), index=False)

# Write to a log file
with open(os.path.join(output_path, log_filename), 'w', encoding='utf-8') as f:
    f.write('Execution log:\n')
    f.write('Start time: ' + start_time.strftime('%Y-%m-%d %H:%M:%S') + '\n')
    f.write('End time: ' + end_time.strftime('%Y-%m-%d %H:%M:%S') + '\n')
    f.write('Total time: ' + str((end_time - start_time).seconds) + ' seconds\n')
    f.write('Input file path: ' + file_path + '\n')
    f.write('Output file path: ' + os.path.join(output_path, output_filename) + '\n')
    f.write('Keywords list: ' + ', '.join(keywords_list) + '\n')
