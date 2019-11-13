# '''
# This program aggregates data from sheets (of the same name) over multiple excel files.
# Output is a excel for each sheet name with the respective aggregated data.
# '''

import os
import pandas as pd

# look up all files within project scope (file path)
# define file path
path = os.getcwd()
# TODO change file path, must be where excel files are located
files = os.listdir(r'C:\Users\username\filepath')

# get all xlsx files in file path
# TODO change to f[-3] and 'xls' to work with .xls files
files_xlsx = [f for f in files if f[-4:] == 'xlsx']
print(f'Excel documents found ({len(files_xlsx)}): {files_xlsx}')

# name of sheets to be aggregated from each file
# TODO change sheet names that you want to aggregate
sheets_list = ['Sheet 1', 'Sheet 2', 'Sheet 3', 'Sheet 4']

# name of columns that are to be dropped in output, i.e. not needed in aggregated files
# TODO change column names accordingly
drop_columns = ['Comments', 'Remarks', 'Test']

# counter for sheets analyzed and as index for dynamic file names
sheets_counter = 0

# iterate through sheet names from list and aggregate all files into a new excel
for sheet in sheets_list:
    df = pd.DataFrame()
    files_analyzed = 0 # variable to count number of files analyzed for cross-check

    no_data = 0
    for f in files_xlsx:
        try:
            data = pd.read_excel(f, sheet, header=1)  # read data from sheet with column header in second row
            data.dropna(inplace=True, how='all')  # drop empty rows
            data.dropna(inplace=True, how='all', axis=1) # drop empty columns
            data['Source'] = str(f)  # add name of file as 'source' column to data from current file

            # add 'source date, as specified in file name, to data
            data['Source Date'] \
                = f.split(sep='_')[-1]  # first: split name and date, keep date (which includes file type ending)
            temp_date_df = data['Source Date'].str.split('.', expand=True)  # second: split file type from date
            data['Source Date'] = temp_date_df[0]  # rename source date column to only include date

            df = df.append(data)  # append data to aggregated data frame
        except:
            print(f"{f} does not include {sheet}") # print name of excel that doesn't include current sheet / is empty
            no_data += 1
        files_analyzed += 1  # to count files actually analyzed (even if data None)
    print(f'{no_data} files did not include any data for {str(sheets_list[sheets_counter])}')

    # get 'source' and 'source date' to front of data frame
    source_date = df['Source Date']
    source = df['Source']
    df.drop(labels=['Source', 'Source Date'], axis=1, inplace=True)
    df.insert(0, 'Source Date', source_date)
    df.insert(0, 'Source', source)

    df.reset_index(inplace=True, drop=True)

    # drop specified columns not needed
    column_drop_counter = 0
    for column in drop_columns:
        try:
            df.drop(
                inplace=True,
                columns=[column])
            column_drop_counter += 1
        except:
            pass
    print(f'{column_drop_counter} column dropped.')

    unique_sources = df['Source'].nunique()  # check number of aggregated files that have had data in respective sheet
    print(f'\n##### {str(sheets_list[sheets_counter]).upper()} #####\nFiles with data: {unique_sources}')

    # print number of included files of all found files included in aggregation (even if empty)
    # / or print number of missing files
    print(f'Files analyzed: {files_analyzed}')
    if files_analyzed == len(files_xlsx):
        print('All found files included in aggregation.')
    else:
        print(f'{str(len(files_xlsx)- files_analyzed)} FILES MISSING IN AGGREGATION!')
    print(f'Data frame shape (rows, columns): {df.shape}')

    # write aggregated data to new excel file which is named after sheets analyzed
    # TODO change file names of output
    df.to_excel(f'Your project - {str(sheets_list[sheets_counter])} .xlsx')
    print(f'{str(sheets_list[sheets_counter])} Excel printed!\n')

    # increase counter of aggregated sheets and delete temporary data frame
    sheets_counter += 1
    del (df)

print('Aggregation completed.'.upper())
