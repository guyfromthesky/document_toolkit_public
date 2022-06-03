# Module for processing dataframe
import pandas as pd
# System module
import os
# Get current timestamp
import datetime

def main():
    while True:
        file_a = input('Input file A path: ')
        # Raise error if file does not exist
        if not os.path.isfile(file_a):
            print('File does not exist')
            continue
        file_b = input('Input file B path: ')
        # Raise error if file does not exist
        if not os.path.isfile(file_b):
            print('File does not exist')
            continue

        key_a = input('Input key column for dataframe A: ')
        if key_a == '':
            print('Key column cannot be empty')
            continue
        key_b = input('Input key column for dataframe B: ')
        if key_b == '':
            print('Key column cannot be empty')
            continue
        #pd.read_excel(file_a, engine='openpyxl', sheet_name= Sheet_Name, skiprows = Key_Row-1, usecols = List_Col)
        # Raise error if cannot read file
        try:
            df_a = pd.read_excel(file_a, engine='openpyxl', sheet_name= 'Data')
        except Exception as e:
            print('Error reading file A: ', e)

        try:
            df_b = pd.read_excel(file_b, engine='openpyxl', sheet_name= 'Data')
        except Exception as e:
            print('Error reading file B: ', e)

        # Raise error if cannot merge
        try:
            df_merged = pd.merge(df_a, df_b, left_on=key_a, right_on=key_b, how="outer")
            df_merged.dropna(how='all', axis=1, inplace=True)
        except Exception as e:
            print('Error merging dataframes: ', e)
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_file_name = 'merged_' + timestamp + '.xlsx'
        df_merged.to_excel(output_file_name, engine='openpyxl', sheet_name='Data')  
        print('Done')

if __name__ == '__main__':
	main()
